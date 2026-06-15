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
	/// Descripción breve de PagoLotes.
	/// </summary>
	public class PagoLotes : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCompra;
		private System.ComponentModel.IContainer components;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ImageList imageList;
		private Infragistics.Win.Misc.UltraButton btAgregar;
		private Infragistics.Win.Misc.UltraButton btAgregarTodo;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsPagoLote;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor idPagoLote;
		private C1.Data.C1DataSet cdsFormaPago;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorr;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDesglosa;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAsiento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkIncluyenombre;
		private Infragistics.Win.Misc.UltraButton btAsiento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdAsiento;
		private C1.Data.C1DataSet cdsAsiento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.SimpleButton btRetencion;
		private C1.Data.C1DataSet cdsCompraTabla;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbNumCuenta;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago1;
		private C1.Data.C1DataSet cdsAsignaCuenta;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCheque;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCobrador1;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBanco1;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAsiento;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReferencia;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDeseado;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbVencimiento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPendiente;
		private Infragistics.Win.Misc.UltraButton btAsiRet;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdAsiRetL;
		private Infragistics.Win.Misc.UltraButton btPago;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkReceptor;
		private C1.Data.C1DataSet cdsSeteoC;
		private Infragistics.Win.Misc.UltraButton btFacturaD;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.Misc.UltraButton btDeposito;
		private Infragistics.Win.Misc.UltraButton btCash;

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
			this.components = new System.ComponentModel.Container();
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 38, true, "Compra", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 38, true);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 26, true, "Compra", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 26, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SujetoRetencion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(PagoLotes));
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_Pago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco", -1, "cmbBanco1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobrador", -1, "cmbCobrador1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente", -1, "cmbNumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pago");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pendiente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PendLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PendLote1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AsiEgr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDeposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloquear");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn192 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn193 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InteresMora");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Pago", 10, true, "_Pago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Pago", 10, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraFormaPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn194 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn195 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn196 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn197 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn198 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn199 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn200 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn201 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn202 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn203 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn204 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn205 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn206 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn207 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn208 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn209 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn210 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn211 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn212 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn213 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn214 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn215 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn216 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn217 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn218 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn219 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn220 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn221 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn222 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn223 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn224 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn225 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn226 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn227 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn228 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn229 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn230 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn231 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn232 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn233 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn234 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn235 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn236 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn237 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn238 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn239 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn240 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn241 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn242 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn243 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn244 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn245 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn246 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn247 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn248 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn249 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn250 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn251 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn252 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn253 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn254 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn255 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn256 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn257 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn258 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn259 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn260 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn261 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn262 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn263 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn264 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn265 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn266 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn267 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn268 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn269 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn270 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn271 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn272 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn273 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn274 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn275 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn276 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn277 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn278 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn279 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn280 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn281 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn282 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn283 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn284 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn285 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn286 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn287 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn288 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn289 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn290 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn291 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn292 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn293 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn294 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn295 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn296 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn297 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn298 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn299 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn300 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn301 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn302 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn303 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn304 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn305 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn306 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn307 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn308 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn309 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("E_Profesor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn310 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias360");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn311 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasXIII");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn312 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idJefe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn313 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PoseeDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn314 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn315 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn316 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumConadis");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn317 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn318 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vehiculo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn319 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn320 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn321 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn322 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn323 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn324 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn325 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn326 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn327 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn328 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn329 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn330 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraFormaPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn331 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn332 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn333 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn334 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn335 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn336 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn337 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn338 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn339 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn340 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn341 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn342 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn343 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn344 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn345 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn346 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn347 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn348 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn349 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn350 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn351 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn352 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn353 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn354 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn355 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn356 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn357 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn358 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn359 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn360 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn361 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn362 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn363 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn364 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn365 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn366 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn367 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn368 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn369 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn370 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn371 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn372 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn373 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn374 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn375 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn376 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn377 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn378 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn379 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn380 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn381 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn382 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn383 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn384 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn385 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn386 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn387 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn388 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn389 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn390 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn391 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn392 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn393 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn394 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn395 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn396 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn397 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn398 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn399 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn400 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn401 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn402 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn403 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn404 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn405 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn406 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn407 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn408 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn409 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn410 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn411 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn412 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn413 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn414 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn415 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn416 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn417 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn418 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn419 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn420 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn421 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn422 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn423 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn424 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn425 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn426 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn427 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn428 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn429 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn430 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn431 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn432 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn433 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn434 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("E_Profesor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn435 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias360");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn436 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasXIII");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn437 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idJefe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn438 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PoseeDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn439 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn440 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn441 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumConadis");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn442 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn443 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vehiculo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn444 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn445 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn446 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn447 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn448 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn449 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hoja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn450 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProcesoOk");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn451 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn452 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn453 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn454 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn455 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn456 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hoja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn457 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProcesoOk");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsPagoLote = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.btAgregar = new Infragistics.Win.Misc.UltraButton();
			this.btAgregarTodo = new Infragistics.Win.Misc.UltraButton();
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.idPagoLote = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsFormaPago = new C1.Data.C1DataSet();
			this.label3 = new System.Windows.Forms.Label();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.cmbCuentaCorr = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbCobrador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.label6 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.txtAsiento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkDesglosa = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkIncluyenombre = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btAsiento = new Infragistics.Win.Misc.UltraButton();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btRetencion = new DevExpress.XtraEditors.SimpleButton();
			this.cmbNumCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsCompraTabla = new C1.Data.C1DataSet();
			this.cmbFormaPago1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsAsignaCuenta = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCobrador1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbBanco1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.chkAsiento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDeseado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkPendiente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbVencimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btAsiRet = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btPago = new Infragistics.Win.Misc.UltraButton();
			this.btCash = new Infragistics.Win.Misc.UltraButton();
			this.btFacturaD = new Infragistics.Win.Misc.UltraButton();
			this.btDeposito = new Infragistics.Win.Misc.UltraButton();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.txtIdAsiRetL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkReceptor = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagoLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.idPagoLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNumCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsignaCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeseado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVencimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiRetL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			this.SuspendLayout();
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
			ultraGridColumn1.Header.VisiblePosition = 38;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 96;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 85;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 88;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 37;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 35;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 87;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 43;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 44;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 77;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 42;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 41;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 45;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 36;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 47;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 46;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 39;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 40;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 83;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 79;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Width = 66;
			ultraGridColumn22.Header.VisiblePosition = 29;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 6;
			ultraGridColumn24.Header.VisiblePosition = 70;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 65;
			ultraGridColumn26.Header.VisiblePosition = 55;
			ultraGridColumn26.Hidden = true;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance2;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 3;
			ultraGridColumn27.Width = 61;
			ultraGridColumn28.Header.VisiblePosition = 19;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 20;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 21;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 22;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 25;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 76;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 50;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 49;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 26;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 14;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 16;
			ultraGridColumn38.Hidden = true;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance3;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 2;
			ultraGridColumn39.Width = 65;
			ultraGridColumn40.Header.VisiblePosition = 59;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 57;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 15;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 58;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 71;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 69;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 72;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 62;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 73;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 31;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 75;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 74;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 30;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 60;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 68;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 67;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 65;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 18;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 34;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 64;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 7;
			ultraGridColumn61.Header.VisiblePosition = 12;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 32;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 61;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 27;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 9;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 8;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 66;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 24;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 23;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 17;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 63;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 11;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 10;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 51;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 52;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 53;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 54;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 33;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 48;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 56;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 28;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 13;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 4;
			ultraGridColumn83.Width = 59;
			ultraGridColumn84.Header.VisiblePosition = 78;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 80;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 81;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 82;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 84;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 86;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 89;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 90;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 91;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 95;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 94;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 92;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 93;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 97;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 98;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 99;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 5;
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
																										 ultraGridColumn145});
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance4;
			summarySettings1.DisplayFormat = "{0:n2}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance5;
			summarySettings2.DisplayFormat = "{0:n2}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand1.SummaryFooterCaption = "";
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
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 172);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(352, 169);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
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
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idCliente"));
			this.cmbCliente.DataSource = this.cdvCliente;
			ultraGridColumn146.Header.VisiblePosition = 1;
			ultraGridColumn146.Width = 137;
			ultraGridColumn147.Header.VisiblePosition = 3;
			ultraGridColumn147.Hidden = true;
			ultraGridColumn148.Header.VisiblePosition = 0;
			ultraGridColumn148.Width = 322;
			ultraGridColumn149.Header.VisiblePosition = 4;
			ultraGridColumn149.Hidden = true;
			ultraGridColumn150.Header.VisiblePosition = 5;
			ultraGridColumn150.Hidden = true;
			ultraGridColumn151.Header.VisiblePosition = 2;
			ultraGridColumn151.Hidden = true;
			ultraGridColumn152.Header.VisiblePosition = 6;
			ultraGridColumn152.Hidden = true;
			ultraGridColumn153.Header.VisiblePosition = 7;
			ultraGridColumn153.Hidden = true;
			ultraGridColumn154.Header.VisiblePosition = 8;
			ultraGridColumn154.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn146,
																										 ultraGridColumn147,
																										 ultraGridColumn148,
																										 ultraGridColumn149,
																										 ultraGridColumn150,
																										 ultraGridColumn151,
																										 ultraGridColumn152,
																										 ultraGridColumn153,
																										 ultraGridColumn154});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(72, 9);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(272, 21);
			this.cmbCliente.TabIndex = 1;
			this.cmbCliente.ValueMember = "idCliente";
			this.cmbCliente.ValueChanged += new System.EventHandler(this.cmbCliente_ValueChanged);
			this.cmbCliente.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCliente_MouseUp);
			// 
			// cdsPagoLote
			// 
			this.cdsPagoLote.BindingContextCtrl = this;
			this.cdsPagoLote.DataLibrary = "LibFacturacion";
			this.cdsPagoLote.DataLibraryUrl = "";
			this.cdsPagoLote.DataSetDef = "dsPagoLote";
			this.cdsPagoLote.FillOnRequest = false;
			this.cdsPagoLote.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPagoLote.Name = "cdsPagoLote";
			this.cdsPagoLote.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPagoLote.SchemaDef = null;
			this.cdsPagoLote.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsPagoLote_PositionChanged);
			this.cdsPagoLote.BeforeFill += new C1.Data.FillEventHandler(this.cdsPagoLote_BeforeFill);
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.FillOnRequest = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(38, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Cliente";
			// 
			// imageList
			// 
			this.imageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth16Bit;
			this.imageList.ImageSize = new System.Drawing.Size(18, 18);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btAgregar
			// 
			appearance10.Image = 2;
			this.btAgregar.Appearance = appearance10;
			this.btAgregar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAgregar.Enabled = false;
			this.btAgregar.ImageList = this.imageList;
			this.btAgregar.Location = new System.Drawing.Point(368, 224);
			this.btAgregar.Name = "btAgregar";
			this.btAgregar.Size = new System.Drawing.Size(24, 25);
			this.btAgregar.TabIndex = 5;
			this.btAgregar.Click += new System.EventHandler(this.btAgregar_Click);
			// 
			// btAgregarTodo
			// 
			appearance11.Image = 3;
			this.btAgregarTodo.Appearance = appearance11;
			this.btAgregarTodo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAgregarTodo.Enabled = false;
			this.btAgregarTodo.ImageList = this.imageList;
			this.btAgregarTodo.Location = new System.Drawing.Point(368, 172);
			this.btAgregarTodo.Name = "btAgregarTodo";
			this.btAgregarTodo.Size = new System.Drawing.Size(24, 25);
			this.btAgregarTodo.TabIndex = 6;
			this.btAgregarTodo.Visible = false;
			this.btAgregarTodo.Click += new System.EventHandler(this.btAgregarTodo_Click);
			// 
			// btFiltro
			// 
			appearance12.Image = 5;
			this.btFiltro.Appearance = appearance12;
			this.btFiltro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFiltro.ImageList = this.imageList;
			this.btFiltro.Location = new System.Drawing.Point(352, 9);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(24, 25);
			this.btFiltro.TabIndex = 7;
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "_PagoLote.PagoLote - Pago";
			this.ultraGrid2.DataSource = this.cdsPagoLote;
			appearance13.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance13;
			ultraGridColumn155.Header.VisiblePosition = 19;
			ultraGridColumn155.Hidden = true;
			ultraGridColumn156.Header.VisiblePosition = 14;
			ultraGridColumn156.Hidden = true;
			ultraGridColumn157.Header.VisiblePosition = 18;
			ultraGridColumn157.Hidden = true;
			ultraGridColumn158.Header.VisiblePosition = 16;
			ultraGridColumn158.Hidden = true;
			ultraGridColumn159.Header.Caption = "Forma Pago";
			ultraGridColumn159.Header.VisiblePosition = 5;
			ultraGridColumn159.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn159.Width = 107;
			ultraGridColumn160.Header.Caption = "Banco";
			ultraGridColumn160.Header.VisiblePosition = 9;
			ultraGridColumn160.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn160.Width = 106;
			ultraGridColumn161.Header.Caption = "Cobrador";
			ultraGridColumn161.Header.VisiblePosition = 8;
			ultraGridColumn161.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn161.Width = 87;
			ultraGridColumn162.Header.Caption = "Cuenta";
			ultraGridColumn162.Header.VisiblePosition = 4;
			ultraGridColumn162.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn162.Width = 83;
			ultraGridColumn163.Header.VisiblePosition = 26;
			ultraGridColumn163.Hidden = true;
			ultraGridColumn164.Header.VisiblePosition = 6;
			ultraGridColumn164.Hidden = true;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn165.CellAppearance = appearance14;
			ultraGridColumn165.Format = "#,##0.00";
			ultraGridColumn165.Header.VisiblePosition = 0;
			ultraGridColumn165.Width = 70;
			ultraGridColumn166.Header.VisiblePosition = 20;
			ultraGridColumn167.Header.VisiblePosition = 7;
			ultraGridColumn168.Header.VisiblePosition = 23;
			ultraGridColumn168.Hidden = true;
			ultraGridColumn169.Header.VisiblePosition = 1;
			ultraGridColumn169.Width = 87;
			ultraGridColumn170.Header.VisiblePosition = 12;
			ultraGridColumn170.Hidden = true;
			ultraGridColumn171.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn171.Header.VisiblePosition = 3;
			ultraGridColumn172.Header.VisiblePosition = 13;
			ultraGridColumn172.Hidden = true;
			ultraGridColumn173.Header.VisiblePosition = 22;
			ultraGridColumn173.Hidden = true;
			ultraGridColumn174.Header.VisiblePosition = 24;
			ultraGridColumn174.Hidden = true;
			ultraGridColumn175.Header.VisiblePosition = 10;
			ultraGridColumn176.Header.VisiblePosition = 25;
			ultraGridColumn176.Hidden = true;
			ultraGridColumn177.Header.VisiblePosition = 15;
			ultraGridColumn177.Hidden = true;
			ultraGridColumn178.Header.VisiblePosition = 21;
			ultraGridColumn178.Hidden = true;
			ultraGridColumn179.Header.VisiblePosition = 17;
			ultraGridColumn179.Hidden = true;
			ultraGridColumn180.Header.VisiblePosition = 27;
			ultraGridColumn180.Hidden = true;
			ultraGridColumn180.Width = 56;
			ultraGridColumn181.Header.VisiblePosition = 28;
			ultraGridColumn181.Hidden = true;
			ultraGridColumn182.Header.VisiblePosition = 29;
			ultraGridColumn182.Hidden = true;
			ultraGridColumn183.Header.VisiblePosition = 30;
			ultraGridColumn183.Hidden = true;
			ultraGridColumn184.Header.VisiblePosition = 31;
			ultraGridColumn184.Hidden = true;
			ultraGridColumn185.Header.Caption = "N° Cuenta";
			ultraGridColumn185.Header.VisiblePosition = 2;
			ultraGridColumn186.Header.VisiblePosition = 11;
			ultraGridColumn187.Header.VisiblePosition = 32;
			ultraGridColumn187.Hidden = true;
			ultraGridColumn188.Header.VisiblePosition = 33;
			ultraGridColumn188.Hidden = true;
			ultraGridColumn189.Header.VisiblePosition = 34;
			ultraGridColumn189.Hidden = true;
			ultraGridColumn190.Header.VisiblePosition = 35;
			ultraGridColumn190.Hidden = true;
			ultraGridColumn191.Header.VisiblePosition = 36;
			ultraGridColumn192.Header.VisiblePosition = 37;
			ultraGridColumn193.Header.VisiblePosition = 38;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn155,
																										 ultraGridColumn156,
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
																										 ultraGridColumn171,
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
																										 ultraGridColumn186,
																										 ultraGridColumn187,
																										 ultraGridColumn188,
																										 ultraGridColumn189,
																										 ultraGridColumn190,
																										 ultraGridColumn191,
																										 ultraGridColumn192,
																										 ultraGridColumn193});
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance15;
			summarySettings3.DisplayFormat = "{0:n2}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings3});
			ultraGridBand3.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 10F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance17;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(400, 164);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(520, 177);
			this.ultraGrid2.TabIndex = 8;
			this.ultraGrid2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid2_KeyDown);
			this.ultraGrid2.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_AfterCellUpdate);
			// 
			// txtFecha
			// 
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.Fecha"));
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Format = "dd/MMM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(448, 9);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(128, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 9;
			this.txtFecha.Value = ((object)(resources.GetObject("txtFecha.Value")));
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(400, 9);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(38, 16);
			this.label2.TabIndex = 10;
			this.label2.Text = "Fecha:";
			// 
			// idPagoLote
			// 
			this.idPagoLote.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idPagoLote"));
			this.idPagoLote.Location = new System.Drawing.Point(552, 302);
			this.idPagoLote.Name = "idPagoLote";
			this.idPagoLote.Size = new System.Drawing.Size(40, 22);
			this.idPagoLote.TabIndex = 12;
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idFormaPago"));
			this.cmbFormaPago.DataMember = "CompraFormaPago";
			this.cmbFormaPago.DataSource = this.cdsFormaPago;
			this.cmbFormaPago.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "CompraComprobante";
			ultraGridColumn194.Header.VisiblePosition = 3;
			ultraGridColumn194.Hidden = true;
			ultraGridColumn195.Header.VisiblePosition = 4;
			ultraGridColumn195.Hidden = true;
			ultraGridColumn196.Header.VisiblePosition = 0;
			ultraGridColumn196.Width = 157;
			ultraGridColumn197.Header.VisiblePosition = 2;
			ultraGridColumn197.Hidden = true;
			ultraGridColumn198.Header.VisiblePosition = 1;
			ultraGridColumn198.Hidden = true;
			ultraGridColumn199.Header.VisiblePosition = 5;
			ultraGridColumn199.Hidden = true;
			ultraGridColumn199.Width = 54;
			ultraGridColumn200.Header.VisiblePosition = 6;
			ultraGridColumn200.Hidden = true;
			ultraGridColumn200.Width = 62;
			ultraGridColumn201.Header.VisiblePosition = 7;
			ultraGridColumn201.Hidden = true;
			ultraGridColumn201.Width = 43;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn194,
																										 ultraGridColumn195,
																										 ultraGridColumn196,
																										 ultraGridColumn197,
																										 ultraGridColumn198,
																										 ultraGridColumn199,
																										 ultraGridColumn200,
																										 ultraGridColumn201});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(112, 86);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(176, 21);
			this.cmbFormaPago.TabIndex = 13;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			// 
			// cdsFormaPago
			// 
			this.cdsFormaPago.BindingContextCtrl = this;
			this.cdsFormaPago.DataLibrary = "LibFacturacion";
			this.cdsFormaPago.DataLibraryUrl = "";
			this.cdsFormaPago.DataSetDef = "dsCompraTabla";
			this.cdsFormaPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsFormaPago.Name = "cdsFormaPago";
			this.cdsFormaPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFormaPago.SchemaDef = null;
			this.cdsFormaPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsFormaPago_BeforeFill);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 86);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(80, 16);
			this.label3.TabIndex = 14;
			this.label3.Text = "Forma de Pago";
			// 
			// cdsCuentaCorr
			// 
			this.cdsCuentaCorr.BindingContextCtrl = this;
			this.cdsCuentaCorr.DataLibrary = "LibContabilidad";
			this.cdsCuentaCorr.DataLibraryUrl = "";
			this.cdsCuentaCorr.DataSetDef = "dsCuentaCorriente";
			this.cdsCuentaCorr.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentaCorr.Name = "cdsCuentaCorr";
			this.cdsCuentaCorr.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaCorr.SchemaDef = null;
			this.cdsCuentaCorr.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuentaCorr_BeforeFill);
			// 
			// cmbCuentaCorr
			// 
			this.cmbCuentaCorr.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorr.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorr.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idCuentaCorriente"));
			this.cmbCuentaCorr.DataMember = "CuentaCorriente";
			this.cmbCuentaCorr.DataSource = this.cdsCuentaCorr;
			this.cmbCuentaCorr.DisplayLayout.AutoFitColumns = true;
			ultraGridBand5.AddButtonCaption = "CompraComprobante";
			ultraGridColumn202.Header.VisiblePosition = 7;
			ultraGridColumn202.Hidden = true;
			ultraGridColumn203.Header.VisiblePosition = 9;
			ultraGridColumn203.Hidden = true;
			ultraGridColumn203.Width = 155;
			ultraGridColumn204.Header.VisiblePosition = 6;
			ultraGridColumn204.Hidden = true;
			ultraGridColumn205.Header.VisiblePosition = 5;
			ultraGridColumn205.Hidden = true;
			ultraGridColumn206.Header.VisiblePosition = 0;
			ultraGridColumn206.Width = 141;
			ultraGridColumn207.Header.VisiblePosition = 1;
			ultraGridColumn207.Hidden = true;
			ultraGridColumn207.Width = 61;
			ultraGridColumn208.Header.VisiblePosition = 2;
			ultraGridColumn208.Hidden = true;
			ultraGridColumn209.Header.VisiblePosition = 8;
			ultraGridColumn209.Hidden = true;
			ultraGridColumn210.Header.VisiblePosition = 3;
			ultraGridColumn210.Hidden = true;
			ultraGridColumn211.Header.VisiblePosition = 4;
			ultraGridColumn211.Hidden = true;
			ultraGridColumn212.Header.VisiblePosition = 10;
			ultraGridColumn212.Hidden = true;
			ultraGridColumn212.Width = 40;
			ultraGridColumn213.Header.VisiblePosition = 11;
			ultraGridColumn213.Hidden = true;
			ultraGridColumn213.Width = 54;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn213});
			this.cmbCuentaCorr.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCuentaCorr.DisplayMember = "NumCuenta";
			this.cmbCuentaCorr.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCorr.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCorr.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorr.Location = new System.Drawing.Point(384, 60);
			this.cmbCuentaCorr.Name = "cmbCuentaCorr";
			this.cmbCuentaCorr.Size = new System.Drawing.Size(160, 21);
			this.cmbCuentaCorr.TabIndex = 15;
			this.cmbCuentaCorr.ValueMember = "idCuentaCorriente";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(312, 60);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(67, 16);
			this.label4.TabIndex = 16;
			this.label4.Text = "Cuenta Corr.";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(16, 60);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(50, 16);
			this.label5.TabIndex = 18;
			this.label5.Text = "Cobrador";
			// 
			// cmbCobrador
			// 
			this.cmbCobrador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobrador.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idVendedor"));
			this.cmbCobrador.DataSource = this.cdvPersonal;
			this.cmbCobrador.DisplayLayout.AutoFitColumns = true;
			ultraGridBand6.AddButtonCaption = "CompraComprobante";
			ultraGridColumn214.Header.VisiblePosition = 21;
			ultraGridColumn214.Hidden = true;
			ultraGridColumn214.Width = 8;
			ultraGridColumn215.Header.VisiblePosition = 59;
			ultraGridColumn215.Hidden = true;
			ultraGridColumn215.Width = 48;
			ultraGridColumn216.Header.VisiblePosition = 19;
			ultraGridColumn216.Hidden = true;
			ultraGridColumn216.Width = 8;
			ultraGridColumn217.Header.VisiblePosition = 20;
			ultraGridColumn217.Hidden = true;
			ultraGridColumn217.Width = 8;
			ultraGridColumn218.Header.VisiblePosition = 18;
			ultraGridColumn218.Hidden = true;
			ultraGridColumn218.Width = 8;
			ultraGridColumn219.Header.VisiblePosition = 22;
			ultraGridColumn219.Hidden = true;
			ultraGridColumn219.Width = 8;
			ultraGridColumn220.Header.VisiblePosition = 36;
			ultraGridColumn220.Hidden = true;
			ultraGridColumn220.Width = 8;
			ultraGridColumn221.Header.VisiblePosition = 25;
			ultraGridColumn221.Hidden = true;
			ultraGridColumn221.Width = 8;
			ultraGridColumn222.Header.VisiblePosition = 0;
			ultraGridColumn222.Width = 16;
			ultraGridColumn223.Header.VisiblePosition = 4;
			ultraGridColumn223.Hidden = true;
			ultraGridColumn223.Width = 8;
			ultraGridColumn224.Header.VisiblePosition = 35;
			ultraGridColumn224.Hidden = true;
			ultraGridColumn224.Width = 8;
			ultraGridColumn225.Header.VisiblePosition = 24;
			ultraGridColumn225.Hidden = true;
			ultraGridColumn225.Width = 8;
			ultraGridColumn226.Header.VisiblePosition = 13;
			ultraGridColumn226.Hidden = true;
			ultraGridColumn226.Width = 8;
			ultraGridColumn227.Header.VisiblePosition = 32;
			ultraGridColumn227.Hidden = true;
			ultraGridColumn227.Width = 8;
			ultraGridColumn228.Header.VisiblePosition = 7;
			ultraGridColumn228.Hidden = true;
			ultraGridColumn228.Width = 8;
			ultraGridColumn229.Header.VisiblePosition = 6;
			ultraGridColumn229.Hidden = true;
			ultraGridColumn229.Width = 8;
			ultraGridColumn230.Header.VisiblePosition = 2;
			ultraGridColumn230.Hidden = true;
			ultraGridColumn230.Width = 8;
			ultraGridColumn231.Header.VisiblePosition = 29;
			ultraGridColumn231.Hidden = true;
			ultraGridColumn231.Width = 17;
			ultraGridColumn232.Header.VisiblePosition = 28;
			ultraGridColumn232.Hidden = true;
			ultraGridColumn232.Width = 17;
			ultraGridColumn233.Header.VisiblePosition = 23;
			ultraGridColumn233.Hidden = true;
			ultraGridColumn233.Width = 17;
			ultraGridColumn234.Header.VisiblePosition = 31;
			ultraGridColumn234.Hidden = true;
			ultraGridColumn234.Width = 8;
			ultraGridColumn235.Header.VisiblePosition = 10;
			ultraGridColumn235.Hidden = true;
			ultraGridColumn235.Width = 8;
			ultraGridColumn236.Header.VisiblePosition = 33;
			ultraGridColumn236.Hidden = true;
			ultraGridColumn236.Width = 8;
			ultraGridColumn237.Header.VisiblePosition = 14;
			ultraGridColumn237.Hidden = true;
			ultraGridColumn237.Width = 17;
			ultraGridColumn238.Header.VisiblePosition = 15;
			ultraGridColumn238.Hidden = true;
			ultraGridColumn238.Width = 17;
			ultraGridColumn239.Header.VisiblePosition = 17;
			ultraGridColumn239.Hidden = true;
			ultraGridColumn239.Width = 17;
			ultraGridColumn240.Header.VisiblePosition = 26;
			ultraGridColumn240.Hidden = true;
			ultraGridColumn240.Width = 8;
			ultraGridColumn241.Header.VisiblePosition = 5;
			ultraGridColumn241.Hidden = true;
			ultraGridColumn241.Width = 8;
			ultraGridColumn242.Header.VisiblePosition = 37;
			ultraGridColumn242.Hidden = true;
			ultraGridColumn242.Width = 8;
			ultraGridColumn243.Header.VisiblePosition = 38;
			ultraGridColumn243.Hidden = true;
			ultraGridColumn243.Width = 8;
			ultraGridColumn244.Header.VisiblePosition = 27;
			ultraGridColumn244.Hidden = true;
			ultraGridColumn244.Width = 8;
			ultraGridColumn245.Header.VisiblePosition = 30;
			ultraGridColumn245.Hidden = true;
			ultraGridColumn245.Width = 8;
			ultraGridColumn246.Header.VisiblePosition = 12;
			ultraGridColumn246.Hidden = true;
			ultraGridColumn246.Width = 8;
			ultraGridColumn247.Header.VisiblePosition = 9;
			ultraGridColumn247.Hidden = true;
			ultraGridColumn247.Width = 8;
			ultraGridColumn248.Header.VisiblePosition = 11;
			ultraGridColumn248.Hidden = true;
			ultraGridColumn248.Width = 8;
			ultraGridColumn249.Header.VisiblePosition = 16;
			ultraGridColumn249.Hidden = true;
			ultraGridColumn249.Width = 17;
			ultraGridColumn250.Header.VisiblePosition = 1;
			ultraGridColumn250.Hidden = true;
			ultraGridColumn250.Width = 17;
			ultraGridColumn251.Header.VisiblePosition = 34;
			ultraGridColumn251.Hidden = true;
			ultraGridColumn251.Width = 17;
			ultraGridColumn252.Header.VisiblePosition = 3;
			ultraGridColumn252.Hidden = true;
			ultraGridColumn253.Header.VisiblePosition = 8;
			ultraGridColumn253.Hidden = true;
			ultraGridColumn253.Width = 17;
			ultraGridColumn254.Header.VisiblePosition = 39;
			ultraGridColumn254.Hidden = true;
			ultraGridColumn254.Width = 8;
			ultraGridColumn255.Header.VisiblePosition = 40;
			ultraGridColumn255.Hidden = true;
			ultraGridColumn255.Width = 8;
			ultraGridColumn256.Header.VisiblePosition = 41;
			ultraGridColumn256.Hidden = true;
			ultraGridColumn256.Width = 8;
			ultraGridColumn257.Header.VisiblePosition = 42;
			ultraGridColumn257.Hidden = true;
			ultraGridColumn257.Width = 8;
			ultraGridColumn258.Header.VisiblePosition = 43;
			ultraGridColumn258.Hidden = true;
			ultraGridColumn258.Width = 8;
			ultraGridColumn259.Header.VisiblePosition = 44;
			ultraGridColumn259.Hidden = true;
			ultraGridColumn259.Width = 8;
			ultraGridColumn260.Header.VisiblePosition = 45;
			ultraGridColumn260.Hidden = true;
			ultraGridColumn260.Width = 8;
			ultraGridColumn261.Header.VisiblePosition = 46;
			ultraGridColumn261.Hidden = true;
			ultraGridColumn261.Width = 8;
			ultraGridColumn262.Header.VisiblePosition = 47;
			ultraGridColumn262.Hidden = true;
			ultraGridColumn262.Width = 8;
			ultraGridColumn263.Header.VisiblePosition = 48;
			ultraGridColumn263.Hidden = true;
			ultraGridColumn263.Width = 11;
			ultraGridColumn264.Header.VisiblePosition = 49;
			ultraGridColumn264.Hidden = true;
			ultraGridColumn264.Width = 22;
			ultraGridColumn265.Header.VisiblePosition = 50;
			ultraGridColumn265.Hidden = true;
			ultraGridColumn265.Width = 22;
			ultraGridColumn266.Header.VisiblePosition = 51;
			ultraGridColumn266.Hidden = true;
			ultraGridColumn266.Width = 22;
			ultraGridColumn267.Header.VisiblePosition = 52;
			ultraGridColumn267.Hidden = true;
			ultraGridColumn267.Width = 38;
			ultraGridColumn268.Header.VisiblePosition = 53;
			ultraGridColumn268.Hidden = true;
			ultraGridColumn268.Width = 34;
			ultraGridColumn269.Header.VisiblePosition = 54;
			ultraGridColumn269.Hidden = true;
			ultraGridColumn269.Width = 10;
			ultraGridColumn270.Header.VisiblePosition = 55;
			ultraGridColumn270.Hidden = true;
			ultraGridColumn270.Width = 10;
			ultraGridColumn271.Header.VisiblePosition = 56;
			ultraGridColumn271.Hidden = true;
			ultraGridColumn271.Width = 10;
			ultraGridColumn272.Header.VisiblePosition = 57;
			ultraGridColumn272.Hidden = true;
			ultraGridColumn272.Width = 10;
			ultraGridColumn273.Header.VisiblePosition = 58;
			ultraGridColumn273.Hidden = true;
			ultraGridColumn273.Width = 16;
			ultraGridColumn274.Header.VisiblePosition = 60;
			ultraGridColumn274.Hidden = true;
			ultraGridColumn274.Width = 36;
			ultraGridColumn275.Header.VisiblePosition = 61;
			ultraGridColumn275.Hidden = true;
			ultraGridColumn275.Width = 129;
			ultraGridColumn276.Header.VisiblePosition = 62;
			ultraGridColumn276.Hidden = true;
			ultraGridColumn276.Width = 37;
			ultraGridColumn277.Header.VisiblePosition = 63;
			ultraGridColumn277.Hidden = true;
			ultraGridColumn277.Width = 26;
			ultraGridColumn278.Header.VisiblePosition = 64;
			ultraGridColumn278.Hidden = true;
			ultraGridColumn278.Width = 30;
			ultraGridColumn279.Header.VisiblePosition = 65;
			ultraGridColumn279.Hidden = true;
			ultraGridColumn279.Width = 37;
			ultraGridColumn280.Header.VisiblePosition = 66;
			ultraGridColumn280.Hidden = true;
			ultraGridColumn280.Width = 45;
			ultraGridColumn281.Header.VisiblePosition = 67;
			ultraGridColumn281.Hidden = true;
			ultraGridColumn281.Width = 28;
			ultraGridColumn282.Header.VisiblePosition = 68;
			ultraGridColumn282.Hidden = true;
			ultraGridColumn282.Width = 32;
			ultraGridColumn283.Header.VisiblePosition = 69;
			ultraGridColumn283.Hidden = true;
			ultraGridColumn283.Width = 33;
			ultraGridColumn284.Header.VisiblePosition = 70;
			ultraGridColumn284.Hidden = true;
			ultraGridColumn284.Width = 9;
			ultraGridColumn285.Header.VisiblePosition = 71;
			ultraGridColumn285.Hidden = true;
			ultraGridColumn285.Width = 9;
			ultraGridColumn286.Header.VisiblePosition = 72;
			ultraGridColumn286.Hidden = true;
			ultraGridColumn286.Width = 9;
			ultraGridColumn287.Header.VisiblePosition = 73;
			ultraGridColumn287.Hidden = true;
			ultraGridColumn287.Width = 9;
			ultraGridColumn288.Header.VisiblePosition = 74;
			ultraGridColumn288.Hidden = true;
			ultraGridColumn288.Width = 10;
			ultraGridColumn289.Header.VisiblePosition = 75;
			ultraGridColumn289.Hidden = true;
			ultraGridColumn289.Width = 11;
			ultraGridColumn290.Header.VisiblePosition = 76;
			ultraGridColumn290.Hidden = true;
			ultraGridColumn290.Width = 38;
			ultraGridColumn291.Header.VisiblePosition = 77;
			ultraGridColumn291.Hidden = true;
			ultraGridColumn291.Width = 54;
			ultraGridColumn292.Header.VisiblePosition = 78;
			ultraGridColumn292.Hidden = true;
			ultraGridColumn292.Width = 13;
			ultraGridColumn293.Header.VisiblePosition = 79;
			ultraGridColumn293.Hidden = true;
			ultraGridColumn293.Width = 14;
			ultraGridColumn294.Header.VisiblePosition = 80;
			ultraGridColumn294.Hidden = true;
			ultraGridColumn294.Width = 13;
			ultraGridColumn295.Header.VisiblePosition = 81;
			ultraGridColumn295.Hidden = true;
			ultraGridColumn295.Width = 14;
			ultraGridColumn296.Header.VisiblePosition = 82;
			ultraGridColumn296.Hidden = true;
			ultraGridColumn296.Width = 13;
			ultraGridColumn297.Header.VisiblePosition = 83;
			ultraGridColumn297.Hidden = true;
			ultraGridColumn297.Width = 17;
			ultraGridColumn298.Header.VisiblePosition = 84;
			ultraGridColumn298.Hidden = true;
			ultraGridColumn298.Width = 15;
			ultraGridColumn299.Header.VisiblePosition = 85;
			ultraGridColumn299.Hidden = true;
			ultraGridColumn299.Width = 17;
			ultraGridColumn300.Header.VisiblePosition = 86;
			ultraGridColumn300.Hidden = true;
			ultraGridColumn300.Width = 17;
			ultraGridColumn301.Header.VisiblePosition = 87;
			ultraGridColumn301.Hidden = true;
			ultraGridColumn301.Width = 8;
			ultraGridColumn302.Header.VisiblePosition = 88;
			ultraGridColumn302.Hidden = true;
			ultraGridColumn302.Width = 8;
			ultraGridColumn303.Header.VisiblePosition = 89;
			ultraGridColumn303.Hidden = true;
			ultraGridColumn303.Width = 8;
			ultraGridColumn304.Header.VisiblePosition = 90;
			ultraGridColumn304.Hidden = true;
			ultraGridColumn304.Width = 9;
			ultraGridColumn305.Header.VisiblePosition = 91;
			ultraGridColumn305.Hidden = true;
			ultraGridColumn305.Width = 27;
			ultraGridColumn306.Header.VisiblePosition = 92;
			ultraGridColumn306.Hidden = true;
			ultraGridColumn306.Width = 31;
			ultraGridColumn307.Header.VisiblePosition = 93;
			ultraGridColumn307.Hidden = true;
			ultraGridColumn307.Width = 24;
			ultraGridColumn308.Header.VisiblePosition = 94;
			ultraGridColumn308.Hidden = true;
			ultraGridColumn308.Width = 28;
			ultraGridColumn309.Header.VisiblePosition = 95;
			ultraGridColumn309.Width = 17;
			ultraGridColumn310.Header.VisiblePosition = 96;
			ultraGridColumn310.Width = 9;
			ultraGridColumn311.Header.VisiblePosition = 97;
			ultraGridColumn311.Width = 9;
			ultraGridColumn312.Header.VisiblePosition = 98;
			ultraGridColumn312.Width = 9;
			ultraGridColumn313.Header.VisiblePosition = 99;
			ultraGridColumn313.Width = 18;
			ultraGridColumn314.Header.VisiblePosition = 100;
			ultraGridColumn314.Width = 17;
			ultraGridColumn315.Header.VisiblePosition = 101;
			ultraGridColumn315.Width = 16;
			ultraGridColumn316.Header.VisiblePosition = 102;
			ultraGridColumn316.Width = 16;
			ultraGridColumn317.Header.VisiblePosition = 103;
			ultraGridColumn317.Width = 13;
			ultraGridColumn318.Header.VisiblePosition = 104;
			ultraGridColumn318.Width = 17;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn291,
																										 ultraGridColumn292,
																										 ultraGridColumn293,
																										 ultraGridColumn294,
																										 ultraGridColumn295,
																										 ultraGridColumn296,
																										 ultraGridColumn297,
																										 ultraGridColumn298,
																										 ultraGridColumn299,
																										 ultraGridColumn300,
																										 ultraGridColumn301,
																										 ultraGridColumn302,
																										 ultraGridColumn303,
																										 ultraGridColumn304,
																										 ultraGridColumn305,
																										 ultraGridColumn306,
																										 ultraGridColumn307,
																										 ultraGridColumn308,
																										 ultraGridColumn309,
																										 ultraGridColumn310,
																										 ultraGridColumn311,
																										 ultraGridColumn312,
																										 ultraGridColumn313,
																										 ultraGridColumn314,
																										 ultraGridColumn315,
																										 ultraGridColumn316,
																										 ultraGridColumn317,
																										 ultraGridColumn318});
			this.cmbCobrador.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCobrador.DisplayMember = "Nombre";
			this.cmbCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCobrador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador.Location = new System.Drawing.Point(112, 60);
			this.cmbCobrador.Name = "cmbCobrador";
			this.cmbCobrador.Size = new System.Drawing.Size(176, 21);
			this.cmbCobrador.TabIndex = 17;
			this.cmbCobrador.ValueMember = "idPersonal";
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.RowFilter = "Cobrador = 1";
			this.cdvPersonal.Sort = "Nombre";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
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
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(312, 86);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(28, 16);
			this.label6.TabIndex = 19;
			this.label6.Text = "Nota";
			// 
			// txtNotas
			// 
			this.txtNotas.AlwaysInEditMode = true;
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.Nota"));
			this.txtNotas.Location = new System.Drawing.Point(368, 86);
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(248, 22);
			this.txtNotas.TabIndex = 20;
			// 
			// txtCodigo
			// 
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.Codigo"));
			this.txtCodigo.Location = new System.Drawing.Point(72, 34);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(80, 22);
			this.txtCodigo.TabIndex = 21;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(16, 34);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(42, 16);
			this.label7.TabIndex = 22;
			this.label7.Text = "Código:";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(160, 34);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(45, 16);
			this.label8.TabIndex = 24;
			this.label8.Text = "Asiento:";
			// 
			// txtAsiento
			// 
			this.txtAsiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.CodAsiento"));
			this.txtAsiento.Location = new System.Drawing.Point(208, 34);
			this.txtAsiento.Name = "txtAsiento";
			this.txtAsiento.ReadOnly = true;
			this.txtAsiento.Size = new System.Drawing.Size(96, 22);
			this.txtAsiento.TabIndex = 23;
			// 
			// chkDesglosa
			// 
			this.chkDesglosa.Checked = true;
			this.chkDesglosa.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			this.chkDesglosa.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPagoLote, "_PagoLote.DesglosarPagos"));
			this.chkDesglosa.Location = new System.Drawing.Point(320, 34);
			this.chkDesglosa.Name = "chkDesglosa";
			this.chkDesglosa.Size = new System.Drawing.Size(104, 22);
			this.chkDesglosa.TabIndex = 25;
			this.chkDesglosa.Text = "Desglosar Pago";
			// 
			// chkIncluyenombre
			// 
			this.chkIncluyenombre.Checked = true;
			this.chkIncluyenombre.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			this.chkIncluyenombre.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPagoLote, "_PagoLote.IncluyeNombre"));
			this.chkIncluyenombre.Location = new System.Drawing.Point(448, 34);
			this.chkIncluyenombre.Name = "chkIncluyenombre";
			this.chkIncluyenombre.Size = new System.Drawing.Size(96, 22);
			this.chkIncluyenombre.TabIndex = 26;
			this.chkIncluyenombre.Text = "Incluir Nombre";
			// 
			// btAsiento
			// 
			this.btAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsiento.Location = new System.Drawing.Point(328, 350);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(24, 25);
			this.btAsiento.TabIndex = 27;
			this.btAsiento.Text = "A";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			this.btAsiento.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAsiento_MouseDown);
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idAsiento"));
			this.txtIdAsiento.Location = new System.Drawing.Point(512, 302);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.Size = new System.Drawing.Size(32, 22);
			this.txtIdAsiento.TabIndex = 28;
			// 
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsAsiento";
			this.cdsAsiento.FillOnRequest = false;
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			this.cdsAsiento.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsiento_BeforeFill);
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
			// btRetencion
			// 
			this.btRetencion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btRetencion.ImageIndex = 6;
			this.btRetencion.ImageList = this.imageList;
			this.btRetencion.Location = new System.Drawing.Point(352, 350);
			this.btRetencion.Name = "btRetencion";
			this.btRetencion.Size = new System.Drawing.Size(24, 25);
			this.btRetencion.TabIndex = 29;
			this.btRetencion.ToolTip = "Retención en la fuente";
			this.btRetencion.Click += new System.EventHandler(this.btRetencion_Click);
			this.btRetencion.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btRetencion_MouseDown);
			// 
			// cmbNumCuenta
			// 
			this.cmbNumCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNumCuenta.DataMember = "CuentaCorriente";
			this.cmbNumCuenta.DataSource = this.cdsCuentaCorr;
			this.cmbNumCuenta.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn319.Header.VisiblePosition = 6;
			ultraGridColumn319.Hidden = true;
			ultraGridColumn320.Header.VisiblePosition = 9;
			ultraGridColumn320.Hidden = true;
			ultraGridColumn320.Width = 47;
			ultraGridColumn321.Header.VisiblePosition = 4;
			ultraGridColumn321.Hidden = true;
			ultraGridColumn322.Header.VisiblePosition = 5;
			ultraGridColumn322.Hidden = true;
			ultraGridColumn323.Header.VisiblePosition = 0;
			ultraGridColumn323.Width = 181;
			ultraGridColumn324.Header.VisiblePosition = 7;
			ultraGridColumn324.Hidden = true;
			ultraGridColumn325.Header.VisiblePosition = 1;
			ultraGridColumn325.Hidden = true;
			ultraGridColumn326.Header.VisiblePosition = 8;
			ultraGridColumn326.Hidden = true;
			ultraGridColumn327.Header.VisiblePosition = 2;
			ultraGridColumn327.Hidden = true;
			ultraGridColumn328.Header.VisiblePosition = 3;
			ultraGridColumn328.Hidden = true;
			ultraGridColumn329.Header.VisiblePosition = 10;
			ultraGridColumn329.Hidden = true;
			ultraGridColumn329.Width = 42;
			ultraGridColumn330.Header.VisiblePosition = 11;
			ultraGridColumn330.Hidden = true;
			ultraGridColumn330.Width = 58;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn319,
																										 ultraGridColumn320,
																										 ultraGridColumn321,
																										 ultraGridColumn322,
																										 ultraGridColumn323,
																										 ultraGridColumn324,
																										 ultraGridColumn325,
																										 ultraGridColumn326,
																										 ultraGridColumn327,
																										 ultraGridColumn328,
																										 ultraGridColumn329,
																										 ultraGridColumn330});
			this.cmbNumCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbNumCuenta.DisplayMember = "NumCuenta";
			this.cmbNumCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNumCuenta.Location = new System.Drawing.Point(24, 184);
			this.cmbNumCuenta.Name = "cmbNumCuenta";
			this.cmbNumCuenta.Size = new System.Drawing.Size(200, 86);
			this.cmbNumCuenta.TabIndex = 30;
			this.cmbNumCuenta.ValueMember = "idCuentaCorriente";
			this.cmbNumCuenta.Visible = false;
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
			// cmbFormaPago1
			// 
			this.cmbFormaPago1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago1.DataMember = "CompraFormaPago";
			this.cmbFormaPago1.DataSource = this.cdsFormaPago;
			ultraGridBand8.AddButtonCaption = "CompraComprobante";
			ultraGridColumn331.Header.VisiblePosition = 3;
			ultraGridColumn331.Hidden = true;
			ultraGridColumn332.Header.VisiblePosition = 4;
			ultraGridColumn332.Hidden = true;
			ultraGridColumn333.Header.VisiblePosition = 0;
			ultraGridColumn333.Width = 118;
			ultraGridColumn334.Header.VisiblePosition = 2;
			ultraGridColumn334.Hidden = true;
			ultraGridColumn335.Header.VisiblePosition = 1;
			ultraGridColumn335.Hidden = true;
			ultraGridColumn336.Header.VisiblePosition = 5;
			ultraGridColumn336.Hidden = true;
			ultraGridColumn337.Header.VisiblePosition = 6;
			ultraGridColumn337.Hidden = true;
			ultraGridColumn338.Header.VisiblePosition = 7;
			ultraGridColumn338.Hidden = true;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn331,
																										 ultraGridColumn332,
																										 ultraGridColumn333,
																										 ultraGridColumn334,
																										 ultraGridColumn335,
																										 ultraGridColumn336,
																										 ultraGridColumn337,
																										 ultraGridColumn338});
			this.cmbFormaPago1.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbFormaPago1.DisplayMember = "FormaPago";
			this.cmbFormaPago1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago1.Location = new System.Drawing.Point(672, 200);
			this.cmbFormaPago1.Name = "cmbFormaPago1";
			this.cmbFormaPago1.Size = new System.Drawing.Size(144, 86);
			this.cmbFormaPago1.TabIndex = 31;
			this.cmbFormaPago1.ValueMember = "idFormaPago";
			this.cmbFormaPago1.Visible = false;
			// 
			// cdsAsignaCuenta
			// 
			this.cdsAsignaCuenta.BindingContextCtrl = this;
			this.cdsAsignaCuenta.DataLibrary = "LibContabilidad";
			this.cdsAsignaCuenta.DataLibraryUrl = "";
			this.cdsAsignaCuenta.DataSetDef = "dsAsignaCuenta";
			this.cdsAsignaCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsignaCuenta.Name = "cdsAsignaCuenta";
			this.cdsAsignaCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsignaCuenta.SchemaDef = null;
			this.cdsAsignaCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsignaCuenta_BeforeFill);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_PagoLote";
			this.barraDato1.DataNombreId = "idPagoLote";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsPagoLote;
			this.barraDato1.DataTabla = "PagoLote";
			this.barraDato1.DataTablaHija = "Pago";
			this.barraDato1.Location = new System.Drawing.Point(16, 349);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 32;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtCheque
			// 
			this.txtCheque.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.Cheque"));
			this.txtCheque.Location = new System.Drawing.Point(544, 60);
			this.txtCheque.Name = "txtCheque";
			this.txtCheque.Size = new System.Drawing.Size(72, 22);
			this.txtCheque.TabIndex = 33;
			this.toolTip1.SetToolTip(this.txtCheque, "Número de Cheque");
			// 
			// cmbCobrador1
			// 
			this.cmbCobrador1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobrador1.DataSource = this.cdvPersonal;
			this.cmbCobrador1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn339.Header.VisiblePosition = 21;
			ultraGridColumn339.Hidden = true;
			ultraGridColumn340.Header.VisiblePosition = 59;
			ultraGridColumn340.Hidden = true;
			ultraGridColumn340.Width = 52;
			ultraGridColumn341.Header.VisiblePosition = 19;
			ultraGridColumn341.Hidden = true;
			ultraGridColumn342.Header.VisiblePosition = 20;
			ultraGridColumn342.Hidden = true;
			ultraGridColumn343.Header.VisiblePosition = 18;
			ultraGridColumn343.Hidden = true;
			ultraGridColumn344.Header.VisiblePosition = 22;
			ultraGridColumn344.Hidden = true;
			ultraGridColumn345.Header.VisiblePosition = 36;
			ultraGridColumn345.Hidden = true;
			ultraGridColumn346.Header.VisiblePosition = 25;
			ultraGridColumn346.Hidden = true;
			ultraGridColumn347.Header.VisiblePosition = 0;
			ultraGridColumn347.Width = 117;
			ultraGridColumn348.Header.VisiblePosition = 3;
			ultraGridColumn348.Hidden = true;
			ultraGridColumn349.Header.VisiblePosition = 35;
			ultraGridColumn349.Hidden = true;
			ultraGridColumn350.Header.VisiblePosition = 24;
			ultraGridColumn350.Hidden = true;
			ultraGridColumn351.Header.VisiblePosition = 13;
			ultraGridColumn351.Hidden = true;
			ultraGridColumn352.Header.VisiblePosition = 32;
			ultraGridColumn352.Hidden = true;
			ultraGridColumn353.Header.VisiblePosition = 7;
			ultraGridColumn353.Hidden = true;
			ultraGridColumn354.Header.VisiblePosition = 6;
			ultraGridColumn354.Hidden = true;
			ultraGridColumn355.Header.VisiblePosition = 2;
			ultraGridColumn355.Hidden = true;
			ultraGridColumn356.Header.VisiblePosition = 29;
			ultraGridColumn356.Hidden = true;
			ultraGridColumn357.Header.VisiblePosition = 28;
			ultraGridColumn357.Hidden = true;
			ultraGridColumn358.Header.VisiblePosition = 23;
			ultraGridColumn358.Hidden = true;
			ultraGridColumn359.Header.VisiblePosition = 31;
			ultraGridColumn359.Hidden = true;
			ultraGridColumn360.Header.VisiblePosition = 10;
			ultraGridColumn360.Hidden = true;
			ultraGridColumn361.Header.VisiblePosition = 33;
			ultraGridColumn361.Hidden = true;
			ultraGridColumn362.Header.VisiblePosition = 14;
			ultraGridColumn362.Hidden = true;
			ultraGridColumn363.Header.VisiblePosition = 15;
			ultraGridColumn363.Hidden = true;
			ultraGridColumn364.Header.VisiblePosition = 17;
			ultraGridColumn364.Hidden = true;
			ultraGridColumn365.Header.VisiblePosition = 26;
			ultraGridColumn365.Hidden = true;
			ultraGridColumn366.Header.VisiblePosition = 5;
			ultraGridColumn366.Hidden = true;
			ultraGridColumn367.Header.VisiblePosition = 37;
			ultraGridColumn367.Hidden = true;
			ultraGridColumn368.Header.VisiblePosition = 38;
			ultraGridColumn368.Hidden = true;
			ultraGridColumn369.Header.VisiblePosition = 27;
			ultraGridColumn369.Hidden = true;
			ultraGridColumn370.Header.VisiblePosition = 30;
			ultraGridColumn370.Hidden = true;
			ultraGridColumn371.Header.VisiblePosition = 12;
			ultraGridColumn371.Hidden = true;
			ultraGridColumn372.Header.VisiblePosition = 9;
			ultraGridColumn372.Hidden = true;
			ultraGridColumn373.Header.VisiblePosition = 11;
			ultraGridColumn373.Hidden = true;
			ultraGridColumn374.Header.VisiblePosition = 16;
			ultraGridColumn374.Hidden = true;
			ultraGridColumn375.Header.VisiblePosition = 1;
			ultraGridColumn375.Hidden = true;
			ultraGridColumn376.Header.VisiblePosition = 34;
			ultraGridColumn376.Hidden = true;
			ultraGridColumn377.Header.VisiblePosition = 4;
			ultraGridColumn377.Hidden = true;
			ultraGridColumn378.Header.VisiblePosition = 8;
			ultraGridColumn378.Hidden = true;
			ultraGridColumn379.Header.VisiblePosition = 39;
			ultraGridColumn379.Hidden = true;
			ultraGridColumn379.Width = 8;
			ultraGridColumn380.Header.VisiblePosition = 40;
			ultraGridColumn380.Hidden = true;
			ultraGridColumn380.Width = 16;
			ultraGridColumn381.Header.VisiblePosition = 41;
			ultraGridColumn381.Hidden = true;
			ultraGridColumn381.Width = 22;
			ultraGridColumn382.Header.VisiblePosition = 42;
			ultraGridColumn382.Hidden = true;
			ultraGridColumn382.Width = 24;
			ultraGridColumn383.Header.VisiblePosition = 43;
			ultraGridColumn383.Hidden = true;
			ultraGridColumn383.Width = 22;
			ultraGridColumn384.Header.VisiblePosition = 44;
			ultraGridColumn384.Hidden = true;
			ultraGridColumn384.Width = 44;
			ultraGridColumn385.Header.VisiblePosition = 45;
			ultraGridColumn385.Hidden = true;
			ultraGridColumn385.Width = 24;
			ultraGridColumn386.Header.VisiblePosition = 46;
			ultraGridColumn386.Hidden = true;
			ultraGridColumn386.Width = 9;
			ultraGridColumn387.Header.VisiblePosition = 47;
			ultraGridColumn387.Hidden = true;
			ultraGridColumn387.Width = 19;
			ultraGridColumn388.Header.VisiblePosition = 48;
			ultraGridColumn388.Hidden = true;
			ultraGridColumn388.Width = 25;
			ultraGridColumn389.Header.VisiblePosition = 49;
			ultraGridColumn389.Hidden = true;
			ultraGridColumn389.Width = 42;
			ultraGridColumn390.Header.VisiblePosition = 50;
			ultraGridColumn390.Hidden = true;
			ultraGridColumn390.Width = 62;
			ultraGridColumn391.Header.VisiblePosition = 51;
			ultraGridColumn391.Hidden = true;
			ultraGridColumn391.Width = 35;
			ultraGridColumn392.Header.VisiblePosition = 52;
			ultraGridColumn392.Hidden = true;
			ultraGridColumn392.Width = 32;
			ultraGridColumn393.Header.VisiblePosition = 53;
			ultraGridColumn393.Hidden = true;
			ultraGridColumn393.Width = 42;
			ultraGridColumn394.Header.VisiblePosition = 54;
			ultraGridColumn394.Hidden = true;
			ultraGridColumn394.Width = 47;
			ultraGridColumn395.Header.VisiblePosition = 55;
			ultraGridColumn395.Hidden = true;
			ultraGridColumn395.Width = 33;
			ultraGridColumn396.Header.VisiblePosition = 56;
			ultraGridColumn396.Hidden = true;
			ultraGridColumn396.Width = 25;
			ultraGridColumn397.Header.VisiblePosition = 57;
			ultraGridColumn397.Hidden = true;
			ultraGridColumn397.Width = 21;
			ultraGridColumn398.Header.VisiblePosition = 58;
			ultraGridColumn398.Hidden = true;
			ultraGridColumn398.Width = 24;
			ultraGridColumn399.Header.VisiblePosition = 60;
			ultraGridColumn399.Hidden = true;
			ultraGridColumn399.Width = 20;
			ultraGridColumn400.Header.VisiblePosition = 61;
			ultraGridColumn400.Hidden = true;
			ultraGridColumn400.Width = 49;
			ultraGridColumn401.Header.VisiblePosition = 62;
			ultraGridColumn401.Hidden = true;
			ultraGridColumn401.Width = 31;
			ultraGridColumn402.Header.VisiblePosition = 63;
			ultraGridColumn402.Hidden = true;
			ultraGridColumn402.Width = 21;
			ultraGridColumn403.Header.VisiblePosition = 64;
			ultraGridColumn403.Hidden = true;
			ultraGridColumn403.Width = 26;
			ultraGridColumn404.Header.VisiblePosition = 65;
			ultraGridColumn404.Hidden = true;
			ultraGridColumn404.Width = 32;
			ultraGridColumn405.Header.VisiblePosition = 66;
			ultraGridColumn405.Hidden = true;
			ultraGridColumn405.Width = 52;
			ultraGridColumn406.Header.VisiblePosition = 67;
			ultraGridColumn406.Hidden = true;
			ultraGridColumn406.Width = 58;
			ultraGridColumn407.Header.VisiblePosition = 68;
			ultraGridColumn407.Hidden = true;
			ultraGridColumn407.Width = 36;
			ultraGridColumn408.Header.VisiblePosition = 69;
			ultraGridColumn408.Hidden = true;
			ultraGridColumn408.Width = 27;
			ultraGridColumn409.Header.VisiblePosition = 70;
			ultraGridColumn409.Hidden = true;
			ultraGridColumn409.Width = 23;
			ultraGridColumn410.Header.VisiblePosition = 71;
			ultraGridColumn410.Hidden = true;
			ultraGridColumn410.Width = 34;
			ultraGridColumn411.Header.VisiblePosition = 72;
			ultraGridColumn411.Hidden = true;
			ultraGridColumn411.Width = 28;
			ultraGridColumn412.Header.VisiblePosition = 73;
			ultraGridColumn412.Hidden = true;
			ultraGridColumn412.Width = 20;
			ultraGridColumn413.Header.VisiblePosition = 74;
			ultraGridColumn413.Hidden = true;
			ultraGridColumn413.Width = 18;
			ultraGridColumn414.Header.VisiblePosition = 75;
			ultraGridColumn414.Hidden = true;
			ultraGridColumn414.Width = 15;
			ultraGridColumn415.Header.VisiblePosition = 76;
			ultraGridColumn415.Hidden = true;
			ultraGridColumn415.Width = 50;
			ultraGridColumn416.Header.VisiblePosition = 77;
			ultraGridColumn416.Hidden = true;
			ultraGridColumn416.Width = 43;
			ultraGridColumn417.Header.VisiblePosition = 78;
			ultraGridColumn417.Hidden = true;
			ultraGridColumn417.Width = 10;
			ultraGridColumn418.Header.VisiblePosition = 79;
			ultraGridColumn418.Hidden = true;
			ultraGridColumn418.Width = 11;
			ultraGridColumn419.Header.VisiblePosition = 80;
			ultraGridColumn419.Hidden = true;
			ultraGridColumn419.Width = 10;
			ultraGridColumn420.Header.VisiblePosition = 81;
			ultraGridColumn420.Hidden = true;
			ultraGridColumn420.Width = 10;
			ultraGridColumn421.Header.VisiblePosition = 82;
			ultraGridColumn421.Hidden = true;
			ultraGridColumn421.Width = 10;
			ultraGridColumn422.Header.VisiblePosition = 83;
			ultraGridColumn422.Hidden = true;
			ultraGridColumn422.Width = 13;
			ultraGridColumn423.Header.VisiblePosition = 84;
			ultraGridColumn423.Hidden = true;
			ultraGridColumn423.Width = 11;
			ultraGridColumn424.Header.VisiblePosition = 85;
			ultraGridColumn424.Hidden = true;
			ultraGridColumn424.Width = 13;
			ultraGridColumn425.Header.VisiblePosition = 86;
			ultraGridColumn425.Hidden = true;
			ultraGridColumn425.Width = 16;
			ultraGridColumn426.Header.VisiblePosition = 87;
			ultraGridColumn426.Hidden = true;
			ultraGridColumn426.Width = 23;
			ultraGridColumn427.Header.VisiblePosition = 88;
			ultraGridColumn427.Hidden = true;
			ultraGridColumn427.Width = 23;
			ultraGridColumn428.Header.VisiblePosition = 89;
			ultraGridColumn428.Hidden = true;
			ultraGridColumn428.Width = 23;
			ultraGridColumn429.Header.VisiblePosition = 90;
			ultraGridColumn429.Hidden = true;
			ultraGridColumn429.Width = 22;
			ultraGridColumn430.Header.VisiblePosition = 91;
			ultraGridColumn430.Hidden = true;
			ultraGridColumn430.Width = 48;
			ultraGridColumn431.Header.VisiblePosition = 92;
			ultraGridColumn431.Hidden = true;
			ultraGridColumn431.Width = 37;
			ultraGridColumn432.Header.VisiblePosition = 93;
			ultraGridColumn432.Hidden = true;
			ultraGridColumn432.Width = 24;
			ultraGridColumn433.Header.VisiblePosition = 94;
			ultraGridColumn433.Hidden = true;
			ultraGridColumn433.Width = 21;
			ultraGridColumn434.Header.VisiblePosition = 95;
			ultraGridColumn434.Hidden = true;
			ultraGridColumn434.Width = 76;
			ultraGridColumn435.Header.VisiblePosition = 96;
			ultraGridColumn435.Hidden = true;
			ultraGridColumn435.Width = 27;
			ultraGridColumn436.Header.VisiblePosition = 97;
			ultraGridColumn436.Hidden = true;
			ultraGridColumn436.Width = 20;
			ultraGridColumn437.Header.VisiblePosition = 98;
			ultraGridColumn437.Hidden = true;
			ultraGridColumn437.Width = 14;
			ultraGridColumn438.Header.VisiblePosition = 99;
			ultraGridColumn438.Hidden = true;
			ultraGridColumn438.Width = 28;
			ultraGridColumn439.Header.VisiblePosition = 100;
			ultraGridColumn439.Hidden = true;
			ultraGridColumn439.Width = 10;
			ultraGridColumn440.Header.VisiblePosition = 101;
			ultraGridColumn440.Hidden = true;
			ultraGridColumn440.Width = 10;
			ultraGridColumn441.Header.VisiblePosition = 102;
			ultraGridColumn441.Hidden = true;
			ultraGridColumn441.Width = 10;
			ultraGridColumn442.Header.VisiblePosition = 103;
			ultraGridColumn442.Hidden = true;
			ultraGridColumn442.Width = 16;
			ultraGridColumn443.Header.VisiblePosition = 104;
			ultraGridColumn443.Hidden = true;
			ultraGridColumn443.Width = 17;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn339,
																										 ultraGridColumn340,
																										 ultraGridColumn341,
																										 ultraGridColumn342,
																										 ultraGridColumn343,
																										 ultraGridColumn344,
																										 ultraGridColumn345,
																										 ultraGridColumn346,
																										 ultraGridColumn347,
																										 ultraGridColumn348,
																										 ultraGridColumn349,
																										 ultraGridColumn350,
																										 ultraGridColumn351,
																										 ultraGridColumn352,
																										 ultraGridColumn353,
																										 ultraGridColumn354,
																										 ultraGridColumn355,
																										 ultraGridColumn356,
																										 ultraGridColumn357,
																										 ultraGridColumn358,
																										 ultraGridColumn359,
																										 ultraGridColumn360,
																										 ultraGridColumn361,
																										 ultraGridColumn362,
																										 ultraGridColumn363,
																										 ultraGridColumn364,
																										 ultraGridColumn365,
																										 ultraGridColumn366,
																										 ultraGridColumn367,
																										 ultraGridColumn368,
																										 ultraGridColumn369,
																										 ultraGridColumn370,
																										 ultraGridColumn371,
																										 ultraGridColumn372,
																										 ultraGridColumn373,
																										 ultraGridColumn374,
																										 ultraGridColumn375,
																										 ultraGridColumn376,
																										 ultraGridColumn377,
																										 ultraGridColumn378,
																										 ultraGridColumn379,
																										 ultraGridColumn380,
																										 ultraGridColumn381,
																										 ultraGridColumn382,
																										 ultraGridColumn383,
																										 ultraGridColumn384,
																										 ultraGridColumn385,
																										 ultraGridColumn386,
																										 ultraGridColumn387,
																										 ultraGridColumn388,
																										 ultraGridColumn389,
																										 ultraGridColumn390,
																										 ultraGridColumn391,
																										 ultraGridColumn392,
																										 ultraGridColumn393,
																										 ultraGridColumn394,
																										 ultraGridColumn395,
																										 ultraGridColumn396,
																										 ultraGridColumn397,
																										 ultraGridColumn398,
																										 ultraGridColumn399,
																										 ultraGridColumn400,
																										 ultraGridColumn401,
																										 ultraGridColumn402,
																										 ultraGridColumn403,
																										 ultraGridColumn404,
																										 ultraGridColumn405,
																										 ultraGridColumn406,
																										 ultraGridColumn407,
																										 ultraGridColumn408,
																										 ultraGridColumn409,
																										 ultraGridColumn410,
																										 ultraGridColumn411,
																										 ultraGridColumn412,
																										 ultraGridColumn413,
																										 ultraGridColumn414,
																										 ultraGridColumn415,
																										 ultraGridColumn416,
																										 ultraGridColumn417,
																										 ultraGridColumn418,
																										 ultraGridColumn419,
																										 ultraGridColumn420,
																										 ultraGridColumn421,
																										 ultraGridColumn422,
																										 ultraGridColumn423,
																										 ultraGridColumn424,
																										 ultraGridColumn425,
																										 ultraGridColumn426,
																										 ultraGridColumn427,
																										 ultraGridColumn428,
																										 ultraGridColumn429,
																										 ultraGridColumn430,
																										 ultraGridColumn431,
																										 ultraGridColumn432,
																										 ultraGridColumn433,
																										 ultraGridColumn434,
																										 ultraGridColumn435,
																										 ultraGridColumn436,
																										 ultraGridColumn437,
																										 ultraGridColumn438,
																										 ultraGridColumn439,
																										 ultraGridColumn440,
																										 ultraGridColumn441,
																										 ultraGridColumn442,
																										 ultraGridColumn443});
			this.cmbCobrador1.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbCobrador1.DisplayMember = "Nombre";
			this.cmbCobrador1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador1.Location = new System.Drawing.Point(496, 172);
			this.cmbCobrador1.Name = "cmbCobrador1";
			this.cmbCobrador1.Size = new System.Drawing.Size(136, 69);
			this.cmbCobrador1.TabIndex = 34;
			this.cmbCobrador1.ValueMember = "idPersonal";
			this.cmbCobrador1.Visible = false;
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
			// btFactura
			// 
			this.btFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btFactura.Location = new System.Drawing.Point(440, 350);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(32, 25);
			this.btFactura.TabIndex = 35;
			this.btFactura.Text = "FI";
			this.toolTip1.SetToolTip(this.btFactura, "Factura de la izquierda");
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// cmbBanco
			// 
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idBanco"));
			this.cmbBanco.DataMember = "Banco";
			this.cmbBanco.DataSource = this.cdsCuentaCorr;
			this.cmbBanco.DisplayLayout.AutoFitColumns = true;
			ultraGridBand10.AddButtonCaption = "CompraFormaPago";
			ultraGridColumn444.Header.VisiblePosition = 4;
			ultraGridColumn444.Hidden = true;
			ultraGridColumn445.Header.VisiblePosition = 0;
			ultraGridColumn445.Width = 157;
			ultraGridColumn446.Header.VisiblePosition = 1;
			ultraGridColumn446.Hidden = true;
			ultraGridColumn446.Width = 81;
			ultraGridColumn447.Header.VisiblePosition = 3;
			ultraGridColumn447.Hidden = true;
			ultraGridColumn447.Width = 83;
			ultraGridColumn448.Header.VisiblePosition = 2;
			ultraGridColumn448.Hidden = true;
			ultraGridColumn449.Header.VisiblePosition = 5;
			ultraGridColumn449.Hidden = true;
			ultraGridColumn449.Width = 43;
			ultraGridColumn450.Header.VisiblePosition = 6;
			ultraGridColumn450.Hidden = true;
			ultraGridColumn450.Width = 45;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn444,
																											ultraGridColumn445,
																											ultraGridColumn446,
																											ultraGridColumn447,
																											ultraGridColumn448,
																											ultraGridColumn449,
																											ultraGridColumn450});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(112, 112);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(176, 21);
			this.cmbBanco.TabIndex = 36;
			this.cmbBanco.ValueMember = "idBanco";
			// 
			// cmbBanco1
			// 
			this.cmbBanco1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco1.DataMember = "Banco";
			this.cmbBanco1.DataSource = this.cdsCuentaCorr;
			ultraGridColumn451.Header.VisiblePosition = 3;
			ultraGridColumn451.Hidden = true;
			ultraGridColumn452.Header.VisiblePosition = 0;
			ultraGridColumn452.Width = 142;
			ultraGridColumn453.Header.VisiblePosition = 1;
			ultraGridColumn453.Hidden = true;
			ultraGridColumn454.Header.VisiblePosition = 4;
			ultraGridColumn454.Hidden = true;
			ultraGridColumn455.Header.VisiblePosition = 2;
			ultraGridColumn455.Hidden = true;
			ultraGridColumn456.Header.VisiblePosition = 5;
			ultraGridColumn456.Hidden = true;
			ultraGridColumn457.Header.VisiblePosition = 6;
			ultraGridColumn457.Hidden = true;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn451,
																											ultraGridColumn452,
																											ultraGridColumn453,
																											ultraGridColumn454,
																											ultraGridColumn455,
																											ultraGridColumn456,
																											ultraGridColumn457});
			this.cmbBanco1.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbBanco1.DisplayMember = "Nombre";
			this.cmbBanco1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco1.Location = new System.Drawing.Point(448, 224);
			this.cmbBanco1.Name = "cmbBanco1";
			this.cmbBanco1.Size = new System.Drawing.Size(160, 86);
			this.cmbBanco1.TabIndex = 37;
			this.cmbBanco1.ValueMember = "idBanco";
			this.cmbBanco1.Visible = false;
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
			// chkAsiento
			// 
			this.chkAsiento.Checked = true;
			this.chkAsiento.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			this.chkAsiento.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPagoLote, "_PagoLote.AsiEgr"));
			this.chkAsiento.Location = new System.Drawing.Point(560, 34);
			this.chkAsiento.Name = "chkAsiento";
			this.chkAsiento.Size = new System.Drawing.Size(72, 22);
			this.chkAsiento.TabIndex = 38;
			this.chkAsiento.Text = "Asiento";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(312, 112);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(34, 16);
			this.label9.TabIndex = 39;
			this.label9.Text = "Refer.";
			// 
			// txtReferencia
			// 
			this.txtReferencia.AlwaysInEditMode = true;
			this.txtReferencia.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.Referencia"));
			this.txtReferencia.Location = new System.Drawing.Point(368, 112);
			this.txtReferencia.Name = "txtReferencia";
			this.txtReferencia.Size = new System.Drawing.Size(248, 22);
			this.txtReferencia.TabIndex = 40;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(16, 112);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(34, 16);
			this.label10.TabIndex = 41;
			this.label10.Text = "Banco";
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.Total"));
			this.txtTotal.FormatString = "#,###0.00";
			this.txtTotal.Location = new System.Drawing.Point(744, 350);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(88, 22);
			this.txtTotal.TabIndex = 42;
			this.txtTotal.DoubleClick += new System.EventHandler(this.txtTotal_DoubleClick);
			// 
			// txtDeseado
			// 
			this.txtDeseado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtDeseado.FormatString = "#,###0.00";
			this.txtDeseado.Location = new System.Drawing.Point(832, 350);
			this.txtDeseado.Name = "txtDeseado";
			this.txtDeseado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDeseado.PromptChar = ' ';
			this.txtDeseado.Size = new System.Drawing.Size(88, 22);
			this.txtDeseado.TabIndex = 43;
			// 
			// chkPendiente
			// 
			this.chkPendiente.Location = new System.Drawing.Point(320, 138);
			this.chkPendiente.Name = "chkPendiente";
			this.chkPendiente.Size = new System.Drawing.Size(88, 21);
			this.chkPendiente.TabIndex = 44;
			this.chkPendiente.Text = "Pendiente";
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(16, 138);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(65, 16);
			this.label11.TabIndex = 46;
			this.label11.Text = "Vencimiento";
			// 
			// cmbVencimiento
			// 
			this.cmbVencimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbVencimiento.DateButtons.Add(dateButton2);
			this.cmbVencimiento.Format = "dd/MMM/yyyy";
			this.cmbVencimiento.Location = new System.Drawing.Point(112, 138);
			this.cmbVencimiento.Name = "cmbVencimiento";
			this.cmbVencimiento.NonAutoSizeHeight = 23;
			this.cmbVencimiento.Size = new System.Drawing.Size(128, 21);
			this.cmbVencimiento.SpinButtonsVisible = true;
			this.cmbVencimiento.TabIndex = 45;
			this.cmbVencimiento.Value = ((object)(resources.GetObject("cmbVencimiento.Value")));
			// 
			// btAsiRet
			// 
			this.btAsiRet.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsiRet.Location = new System.Drawing.Point(376, 350);
			this.btAsiRet.Name = "btAsiRet";
			this.btAsiRet.Size = new System.Drawing.Size(24, 25);
			this.btAsiRet.TabIndex = 47;
			this.btAsiRet.Text = "R";
			this.toolTip1.SetToolTip(this.btAsiRet, "Crea Asiento de Retención / Derecho: Mostrar Asiento de Retención");
			this.btAsiRet.Click += new System.EventHandler(this.btAsiRet_Click);
			this.btAsiRet.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btAsiRet_MouseDown);
			// 
			// btPago
			// 
			this.btPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btPago.Location = new System.Drawing.Point(504, 350);
			this.btPago.Name = "btPago";
			this.btPago.Size = new System.Drawing.Size(32, 25);
			this.btPago.TabIndex = 49;
			this.btPago.Text = "Pg";
			this.toolTip1.SetToolTip(this.btPago, "Muestra el pago");
			this.btPago.Click += new System.EventHandler(this.btPago_Click);
			this.btPago.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btPago_MouseDown);
			// 
			// btCash
			// 
			this.btCash.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCash.Location = new System.Drawing.Point(536, 350);
			this.btCash.Name = "btCash";
			this.btCash.Size = new System.Drawing.Size(40, 25);
			this.btCash.TabIndex = 51;
			this.btCash.Text = "Cash";
			this.toolTip1.SetToolTip(this.btCash, "Crea Reporte Cash");
			this.btCash.Click += new System.EventHandler(this.btCash_Click);
			// 
			// btFacturaD
			// 
			this.btFacturaD.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btFacturaD.Location = new System.Drawing.Point(472, 350);
			this.btFacturaD.Name = "btFacturaD";
			this.btFacturaD.Size = new System.Drawing.Size(32, 25);
			this.btFacturaD.TabIndex = 52;
			this.btFacturaD.Text = "FD";
			this.toolTip1.SetToolTip(this.btFacturaD, "Factura de la Derecha");
			this.btFacturaD.Click += new System.EventHandler(this.btFacturaD_Click);
			// 
			// btDeposito
			// 
			this.btDeposito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance20.Image = 7;
			this.btDeposito.Appearance = appearance20;
			this.btDeposito.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btDeposito.ImageList = this.imageList1;
			this.btDeposito.Location = new System.Drawing.Point(400, 350);
			this.btDeposito.Name = "btDeposito";
			this.btDeposito.Size = new System.Drawing.Size(26, 25);
			this.btDeposito.TabIndex = 53;
			this.toolTip1.SetToolTip(this.btDeposito, "Impresión de Depósito");
			this.btDeposito.UseMnemonic = false;
			this.btDeposito.Click += new System.EventHandler(this.btDeposito_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// txtIdAsiRetL
			// 
			this.txtIdAsiRetL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idAsiRetL"));
			this.txtIdAsiRetL.Location = new System.Drawing.Point(616, 9);
			this.txtIdAsiRetL.Name = "txtIdAsiRetL";
			this.txtIdAsiRetL.Size = new System.Drawing.Size(40, 22);
			this.txtIdAsiRetL.TabIndex = 48;
			// 
			// chkReceptor
			// 
			this.chkReceptor.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPagoLote, "_PagoLote.Receptor"));
			this.chkReceptor.Location = new System.Drawing.Point(464, 138);
			this.chkReceptor.Name = "chkReceptor";
			this.chkReceptor.Size = new System.Drawing.Size(136, 21);
			this.chkReceptor.TabIndex = 50;
			this.chkReceptor.Text = "Filtro por Receptor";
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
			// PagoLotes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(936, 390);
			this.Controls.Add(this.btDeposito);
			this.Controls.Add(this.btFacturaD);
			this.Controls.Add(this.btCash);
			this.Controls.Add(this.chkReceptor);
			this.Controls.Add(this.btPago);
			this.Controls.Add(this.txtIdAsiRetL);
			this.Controls.Add(this.btAsiRet);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbVencimiento);
			this.Controls.Add(this.chkPendiente);
			this.Controls.Add(this.txtDeseado);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtReferencia);
			this.Controls.Add(this.chkAsiento);
			this.Controls.Add(this.cmbBanco1);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.cmbCobrador1);
			this.Controls.Add(this.txtCheque);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbFormaPago1);
			this.Controls.Add(this.cmbNumCuenta);
			this.Controls.Add(this.btRetencion);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.chkIncluyenombre);
			this.Controls.Add(this.chkDesglosa);
			this.Controls.Add(this.txtAsiento);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.cmbCobrador);
			this.Controls.Add(this.cmbCuentaCorr);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.idPagoLote);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.btAgregarTodo);
			this.Controls.Add(this.btAgregar);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "PagoLotes";
			this.Text = "Pagos por Lotes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.PagoLotes_Closing);
			this.Load += new System.EventHandler(this.PagoLotes_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagoLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.idPagoLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNumCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsignaCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeseado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVencimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiRetL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables Globales
		int IdTipoFactura = 0;
		int iIdPagoLote = 0;
		C1.Data.C1DataRow drPagoLote;
		C1.Data.C1DataRow drSeteoF;
		CultureInfo us = new CultureInfo("en-US");
		C1.Data.C1DataRow drSeteo;
		Acceso miAcceso;
		#endregion Variables Globales
		#region Pantalla PagoLote
		public PagoLotes(int idTipoFactura)
		{
			InitializeComponent();
			IdTipoFactura = idTipoFactura;
		}

		public PagoLotes(int idTipoFactura, int idPagoLote)
		{
			InitializeComponent();
			IdTipoFactura = idTipoFactura;
			iIdPagoLote = idPagoLote;
		}

		bool bAsientoAutomatico = false;
		private void PagoLotes_Load(object sender, System.EventArgs e)
		{
			#region Auditoria
			string stAudita = "Exec AuditaCrear 53, 6, 'PagoLote'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
			#endregion 
			#region Seguridad
			string stVerifica = "Exec PagoSeguridadCxC " + IdTipoFactura;
			if (Funcion.bEjecutaSQL(cdsCliente, stVerifica))
				miAcceso =new Acceso(cdsSeteo, "0417");
			else
				miAcceso =new Acceso(cdsSeteo, "0418");
			#endregion
			#region Filtro
			barraDato1.stFiltroExt = " And idTipoFactura = " + IdTipoFactura + " And IsNull(idDeposito, 0) = 0";
			string stFiltro="";
			if (IdTipoFactura == 1) 
			{
				Text = "Cuentas por Cobrar por Lotes";
				stFiltro = "Proveedor = 0 Or Ambos <> 0";
				txtCodigo.ReadOnly = Funcion.bEjecutaSQL(cdsAsiento, 
					"Select IsNull(Automatico, 0) From CompraNumero Where idTipoFactura = 21");
			}
			else 
			{
				Text = "Cuentas por Pagar por Lotes";
				toolTip1.SetToolTip(txtCheque, "Número de Depósito");
				stFiltro = " Proveedor <> 0 Or Ambos <> 0";
				txtCodigo.ReadOnly = Funcion.bEjecutaSQL(cdsAsiento, 
					"Select IsNull(Automatico, 0) From CompraNumero Where idTipoFactura = 22");
				label1.Text = "Prov.";
			}
			C1.Data.FilterCondition fcCliente = new 
				C1.Data.FilterCondition(cdsCliente, "Cliente", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcCliente);
			cdsCliente.Fill(fcTotal, false);
			#endregion
			#region Buscar
			if (iIdPagoLote > 0)
			{
				barraDato1.IdRegistro = iIdPagoLote;
				barraDato1.ProximoId(5);
			}
			else
				barraDato1.PosUltima();
			#endregion Buscar
			#region Inicio
			barraDato1.HabilitaControles(false);
			txtIdAsiRetL.Width = 0;
			idPagoLote.Width = 0;
			txtIdAsiento.Width = 0;
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			drSeteo = cdsSeteo.TableViews["Seteo"].Rows[0];
			ultraGrid1.Enabled = true;
			ultraGrid2.Enabled = true;
			Cursor = Cursors.Default;
			#endregion Inicio
			#region Asiento Automatico
			string stExec = "Exec SeteoGExiste 'FALC'";
			if (IdTipoFactura == 4) stExec = "Exec SeteoGExiste 'FALP'";
			bAsientoAutomatico = Funcion.bEjecutaSQL(cdsPagoLote, stExec);
			if (bAsientoAutomatico) btAsiento.Text = "a";
			#endregion Asiento Automatico
		}

		private void PagoLotes_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (barraDato1.bEditar)
			{
				e.Cancel = true;
				MessageBox.Show("No puede Salir de esta pantalla mientras edita un registro");
			}
			string stAudita = "Exec AuditaCrear 53, 7, 'PagoLote'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
		}


		#endregion Pantalla PagoLote
		#region cds
		private void cdsPagoLote_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			if (e.DataMember.ToString() == "_PagoLote")
			{
				drPagoLote = e.Row;
			}
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoC.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

			string stbProveedor = "";
			switch(IdTipoFactura)
			{
				case 1:
				case 3:
				case 5:
				case 13:
				case 15:
					stbProveedor = "False";
					break;
				case 2:
				case 4:
				case 6:
				case 14:
				case 16:
					stbProveedor = "True";
					break;
			}
			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsClientes"];

			if (stbProveedor == "True") stbProveedor = " <> 0";
			else stbProveedor = " = 0";

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"],
				"[Ambos] <> 0 OR [Proveedor] " + stbProveedor));
		}

		private void cdsPagoLote_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPagoLote.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsFormaPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsFormaPago.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCuentaCorr_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuentaCorr.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsAsiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAsiento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCompraTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompraTabla.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsAsignaCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAsignaCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}


		#endregion cds
		#region Barra Standard
		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Eliminar Pagos", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea Borrar este pago por lotes?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, 
				MessageBoxDefaultButton.Button2)) return;

			int IdPagoLote = int.Parse(this.idPagoLote.Value.ToString());
			string stMensaje = "Exec PagoLoteVerificaBorrar " + IdPagoLote;
			stMensaje = Funcion.sEscalarSQL(cdsPagoLote, stMensaje);
			if (stMensaje.Length > 0)
			{
				if (stMensaje.StartsWith("Precauci")) // Confirmacion para anular asiento
				{
					if (DialogResult.No == MessageBox.Show(stMensaje + "¿Desea Anularlo?", 
						"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk)) return;
				}
				else
				{
					MessageBox.Show(stMensaje, "Información",
						MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}
			stMensaje = "Exec PagoLoteBorrar " + IdPagoLote;
			Funcion.EjecutaSQL(cdsPagoLote, stMensaje);
			barraDato1.ProximoId(2);

			MessageBox.Show("Registro Borrado",
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using(BuscaPagoLotes miPagoLote = new BuscaPagoLotes(IdTipoFactura))
			{
				if (DialogResult.OK == miPagoLote.ShowDialog())
				{
					try
					{
						int IdBusca = int.Parse(miPagoLote.idBusca.Text);
						barraDato1.IdRegistro = IdBusca;
						barraDato1.ProximoId(5);
					}
					catch{};
				}
			}
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene Acceso a Crear Pagos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			barraDato1.CrearRegistro();
			drPagoLote["idTipoFactura"] = IdTipoFactura;
			drPagoLote["FechaIngreso"] = Funcion.dtEscalarSQL(cdsAsiento, "Select GetDate()");
			int IdBusca = int.Parse(idPagoLote.Value.ToString());
			barraDato1.GrabaRegistro();
			barraDato1.IdRegistro = IdBusca;
			barraDato1.ProximoId(5);
			barraDato1.EditarRegistro();
			barraDato1.bNuevo = true;
			ValoresIniciales();
		}
		private void ValoresIniciales()
		{
			drPagoLote = cdsPagoLote.TableViews["PagoLote"].Rows.Find
				(int.Parse(idPagoLote.Value.ToString()));
			drPagoLote["Fecha"] = DateTime.Today;
			drPagoLote["idTipoFactura"] = IdTipoFactura;
			Habilita(true);
		}


		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			cdsPagoLote.RejectChanges();
			barraDato1.DeshacerRegistro();
			Habilita(false);
			BorraErrores();
			ultraGrid1.Enabled = true;
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (drPagoLote == null)
			{
				MessageBox.Show("No existe un registro para editar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene Acceso a Editar Pagos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			string stSelect = "Select Count(*) From Pago Where Bloquear = 1 And idPagoLote = "
				+ idPagoLote.Value.ToString();
			int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			if (iCuenta > 0)
			{
				MessageBox.Show("No puede editar registro porque existe algún pago bloqueado",
					"Información");
				return;
			}
			barraDato1.EditarRegistro();
			drPagoLote = cdsPagoLote.TableViews["PagoLote"].Rows.Find(int.Parse(idPagoLote.Value.ToString()));
			Habilita(true);
			chkAsiento.Enabled = false;
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Retencion.idPagoGrupo} = " + idPagoLote.Value.ToString();
			string stRep = "RetencionLote.Rpt";
			Reporte miRep = new Reporte(stRep, stFiltro);
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;      
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Pago.idPagoLote} = " + idPagoLote.Value.ToString();
			string stRep = "AbonoLotes.Rpt";
			if (IdTipoFactura == 1) stRep = "CobroLotes.Rpt";
				Reporte miRep = new Reporte(stRep, stFiltro);
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
      		
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			double dTotal = (double) txtTotal.Value;
			if (!Verifica()) return;

			drPagoLote = cdsPagoLote.TableViews["PagoLote"].Rows[0];
			NumeroPago();
		
			if (!barraDato1.bNuevo)
			{
				drPagoLote["Editar"] = true;
				drPagoLote["FechaEditado"] = DateTime.Now;
			}
			drPagoLote["idTipoFactura"] = IdTipoFactura;
			drPagoLote["Total"] = dTotal;
			barraDato1.GrabaRegistro();
			string stPagoLote = string.Format("Exec PagoLotePagoAct {0}, {1}",
				idPagoLote.Value, IdTipoFactura);
			Funcion.EjecutaSQL(cdsAsiento, stPagoLote, true);

			Habilita(false);
			ultraGrid1.Enabled = true;
			ultraGrid2.Enabled = true;
			btFiltro_Click(this, e);
			if (bAsientoAutomatico)
			{
				barraDato1.ProximoId(5);
				string stProcPagoLote = "Exec PagoLoteAsiento " 
					+ idPagoLote.Value.ToString() + ", " + IdTipoFactura.ToString();
				Funcion.EjecutaSQL(cdsAsiento, stProcPagoLote, true);
				barraDato1.ProximoId(5);
			}
			barraDato1.bEditar = true;
			btFiltro_Click(this, e);
			barraDato1.bEditar = false;
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
		
		}


		#endregion Barra Standard
		#region Barra Extendida
		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			#region Asiento Automatico
			if (bAsientoAutomatico)
			{
				AsientoAbrir();
				return;
			}
			#endregion Asiento Automatico

			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene acceso a Crear Asientos", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (barraDato1.bEditar) 
			{
				MessageBox.Show("No puede crear un asiento mientras está editando",
					"Información");
				return;
			}

			string stVerifica = "Exec PagoLoteVerificaAsiento " + idPagoLote.Text;
			string stMensaje = Funcion.sEscalarSQL(cdsCliente, stVerifica);
			if (stMensaje.Length > 0)
			{
				if (stMensaje.StartsWith("Precau"))
				{
					if (DialogResult.No == MessageBox.Show(stMensaje + "¿Desea Sobreescribirlo?", "Confirmación",
						MessageBoxButtons.YesNo, MessageBoxIcon.Question) ) return;
				}
				else
				{
					MessageBox.Show(stMensaje, "Información",
						MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}

			string stProcPagoLote = "Exec PagoLoteAsiento " 
				+ idPagoLote.Value.ToString() + ", " + IdTipoFactura.ToString();
			Funcion.EjecutaSQL(cdsAsiento, stProcPagoLote, true);
			barraDato1.ProximoId(5);
			Cursor = Cursors.Default;
			MessageBox.Show("Asiento Generado", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);

			#region Crea asiento antes
			//			C1.Data.C1DataRow drAsiento;
			//			if (bRepetido)
			//			{
			//				string stFiltro = "idAsiento = " + txtIdAsiento.Text.ToString();
			//				C1.Data.FilterCondition fcPadre = new 
			//					C1.Data.FilterCondition(cdsAsiento, "Asiento", stFiltro);
			//				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			//				fcTotal.Add(fcPadre);
			//
			//				C1.Data.FilterCondition fcHija = new 
			//					C1.Data.FilterCondition(cdsAsiento, "AsientoDetalle", "idAsiento = 0");
			//				fcTotal.Add(fcHija);
			//			
			//				this.cdsAsiento.Fill(fcTotal, false);
			//
			//				drAsiento = cdsAsiento.TableViews["Asiento"].Rows[0];
			//			}
			//			else
			//			{
			//				drAsiento = cdsAsiento.TableViews["Asiento"].AddNew();
			//			}
			//
			//			drAsiento["Fecha"] = (DateTime) this.txtFecha.Value;
			//
			//			string stBuscaTipoA_E = "SELECT Asiento FROM CuentaCorriente "
			//				+ "WHERE idCuentaCorriente = "
			//				+ cmbCuentaCorr.Value.ToString();
			//			bool bAsiento = Funcion.bEscalarSQL(cdsPagoLote, stBuscaTipoA_E, false);
			//			if (bAsiento == true) 
			//			{
			//				drAsiento["idTipoAsiento"] = 1; // asiento
			//			}
			//			else
			//			{
			//				if (IdTipoFactura == 1)	drAsiento["idTipoAsiento"] = 2; // Ingreso
			//				else drAsiento["idTipoAsiento"] = 3; // Egreso
			//			}
			//			if (chkAsiento.Checked) 
			//				drAsiento["idTipoAsiento"] = 1;
			//
			//			drAsiento["Descripcion"] = this.cmbCliente.Text;
			//			//			drAsiento["idProyecto"] = (int) this.cmbProyecto.EditValue;
			//			//			drAsiento["idSubProyecto"] = (int) this.cmbSubProyecto.EditValue;
			//
			//			drAsiento["Notas"] = "";
			//			if ((bool) drSeteoF["CopiarNota"] == true && this.txtNotas.Value != null)
			//				drAsiento["Notas"] = this.txtNotas.Value;
			//
			//			if ((bool) drSeteoF["CopiarNumero"] == true && this.txtCodigo.Value != null)
			//				drAsiento["Numero"] = this.txtCodigo.Value;
			//
			//			// Crea el codigo del asiento
			//			if (!bRepetido)
			//				drAsiento["CodAsiento"] = Contabilidad.NumeroAsiento(this.cdsAsiento, (int) drAsiento["idTipoAsiento"]);
			//			string stFactura = "";
			//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			//				stFactura += dr.Cells["Notas"].Value.ToString().Trim() + " ";
			//			if (txtNotas.Text.Trim().Length == 0)
			//				drAsiento["Notas"] += stFactura;
			//
			//			if (this.cmbCliente.Text.Length == 0)
			//				drAsiento["Descripcion"] = "Pago por Lotes";
			//			else
			//				drAsiento["Descripcion"] = cmbCliente.Text;
			//
			//			// Actualiza el codigo del asiento en cada uno de los pagos
			//			double dTotal = 0;
			//			string stNota = "";
			//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			//			{
			//				dTotal += (double) dr.Cells["Pago"].Value;
			//				stNota += dr.Cells["Concepto"].Value.ToString() + " ";
			//			}
			//			dTotal = Math.Round(dTotal, 2);
			//			drAsiento["Total"] = dTotal;
			//			drAsiento["TotalH"] = dTotal;
			//			drAsiento["Notas"] = stNota;
			//			drAsiento["Cheque"] = txtCheque.Text;
			//			this.cdsAsiento.Update();
			//
			//			// Creacion de Detalle de Asiento
			//			this.GrupoCxC((int) drAsiento["idAsiento"], dTotal);
			//
			//			drPagoLote = cdsPagoLote.TableViews["PagoLote"].Rows.Find(int.Parse(idPagoLote.Value.ToString()));
			//			drPagoLote["idAsiento"] = drAsiento["idAsiento"];
			//			drPagoLote["CodAsiento"] = drAsiento["CodAsiento"];
			//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			//			{
			//				dr.Cells["NumeroIngreso"].Value = drAsiento["CodAsiento"];
			//				dr.Cells["idAsiento"].Value = drAsiento["idAsiento"];
			//			}
			//			this.cdsPagoLote.Update();
			//			
			//			// Creacion de Retencion
			//			if (!bRepetido)
			//			{
			//				int IngEgr = 2;
			//				if (IdTipoFactura == 1) IngEgr = 1;
			//
			//				int idPrinc = int.Parse(txtIdAsiento.Value.ToString());
			//				int idCuentaCorr = int.Parse(cmbCuentaCorr.Value.ToString());
			//				int IdCliente = int.Parse(cmbCliente.Value.ToString());
			//				DateTime dtFecha = (DateTime) txtFecha.Value;
			//				string stBeneficiario = cmbCliente.Text.Trim();
			//
			//				string stSelect = "Insert Into Retencion (idAsiento"
			//					+ ", idCuentaCorriente, idCliente, Fecha "
			//					+ ",Cheque,IngresoEgreso, Total, CodAsiento, FacturaIva0, "
			//					+ "Beneficiario) VALUES (" + idPrinc.ToString()
			//					+ ", " + idCuentaCorr.ToString() 
			//					+ ", " + IdCliente.ToString()
			//					+ ", '" + dtFecha.ToString("yyyyMMdd")
			//					+ "', '" + txtCheque.Text.ToString().Trim() + "', " + IngEgr.ToString()
			//					+ ", " + dTotal.ToString("0.00", us)
			//					+ ", '" + drAsiento["CodAsiento"].ToString() + "'"
			//					+ ", " + dTotal.ToString("0.00", us)
			//					+ ", '" + stBeneficiario + "')";
			////				Funcion.EjecutaSQL(cdsAsiento, stSelect);
			//			}

			//			Cursor = Cursors.Default;
			//			MessageBox.Show("Asiento Generado", "Información",
			//				MessageBoxButtons.OK, MessageBoxIcon.Information);
			#endregion Crea asiento
		}

		private void AsientoAbrir()
		{
			if (int.Parse(this.txtIdAsiento.Value.ToString()) == 0)
			{
				MessageBox.Show("No existe un asiento creado de estos Pagos");
				return;
			}
			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(int.Parse(this.txtIdAsiento.Value.ToString()));
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void btAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			AsientoAbrir();
		}

		private void btRetencion_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bNuevo)
			{
				MessageBox.Show("Debe grabar el Registro para generar la retencion",
					"Información");
				return;
			}
			if (ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la factura para crear una retencion",
					"Información");
				return;
			}
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;

			C1.Data.C1DataRow drCompra = cdsCompra.TableViews["Compra"].Rows.Find(IdCompra);
			if (this.IdTipoFactura == 4) // Compra
			{
//				if (drSeteo["FormMarzo"] == System.DBNull.Value 
//					|| (bool) drSeteo["FormMarzo"] == false)
//				{
//					using (Egreso miEgreso = new Egreso())
//					{
//						miEgreso.IdFactura = IdCompra;
//						miEgreso.drFactura = drCompra;
//						miEgreso.ShowDialog();
//					}
//				}
//				else
				{
//					using(EgresoAT miEgreso = new EgresoAT())
//					{
//						miEgreso.IdFactura = IdCompra;
//						miEgreso.drFactura = drCompra;
//						miEgreso.txtIdPagoGrupo.EditValue = int.Parse(idPagoLote.Value.ToString());
//						#region Tratamiento de varias retenciones
//						if (ultraGrid1.Selected.Rows.Count > 1)
//						{
//							#region Verificacion de Retenciones Anteriores
//							double dRetencion = 0;
//							string stErrores = "";
//							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
//							{
//								dRetencion = 0;
//								if (dr.Cells["Retenido"].Value != System.DBNull.Value)
//									dRetencion = (double) dr.Cells["Retenido"].Value;
//								if (dRetencion >= .001)
//								{
//									if (dr.Cells["Numero"].Value != System.DBNull.Value)
//										stErrores += "Factura " + dr.Cells["numero"].Value.ToString() + " ya tiene retencion\n";
//								}
//							}
//							if (stErrores.Trim().Length > 0)
//							{
//								MessageBox.Show("Estas facturas tienen retenciones anteriores: \n" + stErrores, "Información");
//								Cursor = Cursors.Default;
//								return;
//							}
//							#endregion Verificacion de Retenciones Anteriores
//							#region Cambio de Identificacion de IdPagoLote
//							int IdCompraGrupo;
//							string stSelect;
//							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
//							{
//								IdCompraGrupo = 0;
//								if (dr.Cells["idCompra"].Value != System.DBNull.Value)
//									IdCompraGrupo = (int) dr.Cells["idCompra"].Value;
//								stSelect = "Update Compra Set IdPagoLot = " + idPagoLote.Value.ToString()
//									+ " Where IdCompra = " + IdCompraGrupo.ToString();
//								Funcion.EjecutaSQL(cdsCompra, stSelect, true);
//							}
//							miEgreso.txtIdPagoLote.EditValue = int.Parse(idPagoLote.Value.ToString());
//							#endregion Cambio de Identificacion de IdPagoLote
//							if (DialogResult.OK == miEgreso.ShowDialog())
//							{
//								string stDistribuye = "EXEC DistribuyeRet " + idPagoLote.Value.ToString();
//								Funcion.EjecutaSQL(cdsCompra, stDistribuye, true);
//							}
//							else
//							{
//								string stDistribuye = "Update Compra set idPagoLot = 0 Where idPagoLot = " 
//									+ idPagoLote.Value.ToString();
//								Funcion.EjecutaSQL(cdsCompra, stDistribuye, true);
//							}
//						}
//							#endregion Tratamiento de varias retenciones
//						else
//							miEgreso.ShowDialog();
//					}
				}
			}

			if (this.IdTipoFactura == 1) // Venta
			{
//				if (drSeteo["FormMarzo"] == System.DBNull.Value 
//					|| (bool) drSeteo["FormMarzo"] == false)
//				{
//					using (Ingreso miIngreso = new Ingreso())
//					{
//						miIngreso.IdFactura = IdCompra;
//						miIngreso.drFactura = drCompra;
//						miIngreso.ShowDialog();
//					}
//				}
//				else
				{
					using (IngresoAT miIngreso = new IngresoAT())
					{
						miIngreso.IdFactura = IdCompra;
						miIngreso.drFactura = drCompra;
						miIngreso.txtIdPagoGrupo.EditValue = int.Parse(idPagoLote.Value.ToString());
						#region Tratamiento de varias retenciones
						if (ultraGrid1.Selected.Rows.Count > 1)
						{
							#region Verificacion de Retenciones Anteriores
							double dRetencion = 0;
							string stErrores = "";
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
							{
								dRetencion = 0;
								if (dr.Cells["Retenido"].Value != System.DBNull.Value)
									dRetencion = (double) dr.Cells["Retenido"].Value;
								if (dRetencion >= .001)
								{
									if (dr.Cells["Numero"].Value != System.DBNull.Value)
										stErrores += "Factura " + dr.Cells["numero"].Value.ToString() + " ya tiene retencion\n";
								}
							}
							if (stErrores.Trim().Length > 0)
							{
								MessageBox.Show("Estas facturas tienen retenciones anteriores: \n" + stErrores, "Información");
								Cursor = Cursors.Default;
								return;
							}
							#endregion Verificacion de Retenciones Anteriores
							#region Cambio de Identificacion de IdPagoLote
							int IdCompraGrupo;
							string stSelect;
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
							{
								IdCompraGrupo = 0;
								if (dr.Cells["idCompra"].Value != System.DBNull.Value)
									IdCompraGrupo = (int) dr.Cells["idCompra"].Value;
								stSelect = "Update Compra Set IdPagoLot = " + idPagoLote.Value.ToString()
									+ " Where IdCompra = " + IdCompraGrupo.ToString();
								Funcion.EjecutaSQL(cdsCompra, stSelect, true);
							}
							miIngreso.txtIdPagoLote.EditValue = int.Parse(idPagoLote.Value.ToString());
							#endregion Cambio de Identificacion de IdPagoLote
							if (DialogResult.OK == miIngreso.ShowDialog())
							{
								string stDistribuye = "EXEC DistribuyeRet " + idPagoLote.Value.ToString();
								Funcion.EjecutaSQL(cdsCompra, stDistribuye, true);
							}
							else
							{
								string stDistribuye = "Update Compra set idPagoLot = 0 Where idPagoLot = " 
									+ idPagoLote.Value.ToString();
								Funcion.EjecutaSQL(cdsCompra, stDistribuye, true);
							}
						}
							#endregion Tratamiento de varias retenciones
						else
							miIngreso.ShowDialog();
					}
				}
			}
			btFiltro_Click(this, e);
			Cursor = Cursors.Default;
		}

		private void btRetencion_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			AsiendoDeRetencion();
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Factura de la grilla izquierda que desea ver el detalle", "Información");
				return;
			}
			Cursor = Cursors.WaitCursor;
			int IdFact = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
//			Compra miCompra = new Compra(IdTipoFactura, IdFact);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void btAsiRet_Click(object sender, System.EventArgs e)
		{
			AsiendoDeRetencion();
		}

		private void btAsiRet_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (this.txtIdAsiRetL.Text == "" || int.Parse(this.txtIdAsiRetL.Text) == 0)
			{
				MessageBox.Show("No existe un asiento creado de esta Retencion");
				return;
			}
			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(int.Parse(this.txtIdAsiRetL.Text));
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		
		#endregion Barra Extendida
		#region Global
		private void btQuitaTodo_Click(object sender, System.EventArgs e)
		{
			
		}

		private void btAgregarTodo_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btAgregar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
			{
				if (!VerificaAgregar()) return;
				txtTotal_DoubleClick(this, e);
				double dPagoActual = (double) dr.Cells["Saldo"].Value;
				if (!(IdTipoFactura == 1 || IdTipoFactura == 12))
				{
					if (drSeteoF["AutorizaPago"] != DBNull.Value && (bool) drSeteoF["AutorizaPago"] == true)
					{
						if (dr.Cells["PagoAutorizado"].Value == DBNull.Value)
							dPagoActual = 0;
						else
							dPagoActual = (double) dr.Cells["PagoAutorizado"].Value;
						if (dPagoActual < 0.01)
						{
							MessageBox.Show("No tiene autorizacion para realizar el pago", "Información");
							return;
						}
					}
				}
				double dSaldoAntes = (double) this.txtTotal.Value;
				double dDeseado = 0;
				if (this.txtDeseado != null)
					dDeseado = (double) this.txtDeseado.Value;
				if (dDeseado > 0)
				{
					if (dPagoActual + dSaldoAntes > dDeseado)
					{
						dPagoActual = Math.Round(dDeseado - dSaldoAntes, 2);
						if (dPagoActual < 0.01)
						{
							MessageBox.Show("No puede realizar mas pagos porque sobrepasa lo deseado", "Información");
							return;
						}
					}
				}

				// Necesario para refrescar la grilla
				C1.Data.C1DataRow drPago = cdsPagoLote.TableViews["Pago"].AddNew();
				drPago.CancelEdit();
				ultraGrid2.Refresh();

				drPago = cdsPagoLote.TableViews["Pago"].AddNew();
				drPago["Fecha"] = (DateTime) txtFecha.Value;
				drPago["idPagoLote"] = int.Parse(idPagoLote.Value.ToString());
				drPago["idCompra"] = (int) dr.Cells["idCompra"].Value;
				//			drPago["Pago"] = (double) this.ultraGrid1.ActiveRow.Cells["Saldo"].Value;
				drPago["Pago"] = dPagoActual;
				drPago["idFormaPago"] = (int) cmbFormaPago.Value;
				drPago["idBanco"] = (int) cmbBanco.Value;
				drPago["Nombre"] = cmbCliente.Text.Trim();
				drPago["Cuenta"] = txtNotas.Text.Trim();
				drPago["Cheque"] = txtCheque.Text.Trim();
				drPago["Referencia"] = txtReferencia.Text.Trim();
				drPago["idCuentaCorriente"] = (int) cmbCuentaCorr.Value;
				drPago["idCobrador"] = (int) cmbCobrador.Value;
				if (cmbVencimiento.Value != DBNull.Value)
					drPago["Vencimiento"] = (DateTime) cmbVencimiento.Value;
				if (chkPendiente.Checked)
					drPago["Pendiente"] = true;
				if (dr.Cells["Numero"].Value != System.DBNull.Value)
					drPago["Concepto"] = dr.Cells["Numero"].Value.ToString().Trim();
				ultraGrid2.Refresh();
				txtTotal_DoubleClick(this, e);
			}
		}


		#endregion
		private void BorraErrores()
		{
			this.errorProvider.SetError(this.cmbBanco, "");
			this.errorProvider.SetError(this.txtFecha, "");
			this.errorProvider.SetError(this.ultraGrid1, "");
			this.errorProvider.SetError(this.ultraGrid2, "");
			this.errorProvider.SetError(this.cmbCliente, "");
			this.errorProvider.SetError(this.cmbFormaPago, "");
			this.errorProvider.SetError(this.cmbCuentaCorr, "");
			this.errorProvider.SetError(this.txtCodigo, "");
		}

		private bool VerificaAgregar()
		{
			bool bOk = true;
			this.BorraErrores();
			if (this.txtFecha.Value == System.DBNull.Value)
			{
				this.errorProvider.SetError(this.txtFecha, "Ingrese Fecha");
				bOk = false;
			}
			if (ultraGrid1.ActiveRow == null || ultraGrid1.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				this.errorProvider.SetError(this.ultraGrid1, "Seleccione factura que desea realizar el pago");
				bOk = false;
			}
			if ((int) this.cmbFormaPago.Value == 0)
			{
				this.errorProvider.SetError(this.cmbFormaPago, "Ingrese Forma de Pago");
				bOk = false;
			}
			if((int) this.cmbCuentaCorr.Value == 0)
			{
				this.errorProvider.SetError(this.cmbCuentaCorr, "Ingrese Cuenta corriente");
				bOk = false;
			}
			if (IdTipoFactura == 1)
			{
				if(cmbBanco.Value == System.DBNull.Value || (int) this.cmbBanco.Value == 0)
				{
					this.errorProvider.SetError(this.cmbBanco, "Ingrese Banco");
					bOk = false;
				}
			}

			// Verificacion de Factura no repetida
			int IdC = (int) this.ultraGrid1.ActiveRow.Cells["idCompra"].Value;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if (dr.Cells["idCompra"].Value == System.DBNull.Value)
				{
					this.errorProvider.SetError(this.ultraGrid2, "IdCompra no valido");
					bOk = false;
				}
				else
				{
					if (IdC == (int) dr.Cells["idCompra"].Value)
					{
						this.errorProvider.SetError(this.ultraGrid2, "Factura ya existe en lista de pagos");
						bOk = false;
					}
				}
			}
				return bOk;
		}


		#region Grilla
		private void ultraGrid2_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.IsInEditMode)
				if (e.Cell.Column.ToString() == "Pago")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						if (drSeteoF["PagoNegativo"] == System.DBNull.Value ||
							(bool) drSeteoF["PagoNegativo"] == false)
						{
							string stSaldo = "Saldo";
							if (drSeteoF["AutorizaPago"] != DBNull.Value && (bool) drSeteoF["AutorizaPago"] == true)
								stSaldo = "PagoAutorizado";
							string stSelect = "Select " + stSaldo + " From Compra Where idCompra = " 
								+ e.Cell.Row.Cells["idCompra"].Value.ToString();
							double dSaldo = Funcion.dEscalarSQL(cdsCompra, stSelect, true);
							dSaldo = Math.Round(dSaldo, 2);
							if ((double) e.Cell.Value > dSaldo)
							{
								MessageBox.Show("Pago Excede Saldo", "Información");
								e.Cell.Row.Cells["Pago"].Value = dSaldo;
							}
              txtTotal_DoubleClick(this, e);
						}  // Fin de Pago Negativo = falso
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
						//						ultraGrid2.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//				MessageBox.Show(ex.Message, "Tecla");
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


		#endregion
		private void Habilita(bool bActiva)
		{
			this.btAgregar.Enabled = bActiva;
			this.btAgregarTodo.Enabled = bActiva;
//			this.btQuitar.Enabled = bActiva;
//			this.btQuitaTodo.Enabled = bActiva;
		}
		private void NumeroPago()
		{
			if (this.barraDato1.bNuevo)
			{
				int iTipo = 0;
				if (IdTipoFactura != 1) iTipo = 1;
				string stNumero = Funcion.NumeraFactura(this.cdsCompraTabla, 21 + iTipo);
				if (stNumero.Trim().Length > 0) drPagoLote["Codigo"] = stNumero;

				if (iTipo != 1) return;

				if (!chkAsiento.Checked)
				{
					int iNumCheque = Contabilidad.NumeroCheque(
						cdsCuentaCorr, (int) cmbCuentaCorr.Value, true);
					if (iNumCheque != -1)	drPagoLote["Cheque"] = iNumCheque.ToString();
				}
			}
		}

		private bool Verifica()
		{
			double dTotal = (double) txtTotal.Value;
			bool bOk = true;
			BorraErrores();
			if (txtFecha.Value == DBNull.Value)
			{
				errorProvider.SetError(txtFecha, "Ingrese Fecha");
				bOk = false;
			}
			if ((int) cmbFormaPago.Value == 0 && dTotal > 0)
			{
				errorProvider.SetError(cmbFormaPago, "Ingrese Forma de Pago");
				bOk = false;
			}
			if ((int) cmbCliente.Value == 0)
			{
				errorProvider.SetError(cmbCliente, "Ingrese Cliente");
				bOk = false;
			}
			if((int) cmbCuentaCorr.Value == 0 && dTotal > 0)
			{
				errorProvider.SetError(cmbCuentaCorr, "Ingrese Cuenta corriente");
				bOk = false;
			}
			string stCodigo = txtCodigo.Text.Trim();
			string stMensaje = Funcion.sEscalarSQL(cdsAsiento, 
        "Exec PagoLoteVerificaCodigo '" + stCodigo + "', " + idPagoLote.Value.ToString(), true);
			if(stMensaje.Length > 0)
			{
				errorProvider.SetError(txtCodigo, stMensaje);
				bOk = false;
			}
			return bOk;
		}

		private void GrupoCxC(int IdAsiento, double dTotal)
		{
			int IdCuenta = 0;
			double dValor = 0;
			C1.Data.C1DataRow drAsigna;
			C1.Data.C1DataRow drCC = cdsCuentaCorr.TableViews["CuentaCorriente"].Rows.Find((int) cmbCuentaCorr.Value);
			if ((int) cmbCuentaCorr.Value != 0 || drCC["idCuenta"] != null || (int) drCC["idCuenta"] == 0)
				IdCuenta = (int) drCC["idCuenta"];
			else
			{
				drAsigna = this.cdsAsignaCuenta.TableViews["AsignaCuenta"].Rows[0];
				IdCuenta = (int) drAsigna["idCuenta"]; // Caja o Bancos en Ventas
			}

			C1.Data.C1DataRow drDetAsi;
			int iCuenta1 = 0;
			if (IdTipoFactura == 1) 
			{
				if (!chkDesglosa.Checked) 
				{
					drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
					drDetAsi["Debe"] = dTotal;
					drDetAsi["idCuenta"] = IdCuenta;
					drDetAsi["idAsiento"] = IdAsiento;
				}
				else
				{
					string stCuentaGrupo = "select sum(IsNull(pago.pago, 0)), Max(cuentacorriente.idCuenta) "
						+ ", Max(cast(IsNull(Compra.Numero, '') as VarChar(1000))) from pago inner join cuentacorriente on "
						+ "pago.idcuentacorriente = cuentacorriente.idcuentacorriente "
						+ " Inner Join Compra on Pago.idCompra = Compra.idCompra "
						+ "Where idPagoLote = " 
						+ this.idPagoLote.Value.ToString();
					if (drSeteoF["AgrupaReferencia"] != DBNull.Value && (bool) drSeteoF["AgrupaReferencia"] == true)
						stCuentaGrupo += " group by pago.Referencia ";
					else
						stCuentaGrupo += " group by pago.idcuentacorriente ";
				
					SqlDataReader miGrupo = Funcion.rEscalarSQL(cdsCliente, stCuentaGrupo, true);
					while (miGrupo.Read())
					{
						drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
						drDetAsi["Debe"] = miGrupo.GetDouble(0);
						drDetAsi["idCuenta"] = miGrupo.GetInt32(1);
						drDetAsi["Detalle"] = miGrupo.GetString(2);
						drDetAsi["idAsiento"] = IdAsiento;
					}
					miGrupo.Close();
				}
			}
			else iCuenta1 = IdCuenta;

			#region Detalle en asiento por cada linea
			string strSelect = "";
			if (chkDesglosa.Checked) 
			{
				string stNombre = "";
				if (chkIncluyenombre.Checked)
					stNombre += ", Min(Cast(IsNull(Compra.Numero, '') as varchar(1000))), MIN(ISNULL(Cliente.Nombre, '')), MIN(IsNull(Compra.Numero, ''))";
        	strSelect = "SELECT SUM(Pago.Pago), " +
					"MIN(IsNull(ClienteGrupo.idCuenta, 0)) " + stNombre + 
					" FROM Pago " +
					"Inner Join Compra ON Pago.idCompra = Compra.idCompra " +
					"Inner Join Cliente ON Compra.idCliente = Cliente.idCliente " +
					"Full Join ClienteGrupo ON ClienteGrupo.idGrupoCliente = " +
					"Cliente.idGrupoCliente Where idPagoLote = " + 
					this.idPagoLote.Value.ToString() +
					" Group by Pago.idPago ";
			}
			else
			{
				strSelect = "SELECT SUM(Pago.Pago), " +
					"MIN(ClienteGrupo.idCuenta) FROM Pago " +
					"Inner Join Compra ON Pago.idCompra = Compra.idCompra " +
					"Inner Join Cliente ON Compra.idCliente = Cliente.idCliente " +
					"Full Join ClienteGrupo ON ClienteGrupo.idGrupoCliente = " +
					"Cliente.idGrupoCliente Where idPagoLote = " + 
					this.idPagoLote.Value.ToString() +
					" Group by idCuenta Order by idCuenta";
			}
			SqlDataReader miReader = 
				Funcion.rEscalarSQL(cdsCliente, strSelect, true);
			try
			{
				while (miReader.Read())
				{
					dValor = miReader.GetDouble(0);
					IdCuenta = 0;
					if (miReader.GetValue(1) != System.DBNull.Value)
						IdCuenta =  miReader.GetInt32(1);
					if (IdCuenta == 0) 
					{
						int iTipo = 3;
						if (IdTipoFactura != 1) iTipo = 12;
						drAsigna = cdsAsignaCuenta.TableViews["AsignaCuenta"].Rows[iTipo];
						IdCuenta = (int) drAsigna["idCuenta"]; // Cuentas por cobrar
					}

					drDetAsi = this.cdsAsiento.TableViews["AsientoDetalle"].AddNew();
					if (IdTipoFactura == 1) drDetAsi["Haber"] = dValor;
					else drDetAsi["Debe"] = dValor;
					drDetAsi["idCuenta"] = IdCuenta;
					drDetAsi["idAsiento"] = IdAsiento;
					if (chkDesglosa.Checked)
					{
						if (chkIncluyenombre.Checked)
						{
							string stDetalle = miReader.GetString(2).Trim()
								+ " " + miReader.GetString(3).Trim();
							drDetAsi["Detalle"] = stDetalle;
						}
					}
				}
			}
			catch(System.Exception ex) 
			{
				MessageBox.Show(ex.Message, "Detalle de asiento");
			}
			miReader.Close();
		
			#endregion // Fin de Detalle de asiento por pagos no detallados

			if (IdTipoFactura != 1) 
			{
				if (!chkDesglosa.Checked) 
				{
					drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
					drDetAsi["Haber"] = dTotal;
					drDetAsi["idCuenta"] = iCuenta1;
					drDetAsi["idAsiento"] = IdAsiento;
				}
				else
				{
					string stCuentaGrupo = "select sum(IsNull(pago.pago, 0)), Max(cuentacorriente.idCuenta) "
						+ "from pago inner join cuentacorriente on "
						+ "pago.idcuentacorriente = cuentacorriente.idcuentacorriente "
						+ "Where idPagoLote = " 
						+ this.idPagoLote.Value.ToString();
					if (drSeteoF["AgrupaReferencia"] != DBNull.Value && (bool) drSeteoF["AgrupaReferencia"] == true)
						stCuentaGrupo += " group by pago.Referencia ";
					else
						stCuentaGrupo += " group by pago.idcuentacorriente ";
					SqlDataReader miGrupo = Funcion.rEscalarSQL(cdsCliente, stCuentaGrupo, true);
					while (miGrupo.Read())
					{
						drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
						drDetAsi["Haber"] = miGrupo.GetDouble(0);
						drDetAsi["idCuenta"] = miGrupo.GetInt32(1);
						drDetAsi["idAsiento"] = IdAsiento;
					}
					miGrupo.Close();
				}
			}
			this.cdsPagoLote.Schema.Connections[0].Close();
			this.cdsAsiento.Update();
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			Cursor = Cursors.WaitCursor;
			string stFiltro = "";
			if (cmbCliente.Text.Trim().Length > 0 && (int) this.cmbCliente.Value > 0)
			{
				if (chkReceptor.Checked)
					stFiltro = "idRecibe = ";
				else
					stFiltro = "idCliente = ";

				stFiltro += this.cmbCliente.Value.ToString() + " And ";
			}
			else
			{
				if (DialogResult.No == MessageBox.Show("¿Desea Generar Listado sin cliente?", 
					"Confirmación",	MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			}
			stFiltro += " Abs(Saldo) >= 0.005 And Borrar = 0";
			if (IdTipoFactura == 1)
				stFiltro += " And Abs(saldo) > 0 And (idTipoFactura = 1 OR idTipoFactura = 24 Or "
					+ "idTipoFactura = 11 Or idTipoFactura = 27 Or idTipoFactura = 99)";
			else
				stFiltro += " And (idTipoFactura = 4 OR idTipoFactura = 25 Or "
          + "idTipoFactura = 26 Or idTipoFactura = 28)";
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsCompra, "Compra", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

				C1.Data.FilterCondition fcHija = new 
					C1.Data.FilterCondition(this.cdsCompra, "DetCompra", "idCompra = 0");
				fcTotal.Add(fcHija);
			
			this.cdsCompra.Fill(fcTotal, false);
			Cursor = Cursors.Default;
//			MessageBox.Show("Finalizado", "Información");
		}

		private void AsiendoDeRetencion()
		{
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Crear Asientos", "Seguridad", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea crear Asiento de Retención?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;

			Cursor = Cursors.WaitCursor;

			string stExec = "EXEC AsientoRetencion 0, "
				+ idPagoLote.Value.ToString() + ", " + IdTipoFactura.ToString();
			Funcion.EjecutaSQL(cdsAsiento, stExec, true);
			stExec = "Select idAsiRetL FROM PagoLote WHERE idPagoLote = " + idPagoLote.Value.ToString();
			int IdAsiRetL = Funcion.iEscalarSQL(cdsCliente, stExec);
			this.txtIdAsiRetL.Value = IdAsiRetL;
			Cursor = Cursors.Default;
			MessageBox.Show("Asiento de Retención Generado", "Información");
		}

		private void txtTotal_DoubleClick(object sender, System.EventArgs e)
		{
			double dPago = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if (dr.Cells["Pago"].Value != System.DBNull.Value)
					dPago += (double) dr.Cells["Pago"].Value;
			}
			txtTotal.Value = dPago;	
		}

		private void cmbCliente_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				string stPrimero = "Nombre";
				if (cmbCliente.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition == 0)
					stPrimero = "Codigo";
				cdvCliente.Sort = stPrimero;
				cmbCliente.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition = 0;
				cmbCliente.DisplayMember = stPrimero;
			}
		}

		private void btPago_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid2.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Pago que desea ver el detalle", "Información");
				return;
			}
			Cursor = Cursors.WaitCursor;
			int IdPago = (int) ultraGrid2.ActiveRow.Cells["idCompra"].Value;

//			Pagos miPago = new Pagos(IdTipoFactura, IdPago);
//			miPago.MdiParent = this.MdiParent;
//			miPago.Show();
			Cursor = Cursors.Default;
		}

		private void btCash_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Pago.idPagoLote} = " + idPagoLote.Value.ToString();
			string stRep = "AbonoCash.Rpt";
			if (IdTipoFactura == 1) stRep = "CobroCash.Rpt";
			Reporte miRep = new Reporte(stRep, stFiltro);
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void btPago_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (barraDato1.bEditar == true)
			{
				MessageBox.Show("Debe tener el archivo grabado para calcular el interes", "Información");
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea Generar el Interes de los pagos?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
      string stExec = "Exec PagoLoteInteres " + idPagoLote.Value;
			double dTotal = Funcion.dEscalarSQL(cdsAsiento, stExec, true);
			MessageBox.Show("Pago Total: " + dTotal.ToString("0.00"));
		}

		private void btFacturaD_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid2.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la Factura de la grilla derecha que desea ver el detalle", "Información");
				return;
			}
			Cursor = Cursors.WaitCursor;
			int IdFact = (int) ultraGrid2.ActiveRow.Cells["idCompra"].Value;
//			Compra miCompra = new Compra(IdTipoFactura, IdFact);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void btDeposito_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Pago.idPagoLote} = " + idPagoLote.Value.ToString();
			string stRep = "Deposito.Rpt";
			if (IdTipoFactura == 1) stRep = "ReciboIngreso.Rpt";
			Reporte miRep = new Reporte(stRep, stFiltro);
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void cmbCliente_ValueChanged(object sender, System.EventArgs e)
		{
			if(barraDato1.bEditar)
			{
			
				string stFiltro = " idcompra = 0";
				C1.Data.FilterCondition fcPadre = new 
					C1.Data.FilterCondition(this.cdsCompra, "Compra", stFiltro);
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcPadre);

				C1.Data.FilterCondition fcHija = new 
					C1.Data.FilterCondition(this.cdsCompra, "DetCompra", "idCompra = 0");
				fcTotal.Add(fcHija);
			
				this.cdsCompra.Fill(fcTotal, false);
			}
		}

	}
}
