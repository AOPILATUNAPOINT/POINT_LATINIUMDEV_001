using System;
using System.Globalization;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions.Shared;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de PagoLotes.
	/// </summary>
	public class PagosAutoriza : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCompra;
		private System.ComponentModel.IContainer components;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private System.Windows.Forms.ImageList imageList;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private C1.Data.C1DataSet cdsSeteoF;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbVencimiento;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor ultraComboEditor1;
		private C1.Data.C1DataSet cdsPagoAutoriza;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor idPagoLote;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDeseado;
		private Infragistics.Win.Misc.UltraButton btAgregar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoCliente;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmdClientes;
		private Infragistics.Win.Misc.UltraButton btEstimado;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFechaVence;
		private Infragistics.Win.Misc.UltraButton btAgregarTodo;

		public PagosAutoriza()
		{
			InitializeComponent();
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
			this.components = new System.ComponentModel.Container();
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Compra", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoLot");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoPer");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente", -1, "cmdClientes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPorcIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenProd");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpleado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoMoneda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRecibe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContadoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorDolar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorInteres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRevision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Revisado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Otro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalExcento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DerechoDevolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Refrendo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flete");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagosCada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasInicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutImprenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcIva0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Prioridad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutPrincipal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutPresupuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FletePorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BienServicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConvenioInternac");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vuelto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCaducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaServicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaRetDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprob1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depart");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pedido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bulto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoAutorizado");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoAut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloquear");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cif");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CifIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Propina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Servicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutoImpSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNiffIngProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NiifValorIng");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IntImplicito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TasaCambio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoDif");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("subTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComisionConfiteca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasAux");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Garantia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipoObra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nivePrioridadRuta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sistemaMonitoreado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sistemaSeguridad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoNIIF");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FactAfectaCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_AnticipoV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_MultasV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FondoGV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_AnticipoP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_MultasP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FondoGP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_AnticipoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_MultasT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FondoGT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngAfectaCostoNiif");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jmc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FaltaRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FaltaOriginal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IRBP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_IdOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Procesado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrimerPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotaEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotaAsignada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotaUsada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotaDisponible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPrimerPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaTentativaLlegada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Manual");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 38, true, "Compra", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 38, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 26, true, "Compra", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 26, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PagoAutorizado", 99, true, "Compra", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PagoAutorizado", 99, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(PagosAutoriza));
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_PagoAutDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoAutDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoAutoriza");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pago");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Pago", 4, true, "_PagoAutDetalle", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Pago", 4, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn192 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn193 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn194 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn195 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn196 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn197 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn198 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn199 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn200 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn201 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn202 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn203 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn204 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn205 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn206 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn207 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn208 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn209 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn210 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn211 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ClteCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn212 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn213 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn214 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointAutorizado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn215 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn216 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn217 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucionFinanaciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn218 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idObservaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn219 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn220 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HaceRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn221 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaConsignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn222 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn223 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn224 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito2");
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsPagoAutoriza = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.chkTodoCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbVencimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraComboEditor1 = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.idPagoLote = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDeseado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btAgregar = new Infragistics.Win.Misc.UltraButton();
			this.btAgregarTodo = new Infragistics.Win.Misc.UltraButton();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmdClientes = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btEstimado = new Infragistics.Win.Misc.UltraButton();
			this.chkFechaVence = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagoAutoriza)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVencimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraComboEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.idPagoLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeseado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdClientes)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoAutoriza, "_PagoAutoriza.idCliente"));
			this.cmbCliente.DataSource = this.cdvCliente;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Width = 137;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 322;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn24.Header.VisiblePosition = 23;
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
																										 ultraGridColumn24});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.Enabled = false;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(128, 9);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(272, 21);
			this.cmbCliente.TabIndex = 1;
			this.cmbCliente.ValueMember = "idCliente";
			this.cmbCliente.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCliente_MouseUp);
			// 
			// cdsPagoAutoriza
			// 
			this.cdsPagoAutoriza.BindingContextCtrl = this;
			this.cdsPagoAutoriza.DataLibrary = "LibFacturacion";
			this.cdsPagoAutoriza.DataLibraryUrl = "";
			this.cdsPagoAutoriza.DataSetDef = "dsPagoAutoriza";
			this.cdsPagoAutoriza.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPagoAutoriza.Name = "cdsPagoAutoriza";
			this.cdsPagoAutoriza.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPagoAutoriza.SchemaDef = null;
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
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Compra";
			this.ultraGrid1.DataSource = this.cdsCompra;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn25.Header.VisiblePosition = 40;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 96;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 85;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 88;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.Caption = "Proveedor";
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn30.Header.VisiblePosition = 38;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 87;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.Caption = "Centro_Costo";
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn33.Header.VisiblePosition = 10;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 77;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 44;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 43;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 45;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 39;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 47;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 46;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 41;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 42;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 83;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 79;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 2;
			ultraGridColumn45.Width = 75;
			ultraGridColumn46.Header.VisiblePosition = 32;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 3;
			ultraGridColumn48.Header.VisiblePosition = 70;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 4;
			ultraGridColumn49.Width = 70;
			ultraGridColumn50.Header.VisiblePosition = 55;
			ultraGridColumn50.Hidden = true;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance2;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.VisiblePosition = 6;
			ultraGridColumn51.Width = 63;
			ultraGridColumn52.Header.VisiblePosition = 22;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 23;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 24;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 25;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 28;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 76;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 50;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 49;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 29;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 17;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 19;
			ultraGridColumn62.Hidden = true;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance3;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.VisiblePosition = 5;
			ultraGridColumn63.Width = 73;
			ultraGridColumn64.Header.VisiblePosition = 59;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 57;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 18;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 58;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 71;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 69;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 72;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 62;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 73;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 34;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 75;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 74;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 33;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 60;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 68;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 67;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 65;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 21;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 37;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 64;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 16;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 14;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 35;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 61;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 30;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 11;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 9;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 66;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 27;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 26;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 20;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 63;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 13;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 12;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 51;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 52;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 53;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 54;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 36;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 48;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 56;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 31;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 15;
			ultraGridColumn106.Hidden = true;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn107.CellAppearance = appearance4;
			ultraGridColumn107.Format = "#,##0.00";
			ultraGridColumn107.Header.VisiblePosition = 8;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn107.Width = 67;
			ultraGridColumn108.Header.VisiblePosition = 78;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 80;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 81;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 82;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 84;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 86;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 89;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 90;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 91;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 95;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn118.Header.VisiblePosition = 94;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 92;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 93;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 97;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 98;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn123.Header.VisiblePosition = 99;
			ultraGridColumn123.Hidden = true;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn124.CellAppearance = appearance5;
			ultraGridColumn124.Format = "#,##0.00";
			ultraGridColumn124.Header.Caption = "Autorizado";
			ultraGridColumn124.Header.VisiblePosition = 7;
			ultraGridColumn124.Width = 80;
			ultraGridColumn125.Header.VisiblePosition = 100;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 101;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn127.Header.VisiblePosition = 102;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn128.Header.VisiblePosition = 103;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn129.Header.VisiblePosition = 104;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn130.Header.VisiblePosition = 105;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn131.Header.VisiblePosition = 106;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn132.Header.VisiblePosition = 107;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn133.Header.VisiblePosition = 108;
			ultraGridColumn134.Header.VisiblePosition = 109;
			ultraGridColumn135.Header.VisiblePosition = 110;
			ultraGridColumn136.Header.VisiblePosition = 111;
			ultraGridColumn137.Header.VisiblePosition = 112;
			ultraGridColumn138.Header.VisiblePosition = 113;
			ultraGridColumn139.Header.VisiblePosition = 114;
			ultraGridColumn140.Header.VisiblePosition = 115;
			ultraGridColumn141.Header.VisiblePosition = 116;
			ultraGridColumn142.Header.VisiblePosition = 117;
			ultraGridColumn143.Header.VisiblePosition = 118;
			ultraGridColumn144.Header.VisiblePosition = 119;
			ultraGridColumn145.Header.VisiblePosition = 120;
			ultraGridColumn146.Header.VisiblePosition = 121;
			ultraGridColumn147.Header.VisiblePosition = 122;
			ultraGridColumn148.Header.VisiblePosition = 123;
			ultraGridColumn149.Header.VisiblePosition = 124;
			ultraGridColumn150.Header.VisiblePosition = 125;
			ultraGridColumn151.Header.VisiblePosition = 126;
			ultraGridColumn152.Header.VisiblePosition = 127;
			ultraGridColumn153.Header.VisiblePosition = 128;
			ultraGridColumn154.Header.VisiblePosition = 129;
			ultraGridColumn155.Header.VisiblePosition = 130;
			ultraGridColumn156.Header.VisiblePosition = 131;
			ultraGridColumn157.Header.VisiblePosition = 132;
			ultraGridColumn158.Header.VisiblePosition = 133;
			ultraGridColumn159.Header.VisiblePosition = 134;
			ultraGridColumn160.Header.VisiblePosition = 135;
			ultraGridColumn161.Header.VisiblePosition = 136;
			ultraGridColumn162.Header.VisiblePosition = 137;
			ultraGridColumn163.Header.VisiblePosition = 138;
			ultraGridColumn164.Header.VisiblePosition = 139;
			ultraGridColumn165.Header.VisiblePosition = 140;
			ultraGridColumn166.Header.VisiblePosition = 141;
			ultraGridColumn167.Header.VisiblePosition = 142;
			ultraGridColumn168.Header.VisiblePosition = 143;
			ultraGridColumn169.Header.VisiblePosition = 144;
			ultraGridColumn170.Header.VisiblePosition = 145;
			ultraGridColumn171.Header.VisiblePosition = 146;
			ultraGridColumn172.Header.VisiblePosition = 147;
			ultraGridColumn173.Header.VisiblePosition = 148;
			ultraGridColumn174.Header.VisiblePosition = 149;
			ultraGridColumn175.Header.VisiblePosition = 150;
			ultraGridColumn176.Header.VisiblePosition = 151;
			ultraGridColumn177.Header.VisiblePosition = 152;
			ultraGridColumn178.Header.VisiblePosition = 153;
			ultraGridColumn179.Header.VisiblePosition = 154;
			ultraGridColumn180.Header.VisiblePosition = 155;
			ultraGridColumn181.Header.VisiblePosition = 156;
			ultraGridColumn182.Header.VisiblePosition = 157;
			ultraGridColumn183.Header.VisiblePosition = 158;
			ultraGridColumn184.Header.VisiblePosition = 159;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn147,
																										 ultraGridColumn148,
																										 ultraGridColumn149,
																										 ultraGridColumn150,
																										 ultraGridColumn151,
																										 ultraGridColumn152,
																										 ultraGridColumn153,
																										 ultraGridColumn154,
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
																										 ultraGridColumn184});
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance6;
			summarySettings1.DisplayFormat = "{0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance7;
			summarySettings2.DisplayFormat = "{0}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance8;
			summarySettings3.DisplayFormat = "{0}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand2.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 78);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(432, 241);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
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
			// imageList
			// 
			this.imageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth16Bit;
			this.imageList.ImageSize = new System.Drawing.Size(18, 18);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btFiltro
			// 
			appearance13.Image = 5;
			appearance13.ImageHAlign = Infragistics.Win.HAlign.Right;
			this.btFiltro.Appearance = appearance13;
			this.btFiltro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFiltro.ImageList = this.imageList;
			this.btFiltro.Location = new System.Drawing.Point(408, 9);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(80, 25);
			this.btFiltro.TabIndex = 7;
			this.btFiltro.Text = "Filtrar";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// txtFecha
			// 
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Format = "dd/MMM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(232, 43);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(120, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 9;
			this.txtFecha.Value = new System.DateTime(2009, 8, 31, 0, 0, 0, 0);
			this.txtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.txtFecha_BeforeDropDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(152, 43);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(72, 16);
			this.label2.TabIndex = 10;
			this.label2.Text = "Fecha Desde:";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
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
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_PagoAutoriza";
			this.barraDato1.DataNombreId = "idPagoAutoriza";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsPagoAutoriza;
			this.barraDato1.DataTabla = "PagoAutoriza";
			this.barraDato1.DataTablaHija = "PagoAutDetalle";
			this.barraDato1.Location = new System.Drawing.Point(16, 327);
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
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
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
			this.btFactura.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFactura.Location = new System.Drawing.Point(320, 327);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(24, 25);
			this.btFactura.TabIndex = 35;
			this.btFactura.Text = "F";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// chkTodoCliente
			// 
			this.chkTodoCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkTodoCliente.Checked = true;
			this.chkTodoCliente.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodoCliente.Location = new System.Drawing.Point(16, 9);
			this.chkTodoCliente.Name = "chkTodoCliente";
			this.chkTodoCliente.Size = new System.Drawing.Size(104, 21);
			this.chkTodoCliente.TabIndex = 44;
			this.chkTodoCliente.Text = "Todo Proveedor";
			this.chkTodoCliente.CheckedChanged += new System.EventHandler(this.chkPendiente_CheckedChanged);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(368, 43);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(65, 16);
			this.label11.TabIndex = 46;
			this.label11.Text = "Fecha Hasta";
			// 
			// cmbVencimiento
			// 
			this.cmbVencimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbVencimiento.DateButtons.Add(dateButton2);
			this.cmbVencimiento.Format = "dd/MMM/yyyy";
			this.cmbVencimiento.Location = new System.Drawing.Point(448, 43);
			this.cmbVencimiento.Name = "cmbVencimiento";
			this.cmbVencimiento.NonAutoSizeHeight = 23;
			this.cmbVencimiento.Size = new System.Drawing.Size(120, 21);
			this.cmbVencimiento.SpinButtonsVisible = true;
			this.cmbVencimiento.TabIndex = 45;
			this.cmbVencimiento.Value = new System.DateTime(2009, 9, 1, 0, 0, 0, 0);
			// 
			// ultraComboEditor1
			// 
			this.ultraComboEditor1.AutoComplete = true;
			this.ultraComboEditor1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Emision";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Vencimiento";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Entrega";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Toda Fecha";
			this.ultraComboEditor1.Items.Add(valueListItem1);
			this.ultraComboEditor1.Items.Add(valueListItem2);
			this.ultraComboEditor1.Items.Add(valueListItem3);
			this.ultraComboEditor1.Items.Add(valueListItem4);
			this.ultraComboEditor1.Location = new System.Drawing.Point(24, 43);
			this.ultraComboEditor1.Name = "ultraComboEditor1";
			this.ultraComboEditor1.Size = new System.Drawing.Size(120, 22);
			this.ultraComboEditor1.TabIndex = 47;
			this.ultraComboEditor1.Text = "Emision";
			this.ultraComboEditor1.Visible = false;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "_PagoAutoriza.PagoAutoriza - PagoAutDetalle";
			this.ultraGrid2.DataSource = this.cdsPagoAutoriza;
			appearance14.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance14;
			ultraGridColumn185.Header.VisiblePosition = 2;
			ultraGridColumn185.Hidden = true;
			ultraGridColumn186.Header.VisiblePosition = 3;
			ultraGridColumn186.Hidden = true;
			ultraGridColumn187.Header.VisiblePosition = 4;
			ultraGridColumn187.Hidden = true;
			ultraGridColumn188.Header.VisiblePosition = 1;
			ultraGridColumn188.Width = 91;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn189.CellAppearance = appearance15;
			ultraGridColumn189.Format = "#,##0.00";
			ultraGridColumn189.Header.VisiblePosition = 0;
			ultraGridColumn190.Header.VisiblePosition = 5;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn185,
																										 ultraGridColumn186,
																										 ultraGridColumn187,
																										 ultraGridColumn188,
																										 ultraGridColumn189,
																										 ultraGridColumn190});
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance16;
			summarySettings4.DisplayFormat = "{0}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4});
			ultraGridBand3.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 10F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.ultraGrid2.Location = new System.Drawing.Point(488, 78);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(304, 241);
			this.ultraGrid2.TabIndex = 48;
			this.ultraGrid2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid2_KeyDown_1);
			// 
			// idPagoLote
			// 
			this.idPagoLote.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoAutoriza, "_PagoAutoriza.idPagoAutoriza"));
			this.idPagoLote.Location = new System.Drawing.Point(496, 9);
			this.idPagoLote.Name = "idPagoLote";
			this.idPagoLote.PromptChar = ' ';
			this.idPagoLote.Size = new System.Drawing.Size(72, 22);
			this.idPagoLote.TabIndex = 49;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoAutoriza, "_PagoAutoriza.Total"));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(504, 327);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(88, 22);
			this.txtTotal.TabIndex = 50;
			this.txtTotal.DoubleClick += new System.EventHandler(this.txtTotal_DoubleClick);
			this.txtTotal.ValueChanged += new System.EventHandler(this.txtTotal_ValueChanged);
			// 
			// txtDeseado
			// 
			this.txtDeseado.FormatString = "#,##0.00";
			this.txtDeseado.Location = new System.Drawing.Point(592, 327);
			this.txtDeseado.Name = "txtDeseado";
			this.txtDeseado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDeseado.PromptChar = ' ';
			this.txtDeseado.Size = new System.Drawing.Size(88, 22);
			this.txtDeseado.TabIndex = 51;
			// 
			// btAgregar
			// 
			this.btAgregar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btAgregar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAgregar.Location = new System.Drawing.Point(448, 112);
			this.btAgregar.Name = "btAgregar";
			this.btAgregar.Size = new System.Drawing.Size(32, 25);
			this.btAgregar.TabIndex = 52;
			this.btAgregar.Text = ">";
			this.btAgregar.Click += new System.EventHandler(this.btAgregar_Click);
			// 
			// btAgregarTodo
			// 
			this.btAgregarTodo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btAgregarTodo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAgregarTodo.Location = new System.Drawing.Point(448, 155);
			this.btAgregarTodo.Name = "btAgregarTodo";
			this.btAgregarTodo.Size = new System.Drawing.Size(32, 25);
			this.btAgregarTodo.TabIndex = 53;
			this.btAgregarTodo.Text = ">>";
			this.btAgregarTodo.Visible = false;
			this.btAgregarTodo.Click += new System.EventHandler(this.btAgregarTodo_Click_1);
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible =1";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataSource = this.cdvProyecto;
			ultraGridColumn191.Header.VisiblePosition = 2;
			ultraGridColumn191.Hidden = true;
			ultraGridColumn192.Header.VisiblePosition = 6;
			ultraGridColumn192.Hidden = true;
			ultraGridColumn193.Header.VisiblePosition = 0;
			ultraGridColumn193.Width = 144;
			ultraGridColumn194.Header.VisiblePosition = 3;
			ultraGridColumn194.Hidden = true;
			ultraGridColumn195.Header.VisiblePosition = 7;
			ultraGridColumn195.Hidden = true;
			ultraGridColumn196.Header.VisiblePosition = 4;
			ultraGridColumn196.Hidden = true;
			ultraGridColumn197.Header.VisiblePosition = 1;
			ultraGridColumn198.Header.VisiblePosition = 5;
			ultraGridColumn198.Hidden = true;
			ultraGridColumn199.Header.VisiblePosition = 8;
			ultraGridColumn199.Hidden = true;
			ultraGridColumn200.Header.VisiblePosition = 9;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn191,
																										 ultraGridColumn192,
																										 ultraGridColumn193,
																										 ultraGridColumn194,
																										 ultraGridColumn195,
																										 ultraGridColumn196,
																										 ultraGridColumn197,
																										 ultraGridColumn198,
																										 ultraGridColumn199,
																										 ultraGridColumn200});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(16, 103);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(268, 87);
			this.cmbProyecto.TabIndex = 117;
			this.cmbProyecto.ValueMember = "idProyecto";
			this.cmbProyecto.Visible = false;
			// 
			// cmdClientes
			// 
			this.cmdClientes.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdClientes.DataSource = this.cdvCliente;
			ultraGridBand5.AddButtonCaption = "Proyecto";
			ultraGridColumn201.Header.VisiblePosition = 0;
			ultraGridColumn202.Header.VisiblePosition = 2;
			ultraGridColumn203.Header.VisiblePosition = 1;
			ultraGridColumn203.Width = 144;
			ultraGridColumn204.Header.VisiblePosition = 3;
			ultraGridColumn205.Header.VisiblePosition = 4;
			ultraGridColumn206.Header.VisiblePosition = 5;
			ultraGridColumn207.Header.VisiblePosition = 6;
			ultraGridColumn208.Header.VisiblePosition = 7;
			ultraGridColumn209.Header.VisiblePosition = 8;
			ultraGridColumn210.Header.VisiblePosition = 9;
			ultraGridColumn211.Header.VisiblePosition = 10;
			ultraGridColumn212.Header.VisiblePosition = 11;
			ultraGridColumn213.Header.VisiblePosition = 12;
			ultraGridColumn214.Header.VisiblePosition = 13;
			ultraGridColumn215.Header.VisiblePosition = 14;
			ultraGridColumn216.Header.VisiblePosition = 15;
			ultraGridColumn217.Header.VisiblePosition = 16;
			ultraGridColumn218.Header.VisiblePosition = 17;
			ultraGridColumn219.Header.VisiblePosition = 18;
			ultraGridColumn220.Header.VisiblePosition = 19;
			ultraGridColumn221.Header.VisiblePosition = 20;
			ultraGridColumn222.Header.VisiblePosition = 21;
			ultraGridColumn223.Header.VisiblePosition = 22;
			ultraGridColumn224.Header.VisiblePosition = 23;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn224});
			this.cmdClientes.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmdClientes.DisplayMember = "Nombre";
			this.cmdClientes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdClientes.Location = new System.Drawing.Point(112, 121);
			this.cmdClientes.Name = "cmdClientes";
			this.cmdClientes.Size = new System.Drawing.Size(268, 86);
			this.cmdClientes.TabIndex = 120;
			this.cmdClientes.ValueMember = "idCliente";
			this.cmdClientes.Visible = false;
			// 
			// btEstimado
			// 
			this.btEstimado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btEstimado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btEstimado.Location = new System.Drawing.Point(352, 327);
			this.btEstimado.Name = "btEstimado";
			this.btEstimado.Size = new System.Drawing.Size(24, 25);
			this.btEstimado.TabIndex = 121;
			this.btEstimado.Text = "P";
			this.btEstimado.Click += new System.EventHandler(this.btEstimado_Click);
			// 
			// chkFechaVence
			// 
			this.chkFechaVence.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkFechaVence.Checked = true;
			this.chkFechaVence.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkFechaVence.Location = new System.Drawing.Point(16, 43);
			this.chkFechaVence.Name = "chkFechaVence";
			this.chkFechaVence.Size = new System.Drawing.Size(128, 22);
			this.chkFechaVence.TabIndex = 122;
			this.chkFechaVence.Text = "Fecha Vencimiento";
			// 
			// PagosAutoriza
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(808, 368);
			this.Controls.Add(this.chkFechaVence);
			this.Controls.Add(this.btEstimado);
			this.Controls.Add(this.cmdClientes);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.btAgregarTodo);
			this.Controls.Add(this.btAgregar);
			this.Controls.Add(this.txtDeseado);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.idPagoLote);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.ultraComboEditor1);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbVencimiento);
			this.Controls.Add(this.chkTodoCliente);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "PagosAutoriza";
			this.Text = "Autorización de Pagos";
			this.Load += new System.EventHandler(this.PagoLotes_Load);
			this.Closed += new System.EventHandler(this.PagosAutoriza_Closed);
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagoAutoriza)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVencimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraComboEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.idPagoLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeseado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdClientes)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void BorraErrores()
		{
			this.errorProvider.SetError(this.txtFecha, "");
			this.errorProvider.SetError(this.ultraGrid1, "");
			this.errorProvider.SetError(this.ultraGrid2, "");
			this.errorProvider.SetError(this.cmbCliente, "");
		}

		private bool VerificaAgregar()
		{
			bool bOk = true;
			this.BorraErrores();
//			if (this.txtFecha.Value == System.DBNull.Value)
//			{
//				this.errorProvider.SetError(this.txtFecha, "Ingrese Fecha");
//				bOk = false;
//			}
			if (ultraGrid1.ActiveRow == null || ultraGrid1.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				this.errorProvider.SetError(this.ultraGrid1, "Seleccione factura que desea realizar el pago");
				bOk = false;
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

		private void btAgregar_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
			{
				if (!VerificaAgregar()) return;
				txtTotal_DoubleClick(this, e);
				double dPagoActual = (double) dr.Cells["Saldo"].Value;
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
				C1.Data.C1DataRow drPago = this.cdsPagoAutoriza.TableViews["PagoAutDetalle"].AddNew();
				drPago.CancelEdit();
				this.ultraGrid2.Refresh();

				drPago = cdsPagoAutoriza.TableViews["PagoAutDetalle"].AddNew();
				drPago["idPagoAutoriza"] = idPagoLote.Value;
				drPago["idCompra"] = (int) dr.Cells["idCompra"].Value;
				drPago["Pago"] = dPagoActual;
				drPago["Nombre"] = dr.Cells["idcliente"].Text;
				drPago["Numero"] = dr.Cells["numero"].Value.ToString();
				ultraGrid2.Refresh();
				txtTotal_DoubleClick(this, e);
			}
		}

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
			string stSelect = "Delete Pago Where idPagoLote = " 
				+ IdPagoLote.ToString();
			Funcion.EjecutaSQL(cdsCompra, stSelect);
			barraDato1.BorraRegistro();
			MessageBox.Show("Registro Borrado",
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
			txtTotal_DoubleClick(this, e);

		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using(BuscaPagoLotes miPagoLote = new BuscaPagoLotes())
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


		private void Habilita(bool bActiva)
		{
			this.btAgregar.Enabled = bActiva;
			this.btAgregarTodo.Enabled = bActiva;
//			this.btQuitar.Enabled = bActiva;
//			this.btQuitaTodo.Enabled = bActiva;
		}
		private void ValoresIniciales()
		{
//			drPagoLote = cdsPagoAutoriza.TableViews["PagoLote"].Rows.Find
//			(int.Parse(idPagoLote.Value.ToString()));
//			this.drPagoLote["Fecha"] = DateTime.Today;
////			this.drPagoLote["idTipoFactura"] = IdTipoFactura;
			this.Habilita(true);
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene Acceso a Crear Pagos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			this.barraDato1.CrearRegistro();
			ValoresIniciales();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
//			if (barraDato1.bNuevo)
//			{
//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
//				{
//					if (dr.Cells["idPago"].Value != System.DBNull.Value)
//					{
//						int IdPago = (int) dr.Cells["idPago"].Value;
//						this.cdsPagoLote.TableViews["Pago"].Rows.Find(IdPago).Delete();
//					}
//				}
//				int IdPagoLote = int.Parse(this.idPagoLote.Value.ToString());
//				string stSelect = "Delete Pago Where idPagoLote = " 
//					+ IdPagoLote.ToString();
//				Funcion.EjecutaSQL(cdsCompra, stSelect);
//			}
			this.cdsPagoAutoriza.RejectChanges();
			this.barraDato1.DeshacerRegistro();
			this.Habilita(false);
			this.BorraErrores();
			this.ultraGrid1.Enabled = true;
		}

//		C1.Data.C1DataRow drPagoLote;
		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene Acceso a Editar Pagos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.barraDato1.EditarRegistro();
//			drPagoLote = cdsPagoLote.TableViews["PagoLote"].Rows.Find(int.Parse(idPagoLote.Value.ToString()));
			this.Habilita(true);
		}

		private bool Verifica()
		{
			bool bOk = true;
			this.BorraErrores();
			if (this.txtFecha.Value == System.DBNull.Value)
			{
				this.errorProvider.SetError(this.txtFecha, "Ingrese Fecha");
				bOk = false;
			}
			return bOk;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			double dTotal = 0;
			if (txtTotal.Value != DBNull.Value) dTotal = double.Parse(txtTotal.Value.ToString());
			if (dTotal > .001)
			{
			//	txtTotal_DoubleClick(this, e);

				//if (!Verifica()) return;
			}
			DateTime cmbVencimiento = DateTime.Today;
			DateTime txtFecha = DateTime.Today;
			this.barraDato1.GrabaRegistro();
			string stExec = "Exec PagoAutorizaCompra " + idPagoLote.Value.ToString();
			Funcion.EjecutaSQL(cdsCliente, stExec, true);
			this.Habilita(false);
			this.ultraGrid1.Enabled = true;
			this.ultraGrid2.Enabled = true;
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{PagoAutDetalle.idPagoAutoriza} = " + idPagoLote.Value.ToString();
			string stRep = "PagosAutorizados.Rpt";
			Reporte miRep = new Reporte(stRep, stFiltro);
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
      		
		}

		C1.Data.C1DataRow drSeteoF;
		Acceso miAcceso;
		private void PagoLotes_Load(object sender, System.EventArgs e)
		{
			txtFecha.Value= DateTime.Today;
			cmbVencimiento.Value=DateTime.Today;
			//txtTotal.Width = 0;
			txtDeseado.Width=0;
			string stAudita = "Exec AuditaCrear 53, 6, 'PagoLAut'";
			Funcion.EjecutaSQL(cdsPagoAutoriza, stAudita, true);
			miAcceso =new Acceso("04090LA", this.cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
		
			string stFiltro="Proveedor <> 0 Or Ambos <> 0";

			C1.Data.FilterCondition fcCliente = new 
				C1.Data.FilterCondition(cdsCliente, "Cliente", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcCliente);
			cdsCliente.Fill(fcTotal, false);

			this.barraDato1.PosUltima();
			this.barraDato1.HabilitaControles(false);
			this.idPagoLote.Width = 0;
			drSeteoF = this.cdsSeteoF.TableViews["SeteoF"].Rows[0];
			this.ultraGrid1.Enabled = true;
			this.ultraGrid2.Enabled = true;
			this.cmbCliente.Enabled = false;
			Cursor = Cursors.Default;
//			if (Funcion.bEjecutaSQL(cdsCliente, "Exec SeteoGExiste 'FLAP'")) 				
//			{				
//				cambioNicanor();
//			}

		}

		CultureInfo us = new CultureInfo("en-US");

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			Cursor = Cursors.WaitCursor;
			string stFiltro = "";
				DateTime dtFechaDesde = (DateTime) this.txtFecha.Value;
				DateTime dtFechaHasta = (DateTime) this.cmbVencimiento.Value;

			if (cmbCliente.Text.Trim().Length > 0 && (int) this.cmbCliente.Value > 0)
			{
				stFiltro = "idCliente = " + this.cmbCliente.Value.ToString() + " And ";
			}
			else
			{
				if (DialogResult.No == MessageBox.Show("¿Desea Generar Listado de Todos los Proveedores?", 
					"Confirmación",	MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			}
			if (!chkTodoCliente.Checked) 
			{
				stFiltro += " Fecha >=  '" + dtFechaDesde.ToString("yyyyMMdd") + "' And " 
					+ " Fecha <= '"+ dtFechaHasta.ToString("yyyyMMdd") + "' And ";
			}
			else
			{
				stFiltro += " FechaVencimiento >=  '" + dtFechaDesde.ToString("yyyyMMdd") + "' And " 
					+ " FechaVencimiento <= '"+ dtFechaHasta.ToString("yyyyMMdd") + "' And ";
			}
			stFiltro += " Abs(Saldo) >= 0.005 And Borrar = 0";
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
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsClientes"];

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"],
				"(Ambos <> 0 OR Proveedor <> 0) And DesconProv = 0"));
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void btAgregarTodo_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Factura que desea ver el detalle", "Información");
				return;
			}
			Cursor = Cursors.WaitCursor;
			int IdFact = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
//			Compra miCompra = new Compra(4, IdFact);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void ultraGrid2_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.IsInEditMode)
				if (e.Cell.Column.ToString() == "Pago")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						if (drSeteoF["PagoNegativo"] == System.DBNull.Value ||
							(bool) drSeteoF["PagoNegativo"] == false)
						{
							string stSelect = "Select Saldo From Compra Where idCompra = " 
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

		private void chkPendiente_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCliente.Enabled = !chkTodoCliente.Checked;
			if (!chkTodoCliente.Checked) cmbCliente.Select();
		}

		private void PagosAutoriza_Closed(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid2_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
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

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btEstimado_Click(object sender, System.EventArgs e)
		{
			Double dPago = (double) txtTotal.Value;
			int idPLote = (int) idPagoLote.Value;
			PagoAutorizaEstimado PagoAE = new PagoAutorizaEstimado(dPago,idPLote);
			PagoAE.MdiParent = base.MdiParent;
			PagoAE.Show();
		}

		private void txtTotal_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btAgregarTodo_Click_1(object sender, System.EventArgs e)
		{
		
		}


//		private void btEstimado_Click(object sender, System.EventArgs e)
//		{
//			DateTime dtDesde = (DateTime) txtFecha.Value;
//			DateTime dtHasta = (DateTime) cmbVencimiento.Value;
//
//			ParameterFields paramFields = new ParameterFields ();
//
//			ParameterField paramField = new ParameterField ();
//			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
//			paramField.ParameterFieldName = "@FechaIni";
//			discreteVal.Value = dtDesde.ToString("yyyy-MM-dd");
//			paramField.CurrentValues.Add (discreteVal);
//			paramFields.Add (paramField);
//
//			ParameterField paramField1 = new ParameterField ();
//			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
//			paramField1.ParameterFieldName = "@FechaFin";
//			discreteVal1.Value = dtHasta.ToString("yyyy-MM-dd");
//			paramField1.CurrentValues.Add (discreteVal1);
//					
//			paramFields.Add (paramField1);
//
//			string stReporte = "FlujoEfectivoE.rpt";
//			Reporte miReporte = new Reporte(stReporte, "");
//			miReporte.MdiParent = this.MdiParent;
//			//					if (chkResumen.Checked) miReporte.Resumen();
//			miReporte.crVista.ParameterFieldInfo = paramFields;
//			miReporte.Show();	
//







	}
}
