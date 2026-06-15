using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CRMClientes.
	/// </summary>
	public class CRMClientes : DevExpress.XtraEditors.XtraForm
	{
		protected Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Splitter splitter1;
		private System.Windows.Forms.Panel panel1;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private System.Windows.Forms.Splitter splitter2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataView cdvCompra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private C1.Data.C1DataSet cdsClienteTabla;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.Misc.UltraButton btFacturas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		private Infragistics.Win.UltraWinChart.UltraChart ultraChart1;
		private System.Windows.Forms.Splitter splitter3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsCompra;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid3;
		private Infragistics.Win.UltraWinChart.UltraChart ultraChart2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsCliente;
		private Infragistics.Win.Misc.UltraButton btCliente;
		private Infragistics.Win.Misc.UltraButton btArticulo;
		private Infragistics.Win.Misc.UltraButton btPagos;
		private Infragistics.Win.Misc.UltraButton btBuscaNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsBuscaCliente;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGrupo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSector;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCiudad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSector;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCiudad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CRMClientes()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sector");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_Compra", -1);
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Compra - DetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Compra - DetCompra", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefNumero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flete");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pedido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unidades");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoMedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteCiudad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteSector", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			this.ultraChart1 = new Infragistics.Win.UltraWinChart.UltraChart();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btPagos = new Infragistics.Win.Misc.UltraButton();
			this.btArticulo = new Infragistics.Win.Misc.UltraButton();
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.udsBuscaCliente = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btFacturas = new Infragistics.Win.Misc.UltraButton();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid3 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsCompra = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbCiudad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsClienteTabla = new C1.Data.C1DataSet();
			this.cmbSector = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkCiudad = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSector = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkGrupo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btBuscaNombre = new Infragistics.Win.Misc.UltraButton();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.splitter1 = new System.Windows.Forms.Splitter();
			this.panel1 = new System.Windows.Forms.Panel();
			this.ultraChart2 = new Infragistics.Win.UltraWinChart.UltraChart();
			this.splitter3 = new System.Windows.Forms.Splitter();
			this.splitter2 = new System.Windows.Forms.Splitter();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.cdvCompra = new C1.Data.C1DataView();
			this.udsCliente = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.ultraChart1)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBuscaCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCompra)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClienteTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.panel1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraChart2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraChart1
			// 
			this.ultraChart1.Axis.X.Labels.Flip = false;
			this.ultraChart1.Axis.X.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart1.Axis.X.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.Custom;
			this.ultraChart1.Axis.X.Labels.ItemFormatString = "<ITEM_LABEL:MM/yyyy>";
			this.ultraChart1.Axis.X.Labels.OrientationAngle = 0;
			this.ultraChart1.Axis.X.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart1.Axis.X.ScrollScale.Height = 10;
			this.ultraChart1.Axis.X.ScrollScale.Visible = false;
			this.ultraChart1.Axis.X.ScrollScale.Width = 15;
			this.ultraChart1.Axis.X.TickmarkInterval = 0;
			this.ultraChart1.Axis.X2.Labels.Flip = false;
			this.ultraChart1.Axis.X2.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.ItemLabel;
			this.ultraChart1.Axis.X2.Labels.ItemFormatString = "<ITEM_LABEL>";
			this.ultraChart1.Axis.X2.Labels.OrientationAngle = 0;
			this.ultraChart1.Axis.X2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart1.Axis.X2.ScrollScale.Height = 10;
			this.ultraChart1.Axis.X2.ScrollScale.Visible = false;
			this.ultraChart1.Axis.X2.ScrollScale.Width = 15;
			this.ultraChart1.Axis.X2.TickmarkInterval = 0;
			this.ultraChart1.Axis.Y.Labels.Flip = false;
			this.ultraChart1.Axis.Y.Labels.HorizontalAlign = System.Drawing.StringAlignment.Far;
			this.ultraChart1.Axis.Y.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.DataValue;
			this.ultraChart1.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,#00>";
			this.ultraChart1.Axis.Y.Labels.OrientationAngle = 0;
			this.ultraChart1.Axis.Y.Labels.SeriesFormatString = "";
			this.ultraChart1.Axis.Y.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart1.Axis.Y.ScrollScale.Height = 10;
			this.ultraChart1.Axis.Y.ScrollScale.Visible = false;
			this.ultraChart1.Axis.Y.ScrollScale.Width = 15;
			this.ultraChart1.Axis.Y.TickmarkInterval = 0;
			this.ultraChart1.Axis.Y2.Labels.Flip = false;
			this.ultraChart1.Axis.Y2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart1.Axis.Y2.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.DataValue;
			this.ultraChart1.Axis.Y2.Labels.ItemFormatString = "<DATA_VALUE:00.00>";
			this.ultraChart1.Axis.Y2.Labels.OrientationAngle = 0;
			this.ultraChart1.Axis.Y2.Labels.SeriesFormatString = "";
			this.ultraChart1.Axis.Y2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart1.Axis.Y2.ScrollScale.Height = 10;
			this.ultraChart1.Axis.Y2.ScrollScale.Visible = false;
			this.ultraChart1.Axis.Y2.ScrollScale.Width = 15;
			this.ultraChart1.Axis.Y2.TickmarkInterval = 0;
			this.ultraChart1.Axis.Z.Labels.Flip = false;
			this.ultraChart1.Axis.Z.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart1.Axis.Z.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.ItemLabel;
			this.ultraChart1.Axis.Z.Labels.ItemFormatString = "<ITEM_LABEL>";
			this.ultraChart1.Axis.Z.Labels.OrientationAngle = 0;
			this.ultraChart1.Axis.Z.Labels.SeriesFormatString = "";
			this.ultraChart1.Axis.Z.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart1.Axis.Z.ScrollScale.Height = 10;
			this.ultraChart1.Axis.Z.ScrollScale.Visible = false;
			this.ultraChart1.Axis.Z.ScrollScale.Width = 15;
			this.ultraChart1.Axis.Z.TickmarkInterval = 0;
			this.ultraChart1.Axis.Z2.Labels.Flip = false;
			this.ultraChart1.Axis.Z2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart1.Axis.Z2.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.None;
			this.ultraChart1.Axis.Z2.Labels.ItemFormatString = "";
			this.ultraChart1.Axis.Z2.Labels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal;
			this.ultraChart1.Axis.Z2.Labels.OrientationAngle = 0;
//			this.ultraChart1.Axis.Z2.Labels.FormatString = "";
			this.ultraChart1.Axis.Z2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart1.Axis.Z2.ScrollScale.Height = 10;
			this.ultraChart1.Axis.Z2.ScrollScale.Visible = false;
			this.ultraChart1.Axis.Z2.ScrollScale.Width = 15;
			this.ultraChart1.Axis.Z2.TickmarkInterval = 0;
			this.ultraChart1.Data.SwapRowsAndColumns = true;
			this.ultraChart1.Data.ZeroAligned = true;
			this.ultraChart1.Dock = System.Windows.Forms.DockStyle.Left;
			this.ultraChart1.EmptyChartText = "No hay datos";
			this.ultraChart1.ForeColor = System.Drawing.SystemColors.ControlText;
			this.ultraChart1.Location = new System.Drawing.Point(0, 179);
			this.ultraChart1.Name = "ultraChart1";
			this.ultraChart1.Size = new System.Drawing.Size(304, 243);
			this.ultraChart1.TabIndex = 2;
			this.ultraChart1.TitleBottom.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart1.TitleBottom.Text = "";
			this.ultraChart1.TitleTop.Text = "Ventas Mensuales";
			this.ultraChart1.Tooltips.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F);
			this.ultraChart1.Tooltips.UseControl = false;
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.btPagos);
			this.ultraTabPageControl1.Controls.Add(this.btArticulo);
			this.ultraTabPageControl1.Controls.Add(this.btCliente);
			this.ultraTabPageControl1.Controls.Add(this.txtIdCliente);
			this.ultraTabPageControl1.Controls.Add(this.btFacturas);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor2);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor1);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(2, 24);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(561, 150);
			// 
			// btPagos
			// 
			this.btPagos.Location = new System.Drawing.Point(32, 104);
			this.btPagos.Name = "btPagos";
			this.btPagos.TabIndex = 6;
			this.btPagos.Text = "Pagos";
			this.btPagos.Click += new System.EventHandler(this.btPagos_Click);
			// 
			// btArticulo
			// 
			this.btArticulo.Location = new System.Drawing.Point(312, 104);
			this.btArticulo.Name = "btArticulo";
			this.btArticulo.TabIndex = 5;
			this.btArticulo.Text = "Articulos";
			this.btArticulo.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// btCliente
			// 
			this.btCliente.Location = new System.Drawing.Point(312, 64);
			this.btCliente.Name = "btCliente";
			this.btCliente.TabIndex = 4;
			this.btCliente.Text = "Cliente";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.udsBuscaCliente, "Band 0.idCliente"));
			this.txtIdCliente.Location = new System.Drawing.Point(336, 16);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(48, 21);
			this.txtIdCliente.TabIndex = 3;
			// 
			// udsBuscaCliente
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.udsBuscaCliente.Band.Columns.AddRange(new object[] {
																																ultraDataColumn1,
																																ultraDataColumn2,
																																ultraDataColumn3,
																																ultraDataColumn4,
																																ultraDataColumn5,
																																ultraDataColumn6,
																																ultraDataColumn7});
			// 
			// btFacturas
			// 
			this.btFacturas.Location = new System.Drawing.Point(32, 64);
			this.btFacturas.Name = "btFacturas";
			this.btFacturas.TabIndex = 2;
			this.btFacturas.Text = "Facturas";
			this.btFacturas.Click += new System.EventHandler(this.btFacturas_Click);
			// 
			// ultraTextEditor2
			// 
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.udsBuscaCliente, "Band 0.Ruc"));
			this.ultraTextEditor2.Location = new System.Drawing.Point(224, 16);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(100, 21);
			this.ultraTextEditor2.TabIndex = 1;
			this.ultraTextEditor2.Text = "ultraTextEditor2";
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.udsBuscaCliente, "Band 0.Nombre"));
			this.ultraTextEditor1.Location = new System.Drawing.Point(16, 16);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(192, 21);
			this.ultraTextEditor1.TabIndex = 0;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.ultraGrid2);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(561, 150);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "_Compra";
			this.ultraGrid2.DataSource = this.cdsCompra;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 7;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 8;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 9;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 10;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 11;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 12;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 13;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 14;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 15;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 16;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 17;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 18;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 19;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 20;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 21;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 22;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 23;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 24;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 25;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 26;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn22.Header.VisiblePosition = 27;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn26.Header.VisiblePosition = 5;
			ultraGridColumn27.Header.VisiblePosition = 2;
			ultraGridColumn28.Header.VisiblePosition = 28;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 29;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 30;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 31;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 32;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 33;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 34;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 35;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 36;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 37;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 38;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 6;
			ultraGridColumn40.Header.VisiblePosition = 39;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 41;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 42;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 43;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 44;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 45;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 46;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 47;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 48;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 49;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 50;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 51;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 52;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 53;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 54;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 55;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 56;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 57;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 58;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 59;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 60;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 61;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 62;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 63;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 64;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 65;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 66;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 67;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 68;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 69;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 70;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 71;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 72;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 73;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 74;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 75;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 76;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 77;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 78;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 79;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 80;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 81;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 82;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 83;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 84;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 85;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 86;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 87;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 88;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 89;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 90;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 91;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 92;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 93;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 94;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 95;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 96;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 97;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 98;
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
																										 ultraGridColumn99});
			ultraGridColumn100.Header.VisiblePosition = 4;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 5;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 1;
			ultraGridColumn103.Header.VisiblePosition = 6;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 7;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 8;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 0;
			ultraGridColumn107.Header.VisiblePosition = 2;
			ultraGridColumn108.Header.VisiblePosition = 9;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 10;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 11;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 12;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 3;
			ultraGridColumn113.Header.VisiblePosition = 13;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 14;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 15;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 16;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 17;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn118.Header.VisiblePosition = 18;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 19;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 20;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 21;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 22;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn123.Header.VisiblePosition = 23;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn124.Header.VisiblePosition = 24;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn125.Header.VisiblePosition = 25;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 26;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn127.Header.VisiblePosition = 27;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn128.Header.VisiblePosition = 28;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn129.Header.VisiblePosition = 29;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn130.Header.VisiblePosition = 30;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn131.Header.VisiblePosition = 31;
			ultraGridColumn131.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn131});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.ultraGrid2.Location = new System.Drawing.Point(0, 0);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(561, 150);
			this.ultraGrid2.TabIndex = 0;
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
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.ultraGrid3);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(561, 150);
			// 
			// ultraGrid3
			// 
			this.ultraGrid3.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid3.DataMember = "Band 0";
			this.ultraGrid3.DataSource = this.udsCompra;
			ultraGridColumn132.Header.VisiblePosition = 0;
			ultraGridColumn133.Header.VisiblePosition = 1;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn132,
																										 ultraGridColumn133});
			this.ultraGrid3.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraGrid3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid3.Location = new System.Drawing.Point(48, 24);
			this.ultraGrid3.Name = "ultraGrid3";
			this.ultraGrid3.Size = new System.Drawing.Size(224, 88);
			this.ultraGrid3.TabIndex = 0;
			// 
			// udsCompra
			// 
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(System.Double);
			this.udsCompra.Band.Columns.AddRange(new object[] {
																													ultraDataColumn8,
																													ultraDataColumn9});
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.cmbCiudad);
			this.ultraTabPageControl4.Controls.Add(this.cmbSector);
			this.ultraTabPageControl4.Controls.Add(this.cmbGrupo);
			this.ultraTabPageControl4.Controls.Add(this.chkCiudad);
			this.ultraTabPageControl4.Controls.Add(this.chkSector);
			this.ultraTabPageControl4.Controls.Add(this.chkGrupo);
			this.ultraTabPageControl4.Controls.Add(this.txtNombre);
			this.ultraTabPageControl4.Controls.Add(this.btBuscaNombre);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(561, 150);
			// 
			// cmbCiudad
			// 
			this.cmbCiudad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCiudad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCiudad.DataMember = "ClienteCiudad";
			this.cmbCiudad.DataSource = this.cdsClienteTabla;
			ultraGridBand4.AddButtonCaption = "ClienteGrupo";
			ultraGridColumn134.Header.VisiblePosition = 1;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn135.Header.VisiblePosition = 0;
			ultraGridColumn135.Width = 244;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn134,
																										 ultraGridColumn135});
			this.cmbCiudad.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCiudad.DisplayMember = "Ciudad";
			this.cmbCiudad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCiudad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCiudad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCiudad.Location = new System.Drawing.Point(144, 104);
			this.cmbCiudad.Name = "cmbCiudad";
			this.cmbCiudad.Size = new System.Drawing.Size(200, 21);
			this.cmbCiudad.TabIndex = 7;
			this.cmbCiudad.ValueMember = "idCiudad";
			// 
			// cdsClienteTabla
			// 
			this.cdsClienteTabla.BindingContextCtrl = this;
			this.cdsClienteTabla.DataLibrary = "LibFacturacion";
			this.cdsClienteTabla.DataLibraryUrl = "";
			this.cdsClienteTabla.DataSetDef = "dsClienteTabla";
			this.cdsClienteTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsClienteTabla.Name = "cdsClienteTabla";
			this.cdsClienteTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsClienteTabla.SchemaDef = null;
			// 
			// cmbSector
			// 
			this.cmbSector.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSector.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSector.DataMember = "ClienteSector";
			this.cmbSector.DataSource = this.cdsClienteTabla;
			ultraGridBand5.AddButtonCaption = "ClienteGrupo";
			ultraGridColumn136.Header.VisiblePosition = 1;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 0;
			ultraGridColumn137.Width = 238;
			ultraGridColumn138.Header.VisiblePosition = 2;
			ultraGridColumn138.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn136,
																										 ultraGridColumn137,
																										 ultraGridColumn138});
			this.cmbSector.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbSector.DisplayMember = "Sector";
			this.cmbSector.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSector.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSector.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSector.Location = new System.Drawing.Point(144, 80);
			this.cmbSector.Name = "cmbSector";
			this.cmbSector.Size = new System.Drawing.Size(200, 21);
			this.cmbSector.TabIndex = 6;
			this.cmbSector.ValueMember = "idSector";
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataMember = "ClienteGrupo";
			this.cmbGrupo.DataSource = this.cdsClienteTabla;
			ultraGridColumn139.Header.VisiblePosition = 1;
			ultraGridColumn139.Hidden = true;
			ultraGridColumn140.Header.VisiblePosition = 0;
			ultraGridColumn140.Width = 238;
			ultraGridColumn141.Header.VisiblePosition = 2;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn142.Header.VisiblePosition = 3;
			ultraGridColumn142.Hidden = true;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn139,
																										 ultraGridColumn140,
																										 ultraGridColumn141,
																										 ultraGridColumn142});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(144, 56);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(200, 21);
			this.cmbGrupo.TabIndex = 5;
			this.cmbGrupo.ValueMember = "idGrupoCliente";
			// 
			// chkCiudad
			// 
			this.chkCiudad.BackColor = System.Drawing.Color.Transparent;
			this.chkCiudad.Checked = true;
			this.chkCiudad.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCiudad.Location = new System.Drawing.Point(40, 104);
			this.chkCiudad.Name = "chkCiudad";
			this.chkCiudad.Size = new System.Drawing.Size(96, 20);
			this.chkCiudad.TabIndex = 4;
			this.chkCiudad.Text = "Toda Ciudad";
			// 
			// chkSector
			// 
			this.chkSector.BackColor = System.Drawing.Color.Transparent;
			this.chkSector.Checked = true;
			this.chkSector.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkSector.Location = new System.Drawing.Point(40, 80);
			this.chkSector.Name = "chkSector";
			this.chkSector.Size = new System.Drawing.Size(96, 20);
			this.chkSector.TabIndex = 3;
			this.chkSector.Text = "Todo Sector";
			// 
			// chkGrupo
			// 
			this.chkGrupo.BackColor = System.Drawing.Color.Transparent;
			this.chkGrupo.Checked = true;
			this.chkGrupo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkGrupo.Location = new System.Drawing.Point(40, 56);
			this.chkGrupo.Name = "chkGrupo";
			this.chkGrupo.Size = new System.Drawing.Size(96, 20);
			this.chkGrupo.TabIndex = 2;
			this.chkGrupo.Text = "Todo Grupo";
			// 
			// txtNombre
			// 
			this.txtNombre.Location = new System.Drawing.Point(128, 16);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(264, 21);
			this.txtNombre.TabIndex = 1;
			// 
			// btBuscaNombre
			// 
			this.btBuscaNombre.Location = new System.Drawing.Point(32, 16);
			this.btBuscaNombre.Name = "btBuscaNombre";
			this.btBuscaNombre.TabIndex = 0;
			this.btBuscaNombre.Text = "Nombre";
			this.btBuscaNombre.Click += new System.EventHandler(this.btBuscaNombre_Click);
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.FillOnRequest = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Band 0";
			this.ultraGrid1.DataSource = this.udsBuscaCliente;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			ultraGridBand7.AddButtonCaption = "Cliente";
			ultraGridColumn143.Header.VisiblePosition = 0;
			ultraGridColumn143.Width = 214;
			ultraGridColumn144.Header.VisiblePosition = 4;
			ultraGridColumn144.Width = 71;
			ultraGridColumn145.Header.VisiblePosition = 5;
			ultraGridColumn145.Width = 75;
			ultraGridColumn146.Header.VisiblePosition = 1;
			ultraGridColumn147.Header.VisiblePosition = 2;
			ultraGridColumn148.Header.VisiblePosition = 3;
			ultraGridColumn149.Header.VisiblePosition = 6;
			ultraGridColumn149.Hidden = true;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn143,
																										 ultraGridColumn144,
																										 ultraGridColumn145,
																										 ultraGridColumn146,
																										 ultraGridColumn147,
																										 ultraGridColumn148,
																										 ultraGridColumn149});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Dock = System.Windows.Forms.DockStyle.Left;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(0, 0);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(224, 422);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.ultraGrid1_MouseDown);
			this.ultraGrid1.AfterRowActivate += new System.EventHandler(this.ultraGrid1_AfterRowActivate);
			// 
			// splitter1
			// 
			this.splitter1.Location = new System.Drawing.Point(224, 0);
			this.splitter1.Name = "splitter1";
			this.splitter1.Size = new System.Drawing.Size(3, 422);
			this.splitter1.TabIndex = 1;
			this.splitter1.TabStop = false;
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.ultraChart2);
			this.panel1.Controls.Add(this.splitter3);
			this.panel1.Controls.Add(this.ultraChart1);
			this.panel1.Controls.Add(this.splitter2);
			this.panel1.Controls.Add(this.ultraTabControl1);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel1.Location = new System.Drawing.Point(227, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(565, 422);
			this.panel1.TabIndex = 2;
			// 
			// ultraChart2
			// 
			this.ultraChart2.Axis.X.Labels.Flip = false;
			this.ultraChart2.Axis.X.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart2.Axis.X.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.ItemLabel;
			this.ultraChart2.Axis.X.Labels.ItemFormatString = "<ITEM_LABEL>";
			this.ultraChart2.Axis.X.Labels.OrientationAngle = 0;
			this.ultraChart2.Axis.X.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart2.Axis.X.Labels.Visible = false;
			this.ultraChart2.Axis.X.ScrollScale.Height = 10;
			this.ultraChart2.Axis.X.ScrollScale.Visible = false;
			this.ultraChart2.Axis.X.ScrollScale.Width = 15;
			this.ultraChart2.Axis.X.TickmarkInterval = 0;
			this.ultraChart2.Axis.X2.Labels.Flip = false;
			this.ultraChart2.Axis.X2.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.ItemLabel;
			this.ultraChart2.Axis.X2.Labels.ItemFormatString = "<ITEM_LABEL>";
			this.ultraChart2.Axis.X2.Labels.OrientationAngle = 0;
			this.ultraChart2.Axis.X2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart2.Axis.X2.ScrollScale.Height = 10;
			this.ultraChart2.Axis.X2.ScrollScale.Visible = false;
			this.ultraChart2.Axis.X2.ScrollScale.Width = 15;
			this.ultraChart2.Axis.X2.TickmarkInterval = 0;
			this.ultraChart2.Axis.Y.Labels.Flip = false;
			this.ultraChart2.Axis.Y.Labels.HorizontalAlign = System.Drawing.StringAlignment.Far;
			this.ultraChart2.Axis.Y.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.DataValue;
			this.ultraChart2.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,#00>";
			this.ultraChart2.Axis.Y.Labels.OrientationAngle = 0;
			this.ultraChart2.Axis.Y.Labels.SeriesFormatString = "";
			this.ultraChart2.Axis.Y.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart2.Axis.Y.ScrollScale.Height = 10;
			this.ultraChart2.Axis.Y.ScrollScale.Visible = false;
			this.ultraChart2.Axis.Y.ScrollScale.Width = 15;
			this.ultraChart2.Axis.Y.TickmarkInterval = 0;
			this.ultraChart2.Axis.Y2.Labels.Flip = false;
			this.ultraChart2.Axis.Y2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart2.Axis.Y2.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.DataValue;
			this.ultraChart2.Axis.Y2.Labels.ItemFormatString = "<DATA_VALUE:00.00>";
			this.ultraChart2.Axis.Y2.Labels.OrientationAngle = 0;
			this.ultraChart2.Axis.Y2.Labels.SeriesFormatString = "";
			this.ultraChart2.Axis.Y2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart2.Axis.Y2.ScrollScale.Height = 10;
			this.ultraChart2.Axis.Y2.ScrollScale.Visible = false;
			this.ultraChart2.Axis.Y2.ScrollScale.Width = 15;
			this.ultraChart2.Axis.Y2.TickmarkInterval = 0;
			this.ultraChart2.Axis.Z.Labels.Flip = false;
			this.ultraChart2.Axis.Z.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart2.Axis.Z.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.None;
			this.ultraChart2.Axis.Z.Labels.ItemFormatString = "";
			this.ultraChart2.Axis.Z.Labels.OrientationAngle = 0;
			this.ultraChart2.Axis.Z.Labels.SeriesFormatString = "";
			this.ultraChart2.Axis.Z.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart2.Axis.Z.ScrollScale.Height = 10;
			this.ultraChart2.Axis.Z.ScrollScale.Visible = false;
			this.ultraChart2.Axis.Z.ScrollScale.Width = 15;
			this.ultraChart2.Axis.Z.TickmarkInterval = 0;
			this.ultraChart2.Axis.Z2.Labels.Flip = false;
			this.ultraChart2.Axis.Z2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart2.Axis.Z2.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.None;
			this.ultraChart2.Axis.Z2.Labels.ItemFormatString = "";
			this.ultraChart2.Axis.Z2.Labels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal;
			this.ultraChart2.Axis.Z2.Labels.OrientationAngle = 0;
			this.ultraChart2.Axis.Z2.Labels.SeriesFormatString = "";
			this.ultraChart2.Axis.Z2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart2.Axis.Z2.ScrollScale.Height = 10;
			this.ultraChart2.Axis.Z2.ScrollScale.Visible = false;
			this.ultraChart2.Axis.Z2.ScrollScale.Width = 15;
			this.ultraChart2.Axis.Z2.TickmarkInterval = 0;
			this.ultraChart2.Data.SwapRowsAndColumns = true;
			this.ultraChart2.Data.ZeroAligned = true;
			this.ultraChart2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.ultraChart2.ForeColor = System.Drawing.SystemColors.ControlText;
			this.ultraChart2.Legend.Location = Infragistics.UltraChart.Shared.Styles.LegendLocation.Bottom;
			this.ultraChart2.Legend.Visible = true;
			this.ultraChart2.Location = new System.Drawing.Point(307, 179);
			this.ultraChart2.Name = "ultraChart2";
			this.ultraChart2.Size = new System.Drawing.Size(258, 243);
			this.ultraChart2.TabIndex = 4;
			this.ultraChart2.TitleBottom.Text = "";
			this.ultraChart2.TitleTop.Text = "Comparacion entre Clientes";
			this.ultraChart2.Tooltips.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F);
			this.ultraChart2.Tooltips.UseControl = false;
			// 
			// splitter3
			// 
			this.splitter3.Location = new System.Drawing.Point(304, 179);
			this.splitter3.Name = "splitter3";
			this.splitter3.Size = new System.Drawing.Size(3, 243);
			this.splitter3.TabIndex = 3;
			this.splitter3.TabStop = false;
			// 
			// splitter2
			// 
			this.splitter2.Dock = System.Windows.Forms.DockStyle.Top;
			this.splitter2.Location = new System.Drawing.Point(0, 176);
			this.splitter2.Name = "splitter2";
			this.splitter2.Size = new System.Drawing.Size(565, 3);
			this.splitter2.TabIndex = 1;
			this.splitter2.TabStop = false;
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl3);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl4);
			this.ultraTabControl1.Dock = System.Windows.Forms.DockStyle.Top;
			this.ultraTabControl1.Location = new System.Drawing.Point(0, 0);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(565, 176);
			this.ultraTabControl1.TabIndex = 0;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Información";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Facturas";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Detalle";
			ultraTab4.TabPage = this.ultraTabPageControl4;
			ultraTab4.Text = "Opciones";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2,
																																															 ultraTab3,
																																															 ultraTab4});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(561, 150);
			// 
			// cdvCompra
			// 
			this.cdvCompra.BindingContextCtrl = this;
			this.cdvCompra.DataSet = this.cdsCompra;
			this.cdvCompra.TableName = "";
			this.cdvCompra.TableViewName = "Compra";
			// 
			// udsCliente
			// 
			ultraDataColumn11.DataType = typeof(System.Double);
			this.udsCliente.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn10,
																													 ultraDataColumn11});
			// 
			// CRMClientes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(792, 422);
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.splitter1);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "CRMClientes";
			this.Text = "CRMClientes";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Closing += new System.ComponentModel.CancelEventHandler(this.CRMClientes_Closing);
			this.Load += new System.EventHandler(this.CRMClientes_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraChart1)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBuscaCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCompra)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClienteTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.panel1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraChart2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void CRMClientes_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'CRM Clientes'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
			txtIdCliente.Width = 0;
			btBuscaNombre_Click(this, e);
		}

		private void btFacturas_Click(object sender, System.EventArgs e)
		{
			string stFiltro = "idCliente = " + txtIdCliente.Value.ToString();
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsCompra, "Compra", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			C1.Data.FilterCondition fcHija = new 
				C1.Data.FilterCondition(this.cdsCompra, "DetCompra", "idCompra = 0");
			fcTotal.Add(fcHija);
			cdsCompra.Fill(fcTotal, false);

			string stSelect = "select sum(Compra.Total), Max(Fecha) From Compra "
				+ "Where Compra.idCliente = " + txtIdCliente.Value.ToString() + " "
				+ "Group by Month(Compra.Fecha) Order by Max(Compra.Fecha) asc";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCliente, stSelect, true);

			udsCompra.Rows.Clear();
			int iNumFilas = -1;
			while(miReader.Read())
			{
				iNumFilas ++;
				udsCompra.Rows.Add();
				udsCompra.Rows[iNumFilas]["Valor"]=miReader.GetDouble(0);
				udsCompra.Rows[iNumFilas]["Fecha"]=miReader.GetDateTime(1);
			}
			miReader.Close();
			ultraChart1.TitleTop.Text = "Facturas por meses";
			ultraChart1.DataSource = udsCompra;
			ultraChart1.DataBind();
		}

		private void btCliente_Click(object sender, System.EventArgs e)
		{
			int iNumFilas = 0;
			udsCliente.Rows.Clear();
			string stSelect = "select top 1 sum(IsNull(Compra.Total, 0)), "
				+ "Max(IsNull(Cliente.Nombre, '')) From Compra "
				+ "Inner Join Cliente on Compra.idCliente = Cliente.idCliente "
				+ "Where Compra.idCliente = " + txtIdCliente.Value.ToString();
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCliente, stSelect, true);
			while (miReader.Read())
			{
				udsCliente.Rows.Add();
				udsCliente.Rows[iNumFilas]["Total"]=miReader.GetDouble(0);
				udsCliente.Rows[iNumFilas]["Nombre"]=miReader.GetString(1);
			}
			miReader.Close();
			stSelect = "select top 10 sum(IsNull(Compra.Total, 0)), "
				+ "Max(IsNull(Cliente.Nombre, '')) From Compra "
				+ "Inner Join Cliente on Compra.idCliente = Cliente.idCliente "
				+ "Group by Compra.idCliente Order by sum(Compra.Total) desc ";
			miReader = Funcion.rEscalarSQL(cdsCliente, stSelect, true);
			while (miReader.Read())
			{
				iNumFilas ++;
				udsCliente.Rows.Add();
				udsCliente.Rows[iNumFilas]["Total"]=miReader.GetDouble(0);
				udsCliente.Rows[iNumFilas]["Nombre"]=miReader.GetString(1);
			}
			miReader.Close();
			ultraChart2.TitleTop.Text = "Comparacion de Clientes";
			ultraChart2.DataSource = udsCliente;
			ultraChart2.DataBind();
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			string stSelect = "Select Top 10 sum(DetCompra.Cantidad * DetCompra.Precio), "
				+ "Max(Articulo.Articulo) From DetCompra "
				+ "Inner Join Articulo On DetCompra.idArticulo = Articulo.idArticulo "
				+ "Inner Join Compra On DetCompra.idCompra = Compra.idCompra "
				+ "Where Compra.idCliente = " + txtIdCliente.Value.ToString() + " "
				+ "Group by DetCompra.idArticulo "
				+ "Order by sum(DetCompra.Cantidad * DetCompra.Precio) desc";
			int iNumFilas = -1;
			udsCliente.Rows.Clear();
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCliente, stSelect, true);
			while (miReader.Read())
			{
				iNumFilas ++;
				udsCliente.Rows.Add();
				udsCliente.Rows[iNumFilas]["Total"]=miReader.GetDouble(0);
				udsCliente.Rows[iNumFilas]["Nombre"]=miReader.GetString(1);
			}
			miReader.Close();
			ultraChart2.TitleTop.Text = "Comparacion de Articulos";
			ultraChart2.DataSource = udsCliente;
			ultraChart2.DataBind();
		}

		private void btPagos_Click(object sender, System.EventArgs e)
		{
			string stSelect = "Select Top 10 sum(Pago.Pago), "
				+ "Max(Pago.Fecha) From Pago "
				+ "Inner Join Compra On Pago.idCompra = Compra.idCompra "
				+ "Where Compra.idCliente = " + txtIdCliente.Value.ToString() + " "
				+ "Group by Month(Pago.Fecha) "
				+ "Order by Max(Pago.Fecha) desc";
			int iNumFilas = -1;
			udsCompra.Rows.Clear();
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCliente, stSelect, true);
			while (miReader.Read())
			{
				iNumFilas ++;
				udsCompra.Rows.Add();
				udsCompra.Rows[iNumFilas]["Valor"]=miReader.GetDouble(0);
				udsCompra.Rows[iNumFilas]["Fecha"]=miReader.GetDateTime(1);
			}
			miReader.Close();
			ultraChart1.TitleTop.Text = "Pagos por meses";
			ultraChart1.DataSource = udsCompra;
			ultraChart1.DataBind();	
		}

		private void btBuscaNombre_Click(object sender, System.EventArgs e)
		{
			string stSelect = "SELECT Top 20 Max(IsNull(Cliente.Nombre, '')), "
				+ "Max(IsNull(Cliente.Codigo, '')), "
				+ "Max(IsNull(Cliente.Ruc, '')), Max(IsNull(ClienteGrupo.Grupo, '')), "
				+ "Max(IsNull(ClienteCiudad.Ciudad, '')), Max(IsNull(ClienteSector.Sector, '')), "
				+ "Max(IsNull(Cliente.IdCliente, 0)) "
				+ "FROM Cliente left JOIN "
				+ "ClienteGrupo ON Cliente.idGrupoCliente = ClienteGrupo.idGrupoCliente left JOIN "
				+ "ClienteCiudad ON Cliente.idCiudad = ClienteCiudad.idCiudad left JOIN "
				+ "ClienteSector ON Cliente.idSector = ClienteSector.idSector "
				+ "Inner Join Compra On Compra.idCliente = Cliente.idCliente Where 1=1 ";
			if (txtNombre.Text.Trim().Length > 0)
				stSelect += "And Cliente.Nombre = '" + txtNombre.Text.Trim() + "' ";
			if (!chkGrupo.Checked)
				stSelect += "And ClienteGrupo.idGrupoCliente = " + cmbGrupo.Value.ToString() + " ";
			if (!chkSector.Checked)
				stSelect += "And ClienteSector.idSector = " + cmbSector.Value.ToString() + " ";
			if (!chkCiudad.Checked)
				stSelect += "And ClienteCiudad.idCiudad = " + cmbCiudad.Value.ToString() + " ";
			stSelect +="Group by Compra.idCliente "
				+ "Order by Sum(Compra.Total) desc ";
			udsBuscaCliente.Rows.Clear();
			int iNumFilas = -1;
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCliente, stSelect, true);
			while (miReader.Read())
			{
				iNumFilas ++;
				udsBuscaCliente.Rows.Add();
				udsBuscaCliente.Rows[iNumFilas]["Nombre"]=miReader.GetString(0);
				udsBuscaCliente.Rows[iNumFilas]["Codigo"]=miReader.GetString(1);
				udsBuscaCliente.Rows[iNumFilas]["Ruc"]=miReader.GetString(2);
				udsBuscaCliente.Rows[iNumFilas]["Grupo"]=miReader.GetString(3);
				udsBuscaCliente.Rows[iNumFilas]["Ciudad"]=miReader.GetString(4);
				udsBuscaCliente.Rows[iNumFilas]["Sector"]=miReader.GetString(5);
				udsBuscaCliente.Rows[iNumFilas]["idCliente"]=miReader.GetInt32(6);
			}
			miReader.Close();
			ultraGrid1.Refresh();
		}

		private void ultraGrid1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (ultraGrid1.ActiveRow == null) return;
			btFacturas_Click(this, e);
			ultraButton1_Click(this, e);
		}

		private void ultraGrid1_AfterRowActivate(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null) return;
			btFacturas_Click(this, e);
			ultraButton1_Click(this, e);
		}

		private void CRMClientes_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'CRM Clientes'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}
	}
}
