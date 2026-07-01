using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions.Shared;
using C1.Data;
using System.Drawing.Printing;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFacturacionHabitaciones.
	/// </summary>
	public class frmFacturacionHabitaciones : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCajero;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempo;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label lblHabitacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGuardia;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUnidad;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridArticulos;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTurnoIngreso;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTurnoSalida;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIngreso;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaSalida;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbHabitaciones;
		private System.Windows.Forms.Button btnTemporal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Label lblLocal;
		private System.Windows.Forms.Label lblHabitaciones;
		private System.Windows.Forms.Button btnCliente;
		private System.Windows.Forms.Label lblIngresos;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPersonas;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEncuesta;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEncuesta2;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Button btnClaveDescuento;
		private System.Windows.Forms.Button btnRefacturacion;
		private System.Windows.Forms.Button btnDescEmpAut;
		private System.Windows.Forms.ContextMenu mnuImprimir;
		private System.Windows.Forms.MenuItem mnuImprimirFactura;
		private System.Windows.Forms.MenuItem mnuImprimirRIDE;
		private System.Windows.Forms.MenuItem mnuImprimirComprobanteVenta;
		private System.Windows.Forms.Button btnMitadPrecio;
		private System.Windows.Forms.Button btnExtras;
		private System.Windows.Forms.Button btnCupon;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCupon;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCupon;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempoAdicional;
		private System.Windows.Forms.Button btnTiempoAdicional;
		private System.Windows.Forms.Button btnanticiposPendientes;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAbono;
		private System.Windows.Forms.Button btnOrdenCocina;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTarjetaVIP;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTarjetaVIP;
		private System.Windows.Forms.Button btnInfoTarjetasVIP;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPuntosTarjetaVIP;
		private System.Windows.Forms.Label lblPuntosTarjetaVIP;
		private System.Windows.Forms.Button btnCanjePuntos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVideo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button BtnReImprimirComanda;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private System.Windows.Forms.Button btnReturnPass;
		private System.Windows.Forms.Label lblSaldoFactura;

		Infragistics.Win.UltraWinGrid.UltraGrid UGridFP;

		public frmFacturacionHabitaciones()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		int idGuardia = 0;
		int idHabitacion = 0;
		int iBodega = 0;
		int iTarifaExpress = 0;
		int idCompraV = 0;
		bool bMenuCuadricula = false;
		bool bServicioPagado = false;
		bool bVerificacionPago = false;
		bool bPreventa = false;

		public frmFacturacionHabitaciones(int IdHabitacion, int IBodega, bool BMenuCuadricula, int IdCompraV, bool BServicioPagado, int ITarifaExpress)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idHabitacion = IdHabitacion;
			iBodega = IBodega;
			bMenuCuadricula = BMenuCuadricula;
			idCompraV = IdCompraV;
			bServicioPagado = BServicioPagado;
			iTarifaExpress = ITarifaExpress;

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		public frmFacturacionHabitaciones(int IdHabitacion, int IBodega, bool BMenuCuadricula, int IdCompraV, bool BServicioPagado, int ITarifaExpress, bool BRegistroCobro, bool BVerificacionPago)
		{
			InitializeComponent();

			idHabitacion = IdHabitacion;
			iBodega = IBodega;
			bMenuCuadricula = BMenuCuadricula;
			idCompraV = IdCompraV;
			bServicioPagado = BServicioPagado;
			iTarifaExpress = ITarifaExpress;

			bRegistroCobro = BRegistroCobro;
			bVerificacionPago = BVerificacionPago;

		}
		public frmFacturacionHabitaciones(int IdHabitacion,int IBodega,bool BMenuCuadricula,int IdCompraV, bool BServicioPagado, int ITarifaExpress, bool BRegistroCobro, bool BVerificacionPago, bool BPreventa)
		{
			InitializeComponent();

			idHabitacion = IdHabitacion;
			iBodega = IBodega;
			bMenuCuadricula = BMenuCuadricula;
			idCompraV = IdCompraV;
			bServicioPagado = BServicioPagado;
			iTarifaExpress = ITarifaExpress;
			bRegistroCobro = BRegistroCobro;
			bVerificacionPago = BVerificacionPago;
			bPreventa = BPreventa;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Turno");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFacturacionHabitaciones));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeleccionar", 0);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio2");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			this.cmbCajero = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.txtTiempo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnCliente = new System.Windows.Forms.Button();
			this.cmbCobrador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbGuardia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label18 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label12 = new System.Windows.Forms.Label();
			this.lblHabitacion = new System.Windows.Forms.Label();
			this.uGridArticulos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtUnidad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.btnVer = new System.Windows.Forms.Button();
			this.lblLocal = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblHabitaciones = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.dtFechaIngreso = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.dtFechaSalida = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.optTurnoIngreso = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optTurnoSalida = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbHabitaciones = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnTemporal = new System.Windows.Forms.Button();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.lblIngresos = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtPersonas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEncuesta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtEncuesta2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.btnClaveDescuento = new System.Windows.Forms.Button();
			this.btnRefacturacion = new System.Windows.Forms.Button();
			this.btnDescEmpAut = new System.Windows.Forms.Button();
			this.mnuImprimir = new System.Windows.Forms.ContextMenu();
			this.mnuImprimirFactura = new System.Windows.Forms.MenuItem();
			this.mnuImprimirRIDE = new System.Windows.Forms.MenuItem();
			this.mnuImprimirComprobanteVenta = new System.Windows.Forms.MenuItem();
			this.btnMitadPrecio = new System.Windows.Forms.Button();
			this.btnExtras = new System.Windows.Forms.Button();
			this.btnCupon = new System.Windows.Forms.Button();
			this.chkCupon = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtCupon = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.txtTiempoAdicional = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnTiempoAdicional = new System.Windows.Forms.Button();
			this.btnanticiposPendientes = new System.Windows.Forms.Button();
			this.chkAbono = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnOrdenCocina = new System.Windows.Forms.Button();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.txtTarjetaVIP = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkTarjetaVIP = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnInfoTarjetasVIP = new System.Windows.Forms.Button();
			this.lblPuntosTarjetaVIP = new System.Windows.Forms.Label();
			this.txtPuntosTarjetaVIP = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnCanjePuntos = new System.Windows.Forms.Button();
			this.chkVideo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.BtnReImprimirComanda = new System.Windows.Forms.Button();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnReturnPass = new System.Windows.Forms.Button();
			this.lblSaldoFactura = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGuardia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridArticulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUnidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTurnoIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTurnoSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHabitaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPersonas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEncuesta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEncuesta2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCupon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoAdicional)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTarjetaVIP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPuntosTarjetaVIP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbCajero
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCajero.Appearance = appearance1;
			this.cmbCajero.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCajero.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 250;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 30;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbCajero.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCajero.DisplayMember = "Nombre";
			this.cmbCajero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCajero.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCajero.Enabled = false;
			this.cmbCajero.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCajero.Location = new System.Drawing.Point(368, 48);
			this.cmbCajero.Name = "cmbCajero";
			this.cmbCajero.Size = new System.Drawing.Size(180, 22);
			this.cmbCajero.TabIndex = 267;
			this.cmbCajero.ValueMember = "idPersonal";
			this.cmbCajero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCajero_KeyDown);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label19.Location = new System.Drawing.Point(312, 51);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(36, 17);
			this.label19.TabIndex = 266;
			this.label19.Text = "Cajero";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTiempo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempo.Appearance = appearance2;
			this.txtTiempo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempo.Enabled = false;
			this.txtTiempo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTiempo.FormatString = "#0";
			this.txtTiempo.Location = new System.Drawing.Point(576, 80);
			this.txtTiempo.MaskInput = "nnnnnnnnn";
			this.txtTiempo.MaxValue = 24;
			this.txtTiempo.MinValue = 1;
			this.txtTiempo.Name = "txtTiempo";
			this.txtTiempo.PromptChar = ' ';
			this.txtTiempo.Size = new System.Drawing.Size(40, 22);
			this.txtTiempo.TabIndex = 261;
			this.txtTiempo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTiempo_KeyDown);
			this.txtTiempo.Click += new System.EventHandler(this.txtTiempo_Click);
			this.txtTiempo.Enter += new System.EventHandler(this.txtTiempo_Enter);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(512, 80);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(41, 17);
			this.label14.TabIndex = 258;
			this.label14.Text = "Tiempo";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn4.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn7.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn12.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn13.DefaultValue = 0;
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
			ultraDataColumn15.DataType = typeof(bool);
			ultraDataColumn15.DefaultValue = false;
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn16.DefaultValue = 0;
			ultraDataColumn17.DataType = typeof(bool);
			ultraDataColumn17.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
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
																																 ultraDataColumn17});
			// 
			// btnCliente
			// 
			this.btnCliente.Enabled = false;
			this.btnCliente.Image = ((System.Drawing.Image)(resources.GetObject("btnCliente.Image")));
			this.btnCliente.Location = new System.Drawing.Point(832, 110);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(26, 26);
			this.btnCliente.TabIndex = 250;
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
			// 
			// cmbCobrador
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCobrador.Appearance = appearance3;
			this.cmbCobrador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 335;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbCobrador.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCobrador.DisplayMember = "Nombre";
			this.cmbCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCobrador.Enabled = false;
			this.cmbCobrador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador.Location = new System.Drawing.Point(96, 48);
			this.cmbCobrador.Name = "cmbCobrador";
			this.cmbCobrador.Size = new System.Drawing.Size(180, 22);
			this.cmbCobrador.TabIndex = 265;
			this.cmbCobrador.ValueMember = "idPersonal";
			this.cmbCobrador.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCobrador_KeyDown);
			this.cmbCobrador.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCobrador_InitializeLayout);
			// 
			// cmbGuardia
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGuardia.Appearance = appearance4;
			this.cmbGuardia.CausesValidation = false;
			this.cmbGuardia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGuardia.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 335;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbGuardia.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbGuardia.DisplayMember = "Nombre";
			this.cmbGuardia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGuardia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGuardia.Enabled = false;
			this.cmbGuardia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGuardia.Location = new System.Drawing.Point(648, 48);
			this.cmbGuardia.Name = "cmbGuardia";
			this.cmbGuardia.Size = new System.Drawing.Size(180, 22);
			this.cmbGuardia.TabIndex = 264;
			this.cmbGuardia.ValueMember = "idPersonal";
			this.cmbGuardia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbGuardia_KeyDown);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.Location = new System.Drawing.Point(8, 51);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(50, 17);
			this.label18.TabIndex = 263;
			this.label18.Text = "Cobrador";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(576, 51);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(43, 17);
			this.label17.TabIndex = 262;
			this.label17.Text = "Guardia";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.Location = new System.Drawing.Point(736, 80);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(39, 17);
			this.label15.TabIndex = 260;
			this.label15.Text = "Unidad";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(8, 80);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(34, 17);
			this.label10.TabIndex = 255;
			this.label10.Text = "Fecha";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFormaPago
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance5;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 216;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(96, 144);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(152, 22);
			this.cmbFormaPago.TabIndex = 253;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbFormaPago_KeyDown);
			this.cmbFormaPago.ValueChanged += new System.EventHandler(this.cmbFormaPago_ValueChanged);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(832, 187);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(79, 17);
			this.label12.TabIndex = 251;
			this.label12.Text = "Buscar Articulo";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblHabitacion
			// 
			this.lblHabitacion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblHabitacion.Font = new System.Drawing.Font("Tahoma", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblHabitacion.ForeColor = System.Drawing.Color.Firebrick;
			this.lblHabitacion.Location = new System.Drawing.Point(8, 4);
			this.lblHabitacion.Name = "lblHabitacion";
			this.lblHabitacion.Size = new System.Drawing.Size(664, 30);
			this.lblHabitacion.TabIndex = 249;
			this.lblHabitacion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblHabitacion.Visible = false;
			// 
			// uGridArticulos
			// 
			this.uGridArticulos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridArticulos.DataSource = this.ultraDataSource3;
			appearance6.BackColor = System.Drawing.Color.White;
			this.uGridArticulos.DisplayLayout.Appearance = appearance6;
			this.uGridArticulos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Width = 299;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance7;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Precio";
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Width = 81;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Reserva";
			ultraGridColumn17.Header.VisiblePosition = 5;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 80;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			ultraGridColumn18.CellButtonAppearance = appearance8;
			ultraGridColumn18.Header.Caption = "...";
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 20;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.uGridArticulos.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridArticulos.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.uGridArticulos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridArticulos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridArticulos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridArticulos.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridArticulos.DisplayLayout.Override.CellAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 9F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridArticulos.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.uGridArticulos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridArticulos.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridArticulos.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridArticulos.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.uGridArticulos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridArticulos.Location = new System.Drawing.Point(832, 216);
			this.uGridArticulos.Name = "uGridArticulos";
			this.uGridArticulos.Size = new System.Drawing.Size(421, 296);
			this.uGridArticulos.TabIndex = 247;
			this.uGridArticulos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridArticulos_ClickCellButton);
			this.uGridArticulos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridArticulos_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// dtFecha
			// 
			appearance16.FontData.Name = "Tahoma";
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance16;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(96, 80);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 254;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblNumero.Location = new System.Drawing.Point(264, 80);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 17);
			this.lblNumero.TabIndex = 269;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance17;
			this.txtNumero.CausesValidation = false;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(368, 80);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(128, 22);
			this.txtNumero.TabIndex = 268;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(512, 112);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 17);
			this.label3.TabIndex = 437;
			this.label3.Text = "Nombre";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(8, 147);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 17);
			this.label8.TabIndex = 436;
			this.label8.Text = "Forma de Pago";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(560, 456);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(55, 29);
			this.label1.TabIndex = 435;
			this.label1.Text = "Total";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSubTotal.Location = new System.Drawing.Point(560, 371);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(45, 17);
			this.lblSubTotal.TabIndex = 428;
			this.lblSubTotal.Text = "Subtotal";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva
			// 
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(560, 419);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(32, 17);
			this.lblIva.TabIndex = 427;
			this.lblIva.Text = "I.V.A.";
			this.lblIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(560, 395);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(57, 17);
			this.lblDescuento.TabIndex = 426;
			this.lblDescuento.Text = "Descuento";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 112);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(71, 17);
			this.lblIdentificacion.TabIndex = 423;
			this.lblIdentificacion.Text = "Identificación";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCliente.Location = new System.Drawing.Point(264, 112);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(86, 17);
			this.lblCliente.TabIndex = 422;
			this.lblCliente.Text = "N. Identificación";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance18;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Tahoma", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(648, 453);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(176, 35);
			this.txtTotal.TabIndex = 434;
			this.txtTotal.ValueChanged += new System.EventHandler(this.txtTotal_ValueChanged);
			// 
			// txtIvaTotal
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance19;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(712, 416);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(112, 22);
			this.txtIvaTotal.TabIndex = 433;
			// 
			// txtDescIvaTotal
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance20;
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(712, 392);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(112, 22);
			this.txtDescIvaTotal.TabIndex = 431;
			// 
			// txtIva
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance21;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(712, 368);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(112, 22);
			this.txtIva.TabIndex = 430;
			// 
			// cmbTipoRuc
			// 
			appearance22.FontData.Name = "Tahoma";
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance22;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 233;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(96, 112);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(150, 22);
			this.cmbTipoRuc.TabIndex = 418;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoRuc_KeyDown);
			this.cmbTipoRuc.ValueChanged += new System.EventHandler(this.cmbTipoRuc_ValueChanged);
			// 
			// txtRuc
			// 
			this.txtRuc.AcceptsTab = true;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance23;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtRuc.Location = new System.Drawing.Point(368, 112);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(128, 22);
			this.txtRuc.TabIndex = 419;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			this.txtRuc.InitializeEditorButtonCheckState += new Infragistics.Win.UltraWinEditors.InitializeCheckStateEventHandler(this.txtRuc_InitializeEditorButtonCheckState);
			// 
			// txtNombre
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance24;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNombre.Location = new System.Drawing.Point(576, 112);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(248, 22);
			this.txtNombre.TabIndex = 420;
			// 
			// txtBuscar
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance25;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Enabled = false;
			this.txtBuscar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtBuscar.Location = new System.Drawing.Point(920, 184);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(240, 22);
			this.txtBuscar.TabIndex = 443;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(-8, 168);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1264, 8);
			this.groupBox1.TabIndex = 444;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(864, 96);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(400, 8);
			this.groupBox2.TabIndex = 445;
			this.groupBox2.TabStop = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(512, 144);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 446;
			// 
			// txtUnidad
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUnidad.Appearance = appearance26;
			this.txtUnidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUnidad.Enabled = false;
			this.txtUnidad.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtUnidad.FormatString = "#0";
			this.txtUnidad.Location = new System.Drawing.Point(784, 80);
			this.txtUnidad.MaxValue = 100000;
			this.txtUnidad.MinValue = 0;
			this.txtUnidad.Name = "txtUnidad";
			this.txtUnidad.PromptChar = ' ';
			this.txtUnidad.Size = new System.Drawing.Size(40, 22);
			this.txtUnidad.TabIndex = 447;
			this.txtUnidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtUnidad_KeyDown);
			this.txtUnidad.Click += new System.EventHandler(this.txtUnidad_Click);
			this.txtUnidad.Validated += new System.EventHandler(this.txtUnidad_Validated);
			this.txtUnidad.Enter += new System.EventHandler(this.txtUnidad_Enter);
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
			// btnAsiento
			// 
			this.btnAsiento.BackColor = System.Drawing.SystemColors.Control;
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.Location = new System.Drawing.Point(458, 536);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.TabIndex = 454;
			this.btnAsiento.Text = "Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			this.btnAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnAsiento_MouseDown);
			// 
			// btnNuevo
			// 
			this.btnNuevo.BackColor = System.Drawing.SystemColors.Control;
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 536);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.TabIndex = 448;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.SystemColors.Control;
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(83, 536);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.TabIndex = 449;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.BackColor = System.Drawing.SystemColors.Control;
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(158, 536);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.TabIndex = 452;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.BackColor = System.Drawing.SystemColors.Control;
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(233, 536);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.TabIndex = 451;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.BackColor = System.Drawing.SystemColors.Control;
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(308, 536);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.TabIndex = 450;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.BackColor = System.Drawing.SystemColors.Control;
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(533, 536);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.TabIndex = 455;
			this.btnAnular.Text = "Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.BackColor = System.Drawing.SystemColors.Control;
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(608, 536);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 456;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.BackColor = System.Drawing.SystemColors.Control;
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Enabled = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(683, 536);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 457;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Location = new System.Drawing.Point(0, 520);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1256, 8);
			this.groupBox3.TabIndex = 458;
			this.groupBox3.TabStop = false;
			// 
			// btnVer
			// 
			this.btnVer.Enabled = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(1176, 183);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 694;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// lblLocal
			// 
			this.lblLocal.AutoSize = true;
			this.lblLocal.Location = new System.Drawing.Point(16, 11);
			this.lblLocal.Name = "lblLocal";
			this.lblLocal.Size = new System.Drawing.Size(31, 16);
			this.lblLocal.TabIndex = 696;
			this.lblLocal.Text = "Local";
			this.lblLocal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblLocal.Visible = false;
			// 
			// cmbBodega
			// 
			appearance27.FontData.Name = "Tahoma";
			this.cmbBodega.Appearance = appearance27;
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 180;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(250, 22);
			this.cmbBodega.TabIndex = 695;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.Visible = false;
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// lblHabitaciones
			// 
			this.lblHabitaciones.AutoSize = true;
			this.lblHabitaciones.Location = new System.Drawing.Point(336, 11);
			this.lblHabitaciones.Name = "lblHabitaciones";
			this.lblHabitaciones.Size = new System.Drawing.Size(58, 16);
			this.lblHabitaciones.TabIndex = 698;
			this.lblHabitaciones.Text = "Habitacion";
			this.lblHabitaciones.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblHabitaciones.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(864, 115);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(75, 17);
			this.label4.TabIndex = 700;
			this.label4.Text = "Fecha Ingreso";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaIngreso
			// 
			appearance28.FontData.Name = "Tahoma";
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaIngreso.Appearance = appearance28;
			this.dtFechaIngreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaIngreso.DateButtons.Add(dateButton2);
			this.dtFechaIngreso.Enabled = false;
			this.dtFechaIngreso.Format = "dd/MM/yyyy HH:mm";
			this.dtFechaIngreso.Location = new System.Drawing.Point(952, 113);
			this.dtFechaIngreso.Name = "dtFechaIngreso";
			this.dtFechaIngreso.NonAutoSizeHeight = 23;
			this.dtFechaIngreso.Size = new System.Drawing.Size(144, 21);
			this.dtFechaIngreso.SpinButtonsVisible = true;
			this.dtFechaIngreso.TabIndex = 699;
			this.dtFechaIngreso.Value = ((object)(resources.GetObject("dtFechaIngreso.Value")));
			this.dtFechaIngreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaIngreso_KeyDown);
			this.dtFechaIngreso.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaIngreso_BeforeDropDown);
			this.dtFechaIngreso.ValueChanged += new System.EventHandler(this.dtFechaIngreso_ValueChanged);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(864, 147);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(67, 17);
			this.label5.TabIndex = 702;
			this.label5.Text = "Fecha Salida";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaSalida
			// 
			appearance29.FontData.Name = "Tahoma";
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaSalida.Appearance = appearance29;
			this.dtFechaSalida.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaSalida.DateButtons.Add(dateButton3);
			this.dtFechaSalida.Enabled = false;
			this.dtFechaSalida.Format = "dd/MM/yyyy HH:mm";
			this.dtFechaSalida.Location = new System.Drawing.Point(952, 145);
			this.dtFechaSalida.Name = "dtFechaSalida";
			this.dtFechaSalida.NonAutoSizeHeight = 23;
			this.dtFechaSalida.Size = new System.Drawing.Size(144, 21);
			this.dtFechaSalida.SpinButtonsVisible = true;
			this.dtFechaSalida.TabIndex = 701;
			this.dtFechaSalida.Value = ((object)(resources.GetObject("dtFechaSalida.Value")));
			this.dtFechaSalida.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaSalida_BeforeDropDown);
			this.dtFechaSalida.ValueChanged += new System.EventHandler(this.dtFechaSalida_ValueChanged);
			// 
			// optTurnoIngreso
			// 
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTurnoIngreso.Appearance = appearance30;
			this.optTurnoIngreso.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTurnoIngreso.Enabled = false;
			this.optTurnoIngreso.ItemAppearance = appearance31;
			this.optTurnoIngreso.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Turno 1";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Turno 2";
			this.optTurnoIngreso.Items.Add(valueListItem1);
			this.optTurnoIngreso.Items.Add(valueListItem2);
			this.optTurnoIngreso.ItemSpacingVertical = 10;
			this.optTurnoIngreso.Location = new System.Drawing.Point(1104, 111);
			this.optTurnoIngreso.Name = "optTurnoIngreso";
			this.optTurnoIngreso.Size = new System.Drawing.Size(144, 24);
			this.optTurnoIngreso.TabIndex = 703;
			// 
			// optTurnoSalida
			// 
			appearance32.ForeColor = System.Drawing.Color.Black;
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTurnoSalida.Appearance = appearance32;
			this.optTurnoSalida.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTurnoSalida.Enabled = false;
			this.optTurnoSalida.ItemAppearance = appearance33;
			this.optTurnoSalida.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Turno 1";
			valueListItem4.DataValue = 2;
			valueListItem4.DisplayText = "Turno 2";
			this.optTurnoSalida.Items.Add(valueListItem3);
			this.optTurnoSalida.Items.Add(valueListItem4);
			this.optTurnoSalida.ItemSpacingVertical = 10;
			this.optTurnoSalida.Location = new System.Drawing.Point(1104, 143);
			this.optTurnoSalida.Name = "optTurnoSalida";
			this.optTurnoSalida.Size = new System.Drawing.Size(144, 24);
			this.optTurnoSalida.TabIndex = 704;
			this.optTurnoSalida.ValueChanged += new System.EventHandler(this.optTurnoSalida_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// cmbHabitaciones
			// 
			appearance34.FontData.Name = "Tahoma";
			this.cmbHabitaciones.Appearance = appearance34;
			this.cmbHabitaciones.CausesValidation = false;
			this.cmbHabitaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbHabitaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbHabitaciones.DataSource = this.ultraDataSource1;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 250;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbHabitaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbHabitaciones.DisplayMember = "Nombre";
			this.cmbHabitaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbHabitaciones.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbHabitaciones.Enabled = false;
			this.cmbHabitaciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbHabitaciones.Location = new System.Drawing.Point(408, 8);
			this.cmbHabitaciones.MaxDropDownItems = 30;
			this.cmbHabitaciones.Name = "cmbHabitaciones";
			this.cmbHabitaciones.Size = new System.Drawing.Size(250, 22);
			this.cmbHabitaciones.TabIndex = 705;
			this.cmbHabitaciones.ValueMember = "idHabitacion";
			this.cmbHabitaciones.Visible = false;
			this.cmbHabitaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbHabitaciones_KeyDown);
			this.cmbHabitaciones.ValueChanged += new System.EventHandler(this.cmbHabitaciones_ValueChanged);
			this.cmbHabitaciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbHabitaciones_InitializeLayout);
			// 
			// btnTemporal
			// 
			this.btnTemporal.CausesValidation = false;
			this.btnTemporal.Enabled = false;
			this.btnTemporal.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTemporal.ImageIndex = 66;
			this.btnTemporal.Location = new System.Drawing.Point(383, 536);
			this.btnTemporal.Name = "btnTemporal";
			this.btnTemporal.TabIndex = 706;
			this.btnTemporal.Text = "Temporal";
			this.btnTemporal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnTemporal.Click += new System.EventHandler(this.btnTemporal_Click);
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.AcceptsTab = true;
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance35;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(64, 384);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(416, 120);
			this.txtNotas.TabIndex = 707;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(8, 412);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(33, 17);
			this.label6.TabIndex = 708;
			this.label6.Text = "Notas";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox4
			// 
			this.groupBox4.Location = new System.Drawing.Point(552, 440);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(272, 8);
			this.groupBox4.TabIndex = 709;
			this.groupBox4.TabStop = false;
			// 
			// lblIngresos
			// 
			this.lblIngresos.AutoSize = true;
			this.lblIngresos.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIngresos.ForeColor = System.Drawing.Color.Firebrick;
			this.lblIngresos.Location = new System.Drawing.Point(264, 144);
			this.lblIngresos.Name = "lblIngresos";
			this.lblIngresos.Size = new System.Drawing.Size(0, 23);
			this.lblIngresos.TabIndex = 711;
			this.lblIngresos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(632, 80);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(49, 17);
			this.label2.TabIndex = 712;
			this.label2.Text = "Personas";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPersonas
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPersonas.Appearance = appearance36;
			this.txtPersonas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPersonas.Enabled = false;
			this.txtPersonas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPersonas.FormatString = "#0";
			this.txtPersonas.Location = new System.Drawing.Point(696, 80);
			this.txtPersonas.MaskInput = "nnnnnnn";
			this.txtPersonas.MaxValue = 4;
			this.txtPersonas.MinValue = 2;
			this.txtPersonas.Name = "txtPersonas";
			this.txtPersonas.PromptChar = ' ';
			this.txtPersonas.Size = new System.Drawing.Size(27, 22);
			this.txtPersonas.TabIndex = 713;
			this.txtPersonas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPersonas_KeyDown);
			this.txtPersonas.Click += new System.EventHandler(this.txtPersonas_Click);
			this.txtPersonas.Validated += new System.EventHandler(this.txtPersonas_Validated);
			this.txtPersonas.Enter += new System.EventHandler(this.txtPersonas_Enter);
			// 
			// txtEncuesta
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEncuesta.Appearance = appearance37;
			this.txtEncuesta.CausesValidation = false;
			this.txtEncuesta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEncuesta.Enabled = false;
			this.txtEncuesta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEncuesta.Location = new System.Drawing.Point(336, 456);
			this.txtEncuesta.MaxLength = 9;
			this.txtEncuesta.Name = "txtEncuesta";
			this.txtEncuesta.Size = new System.Drawing.Size(136, 22);
			this.txtEncuesta.TabIndex = 714;
			this.txtEncuesta.Visible = false;
			this.txtEncuesta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtEncuesta_KeyPress);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(240, 456);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(82, 17);
			this.label7.TabIndex = 715;
			this.label7.Text = "N° Encuesta P1";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label7.Visible = false;
			// 
			// txtEncuesta2
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEncuesta2.Appearance = appearance38;
			this.txtEncuesta2.CausesValidation = false;
			this.txtEncuesta2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEncuesta2.Enabled = false;
			this.txtEncuesta2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEncuesta2.Location = new System.Drawing.Point(336, 432);
			this.txtEncuesta2.MaxLength = 9;
			this.txtEncuesta2.Name = "txtEncuesta2";
			this.txtEncuesta2.Size = new System.Drawing.Size(136, 22);
			this.txtEncuesta2.TabIndex = 716;
			this.txtEncuesta2.Visible = false;
			this.txtEncuesta2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtEncuesta2_KeyPress);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(240, 432);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(82, 17);
			this.label9.TabIndex = 717;
			this.label9.Text = "N° Encuesta P2";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label9.Visible = false;
			// 
			// btnClaveDescuento
			// 
			this.btnClaveDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnClaveDescuento.CausesValidation = false;
			this.btnClaveDescuento.Enabled = false;
			this.btnClaveDescuento.Location = new System.Drawing.Point(1152, 536);
			this.btnClaveDescuento.Name = "btnClaveDescuento";
			this.btnClaveDescuento.Size = new System.Drawing.Size(72, 23);
			this.btnClaveDescuento.TabIndex = 718;
			this.btnClaveDescuento.Text = "Descuento";
			this.btnClaveDescuento.Click += new System.EventHandler(this.btnClaveDescuento_Click);
			// 
			// btnRefacturacion
			// 
			this.btnRefacturacion.BackColor = System.Drawing.SystemColors.Control;
			this.btnRefacturacion.Enabled = false;
			this.btnRefacturacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnRefacturacion.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnRefacturacion.Image = ((System.Drawing.Image)(resources.GetObject("btnRefacturacion.Image")));
			this.btnRefacturacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRefacturacion.Location = new System.Drawing.Point(832, 536);
			this.btnRefacturacion.Name = "btnRefacturacion";
			this.btnRefacturacion.Size = new System.Drawing.Size(95, 23);
			this.btnRefacturacion.TabIndex = 719;
			this.btnRefacturacion.Text = "Refacturación";
			this.btnRefacturacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRefacturacion.Click += new System.EventHandler(this.btnRefacturacion_Click);
			// 
			// btnDescEmpAut
			// 
			this.btnDescEmpAut.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnDescEmpAut.CausesValidation = false;
			this.btnDescEmpAut.Enabled = false;
			this.btnDescEmpAut.Location = new System.Drawing.Point(1224, 536);
			this.btnDescEmpAut.Name = "btnDescEmpAut";
			this.btnDescEmpAut.Size = new System.Drawing.Size(24, 23);
			this.btnDescEmpAut.TabIndex = 720;
			this.btnDescEmpAut.Text = "D";
			this.btnDescEmpAut.Click += new System.EventHandler(this.btnDescEmpAut_Click);
			// 
			// mnuImprimir
			// 
			this.mnuImprimir.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuImprimirFactura,
																																								this.mnuImprimirRIDE,
																																								this.mnuImprimirComprobanteVenta});
			// 
			// mnuImprimirFactura
			// 
			this.mnuImprimirFactura.Enabled = false;
			this.mnuImprimirFactura.Index = 0;
			this.mnuImprimirFactura.Text = "Factura";
			this.mnuImprimirFactura.Click += new System.EventHandler(this.mnuImprimirFactura_Click);
			// 
			// mnuImprimirRIDE
			// 
			this.mnuImprimirRIDE.Enabled = false;
			this.mnuImprimirRIDE.Index = 1;
			this.mnuImprimirRIDE.Text = "RIDE";
			this.mnuImprimirRIDE.Click += new System.EventHandler(this.mnuImprimirRIDE_Click);
			// 
			// mnuImprimirComprobanteVenta
			// 
			this.mnuImprimirComprobanteVenta.Enabled = false;
			this.mnuImprimirComprobanteVenta.Index = 2;
			this.mnuImprimirComprobanteVenta.Text = "Comprobante Venta";
			this.mnuImprimirComprobanteVenta.Click += new System.EventHandler(this.mnuImprimirComprobanteVenta_Click);
			// 
			// btnMitadPrecio
			// 
			this.btnMitadPrecio.BackColor = System.Drawing.SystemColors.Control;
			this.btnMitadPrecio.CausesValidation = false;
			this.btnMitadPrecio.Enabled = false;
			this.btnMitadPrecio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnMitadPrecio.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnMitadPrecio.Image = ((System.Drawing.Image)(resources.GetObject("btnMitadPrecio.Image")));
			this.btnMitadPrecio.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnMitadPrecio.Location = new System.Drawing.Point(488, 400);
			this.btnMitadPrecio.Name = "btnMitadPrecio";
			this.btnMitadPrecio.Size = new System.Drawing.Size(64, 23);
			this.btnMitadPrecio.TabIndex = 721;
			this.btnMitadPrecio.Text = "Cupon";
			this.btnMitadPrecio.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnMitadPrecio.Click += new System.EventHandler(this.btnCupon_Click);
			// 
			// btnExtras
			// 
			this.btnExtras.BackColor = System.Drawing.SystemColors.Control;
			this.btnExtras.Enabled = false;
			this.btnExtras.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnExtras.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnExtras.Image = ((System.Drawing.Image)(resources.GetObject("btnExtras.Image")));
			this.btnExtras.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExtras.Location = new System.Drawing.Point(936, 536);
			this.btnExtras.Name = "btnExtras";
			this.btnExtras.Size = new System.Drawing.Size(64, 24);
			this.btnExtras.TabIndex = 722;
			this.btnExtras.Text = "Extras";
			this.btnExtras.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExtras.Click += new System.EventHandler(this.btnExtras_Click);
			// 
			// btnCupon
			// 
			this.btnCupon.BackColor = System.Drawing.SystemColors.Control;
			this.btnCupon.CausesValidation = false;
			this.btnCupon.Enabled = false;
			this.btnCupon.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCupon.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCupon.Image = ((System.Drawing.Image)(resources.GetObject("btnCupon.Image")));
			this.btnCupon.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCupon.Location = new System.Drawing.Point(120, 456);
			this.btnCupon.Name = "btnCupon";
			this.btnCupon.Size = new System.Drawing.Size(64, 24);
			this.btnCupon.TabIndex = 723;
			this.btnCupon.Text = "Cupon";
			this.btnCupon.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCupon.Visible = false;
			this.btnCupon.Click += new System.EventHandler(this.btnCupon_Click_1);
			// 
			// chkCupon
			// 
			this.chkCupon.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCupon.Appearance = appearance39;
			this.chkCupon.BackColor = System.Drawing.Color.Transparent;
			this.chkCupon.Enabled = false;
			this.chkCupon.Location = new System.Drawing.Point(872, 72);
			this.chkCupon.Name = "chkCupon";
			this.chkCupon.Size = new System.Drawing.Size(80, 22);
			this.chkCupon.TabIndex = 725;
			this.chkCupon.Text = "Cupón";
			this.chkCupon.Visible = false;
			this.chkCupon.CheckedChanged += new System.EventHandler(this.chkCupon_CheckedChanged);
			// 
			// txtCupon
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCupon.Appearance = appearance40;
			this.txtCupon.CausesValidation = false;
			this.txtCupon.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCupon.Enabled = false;
			this.txtCupon.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCupon.Location = new System.Drawing.Point(952, 72);
			this.txtCupon.MaxLength = 9;
			this.txtCupon.Name = "txtCupon";
			this.txtCupon.Size = new System.Drawing.Size(112, 22);
			this.txtCupon.TabIndex = 726;
			this.txtCupon.Visible = false;
			this.txtCupon.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCupon_KeyDown);
			this.txtCupon.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCupon_KeyPress);
			this.txtCupon.Validating += new System.ComponentModel.CancelEventHandler(this.txtCupon_Validating);
			this.txtCupon.Validated += new System.EventHandler(this.txtCupon_Validated);
			this.txtCupon.ValueChanged += new System.EventHandler(this.txtCupon_ValueChanged);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.Location = new System.Drawing.Point(680, 8);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(89, 17);
			this.label11.TabIndex = 727;
			this.label11.Text = "Tiempo Adicional";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTiempoAdicional
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempoAdicional.Appearance = appearance41;
			this.txtTiempoAdicional.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempoAdicional.Enabled = false;
			this.txtTiempoAdicional.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTiempoAdicional.FormatString = "#0";
			this.txtTiempoAdicional.Location = new System.Drawing.Point(776, 8);
			this.txtTiempoAdicional.MaskInput = "nnnnnnnnn";
			this.txtTiempoAdicional.MaxValue = 24;
			this.txtTiempoAdicional.MinValue = 0;
			this.txtTiempoAdicional.Name = "txtTiempoAdicional";
			this.txtTiempoAdicional.PromptChar = ' ';
			this.txtTiempoAdicional.Size = new System.Drawing.Size(48, 22);
			this.txtTiempoAdicional.TabIndex = 728;
			this.txtTiempoAdicional.ValueChanged += new System.EventHandler(this.txtTiempoAdicional_ValueChanged);
			// 
			// btnTiempoAdicional
			// 
			this.btnTiempoAdicional.Enabled = false;
			this.btnTiempoAdicional.Image = ((System.Drawing.Image)(resources.GetObject("btnTiempoAdicional.Image")));
			this.btnTiempoAdicional.Location = new System.Drawing.Point(831, 8);
			this.btnTiempoAdicional.Name = "btnTiempoAdicional";
			this.btnTiempoAdicional.Size = new System.Drawing.Size(26, 22);
			this.btnTiempoAdicional.TabIndex = 729;
			this.btnTiempoAdicional.Click += new System.EventHandler(this.btnTiempoAdicional_Click);
			// 
			// btnanticiposPendientes
			// 
			this.btnanticiposPendientes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnanticiposPendientes.BackColor = System.Drawing.SystemColors.Control;
			this.btnanticiposPendientes.CausesValidation = false;
			this.btnanticiposPendientes.Enabled = false;
			this.btnanticiposPendientes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnanticiposPendientes.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnanticiposPendientes.Image = ((System.Drawing.Image)(resources.GetObject("btnanticiposPendientes.Image")));
			this.btnanticiposPendientes.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnanticiposPendientes.Location = new System.Drawing.Point(1032, 536);
			this.btnanticiposPendientes.Name = "btnanticiposPendientes";
			this.btnanticiposPendientes.Size = new System.Drawing.Size(120, 22);
			this.btnanticiposPendientes.TabIndex = 730;
			this.btnanticiposPendientes.Text = "Anticipos Vigentes";
			this.btnanticiposPendientes.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnanticiposPendientes.Click += new System.EventHandler(this.btnanticiposPendientes_Click);
			// 
			// chkAbono
			// 
			this.chkAbono.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAbono.Appearance = appearance42;
			this.chkAbono.BackColor = System.Drawing.Color.Transparent;
			this.chkAbono.Enabled = false;
			this.chkAbono.Location = new System.Drawing.Point(80, 432);
			this.chkAbono.Name = "chkAbono";
			this.chkAbono.Size = new System.Drawing.Size(136, 22);
			this.chkAbono.TabIndex = 731;
			this.chkAbono.Text = "Abono Proxima Visita";
			this.chkAbono.Visible = false;
			// 
			// btnOrdenCocina
			// 
			this.btnOrdenCocina.BackColor = System.Drawing.SystemColors.Control;
			this.btnOrdenCocina.CausesValidation = false;
			this.btnOrdenCocina.Enabled = false;
			this.btnOrdenCocina.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnOrdenCocina.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnOrdenCocina.Image = ((System.Drawing.Image)(resources.GetObject("btnOrdenCocina.Image")));
			this.btnOrdenCocina.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnOrdenCocina.Location = new System.Drawing.Point(1112, 72);
			this.btnOrdenCocina.Name = "btnOrdenCocina";
			this.btnOrdenCocina.Size = new System.Drawing.Size(136, 23);
			this.btnOrdenCocina.TabIndex = 732;
			this.btnOrdenCocina.Text = "Orden Cocina";
			this.btnOrdenCocina.Click += new System.EventHandler(this.btnOrdenCocina_Click);
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
			// txtTarjetaVIP
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTarjetaVIP.Appearance = appearance43;
			this.txtTarjetaVIP.CausesValidation = false;
			this.txtTarjetaVIP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTarjetaVIP.Enabled = false;
			this.txtTarjetaVIP.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTarjetaVIP.Location = new System.Drawing.Point(952, 8);
			this.txtTarjetaVIP.MaxLength = 9;
			this.txtTarjetaVIP.Name = "txtTarjetaVIP";
			this.txtTarjetaVIP.Size = new System.Drawing.Size(96, 22);
			this.txtTarjetaVIP.TabIndex = 733;
			this.txtTarjetaVIP.Visible = false;
			this.txtTarjetaVIP.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTarjetaVIP_KeyDown);
			// 
			// chkTarjetaVIP
			// 
			this.chkTarjetaVIP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTarjetaVIP.Appearance = appearance44;
			this.chkTarjetaVIP.BackColor = System.Drawing.Color.Transparent;
			this.chkTarjetaVIP.Enabled = false;
			this.chkTarjetaVIP.Location = new System.Drawing.Point(872, 16);
			this.chkTarjetaVIP.Name = "chkTarjetaVIP";
			this.chkTarjetaVIP.Size = new System.Drawing.Size(80, 22);
			this.chkTarjetaVIP.TabIndex = 734;
			this.chkTarjetaVIP.Text = "Tarjeta VIP";
			this.chkTarjetaVIP.Visible = false;
			this.chkTarjetaVIP.CheckedChanged += new System.EventHandler(this.chkTarjetaVIP_CheckedChanged);
			// 
			// btnInfoTarjetasVIP
			// 
			this.btnInfoTarjetasVIP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnInfoTarjetasVIP.CausesValidation = false;
			this.btnInfoTarjetasVIP.Enabled = false;
			this.btnInfoTarjetasVIP.Image = ((System.Drawing.Image)(resources.GetObject("btnInfoTarjetasVIP.Image")));
			this.btnInfoTarjetasVIP.Location = new System.Drawing.Point(1040, 8);
			this.btnInfoTarjetasVIP.Name = "btnInfoTarjetasVIP";
			this.btnInfoTarjetasVIP.Size = new System.Drawing.Size(24, 22);
			this.btnInfoTarjetasVIP.TabIndex = 735;
			this.btnInfoTarjetasVIP.Visible = false;
			this.btnInfoTarjetasVIP.Click += new System.EventHandler(this.btnInfoTarjetasVIP_Click);
			// 
			// lblPuntosTarjetaVIP
			// 
			this.lblPuntosTarjetaVIP.AutoSize = true;
			this.lblPuntosTarjetaVIP.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPuntosTarjetaVIP.Location = new System.Drawing.Point(864, 40);
			this.lblPuntosTarjetaVIP.Name = "lblPuntosTarjetaVIP";
			this.lblPuntosTarjetaVIP.Size = new System.Drawing.Size(102, 17);
			this.lblPuntosTarjetaVIP.TabIndex = 736;
			this.lblPuntosTarjetaVIP.Text = "Puntos Tarjeta VIP:";
			this.lblPuntosTarjetaVIP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblPuntosTarjetaVIP.Visible = false;
			// 
			// txtPuntosTarjetaVIP
			// 
			appearance45.BackColor = System.Drawing.SystemColors.Control;
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPuntosTarjetaVIP.Appearance = appearance45;
			this.txtPuntosTarjetaVIP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPuntosTarjetaVIP.Enabled = false;
			this.txtPuntosTarjetaVIP.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPuntosTarjetaVIP.FormatString = "";
			this.txtPuntosTarjetaVIP.Location = new System.Drawing.Point(992, 40);
			this.txtPuntosTarjetaVIP.MaskInput = "nnnnnnn";
			this.txtPuntosTarjetaVIP.MaxValue = 999999;
			this.txtPuntosTarjetaVIP.MinValue = 0;
			this.txtPuntosTarjetaVIP.Name = "txtPuntosTarjetaVIP";
			this.txtPuntosTarjetaVIP.PromptChar = ' ';
			this.txtPuntosTarjetaVIP.Size = new System.Drawing.Size(56, 22);
			this.txtPuntosTarjetaVIP.TabIndex = 737;
			this.txtPuntosTarjetaVIP.Visible = false;
			// 
			// btnCanjePuntos
			// 
			this.btnCanjePuntos.BackColor = System.Drawing.SystemColors.Control;
			this.btnCanjePuntos.CausesValidation = false;
			this.btnCanjePuntos.Enabled = false;
			this.btnCanjePuntos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCanjePuntos.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCanjePuntos.Image = ((System.Drawing.Image)(resources.GetObject("btnCanjePuntos.Image")));
			this.btnCanjePuntos.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCanjePuntos.Location = new System.Drawing.Point(1112, 24);
			this.btnCanjePuntos.Name = "btnCanjePuntos";
			this.btnCanjePuntos.Size = new System.Drawing.Size(136, 23);
			this.btnCanjePuntos.TabIndex = 738;
			this.btnCanjePuntos.Text = "Canjeo Puntos";
			this.btnCanjePuntos.Click += new System.EventHandler(this.btnCanjePuntos_Click);
			// 
			// chkVideo
			// 
			this.chkVideo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkVideo.Appearance = appearance46;
			this.chkVideo.BackColor = System.Drawing.Color.Transparent;
			this.chkVideo.Enabled = false;
			this.chkVideo.Location = new System.Drawing.Point(240, 432);
			this.chkVideo.Name = "chkVideo";
			this.chkVideo.Size = new System.Drawing.Size(136, 22);
			this.chkVideo.TabIndex = 739;
			this.chkVideo.Text = "Promoción Video";
			this.chkVideo.Visible = false;
			this.chkVideo.CheckedChanged += new System.EventHandler(this.chkVideo_CheckedChanged);
			// 
			// BtnReImprimirComanda
			// 
			this.BtnReImprimirComanda.BackColor = System.Drawing.SystemColors.Control;
			this.BtnReImprimirComanda.CausesValidation = false;
			this.BtnReImprimirComanda.Enabled = false;
			this.BtnReImprimirComanda.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.BtnReImprimirComanda.Image = ((System.Drawing.Image)(resources.GetObject("BtnReImprimirComanda.Image")));
			this.BtnReImprimirComanda.ImageAlign = System.Drawing.ContentAlignment.BottomLeft;
			this.BtnReImprimirComanda.Location = new System.Drawing.Point(1112, 48);
			this.BtnReImprimirComanda.Name = "BtnReImprimirComanda";
			this.BtnReImprimirComanda.Size = new System.Drawing.Size(136, 23);
			this.BtnReImprimirComanda.TabIndex = 740;
			this.BtnReImprimirComanda.Text = "ReImprimir Comanda";
			this.BtnReImprimirComanda.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.BtnReImprimirComanda.Click += new System.EventHandler(this.BtnReImprimirComanda_Click);
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource2;
			appearance47.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance47;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 3;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 6;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 4;
			ultraGridColumn31.Width = 357;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance48;
			ultraGridColumn32.Format = "";
			ultraGridColumn32.Header.VisiblePosition = 5;
			ultraGridColumn32.PromptChar = ' ';
			ultraGridColumn32.Width = 62;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance49;
			ultraGridColumn33.Format = "#,##0.0000";
			ultraGridColumn33.Header.VisiblePosition = 7;
			ultraGridColumn33.PromptChar = ' ';
			ultraGridColumn33.Width = 96;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance50;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 8;
			ultraGridColumn34.Width = 78;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance51;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 9;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance52;
			ultraGridColumn36.Format = "#,##0.0000";
			ultraGridColumn36.Header.Caption = "% Desc";
			ultraGridColumn36.Header.VisiblePosition = 10;
			ultraGridColumn36.Width = 86;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance53;
			ultraGridColumn37.Format = "#,##0.0000";
			ultraGridColumn37.Header.Caption = "Subtotal";
			ultraGridColumn37.Header.VisiblePosition = 11;
			ultraGridColumn37.PromptChar = ' ';
			ultraGridColumn37.Width = 116;
			ultraGridColumn38.Header.VisiblePosition = 12;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 13;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 14;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 15;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 16;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 90;
			ultraGridBand9.Columns.AddRange(new object[] {
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
																										 ultraGridColumn42});
			ultraGridBand9.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			appearance54.ForeColor = System.Drawing.Color.Black;
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance54;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance55.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance55;
			appearance56.ForeColor = System.Drawing.Color.Black;
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance56;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance57.FontData.BoldAsString = "True";
			appearance57.FontData.Name = "Arial";
			appearance57.FontData.SizeInPoints = 9F;
			appearance57.ForeColor = System.Drawing.Color.White;
			appearance57.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance57;
			appearance58.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance58.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance58;
			appearance59.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance59.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance59.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance59;
			appearance60.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance60.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance60.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance60;
			this.uGridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 184);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(816, 180);
			this.uGridDetalle.TabIndex = 252;
			this.uGridDetalle.AfterRowsDeleted += new System.EventHandler(this.uGridDetalle_AfterRowsDeleted);
			this.uGridDetalle.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDetalle_BeforeCellUpdate);
			this.uGridDetalle.BeforeRowUpdate += new Infragistics.Win.UltraWinGrid.CancelableRowEventHandler(this.uGridDetalle_BeforeRowUpdate);
			this.uGridDetalle.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalle_BeforeRowsDeleted);
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
			// 
			// btnReturnPass
			// 
			this.btnReturnPass.BackColor = System.Drawing.SystemColors.Control;
			this.btnReturnPass.Enabled = false;
			this.btnReturnPass.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnReturnPass.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnReturnPass.Image = ((System.Drawing.Image)(resources.GetObject("btnReturnPass.Image")));
			this.btnReturnPass.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnReturnPass.Location = new System.Drawing.Point(864, 8);
			this.btnReturnPass.Name = "btnReturnPass";
			this.btnReturnPass.Size = new System.Drawing.Size(216, 56);
			this.btnReturnPass.TabIndex = 741;
			this.btnReturnPass.Text = "RETURN PASS";
			this.btnReturnPass.Click += new System.EventHandler(this.btnReturnPass_Click);
			// 
			// lblSaldoFactura
			// 
			this.lblSaldoFactura.AutoSize = true;
			this.lblSaldoFactura.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSaldoFactura.ForeColor = System.Drawing.Color.Firebrick;
			this.lblSaldoFactura.Location = new System.Drawing.Point(560, 488);
			this.lblSaldoFactura.Name = "lblSaldoFactura";
			this.lblSaldoFactura.Size = new System.Drawing.Size(0, 23);
			this.lblSaldoFactura.TabIndex = 742;
			// 
			// frmFacturacionHabitaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1255, 565);
			this.Controls.Add(this.lblSaldoFactura);
			this.Controls.Add(this.btnReturnPass);
			this.Controls.Add(this.BtnReImprimirComanda);
			this.Controls.Add(this.chkVideo);
			this.Controls.Add(this.btnCanjePuntos);
			this.Controls.Add(this.txtPuntosTarjetaVIP);
			this.Controls.Add(this.lblPuntosTarjetaVIP);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.lblHabitaciones);
			this.Controls.Add(this.lblLocal);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.btnInfoTarjetasVIP);
			this.Controls.Add(this.chkTarjetaVIP);
			this.Controls.Add(this.txtTarjetaVIP);
			this.Controls.Add(this.btnOrdenCocina);
			this.Controls.Add(this.chkAbono);
			this.Controls.Add(this.btnanticiposPendientes);
			this.Controls.Add(this.btnTiempoAdicional);
			this.Controls.Add(this.txtTiempoAdicional);
			this.Controls.Add(this.txtCupon);
			this.Controls.Add(this.chkCupon);
			this.Controls.Add(this.btnCupon);
			this.Controls.Add(this.btnExtras);
			this.Controls.Add(this.btnMitadPrecio);
			this.Controls.Add(this.btnDescEmpAut);
			this.Controls.Add(this.btnRefacturacion);
			this.Controls.Add(this.btnClaveDescuento);
			this.Controls.Add(this.txtEncuesta2);
			this.Controls.Add(this.txtEncuesta);
			this.Controls.Add(this.txtPersonas);
			this.Controls.Add(this.lblIngresos);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.btnTemporal);
			this.Controls.Add(this.cmbHabitaciones);
			this.Controls.Add(this.optTurnoSalida);
			this.Controls.Add(this.optTurnoIngreso);
			this.Controls.Add(this.dtFechaSalida);
			this.Controls.Add(this.dtFechaIngreso);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtUnidad);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.cmbCajero);
			this.Controls.Add(this.txtTiempo);
			this.Controls.Add(this.uGridDetalle);
			this.Controls.Add(this.btnCliente);
			this.Controls.Add(this.cmbCobrador);
			this.Controls.Add(this.cmbGuardia);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.uGridArticulos);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.lblHabitacion);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmFacturacionHabitaciones";
			this.ShowInTaskbar = false;
			this.Text = "Facturación";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFacturacionHabitaciones_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmFacturacionHabitaciones_Closing);
			this.Load += new System.EventHandler(this.frmFacturacionHabitaciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGuardia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridArticulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUnidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTurnoIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTurnoSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHabitaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPersonas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEncuesta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEncuesta2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCupon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoAdicional)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTarjetaVIP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPuntosTarjetaVIP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		bool bNuevo = false;
		bool bEdicion = false;

		int idCompra = 0;
		int idReserva = 0;
		int iEstado = 0;
		int idAsiento = 0;
		int idCliente = 0;
		int idProyecto = 0;
		int idBloqueaTransacciones = 0;
		int idCuponWeb = 0;
				
		bool vValidaRucEmpleado = false;
		bool vValidaFPAgo = false;
		bool bNumeracionAutomatica = false;
		bool bImpreso = false;		
		bool bAnulado = false;
		bool bBtnDesEmpleado = false;
		bool bBtnDesFactura = false;
		bool bEliminaAlValidar = false;
		bool bVal30min = false; // Seguridad
		bool bValRefacturacion = false;
		decimal dTotalFactura = 0.00m;
		decimal dAnticipo = 0.00m;
		bool bUsoAnticipo = false;	
		private bool bActualiza = true;
		int vValFPago = 0;
		int iOrigen = 0;
		DateTime dtFechaPromocion = DateTime.Today;
		decimal dPorcentajeIVA = 0.00m;
		decimal dPorcentajeIVADec = 0.00m;
		int idIngreso = 0;
		bool bNuevoIngreso = false;
		int iIngresos = 0;
		int countFP = 0;
		int iCountFP = 0;
		bool bCuponActivo = false;
		int iPuntos = 0;
		int idOriginal = 0;
		int CoutEditar = 0;
		bool ImprimeRefacturacion = false;
		bool ImprimeConDEs = false;
		string sCorreo = "";
		int iValCobrador = 0;
		bool ValNuevaFactura = false;
		int idCompraDest = 0;
		bool bValHoraDes = false;
		bool TempAnticipo = false;
		bool bRFAnticipo = false;
		int idFe_EstadoDocumento = 0;

		#region Variables Anticipo
		int iCliente = 0;
		int IdAnticipo = 0;
		decimal dSaldo = 0.00m;
		string sMensajeAnticipo = "";
		string sDocumento = "";
		#endregion Variables Anticipo

		int iCompraOriginal = 0;
		int iCompraDividida = 0;
		bool bImpresionDivision = false;
		bool bImpresionPuntoVenta = false;
		bool bFacturaTiempo = false;
		bool bActivaBoton = false;
		bool bImprimirModificacion = false;
		int iActivoMitadPrecion = 0;
		int iActivaChkCupon = 0;
		int iActivoCuponBodega = 0;
		bool bCuponMP = false;
		bool bGrabaComoTemporal = false;

		#region Variables Facturacion Extra
		bool bFacturaExtra = false;
		int iCompraFacturaExtra = 0;
		#endregion Variables Facturacion Extra

		#region Actualizacion Tiempo Cuadricula
		int iTipoModificacion = 0;
		#endregion Actualizacion Tiempo Cuadricula

		#region Toque de Queda
		bool bToqueQueda = false;
		bool bValidaPromocionMensaje = false;
		#endregion Toque de Queda

		#region Valida Anticipos Cliente
		decimal dValorPromActiva = 0.00m;
		bool bUsoPromActiva = false;

		decimal dAnticipoCliente = 0.00m;
		bool bUsoAnticipoCliente = false;
		int iTipoCLiente = 2;
		#endregion Valida Anticipos Cliente

		#region Validacion Gif Card
		bool bUsoGiftCard = false;
		decimal dGiftCard = 0.00m;
		int idHistorialTarjetasVIP = 0;
		#endregion Validacion Gif Card

		#region Variables Comandas
		int iExisComanda = 0;
		int idComanda = 0;
		int iValComandaExis = 0;
		int iValComandaMins = 0;
		int iComandaImp = 0;
		int idComandaOrden = 0;
		int idComandaAutDesc = 0;
		int iExisComandaAuDesc = 0;
		#endregion Variables Comandas

		#region Variables Cupon
		bool bUsoCupon = false;
		decimal dValorCupon = 0.00m;
		#endregion Variables Cupon

		#region Variables Tarjetas VIP
		int iActivaTajetaVIP = 0;
		int iTotalPuntosTarjetaVIP = 0;
		#endregion Variables Tarjetas VIP

		#region Promocion Temporal Video
		bool bPromoVideo = false;
		bool bRFVideoPromocion = false;
		bool bPromoHoras = false;
		#endregion Promocion Temporal Video

		int ValTiempoAdicionalPromo = 0;
		
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void Consulta(int idCompraC)
		{
			#region Maestro Detalle
			int iFormaPago = 0;

			#region Consulta Registro
			string sSQL = string.Format("Exec ConsultaIndividualFacturas {0}", idCompraC);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{				
				idCompra = idCompraC;
				idReserva = dr.GetInt32(2);
				iEstado = dr.GetInt32(3);
				bAnulado = dr.GetBoolean(4);
				bImpreso = dr.GetBoolean(5);
				this.lblEstado.Text = dr.GetString(6);
				iBodega = dr.GetInt32(7);
				idProyecto = dr.GetInt32(8);
				idHabitacion = dr.GetInt32(9);
				if (!dtfechaFact) this.dtFecha.Value = dr.GetDateTime(10);
				this.cmbGuardia.Value = dr.GetInt32(11);
				this.txtNumero.Text = dr.GetString(12);
				this.txtTiempo.Value = dr.GetInt32(13); 
				this.dtFechaIngreso.Value = dr.GetDateTime(14);
				this.optTurnoIngreso.Value = dr.GetInt32(15);
				idCliente = dr.GetInt32(16);
				this.cmbTipoRuc.Value = dr.GetInt32(17);
				this.txtRuc.Text = dr.GetString(18);
				this.txtNombre.Text = dr.GetString(19);
				if (dr.GetValue(20) != System.DBNull.Value) this.dtFechaSalida.Value = dr.GetDateTime(20);
				this.optTurnoSalida.Value = dr.GetInt32(21);
				this.cmbCobrador.Value = dr.GetInt32(22);
				this.cmbCajero.Value = dr.GetInt32(23);
				this.txtUnidad.Value = dr.GetInt32(24);
				iFormaPago = dr.GetInt32(25);
				this.txtNotas.Text = dr.GetString(26);
				this.txtIva.Value = dr.GetDecimal(29);
				this.txtDescIvaTotal.Value = dr.GetDecimal(30);
				this.txtIvaTotal.Value = dr.GetDecimal(31);
				this.txtTotal.Value = dr.GetDecimal(32);
				idAsiento = dr.GetInt32(33);
				iOrigen = dr.GetInt32(34); 
				this.txtPersonas.Value = dr.GetInt32(35);
				if (dr.GetValue(36) != System.DBNull.Value) this.txtEncuesta.Text = dr.GetString(36);
				if (dr.GetValue(37) != System.DBNull.Value) this.txtEncuesta2.Text = dr.GetString(37);				
				iTarifaExpress = dr.GetInt32(38);
				idFe_EstadoDocumento = dr.GetInt32(39);
				this.txtTiempoAdicional.Value = dr.GetInt32(40);
				if(bMenuCuadricula) ValTiempoAdicionalPromo = dr.GetInt32(40);
				bToqueQueda = dr.GetBoolean(41);
			}
			dr.Close();			

			if (dtfechaFact) this.dtFecha.Value = dtFechaRef; 

			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleFacturas {0}", idCompra));
			#endregion Consulta Registro

			//			DateTime ValFecha = Convert.ToDateTime(this.dtFechaIngreso.Value);
			//			if(this.dtFecha.Value == System.DBNull.Value) this.dtFecha.Value = ValFecha;
			
			if (bValRefacturacion) dTotalFactura = Convert.ToDecimal(this.txtTotal.Value);

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));
			this.cmbBodega.Value = iBodega;
			this.cmbFormaPago.Value = iFormaPago;
			this.cmbHabitaciones.Value = idHabitacion;
			idGuardia = Convert.ToInt32(this.cmbGuardia.Value);
			iValCobrador = Convert.ToInt32(this.cmbCobrador.Value);

			if((int)this.cmbFormaPago.Value == 29 || (int)this.cmbFormaPago.Value == 30)
				this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFPagoCarga 3");

			this.mnuImprimirFactura.Enabled = true;

			bImpresionPuntoVenta = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ISNULL (ImpresionPV, 0) From SeteoF"), true);
			
			if (bImpresionPuntoVenta)
				this.mnuImprimirComprobanteVenta.Enabled = true;			

			iIngresos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Ingresos Where idCompra = {0}", idCompra));

			decimal saldoFactura = Funcion.decEscalarSQL(cdsSeteoF,	string.Format("Exec CalculaSaldoFacturas {0}", idCompra));

			this.lblSaldoFactura.Text = "SALDO A COBRAR :          " + saldoFactura.ToString("0.00");

			this.lblIngresos.Text = "INGRESOS : " + iIngresos.ToString();

			string[] sPrmCW = new string[2];
			sPrmCW = Facturacion.sEditarCuponWeb(idCompra, cdsSeteoF);
			idCuponWeb = int.Parse(sPrmCW[0]);

			if (idFe_EstadoDocumento == 2)
			{
				this.btnImprimir.Enabled = true;
				this.mnuImprimirRIDE.Enabled = true;							
			}

			#region Promocion Temporal Video
			//			bool bValPromocionVideo = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ConsultaPromocionVideoIndividual {0}", idCompra), true);
			//			if(bValPromocionVideo)
			//			{
			//				this.chkVideo.Checked = true;
			//			}
			//			else
			//			{
			//				this.chkVideo.Checked = false;
			//			}
			#endregion Promocion Temporal Video
		
			if (!bMenuCuadricula)
			{
				#region Habilita Controles - permisos de usuario
				if (!bAnulado)
				{
					if (iEstado != 6)
					{
						if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
						{
							if (miAcceso.Editar) this.btnEditar.Enabled = true;
							
							if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
							
							if (miAcceso.BDesEmpleado) bBtnDesEmpleado = true;
							if (miAcceso.BAnulaSoloDelDia && !bAnulado)
							{
								if ((DateTime)this.dtFecha.Value == DateTime.Today)
									this.btnAnular.Enabled = true;								
							}							
							if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
							if (miAcceso.BFacturaExtras) this.btnExtras.Enabled = true;
							if (miAcceso.BRefacturacion) 
							{
								bVal30min = true;
								this.btnRefacturacion.Enabled = true;
							}
						}
					}
					else
					{
						if (miAcceso.Editar && (int)this.cmbFormaPago.Value != 6) this.btnEditar.Enabled = true;
						if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
						if (miAcceso.BAnulaSoloDelDia && !bAnulado)
							if ((DateTime)this.dtFecha.Value == DateTime.Today) this.btnAnular.Enabled = true;
						if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;						
						if (miAcceso.BAnulaTiempo)
							this.ValidaTiempoAnulacion();
						else this.btnAnular.Enabled = true;
					}

					#region Impresion
					if (iEstado != 6)
					{
						if (!bImpreso) if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
						
						if (bImpreso) if (miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
					}									
					#endregion Impresion
			
				}				
				#endregion Habilita Controles - permisos de usuario

				this.btnCancelar.Enabled = true;				
			}
			else
			{
				if (!bAnulado)
				{
					if (iEstado == 6)
					{
						if (miAcceso.BDesModifica) bBtnDesFactura = true;
						if (miAcceso.BtnDescuento) this.btnClaveDescuento.Enabled = true;
						this.btnOrdenCocina.Enabled = true;					
					}
				}
			}
			#endregion Maestro Detalle			
		}
		
		bool dtfechaFact = false;
		private bool bFormCargado = false;
		private int idBodegaAnterior = -1;
		int iEstadoHabitacion = 0;
		bool bRegistroCobro = false;
		bool bAnticipoYaAbierto = false;

		private void frmFacturacionHabitaciones_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0703");

			if (!Funcion.Permiso("118", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Facturacion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BDescuento) this.uGridDetalle.DisplayLayout.Bands[0].Columns["DescuentoPorc"].CellActivation = Activation.AllowEdit;
			
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			
			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga 1");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));

			if (this.cmbBodega.Value != null)
				idBodegaAnterior = Convert.ToInt32(this.cmbBodega.Value);

			if (!Funcion.Permiso("118", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Ventas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (!bMenuCuadricula)
			{
				this.lblLocal.Visible = true;
				this.cmbBodega.Visible = true;
			
				this.lblHabitaciones.Visible = true;
				this.cmbHabitaciones.Visible = true;
				
				FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);
				FuncionesProcedimientos.EstadoGrids(false, this.uGridArticulos);

				this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFPagoCarga 1");
				vValFPago = 1;
				vValidaFPAgo = true;
								
				string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
				iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iBodega > 0) this.cmbBodega.Value = iBodega; 							
				
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				this.btnSalir.Enabled = true;

				if (idCompraV > 0) this.Consulta(idCompraV);
			}

			if (bMenuCuadricula)
			{
				this.StartPosition = FormStartPosition.CenterScreen;
				this.ControlBox = false;
				this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
				
				this.lblHabitacion.Visible = true;
				this.lblHabitacion.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec LocalHabitacionTematica {0}", idHabitacion));

				this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFPagoCarga 2");

				vValidaFPAgo = true;
				vValFPago = 2;
				this.btnClaveDescuento.Visible = true;
				this.BtnReImprimirComanda.Enabled= true;
				string sSQL = string.Format("Select ISNULL(Ocupado, 0) From Habitaciones Where idHabitacion = {0}", idHabitacion);
				iEstadoHabitacion = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (bRegistroCobro && idCompraV > 0)
				{
					ValidaTiempo(idCompraV);

					this.Consulta(idCompraV);

					ValidaPromoReturnPass();

					this.btnEditar_Click(sender, e);
				}
				else
				{
					if (iEstadoHabitacion == 1 || iEstadoHabitacion == 6) this.btnNuevo_Click(sender, e);
					else if ((iEstadoHabitacion == 2 || iEstadoHabitacion == 10 || iEstadoHabitacion == 12) && idCompraV > 0)
					{
						ValidaTiempo(idCompraV);

						this.Consulta(idCompraV);

						ValidaPromoReturnPass();

						this.btnEditar_Click(sender, e);
					}
				}

				int c = (int)this.txtTiempo.Value;
				
				int idBodegaHabitacion = Funcion.iEscalarSQL(cdsSeteoF,
					string.Format("SELECT Bodega FROM Habitaciones WHERE idHabitacion = {0}", idHabitacion));

				this.cmbHabitaciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF,
					string.Format("SELECT idHabitacion, Nombre FROM Habitaciones WHERE Bodega = {0} ORDER BY NumeroHabitacion", idBodegaHabitacion));

				#region Personal De Caja
				DateTime fecha = DateTime.Today;
				FuncionesProcedimientos.ListaPersonalFacturas(this.cmbGuardia, fecha, 11, cdsSeteoF);
				FuncionesProcedimientos.ListaPersonalFacturas(this.cmbCobrador, fecha, 6, cdsSeteoF);
				FuncionesProcedimientos.ListaPersonalFacturas(this.cmbCajero, fecha, 4, cdsSeteoF);
        
				DateTime fechaSeleccionada = DateTime.Today;

				if (fechaSeleccionada.Date == DateTime.Today)
				{
					int turnoActual;
					TimeSpan horaActual = DateTime.Now.TimeOfDay;

					if (horaActual >= TimeSpan.FromHours(7) && horaActual < TimeSpan.FromHours(18))
						turnoActual = 1;
					else
						turnoActual = 2;

					DateTime fechaConsulta = DateTime.Today;
			
					if (turnoActual == 2 && horaActual < TimeSpan.FromHours(7))					
						fechaConsulta = fechaConsulta.AddDays(-1);					

					bool bFaltanResponsables = true;

					string sSQLTurno = string.Format("Exec FacturacionPersonalCaja '{0}', {1}, {2}", fechaConsulta.ToString("yyyy-MM-dd"), turnoActual, idBodegaHabitacion);
					SqlDataReader drTurno = Funcion.rEscalarSQL(cdsSeteoF, sSQLTurno, true);
					
					if (drTurno.Read())
					{
						int idGuardia = Convert.ToInt32(drTurno.GetValue(0));
						int idCobrador = Convert.ToInt32(drTurno.GetValue(1));
						int idCajero = Convert.ToInt32(drTurno.GetValue(2));

						if (idGuardia > 0 && idCobrador > 0 && idCajero > 0)
						{
							bFaltanResponsables = false;

							this.cmbGuardia.Value = idGuardia;
							this.cmbCobrador.Value = idCobrador;
							this.cmbCajero.Value = idCajero;

							this.cmbGuardia.Enabled = false;
							this.cmbCobrador.Enabled = false;
							this.cmbCajero.Enabled = false;
						}
					}
					drTurno.Close();				

					if (bFaltanResponsables)
					{
						MessageBox.Show("Aún no se han seleccionado responsables para la caja.",
							"Responsables incompletos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						this.btnCancelar_Click(sender, e);
						this.BeginInvoke(new MethodInvoker(UnloadMe)); 
						return;
					}
				}	
				#endregion Personal De Caja
			}
			
			bFormCargado = true;
			if (bMenuCuadricula && bRegistroCobro && idCompraV > 0 && !bServicioPagado)
			{
				this.BeginInvoke(new MethodInvoker(AbrirAnticipoRegistroCobro));
			}
			if (bMenuCuadricula && bVerificacionPago && idCompraV > 0)
			{
				this.BeginInvoke(new MethodInvoker(AbrirVerificacionPago));
			}
		}

		private void ValidaTiempo(int idCompraC)
		{
			int TiempoEstadia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Jornadas from compra where idCompra = {0}", idCompraC));
			
			if (TiempoEstadia > 8)
				this.txtTiempo.MaxValue = 12;
		}

		private void Total()
		{		
			decimal dDescuentoLineal = 0.00m;			
						
			this.txtIva.Value = 0.0000m;			
			this.txtDescIvaTotal.Value = 0.00m;			
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;
			
			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
			{
				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value && dr.Cells["DescuentoPorc"].Value != DBNull.Value)
				{
					dr.Cells["SubTotal"].Value = Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value);
					dr.Cells["DescuentoArt"].Value = (Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value))/100.00m;

					this.txtIva.Value = Convert.ToDecimal(this.txtIva.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value);
					this.txtDescIvaTotal.Value = Convert.ToDecimal(this.txtDescIvaTotal.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value);
				}
			}

			this.txtIvaTotal.Value = (Convert.ToDecimal(this.txtIva.Value) - Convert.ToDecimal(this.txtDescIvaTotal.Value)) * dPorcentajeIVA / 100.00m;

			this.txtIva.Value = Math.Round(Convert.ToDecimal(this.txtIva.Value), 2);

			this.txtTotal.Value = 
				Convert.ToDecimal(this.txtIva.Value) - 
				Convert.ToDecimal(this.txtDescIvaTotal.Value) + 
				Convert.ToDecimal(this.txtIvaTotal.Value);			
			#endregion Desplazamiento de grilla
		}

		private void ValidaCargaReservas()
		{
			idReserva = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Numero_Reserva {0}", idHabitacion));

			if (idReserva > 0 && DialogResult.Yes == MessageBox.Show("La Habitación esta reservada\n\n¿Desea aplicar la reservación?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DetalleFacturaReservaciones {0}, {1}, {2}, {3}, {4}", idReserva, 3422, idHabitacion, 0, 1));
				this.txtTiempo.MaxValue = 8;
				this.txtTiempo.Value = 5;
			}
			else 
			{
				bool bPromocionDiezHoras = Funcion.bEscalarSQL(cdsSeteoF, "Exec ValidaPromocionDiezHoras", true);

				if (!bPromocionDiezHoras) 
				{					
					this.CargaArticuloHabitacion();
				}
				else if (DialogResult.Yes == MessageBox.Show("¿Desea aplicar la promocion de las 12 horas? Incluye Desayuno", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				{
					this.CargaHabitacionPromocionDiezHoras();
					//this.txtTiempo.Value = 12;
					bPromoHoras = true;
				}
				else 
				{
					//this.CargaArticuloHabitacion();
					bPromoHoras = false;
					if(bMenuCuadricula)
					{
						//if (DialogResult.Yes == MessageBox.Show("¿Desea Cargar el Servicio de hospedaje?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
						//CargaHabitacion(0, 1);
						this.CargaArticuloHabitacion();
						CargarTiempoEstadiaDesdeHabitacion();
					}
				}
			}
		}
		bool cambiacantidad = false;
		private void CargaHabitacion(int idPromocionD, int iCantidad)
		{
			int iFila = uGridDetalle.Rows.Count - 1;
			
			string sSQL = string.Format("Exec DetalleFacturaReservaciones {0}, {1}, {2}, {3}, {4}", 
				0, 3422, (int)this.cmbHabitaciones.Value, idPromocionD, iTarifaExpress);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);

			while (drC.Read())
			{
				iFila++;
				uGridDetalle.DisplayLayout.Bands[0].AddNew();	
				//idArticulo Codigo Articulo, Cantidad, Precio, Impuesto, DescuentoArt, DescuentoPorc, SubTotal, Tipo, Principal, Bloqueado, idCombo
				#region Variables
				int idArtProm = drC.GetInt32(2);
				string sCodProm = drC.GetString(3); 
				string sArtProm = drC.GetString(4);					
				//int iCantidad = drC.GetInt32(5);
				decimal dPrecProm = drC.GetDecimal(6);
				decimal dImpProm = drC.GetDecimal(7);
				decimal dDesArtProm = drC.GetDecimal(8);
				decimal dDesPorProm = drC.GetDecimal(9);
				decimal dSubtotal = drC.GetDecimal(10);					
				int iTipo = drC.GetInt32(11);
				bool bPrincipal = drC.GetBoolean(12);
				bool bBloqueado = drC.GetBoolean(13);
				int idCombo = drC.GetInt32(14);
				#endregion Variables
					
				#region Asigna valores en el grid
				uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;
				uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = idArtProm;
				uGridDetalle.Rows[iFila].Cells["Codigo"].Value = sCodProm;					
				uGridDetalle.Rows[iFila].Cells["Articulo"].Value = sArtProm;
				uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = iCantidad;
				uGridDetalle.Rows[iFila].Cells["Precio"].Value = dPrecProm;
				uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
				uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
				uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
				uGridDetalle.Rows[iFila].Cells["SubTotal"].Value = dPrecProm;
				uGridDetalle.Rows[iFila].Cells["Principal"].Value = bPrincipal;
				uGridDetalle.Rows[iFila].Cells["Tipo"].Value = iTipo;
				uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = bBloqueado;
				uGridDetalle.Rows[iFila].Cells["idCombo"].Value = idCombo;
				uGridDetalle.Rows[iFila].Cells["Entregado"].Value = false;						
				#endregion Asigna valores en el grid
			} 
			drC.Close();
			cambiacantidad = true;
		}

		private void CargaHabitacionPromocionDiezHoras()
		{
			int iFila = uGridDetalle.Rows.Count - 1;
			
			string sSQL = string.Format("Exec DetalleFacturaPromocionPorHoras {0}", 
				(int)this.cmbHabitaciones.Value);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);

			while (drC.Read())
			{
				iFila++;
				uGridDetalle.DisplayLayout.Bands[0].AddNew();	
				//idArticulo Codigo Articulo, Cantidad, Precio, Impuesto, DescuentoArt, DescuentoPorc, SubTotal, Tipo, Principal, Bloqueado, idCombo
				#region Variables
				int idArtProm = drC.GetInt32(2);
				string sCodProm = drC.GetString(3); 
				string sArtProm = drC.GetString(4);					
				int iCantidad = drC.GetInt32(5);
				decimal dPrecProm = drC.GetDecimal(6);
				decimal dImpProm = drC.GetDecimal(7);
				decimal dDesArtProm = drC.GetDecimal(8);
				decimal dDesPorProm = drC.GetDecimal(9);
				decimal dSubtotal = drC.GetDecimal(10);					
				int iTipo = drC.GetInt32(11);
				bool bPrincipal = drC.GetBoolean(12);
				bool bBloqueado = drC.GetBoolean(13);
				int idCombo = drC.GetInt32(14);
				#endregion Variables
					
				#region Asigna valores en el grid
				uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;
				uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = idArtProm;
				uGridDetalle.Rows[iFila].Cells["Codigo"].Value = sCodProm;					
				uGridDetalle.Rows[iFila].Cells["Articulo"].Value = sArtProm;
				uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = iCantidad;
				uGridDetalle.Rows[iFila].Cells["Precio"].Value = dPrecProm;
				uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
				uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
				uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
				uGridDetalle.Rows[iFila].Cells["SubTotal"].Value = dPrecProm;
				uGridDetalle.Rows[iFila].Cells["Principal"].Value = bPrincipal;
				uGridDetalle.Rows[iFila].Cells["Tipo"].Value = iTipo;
				uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = bBloqueado;
				uGridDetalle.Rows[iFila].Cells["idCombo"].Value = idCombo;										
				uGridDetalle.Rows[iFila].Cells["Entregado"].Value = false;
				#endregion Asigna valores en el grid
			} 
			drC.Close();

			this.txtTiempo.MaxValue = 12; //10
			this.txtTiempo.Value = 12; // 10
			bValidaPromocionMensaje = true;
		}

		private void CargaHabitacionTiempoAdicional(int idArticuloHA)
		{
			int iFila = uGridDetalle.Rows.Count - 1;
			
			string sSQL = string.Format("Exec DetalleFacturaReservaciones {0}, {1}, {2}, {3}, {4}, {5}", 
				0, idArticuloHA, (int)this.cmbHabitaciones.Value, 0, 0, (int)this.txtTiempoAdicional.Value);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);

			while (drC.Read())
			{
				iFila++;
				uGridDetalle.DisplayLayout.Bands[0].AddNew();	
				//idArticulo Codigo Articulo, Cantidad, Precio, Impuesto, DescuentoArt, DescuentoPorc, SubTotal, Tipo, Principal, Bloqueado, idCombo
				#region Variables
				int idArtProm = drC.GetInt32(2);
				string sCodProm = drC.GetString(3); 
				string sArtProm = drC.GetString(4);					
				int iCantidad = drC.GetInt32(5);
				decimal dPrecProm = drC.GetDecimal(6);
				decimal dImpProm = drC.GetDecimal(7);
				decimal dDesArtProm = drC.GetDecimal(8);
				decimal dDesPorProm = drC.GetDecimal(9);
				decimal dSubtotal = drC.GetDecimal(10);					
				int iTipo = drC.GetInt32(11);
				bool bPrincipal = drC.GetBoolean(12);
				bool bBloqueado = drC.GetBoolean(13);
				int idCombo = drC.GetInt32(14);
				#endregion Variables
					
				#region Asigna valores en el grid
				uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;
				uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = idArtProm;
				uGridDetalle.Rows[iFila].Cells["Codigo"].Value = sCodProm;					
				uGridDetalle.Rows[iFila].Cells["Articulo"].Value = sArtProm;
				uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = iCantidad;
				uGridDetalle.Rows[iFila].Cells["Precio"].Value = dPrecProm;
				uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
				uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
				uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
				uGridDetalle.Rows[iFila].Cells["SubTotal"].Value = dSubtotal;
				uGridDetalle.Rows[iFila].Cells["Principal"].Value = bPrincipal;
				uGridDetalle.Rows[iFila].Cells["Tipo"].Value = iTipo;
				uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = bBloqueado;
				uGridDetalle.Rows[iFila].Cells["idCombo"].Value = idCombo;	
				uGridDetalle.Rows[iFila].Cells["Entregado"].Value = false;					
				#endregion Asigna valores en el grid
			} 
			drC.Close();
		}

		private void CargaArticuloHabitacion()
		{
			decimal dDescuento = 0.00m;
			int idPromocionD = 0;
			bool bDescuento = false;
			bool bAdicional = false;
			bool bOpcional = false;
			bool bCupon = false;
			bool bAplicaPromocion = true;
			bool bPromocionVIP = false;
			bool bValidaNumeroIngresos = true;
			bool bValidaArtHosp = false;
			int idArticulo = 3422;
			bool bPromocionPreventa = false;

			#region Carga de Hospedaje
			#region Busca Promociones Activas
			DataSet oDSPA = new DataSet();					
			cdsSeteoF.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection Coneccion = cdsSeteoF.Schema.Connections[0];
			Coneccion = cdsSeteoF.Schema.Connections[0];

			string sSQLPA = string.Format("Exec ValidaDescuentoPromocionPeony {0}, {1}, {2}, {3}, '{4}',{5}", 
				(int)this.cmbBodega.Value, 3422, 2, (int)this.cmbHabitaciones.Value, Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"),bPreventa ? 1 : 0);
			SqlDataAdapter oDAPA = new SqlDataAdapter(sSQLPA, (SqlConnection)Coneccion.DbConnection);
			oDAPA.Fill(oDSPA);
					
			if (oDSPA.Tables[0].Rows.Count == 0 && bValidaNumeroIngresos) 
			{
				CargaHabitacion(idPromocionD, 1);	
				bValidaArtHosp = true;						
			}
			else if (oDSPA.Tables[0].Rows.Count > 0)
			{
				#region Validacion Promociones
				if (MenuLatinium.iNivel != 31)
				{
					using (frmListaPromociones LP = new frmListaPromociones(sSQLPA))
					{				
						if (DialogResult.OK == LP.ShowDialog())
						{
							idPromocionD = (int)LP.ultraGrid1.ActiveRow.Cells["idPromocion"].Value;
							bDescuento = (bool)LP.ultraGrid1.ActiveRow.Cells["SumaDescuento"].Value;
							bAdicional = (bool)LP.ultraGrid1.ActiveRow.Cells["Adicional"].Value;
							bOpcional = (bool)LP.ultraGrid1.ActiveRow.Cells["Opcional"].Value;
							bCupon = (bool)LP.ultraGrid1.ActiveRow.Cells["Cupon"].Value;
							bPromocionPreventa = (bool)LP.ultraGrid1.ActiveRow.Cells["PreVenta"].Value;

							#region Valida Cupones
							if (bCupon)
							{
								string[] sPrmCW = new string[2];
								sPrmCW = Facturacion.sCuponWeb(cdsSeteoF);
								idCuponWeb = int.Parse(sPrmCW[0]);

								if (idCuponWeb == 0) 
								{
									bAplicaPromocion = false;
									idPromocionD = 0;
								}

								if (idCuponWeb > 0) bCuponActivo = true;
							}
							#endregion Valida Cupones

							#region Carga Promociones
							if (bAplicaPromocion)
							{
								if (bPromocionPreventa)
								{
									CargaHabitacion(idPromocionD, 1);
								}
								else
								{
									if (bDescuento)
										CargaHabitacion(idPromocionD, 1);

									if (bAdicional)
									{
										CargaHabitacion(idPromocionD, 1);
										this.CargaArticulosPromocionAdicional(idPromocionD);
									}

									if (bOpcional)
									{
										CargaHabitacion(idPromocionD, 1);

										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
										{	
											if ((bool)dr.Cells["Principal"].Value && (int)dr.Cells["idArticulo"].Value == 3422) 
												CargaArticulosPromocionOpcional(idPromocionD);
										}
									}
								}
							}
							else
							{
								CargaHabitacion(0, 1);
							}
							#endregion Carga Promociones
						}
						else
						{
							CargaHabitacion(0, 1);
						}			
					}
				}
				#endregion Validacion Promociones
			}				
			else			
			{
				if (!bAdicional && !bValidaArtHosp)
				{		
					CargaHabitacion(idPromocionD, 1);
				}
			}
			#endregion Busca Promociones Activas
			#endregion Carga de Hospedaje
			
			Total();
		}

		private void CargaHabitacion(int idArticulo, bool bHabitacion)
		{
			decimal dDescuento = 0.00m;
			int idPromocionD = 0;
			bool bDescuento = false;
			bool bAdicional = false;
			bool bOpcional = false;
			bool bCupon = false;
			bool bAplicaPromocion = true;
			bool bPromocionVIP = false;
			bool bValidaNumeroIngresos = true;

			#region Carga de Hospedaje
			if (bHabitacion && !bServicioPagado)
			{												
				if (idReserva == 0) 
				{
					int iContH = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
						if (dr.Cells["idArticulo"].Value.ToString().Length > 0)
							if ((int)dr.Cells["idArticulo"].Value == 3422) iContH++;

					if (iContH == 0)
					{ 				
						int iCantidad = 0;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDet in this.uGridDetalle.Rows.All)
							if (drDet.Cells["idArticulo"].Value.ToString().Length > 0)
								if ((int)drDet.Cells["idArticulo"].Value == 3422) iCantidad = iCantidad + (int)drDet.Cells["Cantidad"].Value;

						if ((iIngresos - iCantidad) > 0)// && bValidaNumeroIngresos)
						{
							CargaHabitacion(idPromocionD, (iIngresos - iCantidad));							
						}
					}
				}

				Total();

				return;
			}
			#endregion Carga de Hospedaje

			Total();
		}
		private void CargaArticulos(int idArticulo, bool bHabitacion)
		{
			decimal dDescuento = 0.00m;
			int idPromocionD = 0;
			bool bDescuento = false;
			bool bAdicional = false;
			bool bOpcional = false;
			bool bCupon = false;
			bool bAplicaPromocion = true;
			bool bPromocionVIP = false;
			bool bValidaNumeroIngresos = true;

			#region Carga de Hospedaje
			if (bHabitacion && !bServicioPagado)
			{												
				if (idReserva == 0) 
				{
					int iContH = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
						if (dr.Cells["idArticulo"].Value.ToString().Length > 0)
							if ((int)dr.Cells["idArticulo"].Value == 3422) iContH++;

					if (iContH == 0)
					{

						#region Valida Promociones Tarjetas VIP

						#region Tarjetas VIP
						//							using (frmValidaNumeroTarjetasVIP VTI = new frmValidaNumeroTarjetasVIP(sSQL, idCompra))//frmValidaTarjetaIngresos VTI = new frmValidaTarjetaIngresos(true, sSQL))
						//							{
						//								if (DialogResult.OK == VTI.ShowDialog())
						//								{
						//									#region Recorre Grid Promociones Tarjeta VIP
						//									int i = 0;
						//									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drVTI in VTI.uGridDetalle.Rows.All)
						//									{							
						//										if ((bool)drVTI.Cells["Sel"].Value)
						//										{
						//											idPromocionD = 0;
						//											bDescuento = false;
						//											bAdicional = false;
						//											bOpcional = false;
						//
						//											idPromocionD = (int)drVTI.Cells["idPromocion"].Value;
						//											bDescuento = (bool)drVTI.Cells["SumaDescuento"].Value;
						//											bAdicional = (bool)drVTI.Cells["Adicional"].Value;
						//											bOpcional = (bool)drVTI.Cells["Opcional"].Value;
						//
						//											#region Promociones
						//											if (bDescuento)
						//											{
						//												if (i == 0) CargaHabitacion(idPromocionD, 1);
						//											}
						//
						//											if (bAdicional)
						//											{
						//												if (i == 0) CargaHabitacion(0, 1);
						//				
						//												this.CargaArticulosPromocionAdicional(idPromocionD);												
						//											}
						//
						//											if (bOpcional)
						//											{
						//												if (i == 0) CargaHabitacion(0, 1);
						//
						//												foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
						//												{	
						//													if ((bool)dr.Cells["Principal"].Value) 
						//														CargaArticulosPromocionOpcional(idPromocionD);
						//												}
						//											}
						//											#endregion Promociones
						//
						//											i++;
						//										}
						//									}
						//									#endregion Recorre Grid Promociones Tarjeta VIP
						//
						//									bValidaNumeroIngresos = true;
						//									bPromocionVIP = true;
						//								}									
						//							}
						#endregion Tarjetas VIP

						#endregion Valida Promociones Tarjetas VIP
						
						#region Si no hay Promociones Tarjetas VIP vigentes valida promociones de descuentos o productos opcionales y adicionales
						if (!bPromocionVIP)
						{
							#region Busca Promociones Activas
							DataSet oDSPA = new DataSet();					
							cdsSeteoF.Schema.Connections[0].Open();
							C1.Data.SchemaObjects.Connection Coneccion = cdsSeteoF.Schema.Connections[0];
							Coneccion = cdsSeteoF.Schema.Connections[0];

							string sSQLPA = string.Format("Exec ValidaDescuentoPromocionPeony {0}, {1}, {2}, {3}, '{4}'", 
								(int)this.cmbBodega.Value, 3422, 2, (int)this.cmbHabitaciones.Value, Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"));
							SqlDataAdapter oDAPA = new SqlDataAdapter(sSQLPA, (SqlConnection)Coneccion.DbConnection);
							oDAPA.Fill(oDSPA);
					
							if (oDSPA.Tables[0].Rows.Count == 0 && bValidaNumeroIngresos) CargaHabitacion(idPromocionD, 1);							
							else if (oDSPA.Tables[0].Rows.Count > 0)
							{
								#region Validacion Promociones
								using (frmListaPromociones LP = new frmListaPromociones(sSQLPA))
								{				
									if (DialogResult.OK == LP.ShowDialog())
									{
										idPromocionD = (int)LP.ultraGrid1.ActiveRow.Cells["idPromocion"].Value;
										bDescuento = (bool)LP.ultraGrid1.ActiveRow.Cells["SumaDescuento"].Value;
										bAdicional = (bool)LP.ultraGrid1.ActiveRow.Cells["Adicional"].Value;
										bOpcional = (bool)LP.ultraGrid1.ActiveRow.Cells["Opcional"].Value;
										bCupon = (bool)LP.ultraGrid1.ActiveRow.Cells["Cupon"].Value;
										
										#region Valida Cupones
										if (bCupon)
										{
											string[] sPrmCW = new string[2];
											sPrmCW = Facturacion.sCuponWeb(cdsSeteoF);
											idCuponWeb = int.Parse(sPrmCW[0]);

											if (idCuponWeb == 0) 
											{
												bAplicaPromocion = false;
												idPromocionD = 0;
											}
											if (idCuponWeb > 0) bCuponActivo = true;
										}
										#endregion Valida Cupones

										#region Carga Promociones
										if (bAplicaPromocion)
										{
											#region Promociones
											if (bDescuento)
												CargaHabitacion(idPromocionD, 1);

											if (bAdicional)
											{
												CargaHabitacion(idPromocionD, 1);
				
												this.CargaArticulosPromocionAdicional(idPromocionD);
											}

											if (bOpcional)
											{
												CargaHabitacion(idPromocionD, 1);

												foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
												{	
													if ((bool)dr.Cells["Principal"].Value  && (int)dr.Cells["idArticulo"].Value == 3422) 
														CargaArticulosPromocionOpcional(idPromocionD);
												}
											}
											#endregion Promociones
										}
										#endregion Carga Promociones
									}
								}
								#endregion Validacion Promociones
							}	
							#endregion Busca Promociones Activas
						}
						#endregion Si no hay Promociones Tarjetas VIP vigentes valida promociones de descuentos o productos opcionales y adicionales

						int iCantidad = 0;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDet in this.uGridDetalle.Rows.All)
							if (drDet.Cells["idArticulo"].Value.ToString().Length > 0)
								if ((int)drDet.Cells["idArticulo"].Value == 3422) iCantidad = iCantidad + (int)drDet.Cells["Cantidad"].Value;

						if ((iIngresos - iCantidad) > 0)// && bValidaNumeroIngresos)
						{
							CargaHabitacion(idPromocionD, (iIngresos - iCantidad));							
						}
					}
				}

				Total();

				return;
			}
			#endregion Carga de Hospedaje

			#region Carga de Productos
			if (!bHabitacion)
			{
				#region variables
				decimal dDescuentaPorc = 0.00m;
				int iFila = uGridDetalle.Rows.Count - 1;
				iFila++;
				uGridDetalle.DisplayLayout.Bands[0].AddNew();	
				#endregion variables

				#region Declara Conexion
				DataSet oDS = new DataSet();					
				cdsSeteoF.Schema.Connections[0].Open();
				C1.Data.SchemaObjects.Connection miConeccion = cdsSeteoF.Schema.Connections[0];
				miConeccion = cdsSeteoF.Schema.Connections[0];
				#endregion Declara Conexion

				string sSQL = string.Format("Exec ValidaDescuentoPromocionPeony {0}, {1}, {2}, {3}, '{4}'", 
					(int)this.cmbBodega.Value, idArticulo, 2, (int)this.cmbHabitaciones.Value, Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"));
				SqlDataAdapter oDA = new SqlDataAdapter(sSQL, (SqlConnection) miConeccion.DbConnection);
				oDA.Fill(oDS);
					
				if (oDS.Tables[0].Rows.Count == 0)
				{	
					#region Informacion Productos sin promocion
					uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;
					uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = (int)this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value;
					uGridDetalle.Rows[iFila].Cells["Codigo"].Value = this.uGridArticulos.ActiveRow.Cells["Codigo"].Value.ToString();				
					uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = 1;
					uGridDetalle.Rows[iFila].Cells["Precio"].Value = Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value)/dPorcentajeIVADec;
					uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = dPorcentajeIVA;
					uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = ((Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value)/dPorcentajeIVADec) * dDescuentaPorc)/100;
					uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = dDescuentaPorc;
					uGridDetalle.Rows[iFila].Cells["SubTotal"].Value = Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value);
					uGridDetalle.Rows[iFila].Cells["Principal"].Value = true;
					uGridDetalle.Rows[iFila].Cells["Tipo"].Value = 0;				
					uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = false;
					uGridDetalle.Rows[iFila].Cells["idCombo"].Value = 0;
					uGridDetalle.Rows[iFila].Cells["Articulo"].Value = this.uGridArticulos.ActiveRow.Cells["Articulo"].Value.ToString();
					uGridDetalle.Rows[iFila].Cells["Entregado"].Value = true;
					#endregion Informacion Productos sin promocion
				}	
				else if (oDS.Tables[0].Rows.Count > 0 )
				{
					#region Lista de Promociones //entra aqui 
					using (frmListaPromociones LP = new frmListaPromociones(sSQL))
					{				
						if (DialogResult.OK == LP.ShowDialog())
						{
							idPromocionD = (int)LP.ultraGrid1.ActiveRow.Cells["idPromocion"].Value;
							bDescuento = (bool)LP.ultraGrid1.ActiveRow.Cells["SumaDescuento"].Value;
							bAdicional = (bool)LP.ultraGrid1.ActiveRow.Cells["Adicional"].Value;
							bOpcional = (bool)LP.ultraGrid1.ActiveRow.Cells["Opcional"].Value;
							bCupon = (bool)LP.ultraGrid1.ActiveRow.Cells["Cupon"].Value;

							if (bCupon)
							{
								string[] sPrmCW = new string[2];
								sPrmCW = Facturacion.sCuponWeb(cdsSeteoF);
								idCuponWeb = int.Parse(sPrmCW[0]);

								if (idCuponWeb == 0) bAplicaPromocion = false;
							}

							if (bAplicaPromocion)
							{
								if (bDescuento)
								{
									dDescuentaPorc = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaDescuentoArticulosPromocionPeony({0}, {1}, {2}, {3})", 
										(int)this.cmbBodega.Value, (int)this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value, 2, (int)this.cmbHabitaciones.Value));
								}
								
								#region Informacion Productos con promocion
								uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;
								uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = (int)this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value;
								uGridDetalle.Rows[iFila].Cells["Codigo"].Value = this.uGridArticulos.ActiveRow.Cells["Codigo"].Value.ToString();				
								uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = 1;
								uGridDetalle.Rows[iFila].Cells["Precio"].Value = Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value)/Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)"));
								uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente(GETDATE(), 1)"));
								uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = ((Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value)/Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)"))) * dDescuentaPorc)/100;
								uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = dDescuentaPorc;
								uGridDetalle.Rows[iFila].Cells["SubTotal"].Value = Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value);
								uGridDetalle.Rows[iFila].Cells["Principal"].Value = true;
								uGridDetalle.Rows[iFila].Cells["Tipo"].Value = 0;				
								uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = false;
								uGridDetalle.Rows[iFila].Cells["idCombo"].Value = 0;
								uGridDetalle.Rows[iFila].Cells["Articulo"].Value = this.uGridArticulos.ActiveRow.Cells["Articulo"].Value.ToString();
								uGridDetalle.Rows[iFila].Cells["Entregado"].Value = false;
								#endregion Informacion Productos con promocion

								#region Valida y Carga Promocion Adicional
								if (bAdicional)
								{				
									this.CargaArticulosPromocionAdicional(idPromocionD);

									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
									{	
										if ((bool)dr.Cells["Principal"].Value && (int)dr.Cells["idArticulo"].Value != 3422) 
											CargaArticulosPromocionOpcional(idPromocionD);
									}
								}
								#endregion Valida y Carga Promocion Adicional

								#region Valida y Carga Promocion Opcional
								if (bOpcional)
								{
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
									{	
										if ((bool)dr.Cells["Principal"].Value && (int)dr.Cells["idArticulo"].Value != 3422) 
											CargaArticulosPromocionOpcional(idPromocionD);
									}
								}
								#endregion Valida y Carga Promocion Opcional
							}
						}
						else 
						{
							#region Informacion Productos cuando no acepta la promocion
							uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;
							uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = (int)this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value;
							uGridDetalle.Rows[iFila].Cells["Codigo"].Value = this.uGridArticulos.ActiveRow.Cells["Codigo"].Value.ToString();				
							uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = 1;
							uGridDetalle.Rows[iFila].Cells["Precio"].Value = Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value)/Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)"));
							uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente(GETDATE(), 1)"));
							uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = ((Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value)/Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)"))) * dDescuentaPorc)/100;
							uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = dDescuentaPorc;
							uGridDetalle.Rows[iFila].Cells["SubTotal"].Value = Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value);
							uGridDetalle.Rows[iFila].Cells["Principal"].Value = true;
							uGridDetalle.Rows[iFila].Cells["Tipo"].Value = 0;				
							uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = false;
							uGridDetalle.Rows[iFila].Cells["idCombo"].Value = 0;
							uGridDetalle.Rows[iFila].Cells["Articulo"].Value = this.uGridArticulos.ActiveRow.Cells["Articulo"].Value.ToString();
							uGridDetalle.Rows[iFila].Cells["Entregado"].Value = false;
							#endregion Informacion Productos cuando no acepta la promocion
						}
					}
					#endregion Lista de Promociones
				}
			}
			#endregion Carga de Productos	

			Total();
		}

		private void CargaArticulosPromocionOpcional(int idPromocion)
		{		
			using (frmListaArticulosCombosPromociones LACP = new frmListaArticulosCombosPromociones(idPromocion, dtFechaPromocion))
			{				
				if (DialogResult.OK == LACP.ShowDialog())
				{
					#region Carga Articulos Promociones
					if (LACP.ultraGrid1.Rows.Count > 0)
					{
						int iFila = uGridDetalle.Rows.Count - 1;

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in LACP.ultraGrid1.Rows.All)
						{
							if ((bool)dr.Cells["Sel"].Value)
							{			
								#region Valida Comanda, Factura Temporal
								ValidaTemporalArticulo((int)dr.Cells["idArticulo"].Value);
								if(iExisComanda == 1 && bMenuCuadricula)
								{
									if (DialogResult.Yes == MessageBox.Show("Esta seguro Ordenar la Comanda Opcional. \n\nSe enviara el pedido a cocina.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
									{
										iFila++;
										uGridDetalle.DisplayLayout.Bands[0].AddNew();
										uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;
										uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = (int)dr.Cells["idArticulo"].Value;											
										uGridDetalle.Rows[iFila].Cells["Codigo"].Value = dr.Cells["Codigo"].Value;
										uGridDetalle.Rows[iFila].Cells["Articulo"].Value = dr.Cells["Articulo"].Value;									
										uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = (int)dr.Cells["Cantidad"].Value;	
										uGridDetalle.Rows[iFila].Cells["Precio"].Value = Convert.ToDecimal(dr.Cells["Precio"].Value);
										uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = Convert.ToDecimal(dr.Cells["Impuesto"].Value);
										uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = Convert.ToDecimal(dr.Cells["DescuentoArt"].Value);
										uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value);
										uGridDetalle.Rows[iFila].Cells["Principal"].Value = (bool)dr.Cells["Principal"].Value;
										uGridDetalle.Rows[iFila].Cells["Tipo"].Value = 4;																		
										uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = false;
										uGridDetalle.Rows[iFila].Cells["idCombo"].Value = idPromocion;
										uGridDetalle.Rows[iFila].Cells["Entregado"].Value = false;

										//										InsertaComanda((int)dr.Cells["idArticulo"].Value, 0, 1, 0, 0, 0, 0, 0, 1, ""); //preuba 
										//										this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleFacturas {0}", idCompra));
									
										#region Imprime Comanda
										string sImpresora = "";
										sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",(int) this.cmbBodega.Value));		
			              
										string sImpresoraCocina = "";
										sImpresoraCocina = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 2 ",(int) this.cmbBodega.Value));		
			
										//string sImpresora = "SLK-TS100"; //Cambiar el nombre de la impresora Cocina  "SLK-TS100"  "RICOH Aficio MP 3350 PCL 6"
										//string sImpresoraCocina = "SLK-TS100_cocina";
										//										ImprimeComanda(iComandaImp, sImpresora);
										//										ImprimeComandaCliente(iComandaImp, sImpresoraCocina);
										iComandaImp = 0;
										iExisComanda = 0;
										idComandaOrden = 0;
										#endregion Imprime Comanda
									}
								}
								else
								{
									uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;
									uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = (int)dr.Cells["idArticulo"].Value;											
									uGridDetalle.Rows[iFila].Cells["Codigo"].Value = dr.Cells["Codigo"].Value;
									uGridDetalle.Rows[iFila].Cells["Articulo"].Value = dr.Cells["Articulo"].Value;									
									uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = (int)dr.Cells["Cantidad"].Value;	
									uGridDetalle.Rows[iFila].Cells["Precio"].Value = Convert.ToDecimal(dr.Cells["Precio"].Value);
									uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = Convert.ToDecimal(dr.Cells["Impuesto"].Value);
									uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = Convert.ToDecimal(dr.Cells["DescuentoArt"].Value);
									uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value);
									uGridDetalle.Rows[iFila].Cells["Principal"].Value = (bool)dr.Cells["Principal"].Value;
									uGridDetalle.Rows[iFila].Cells["Tipo"].Value = 4;																		
									uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = false;
									uGridDetalle.Rows[iFila].Cells["idCombo"].Value = idPromocion;
									uGridDetalle.Rows[iFila].Cells["Entregado"].Value = false;
								}
								#endregion Valida Comanda, Factura Temporal															
							}
						}
					}
					#endregion Carga Articulos Promociones						
				}
			}
		}

		private void CargaArticulosPromocionAdicional(int idPromocion)
		{
			int iFila = uGridDetalle.Rows.Count - 1;
				
			string sSQLProm = string.Format("Exec CargaArticulosPromocionesPeony {0}, '{1}'", 
				idPromocion, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLProm, true);
																	
			while (drC.Read())
			{
				#region Variables
				int idArtProm = drC.GetInt32(0);
				string sCodProm = drC.GetString(1); 
				string sArtProm = drC.GetString(2);					
				int iCantidad = drC.GetInt32(3);
				decimal dPrecProm = drC.GetDecimal(4);
				decimal dImpProm = drC.GetDecimal(5);
				decimal dDesArtProm = drC.GetDecimal(6);
				decimal dDesPorProm = drC.GetDecimal(7);									
				bool bPrincipal = drC.GetBoolean(8);
				#endregion Variables
					
				#region Valida Comanda, Factura Temporal
				ValidaTemporalArticulo(idArtProm);
				if(iExisComanda == 1 && bMenuCuadricula)
				{
					
					if (DialogResult.Yes == MessageBox.Show("Esta seguro realizar la comanda Adicional. \n\nSe enviara el pedido a cocina.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
					{
						iFila++;
						uGridDetalle.DisplayLayout.Bands[0].AddNew();

						#region Asigna valores en el grid
						uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;
						uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = idArtProm;
						uGridDetalle.Rows[iFila].Cells["Codigo"].Value = sCodProm;					
						uGridDetalle.Rows[iFila].Cells["Articulo"].Value = sArtProm;
						uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = iCantidad;
						uGridDetalle.Rows[iFila].Cells["Precio"].Value = dPrecProm;
						uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
						uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
						uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
						uGridDetalle.Rows[iFila].Cells["SubTotal"].Value = dPrecProm;
						uGridDetalle.Rows[iFila].Cells["Principal"].Value = bPrincipal;
						uGridDetalle.Rows[iFila].Cells["Tipo"].Value = 3;
						uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = false;
						uGridDetalle.Rows[iFila].Cells["idCombo"].Value = idPromocion;		
						uGridDetalle.Rows[iFila].Cells["Entregado"].Value = false;		
						#endregion Asigna valores en el grid

						//						InsertaComanda(idArtProm, 0, 1, 0, 0, 0, 0, 0, 1, "");
						//						this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleFacturas {0}", idCompra));	
					}
				}		
				else
				{
					#region Asigna valores en el grid
					uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;
					uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = idArtProm;
					uGridDetalle.Rows[iFila].Cells["Codigo"].Value = sCodProm;					
					uGridDetalle.Rows[iFila].Cells["Articulo"].Value = sArtProm;
					uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = iCantidad;
					uGridDetalle.Rows[iFila].Cells["Precio"].Value = dPrecProm;
					uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
					uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
					uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
					uGridDetalle.Rows[iFila].Cells["SubTotal"].Value = dPrecProm;
					uGridDetalle.Rows[iFila].Cells["Principal"].Value = bPrincipal;
					uGridDetalle.Rows[iFila].Cells["Tipo"].Value = 3;
					uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = false;
					uGridDetalle.Rows[iFila].Cells["idCombo"].Value = idPromocion;										
					uGridDetalle.Rows[iFila].Cells["Entregado"].Value = false;
					#endregion Asigna valores en el grid
				}
				#endregion Valida Comanda, Factura Temporal
			} 
			drC.Close();
			
			//			if(iExisComanda == 1 && bMenuCuadricula)
			//			{
			//				//string sImpresora = "SLK-TS100"; //Cambiar el nombre de la impresora Cocina  "SLK-TS100"  "RICOH Aficio MP 3350 PCL 6"
			//				//string sImpresoraCocina = "SLK-TS100_cocina";
			//				string sImpresora = "";
			//				sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",(int) this.cmbBodega.Value));		
			//			              
			//				string sImpresoraCocina = "";
			//				sImpresoraCocina = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 2 ",(int) this.cmbBodega.Value));		
			//			
			//				ImprimeComanda(iComandaImp, sImpresora);
			//				ImprimeComandaCliente(iComandaImp, sImpresoraCocina);
			//				iComandaImp = 0;
			//				iExisComanda = 0;
			//				idComandaOrden = 0;
			//			}
		}
		
		private void btnNuevo_Click(object sender, System.EventArgs e)
		{	
			if (!bMenuCuadricula) this.btnCancelar_Click(sender, e);

			//FuncionesProcedimientos.RetornaFechaServidor(this.dtFechaIngreso, cdsSeteoF, true);
			this.dtFechaIngreso.Value = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);

			this.optTurnoIngreso.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTurno '{0}'", Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm")));
								
			this.dtFecha.Value = Funcion.dtEscalarSQL(cdsSeteoF, "Select dbo.FacturacionRetornaFecha()");
				
			this.dtFechaPromocion = (DateTime)this.dtFechaIngreso.Value;
			
			iEstado = 6;
			this.txtPersonas.Value = 2;

			bNuevo = true;
			bEdicion = true;

			#region Peticion desde cuadricula
			if (bMenuCuadricula)
			{
				this.cmbBodega.Value = iBodega;
				this.cmbHabitaciones.Value = idHabitacion;
												
				//int idBodegaSeleccionada = Convert.ToInt32(this.cmbBodega.Value);

				bCobroAnticipado = Funcion.bEscalarSQL(cdsSeteoF,string.Format("SELECT CobroAnticipado FROM Bodega WHERE Bodega = {0}", iBodega),true);

				this.btnCancelar.Enabled = true;

				dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
				dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

				ValidaCargaReservas();
				
				if(MenuLatinium.iNivel != 31)
				{					
					this.btnVer_Click(sender, e);
				}
				int a = (int)this.txtTiempo.Value;

				if (!bServicioPagado) this.txtPersonas.Enabled = true;

				this.btnTemporal.Enabled = true;
				this.btnOrdenCocina.Enabled = true;

				this.cmbGuardia.TabIndex = 0;
				//this.cmbGuardia.Enabled = true;
				this.txtEncuesta.Enabled = false;
				this.txtEncuesta.Value = 0;
				this.txtEncuesta2.Enabled = false;
				this.txtEncuesta2.Value = 0;
				this.txtPersonas.Enabled = true;
				this.txtUnidad.Enabled = true;
				//this.cmbCobrador.Enabled = true;
				//this.cmbCajero.Enabled = true;

				if (MenuLatinium.iNivel == 23)
				{
					this.cmbCajero.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalLotteFacturacion {0}, {1}", 4, MenuLatinium.IdUsuario));
					this.cmbCajero.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec PersonalLotteFacturacion {0}, {1}", 4, MenuLatinium.IdUsuario));					
				}

				#region Activar Boton Cupones
				iActivoMitadPrecion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Activo From CuponesMitadPrecio Where Bodega = {0}", (int)this.cmbBodega.Value));
				if (iActivoMitadPrecion == 1 && miAcceso.BCupon) this.btnMitadPrecio.Enabled = true;

				iActivoCuponBodega = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Activo From CuponesEmpresasActivo Where Bodega = {0}", (int)this.cmbBodega.Value));
				if (iActivoCuponBodega == 1 && miAcceso.BCuponEmpresas) this.btnCupon.Enabled = true;

				if(Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ChkActivo From CuponesEmpresasActivo Where Bodega = {0}", (int)this.cmbBodega.Value), true))
				{
					if(idCompra > 0)
					{
						iActivaChkCupon = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaAceptaCupones {0}, {1}", (int)this.cmbBodega.Value, (int)this.cmbHabitaciones.Value));
						if(iActivaChkCupon == 1 && (int)this.cmbTipoRuc.Value != 11) this.chkCupon.Enabled = true;
					}
				}
				#endregion Activar Boton Cupones

				#region Activar Tarjeta VIP
				if(idCompra > 0)
				{
					iActivaTajetaVIP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Activo From ActivaTarjetasVIP Where Bodega = {0}", (int)this.cmbBodega.Value));
					if (iActivaTajetaVIP == 1 && miAcceso.BActivaTarjetaVIP) 
					{
						this.chkTarjetaVIP.Enabled = true;
						this.btnInfoTarjetasVIP.Enabled = true;
					}
				}
				#endregion Activar Tarjeta VIP

				#region Toque de Queda
				bToqueQueda = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec CambiaColorToqueQueda {0}", (int)this.cmbHabitaciones.Value), true);
				if(bValidaPromocionMensaje) bToqueQueda = false;
				#endregion Toque de Queda
			}
			#endregion Peticion desde cuadricula

			#region Peticion desde menu
			if (!bMenuCuadricula)
			{				
				// this.dtFecha.Value = DateTime.Today;
				ValNuevaFactura = true;
									
				#region Habilita controles
	
				#region Validaciones Nuevo
				this.cmbBodega.Enabled = true;
				//this.cmbGuardia.Enabled = true;
				//this.cmbCobrador.Enabled = true;
				this.cmbHabitaciones.Enabled = true;
				#endregion Validaciones Nuevo

				if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;				
				// this.dtFechaIngreso.Value = DateTime.Today;
				if (MenuLatinium.IdUsuario == 279)
				{
					this.dtFechaIngreso.Enabled = true;
				}
				else 
				{
					this.dtFechaIngreso.Enabled = false;
				}
				this.dtFechaSalida.Value = DateTime.Today;
				if (MenuLatinium.IdUsuario == 279)
				{
					this.dtFechaSalida.Enabled = true;
				}
				else
				{
					this.dtFechaSalida.Enabled = false;
				}			
				this.cmbTipoRuc.Value = 5;					
				this.txtUnidad.Enabled = true;
				this.cmbFormaPago.Enabled = true;

				this.txtEncuesta.Enabled = false;
				this.txtEncuesta.Value = 0;
				this.txtEncuesta2.Enabled = false;
				this.txtEncuesta2.Value = 0;
				this.txtUnidad.Enabled = false;
				this.txtUnidad.Value = 0;
				this.txtPersonas.Enabled = false;
				this.txtPersonas.Value = 2;
				this.cmbCajero.Enabled = true;
				
				if (MenuLatinium.iNivel == 23)
				{
					this.cmbCajero.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalLotteFacturacion {0}, {1}", 4, MenuLatinium.IdUsuario));
					this.cmbCajero.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec PersonalLotteFacturacion {0}, {1}", 4, MenuLatinium.IdUsuario));					
				}
				if (MenuLatinium.IdUsuario == 279)
				{
					this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga 3");
					this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFPagoCarga 2");
				}
				else
				{
					this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga 2");
					this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFPagoCarga 3");
				}
				vValFPago = 3;
				vValidaFPAgo = true;
				vValidaRucEmpleado = true;

				this.txtNotas.Enabled = true;
				
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;				
				this.btnGuardar.Enabled = true;
				this.btnTemporal.Enabled = false;			
				this.btnAsiento.Enabled = false;
				this.chkAbono.Enabled = false;
					
				#endregion Habilita controles 
				
				dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
				dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

				// FuncionesProcedimientos.RetornaFechaServidor(this.dtFechaSalida, cdsSeteoF, true);
				this.dtFechaSalida.Value = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);

				this.optTurnoSalida.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTurno '{0}'", Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd HH:mm")));

				this.uGridDetalle.DisplayLayout.Bands[0].Columns["Cantidad"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}
			#endregion Peticion desde menu

			iIngresos = 1;

			if(miAcceso.CambiarNumeracion) this.txtNumero.Enabled = true;
			
			if (miAcceso.BDesModifica) bBtnDesFactura = true;
			if (miAcceso.BDesEmpleado) bBtnDesEmpleado = true;			

			this.cmbTipoRuc.Enabled = true;
			this.txtRuc.Enabled = true;

			this.cmbFormaPago.Enabled = true;
			this.txtNotas.Enabled = true;

			this.txtBuscar.Enabled = true;
			this.btnVer.Enabled = true;
			this.btnOrdenCocina.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
			FuncionesProcedimientos.EstadoGrids(true, this.uGridArticulos);

			this.btnCliente.Enabled = true;
			this.btnAnular.Enabled = true;
			this.btnCancelar.Enabled = true;
			
			if (!bMenuCuadricula)
			{
				if (this.cmbBodega.Enabled) this.cmbBodega.Focus(); else this.cmbHabitaciones.Focus();	
			}
			else this.cmbGuardia.Focus();

			#region Seguridad Anfitriones
			if (MenuLatinium.iNivel == 31)
			{
				this.btnCancelar.Enabled = false;
				this.btnTemporal.Enabled = false;
				this.btnSalir.Enabled = true;
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnAsiento.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnRefacturacion.Enabled = false;
				this.btnCupon.Enabled = false;
				this.btnExtras.Enabled = false;
				this.btnClaveDescuento.Enabled = false;
				this.btnDescEmpAut.Enabled = false;
				this.btnVer.Enabled = false;

				this.cmbTipoRuc.Enabled = false;
				this.cmbCobrador.Enabled = false;
				this.cmbCajero.Enabled = false;
				this.cmbFormaPago.Enabled = false;

				this.txtRuc.Enabled = false;
				this.txtUnidad.Enabled = false;
				this.txtPersonas.Enabled = false;
				this.txtBuscar.Enabled = false;
				this.txtNotas.Enabled = false;

				FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);
				FuncionesProcedimientos.EstadoGrids(false, this.uGridArticulos);
			}
			#endregion Seguridad Anfitriones

			this.ValidaAnticiposClientes();
		}
		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaFacturas Busqueda = new frmBuscaFacturas(1, iBodega, miAcceso.BLimiteBusquedaFechas))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					ValidaTiempo((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
					this.Consulta((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
				}
			}
		}

		private void EdicionRefactura()
		{
			#region Controles
			bNuevo = false;
			bEdicion = true;
			bRFAnticipo = true;
			vValidaFPAgo = true;
			bRFVideoPromocion = true;

			this.txtUnidad.Enabled = true;
			this.txtNotas.Enabled = true;
			this.txtBuscar.Enabled = true;
			//this.dtFechaIngreso.Enabled = true;
			//this.dtFechaIngreso.Value = DateTime.Today;
			this.dtFechaSalida.Enabled = true;

			this.cmbFormaPago.Enabled = true;
			//this.cmbCobrador.Enabled = true;
			this.cmbTipoRuc.Enabled = true;
			//this.cmbCajero.Enabled = true;

			this.txtRuc.Enabled = true;

			this.ValidaAnticiposClientes();

			if (miAcceso.CambiarNumeracion) this.txtNumero.Enabled = true;
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
			FuncionesProcedimientos.EstadoGrids(true, this.uGridArticulos);

			dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
			dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
					
			this.btnVer.Enabled = true;
			this.btnCliente.Enabled = true;
			this.btnGuardar.Enabled = true;
			this.btnMitadPrecio.Enabled = true;
			this.btnCupon.Enabled = true;
			this.btnOrdenCocina.Enabled = true;
			this.btnCancelar.Enabled = false;
			this.btnRefacturacion.Enabled = false;
			this.btnSalir.Enabled = false;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnTemporal.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnClaveDescuento.Enabled = false;
			this.cmbCajero.Enabled = false;			
			this.btnExtras.Enabled = false;
			bFacturaTiempo = false;

			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFPagoCarga 2");
			vValFPago = 2;

			#region Activar Boton Cupones
			iActivoMitadPrecion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Activo From CuponesMitadPrecio Where Bodega = {0}", (int)this.cmbBodega.Value));
			if (iActivoMitadPrecion == 1 && miAcceso.BCupon) this.btnMitadPrecio.Enabled = true;

			iActivoCuponBodega = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Activo From CuponesEmpresasActivo Where Bodega = {0}", (int)this.cmbBodega.Value));
			if (iActivoCuponBodega == 1 && miAcceso.BCuponEmpresas) this.btnCupon.Enabled = true;

			if(Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ChkActivo From CuponesEmpresasActivo Where Bodega = {0}", (int)this.cmbBodega.Value), true))
			{
				if(idCompra > 0)
				{
					iActivaChkCupon = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaAceptaCupones {0}, {1}", (int)this.cmbBodega.Value, (int)this.cmbHabitaciones.Value));
					if(iActivaChkCupon == 1 && (int)this.cmbTipoRuc.Value != 11) this.chkCupon.Enabled = true;					
				}
			}
			#endregion Activar Boton Cupones

			#region Activar Tarjeta VIP
			if(idCompra > 0)
			{
				iActivaTajetaVIP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Activo From ActivaTarjetasVIP Where Bodega = {0}", (int)this.cmbBodega.Value));
				if (iActivaTajetaVIP == 1 && miAcceso.BActivaTarjetaVIP) 
				{
					this.chkTarjetaVIP.Enabled = true;
					this.btnInfoTarjetasVIP.Enabled = true;
				}
			}
			#endregion Activar Tarjeta VIP

			#endregion Controles
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{			
			#region Valida Hora
			if(miAcceso.BEditacionMinutos)
			{
				if (bFacturaTiempo)
				{
					bEdicion = true;
					ValidaEdicionFactura();
					return;
				}
			}
			#endregion Valida Hora

			#region Peticion desde cuadricula
			if (bMenuCuadricula)
			{
				FuncionesProcedimientos.RetornaFechaServidor(this.dtFechaSalida, cdsSeteoF, true);

				//this.ValidaAnticiposClientes();

				this.optTurnoSalida.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTurno '{0}'", Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd HH:mm")));
								
				this.dtFecha.Value = Funcion.dtEscalarSQL(cdsSeteoF, "Select dbo.FacturacionRetornaFecha()");
				
				this.dtFechaPromocion = (DateTime)this.dtFechaIngreso.Value;

				bCobroAnticipado = Funcion.bEscalarSQL(cdsSeteoF, string.Format("SELECT CobroAnticipado FROM Bodega WHERE Bodega = {0}", iBodega), true);

				if (iEstadoHabitacion != 10)
				{
					this.cmbTipoRuc.Enabled = true;
					if ((int)this.cmbTipoRuc.Value != 11) this.txtRuc.Enabled = true;
					this.txtTiempoAdicional.Enabled = true;							
					this.txtBuscar.Enabled = true;
					this.btnVer.Enabled = true;
					this.cmbFormaPago.Enabled = true;
									
					FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
					FuncionesProcedimientos.EstadoGrids(true, this.uGridArticulos);

					this.btnCliente.Enabled = true;											

					if (miAcceso.BAnulaTiempo) 
						this.ValidaTiempoAnulacion();				
					else this.btnAnular.Enabled = true;

					this.btnTemporal.Enabled = true;
					this.btnImprimir.Enabled = true;
					this.btnTiempoAdicional.Enabled = true;
				}

				dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
				dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
				
				//iPuntos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Puntos, 0) From Compra Where idCompra = {0}", idCompra));
				sCorreo = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Correo, '') From Compra Where idCompra = {0}", idCompra));
				
				this.txtEncuesta.Enabled = false;
				this.txtEncuesta.Value = 0;
				this.txtEncuesta2.Enabled = false;
				this.txtEncuesta2.Value = 0;
				
				this.btnVer_Click(sender, e);

				#region Cupon Tarjeta VIP Comentado
				if(iEstado == 6) 
				{
					#region Activar Boton Cupones
//					iActivoMitadPrecion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Activo From CuponesMitadPrecio Where Bodega = {0}", (int)this.cmbBodega.Value));
//					if (iActivoMitadPrecion == 1 && miAcceso.BCupon) this.btnMitadPrecio.Enabled = true;
//
//					iActivoCuponBodega = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Activo From CuponesEmpresasActivo Where Bodega = {0}", (int)this.cmbBodega.Value));
//					if (iActivoCuponBodega == 1 && miAcceso.BCuponEmpresas) this.btnCupon.Enabled = true;
//
//					if(Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ChkActivo From CuponesEmpresasActivo Where Bodega = {0}", (int)this.cmbBodega.Value), true))
//					{
//						if(idCompra > 0)
//						{
//							iActivaChkCupon = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaAceptaCupones {0}, {1}", (int)this.cmbBodega.Value, (int)this.cmbHabitaciones.Value));
//							if(iActivaChkCupon == 1 && (int)this.cmbTipoRuc.Value != 11) this.chkCupon.Enabled = true;					
//						}
//					}
					#endregion Activar Boton Cupones

					#region Activar Tarjeta VIP
//					if(idCompra > 0)
//					{
//						iActivaTajetaVIP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Activo From ActivaTarjetasVIP Where Bodega = {0}", (int)this.cmbBodega.Value));
//						if (iActivaTajetaVIP == 1 && miAcceso.BActivaTarjetaVIP) 
//						{
//							this.chkTarjetaVIP.Enabled = true;
//							this.btnInfoTarjetasVIP.Enabled = true;
//						}
//					}
					#endregion Activar Tarjeta VIP
				}
				#endregion Cupon Tarjeta VIP Comentado 

				this.btnReturnPass.Enabled = true;

				this.CargaArticulos(3422, true);
			}
			#endregion Peticion desde cuadricula

			#region Peticion desde menu
			if (!bMenuCuadricula)
			{	
				if ( idFe_EstadoDocumento == 2 && !miAcceso.BEditarAprobado)
				{
					MessageBox.Show("La Factura esta autorizada por el SRI, no puede ser editada", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (miAcceso.CambiarNumeracion) this.txtNumero.Enabled = true;

//				if (!bImpreso) 
//				{
					if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
					if (miAcceso.BtnDescuento) this.btnClaveDescuento.Enabled = true;
					if (miAcceso.BDesModifica) bBtnDesFactura = true;
					if (miAcceso.BDesEmpleado) bBtnDesEmpleado = true;
					if (miAcceso.BDesEmpleado) this.btnDescEmpAut.Enabled = true;
					
					this.btnCliente.Enabled = true;

					this.cmbTipoRuc.Enabled = true;
					this.txtRuc.Enabled = true;

					this.txtBuscar.Enabled = true;
					this.btnVer.Enabled = true;
					this.btnVer_Click(sender, e);
					
					FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
					FuncionesProcedimientos.EstadoGrids(true, this.uGridArticulos);
//				}

				vValidaFPAgo = true;
				vValFPago = 2;

				//				if (bImpreso) 
				//				{
				if ((int)this.cmbTipoRuc.Value == 11)
				{
					if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;

					this.txtTiempo.Enabled = true;

					this.btnCliente.Enabled = true;

					this.cmbTipoRuc.Enabled = true;
					this.txtRuc.Enabled = true;
					this.txtPersonas.Enabled = true;

					this.txtBuscar.Enabled = true;
					this.btnVer.Enabled = true;
					this.btnVer_Click(sender, e);						
					//}
					FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
					FuncionesProcedimientos.EstadoGrids(true, this.uGridArticulos);
				}

				dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
				dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

				this.dtFechaIngreso.Enabled = true;
				this.dtFechaSalida.Enabled = true;
				//				this.optTurnoIngreso.Enabled = true;
				//				this.optTurnoSalida.Enabled = true;
				//this.cmbCajero.Enabled = true;
				//this.cmbCobrador.Enabled = true;
				this.txtEncuesta.Enabled = true;
				this.txtEncuesta2.Enabled = true;

				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnTemporal.Enabled = false;
				this.btnAsiento.Enabled = false;
				this.btnAnular.Enabled = false;

				this.uGridDetalle.DisplayLayout.Bands[0].Columns["Cantidad"].CellActivation = Activation.AllowEdit;

				#region Forma Pago
				if((int)this.cmbFormaPago.Value != 29 ||(int)this.cmbFormaPago.Value != 30) 			
					this.cmbFormaPago.Enabled = true;			
				#endregion Forma Pago
			}
			#endregion Peticion desde menu

			this.txtUnidad.Enabled = true;

			this.txtNotas.Enabled = true;
      			
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			
			bNuevo = false;
			bEdicion = true;
			
			this.cmbTipoRuc.Focus();			

			#region Anticipos Abono Cliente
			if(idCompra > 0)
			{
				if(idCliente > 0)
				{
					this.ValidaAnticiposClientes(); 
				}
			}
			#endregion Anticipos Abono Cliente
		}

		private void ValidaEdicionFactura()
		{
			if (bEdicion)
			{
				int ValiCuponUtilizado = 0;
				int ValiEdicion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From FacturaEdicion Where Editado = 1 And idCompra = {0}", idCompra));
				int ValiHospedaje = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From FacturaEdicion Where Hospedaje = 1 And idCompra = {0}", idCompra));
				int ValiCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From Compra Where idcompra = {0}", idCompra));
				int ValiFactHospedaje = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From FacturaEdicion Where FacturaDividida = 1 And Hospedaje = 1 And Estado = 1 And idcompra = {0}", idCompra));
				int EstadoCaja = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec EstadoCajaRefacturacion '{0}', {1}, {2}", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm"), (int)this.optTurnoSalida.Value, (int)this.cmbBodega.Value));
				int ValidaAnticipoCruzado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From AnticiposCompras Where idTipoFactura = 1 And idCompra = {0}", idCompra));
				int ValidaTVIPUtilizados = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From TarjetasVIpCompras Where idTipoFactura = 1 And idCompra = {0}", idCompra));
				int iValCompraCupon = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesAnticipoCompra Where idCompraDivision = {0}", idCompra));

				if (ValiCompra == 1)
				{
					int iValEstadoFactura = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From FacturaEdicion Where Estado = 1 And idCompra = {0}", idCompra));
					
					if(iValEstadoFactura == 1)
					{
						#region Validaciones
						if(ValidaTVIPUtilizados > 0)
						{
							MessageBox.Show("No puede Editar la Factura, registra cruce con TARJETA VIP.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;
						}

						if(ValidaAnticipoCruzado > 0)
						{
							MessageBox.Show("No puede Editar la Factura, registra cruce con ANTICIPO.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;
						}

						if (iValCompraCupon >= 1)
						{
							MessageBox.Show("No puede editar la factura, registra cruce con CUPON.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtCupon.Focus();
							return;
						}

						if ((int)this.cmbFormaPago.Value == 14)
						{
							MessageBox.Show("No puede Editar Facturas con Forma Pago Cupon", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;
						}

						if ((int)this.cmbFormaPago.Value == 16)
						{
							MessageBox.Show("No puede Editar Facturas con Forma Pago Prenda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;
						}

						if ((int)this.cmbFormaPago.Value == 29)
						{
							MessageBox.Show("No puede Editar Facturas con Forma Pago Autoconsumo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;
						}

						if ((int)this.cmbFormaPago.Value == 30)
						{
							MessageBox.Show("No puede Editar Facturas con Forma Pago Descuento Empleado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;
						}

						if((int)this.cmbBodega.Value == 2)
						{
							ValiCuponUtilizado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesEmpresas Where idCompraDestino = {0}", idCompra));
							
							if (ValiCuponUtilizado == 1)
							{
								MessageBox.Show("No puede Editar Facturas, ya esta utilizado un cupon.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								bEdicion = false;
								return;
							}
						}

						if((int)this.cmbBodega.Value == 4)
						{
							ValiCuponUtilizado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesGenerados Where idcompra = {0}", idCompra));
							if (ValiCuponUtilizado == 1)
							{
								MessageBox.Show("No puede Editar Facturas, ya esta utilizado un cupon.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								bEdicion = false;
								return;
							}
						}
						
						if (EstadoCaja == 2)
						{
							MessageBox.Show("No se puede modificar si la CAJA ya esta Cerrada", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;
						}

						//						if(ValiHospedaje == 1)
						//						{
						//							MessageBox.Show("No se puede modificar si la Factura es solo Hospedaje", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						//							bEdicion = false;
						//							return;
						//						}
						#endregion Validaciones

						if (ValiEdicion == 1)
						{
							MessageBox.Show("Ya no se puede modificar por segunda vez la factura.", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;
						}
						else
						{
							if(ValiFactHospedaje == 1)
							{
								#region Controles Factura Hospedaje
								FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
								dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
								dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
								this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFPagoCarga 2");

								this.cmbFormaPago.Enabled = true;								
								this.btnGuardar.Enabled = true;
								this.btnCancelar.Enabled = true;
								this.cmbTipoRuc.Enabled = false;
								this.btnVer.Enabled = false;
								this.txtBuscar.Enabled = false;
								bActivaBoton = false;
								this.btnSalir.Enabled = false;
								this.btnNuevo.Enabled = false;
								this.btnConsultar.Enabled = false;
								this.btnEditar.Enabled = false;
								this.btnImprimir.Enabled = false;
								this.btnTemporal.Enabled = false;
								this.btnAsiento.Enabled = false;
								this.btnAnular.Enabled = false;
								this.btnClaveDescuento.Enabled = false;
								this.btnDescEmpAut.Enabled = false;
								this.btnRefacturacion.Enabled = false;
								this.btnMitadPrecio.Enabled = false;
								this.btnCupon.Enabled = false;
								vValFPago = 2;
								CoutEditar = 1;
								#endregion Controles Factura Hospedaje
							}
							else
							{
								#region Controles Factura Consumos
								FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
								dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
								dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
								this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFPagoCarga 2");
								this.cmbFormaPago.Enabled = true;
								this.cmbTipoRuc.Enabled = true;
								this.btnGuardar.Enabled = true;
								this.btnCancelar.Enabled = true;
								this.btnVer.Enabled = true;
								this.txtBuscar.Enabled = true;
								bActivaBoton = true;
								this.btnSalir.Enabled = false;
								this.btnNuevo.Enabled = false;
								this.btnConsultar.Enabled = false;
								this.btnEditar.Enabled = false;
								this.btnImprimir.Enabled = false;
								this.btnTemporal.Enabled = false;
								this.btnAsiento.Enabled = false;
								this.btnAnular.Enabled = false;
								this.btnClaveDescuento.Enabled = false;
								this.btnDescEmpAut.Enabled = false;
								this.btnRefacturacion.Enabled = false;
								this.btnMitadPrecio.Enabled = false;
								this.btnCupon.Enabled = false;
								vValFPago = 2;
								CoutEditar = 1;
								#endregion Controles Factura Consumos
							}
						}
					}
					else
					{
						MessageBox.Show("La Factura Sobre pasa la hora, no puede editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						bEdicion = false;
						return;
					}
				}
			}
		}

		private void ValidaMinFactura()
		{
			if (bEdicion)
			{
				if (this.dtFechaSalida.Value != DBNull.Value)
				{	
					int ValiEdicion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CompraEdicion Where idcompra = {0}", idCompra));

					int ValiCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From Compra Where idcompra = {0}", idCompra));

					int EstadoCaja = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec EstadoCajaRefacturacion '{0}', {1}, {2}", 
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm"), (int)this.optTurnoSalida.Value, (int)this.cmbBodega.Value));

					if (ValiCompra == 1)
					{
						int iValMinutos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaMinutosFactura {0}", idCompra));

						if ((int)this.cmbFormaPago.Value == 16)
						{
							MessageBox.Show("No puede Editar Facturas con Forma Pago Prenda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;	
						}

						if ((int)this.cmbFormaPago.Value == 29)
						{
							MessageBox.Show("No puede Editar Facturas con Forma Pago Autoconsumo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;	
						}

						if ((int)this.cmbFormaPago.Value == 30)
						{
							MessageBox.Show("No puede Editar Facturas con Forma Pago Descuento Empleado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;	
						}						

						if (iValMinutos > 60)
						{
							MessageBox.Show("Pasado los 60 minutos no puede editar la factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;				
						}

						if (EstadoCaja == 2)
						{
							MessageBox.Show("No se puede modificar si la CAJA ya esta Cerrada", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;
						}

						else if(ValiEdicion == 1)
						{
							MessageBox.Show("Ya no puede volver a editar la factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bEdicion = false;
							return;				
						}
						else
						{
							#region controles
							dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
							dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
							this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFPagoCarga 2");
							this.cmbFormaPago.Enabled = true;
							this.btnGuardar.Enabled = true;
							this.btnCancelar.Enabled = true;							
							this.btnSalir.Enabled = false;
							this.btnNuevo.Enabled = false;
							this.btnConsultar.Enabled = false;
							this.btnEditar.Enabled = false;
							this.btnImprimir.Enabled = false;
							this.btnTemporal.Enabled = false;
							this.btnAsiento.Enabled = false;
							this.btnAnular.Enabled = false;
							this.btnClaveDescuento.Enabled = false;
							this.btnDescEmpAut.Enabled = false;
							this.btnRefacturacion.Enabled = false;
							vValFPago = 2;
							CoutEditar = 1;
							#endregion controles
						}
					}
				}
				else
				{
					MessageBox.Show("Ingrese Fecha Salida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFechaSalida.Focus();
					return;
				}
			}			
		}

		private void Imprimir(int iTipo)
		{
			#region Nombre de Reporte y Filtro
			Cursor = Cursors.WaitCursor;
			string stReporte = "Exec CompraReporteNombre " + idCompra.ToString();
			
			stReporte += ", 1";
			string stFiltro = "";
			stReporte = Funcion.sEscalarSQL(cdsSeteoF, stReporte);
			if (iTipo==0)	stFiltro = "{Compra.idCompra} = " + idCompra.ToString();
			else stFiltro = "{RegistroPromociones.idCompra} = " + idCompra.ToString();
			#endregion Nombre de Reporte y Filtro

			#region Imprime el reporte
			//using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			if (iTipo == 0) 
			{
				if (bImpresionDivision)
				{
					#region Impresion Division
					// Reporte 1
					ParameterFields paramFields = new ParameterFields ();
				
					ParameterField pfRet = new ParameterField ();
					ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
					pfRet.ParameterFieldName = "@idCpa";
					idCompraRet.Value = iCompraOriginal;
					pfRet.CurrentValues.Add (idCompraRet);
					paramFields.Add (pfRet);
				
					Reporte miRepor = new Reporte("FacturaLotteServicio.rpt", "");
					miRepor.MdiParent = this.MdiParent;
					miRepor.Titulo("");
					miRepor.crVista.ParameterFieldInfo = paramFields;
					if (bMenuCuadricula) miRepor.ShowDialog();

					// Reporte 2
					ParameterFields paramFields1 = new ParameterFields ();

					ParameterField pfRet1 = new ParameterField ();
					ParameterDiscreteValue idCompraRet1 = new ParameterDiscreteValue ();
					pfRet1.ParameterFieldName = "@idCpa";
					idCompraRet1.Value = iCompraDividida;
					pfRet1.CurrentValues.Add (idCompraRet1);
					paramFields1.Add (pfRet1);
				
					Reporte miReporDivision = new Reporte("FacturaLotteServicio.rpt", "");
					miReporDivision.MdiParent = this.MdiParent;
					miReporDivision.Titulo("");
					miReporDivision.crVista.ParameterFieldInfo = paramFields1;
					if (bMenuCuadricula) miReporDivision.ShowDialog();
					#endregion Impresion Division
				}
				else
				{
					#region Impresion
					ParameterFields paramFields = new ParameterFields ();
				
					ParameterField pfRet = new ParameterField ();
					ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
					pfRet.ParameterFieldName = "@idCpa";
					idCompraRet.Value = idCompra;
					pfRet.CurrentValues.Add (idCompraRet);
					paramFields.Add (pfRet);
				
					Reporte miRepor = new Reporte("FacturaLotteServicio.rpt", "");
					miRepor.MdiParent = this.MdiParent;
					miRepor.Titulo("");
					miRepor.crVista.ParameterFieldInfo = paramFields;

					if (bMenuCuadricula) miRepor.ShowDialog();
					if (!bMenuCuadricula) miRepor.Show();
					#endregion Impresion
				}
			}
			if (iTipo == 1) 
			{
				using (Reporte miReporte = new Reporte("ImprPromocion", stFiltro))
				{		
					if (iTipo == 1) // Muestra en pantalla
						miReporte.ShowDialog();
					if (iTipo >= 2) // Imprime una o mas copias
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
			}
			#endregion Imprime el reporte

			#region Marca como impreso y muestra el tiempo
			string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + idCompra.ToString();
			Funcion.EjecutaSQL(cdsSeteoF, stGrabaImprime);
			bImpreso = true;
			Cursor = Cursors.Default;
			#endregion Marca como impreso y muestra el tiempo
		}

		private void ImprimirRefactura()
		{
			if (bImpresionDivision)
			{
				#region Impresion Division
				// Reporte 1
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfRet = new ParameterField ();
				ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
				pfRet.ParameterFieldName = "@idCpa";
				idCompraRet.Value = iCompraOriginal;
				pfRet.CurrentValues.Add (idCompraRet);
				paramFields.Add (pfRet);
				
				Reporte miRepor = new Reporte("FacturaLotteServicio.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				if (!bMenuCuadricula) miRepor.Show();

				// Reporte 2
				ParameterFields paramFields1 = new ParameterFields ();

				ParameterField pfRet1 = new ParameterField ();
				ParameterDiscreteValue idCompraRet1 = new ParameterDiscreteValue ();
				pfRet1.ParameterFieldName = "@idCpa";
				idCompraRet1.Value = iCompraDividida;
				pfRet1.CurrentValues.Add (idCompraRet1);
				paramFields1.Add (pfRet1);
				
				Reporte miReporDivision = new Reporte("FacturaLotteServicio.rpt", "");
				miReporDivision.MdiParent = this.MdiParent;
				miReporDivision.Titulo("");
				miReporDivision.crVista.ParameterFieldInfo = paramFields1;
				if (!bMenuCuadricula) miReporDivision.Show();
				#endregion Impresion Division
			}
			else
			{
				#region Impresion
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfRet = new ParameterField ();
				ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
				pfRet.ParameterFieldName = "@idCpa";
				idCompraRet.Value = idCompra;
				pfRet.CurrentValues.Add (idCompraRet);
				paramFields.Add (pfRet);
				
				Reporte miRepor = new Reporte("FacturaLotteServicio.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("");
				miRepor.crVista.ParameterFieldInfo = paramFields;

				if (bMenuCuadricula) miRepor.ShowDialog();
				if (!bMenuCuadricula) miRepor.Show();
				#endregion Impresion
			}						
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			if (!bMenuCuadricula)
				this.mnuImprimir.Show(this.btnImprimir, new Point(1, 20));
			else
			{
				if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione el tipo de identificacion del cliente")) return;
				if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;
				if (idCliente == 0)
				{
					MessageBox.Show("El Cliente No Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (DialogResult.Yes == MessageBox.Show("Esta seguro de Imprimir la factura y generar la salida del cliente?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					using (frmMnsajeNumLotte Msje = new frmMnsajeNumLotte(this.txtNumero.Text.ToString(), false))
					{
						if (DialogResult.OK == Msje.ShowDialog())
						{
							this.btnTemporal_Click(sender, e);

							#region Autorizacion Factura
							if (Funcion.bEscalarSQL(cdsSeteoF, "Select FacturaElectronica From SeteoF", false))
							{
								string sSQLClave = string.Format("Exec FE_ClaveDeAcceso {0}, 1", idCompra);
								string sClaveAcceso = Funcion.sEscalarSQL(cdsSeteoF, sSQLClave, false).ToString();

								string cadInvertida = invertirCadena(sClaveAcceso.Substring(0, 48));
								string ValM11 = Modulo11(cadInvertida).ToString().Trim();
								string sClaveSRI = sClaveAcceso.Substring(0, 48) + ValM11;

								string sSQLUpdate = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}",	sClaveSRI, idCompra);

								Funcion.EjecutaSQL(cdsSeteoF, sSQLUpdate, false);
							}
							#endregion Autorizacion Factura

							#region Nombre Impresora
							string sImpresora = "";
							sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",(int) this.cmbBodega.Value));		
							#endregion Nombre Impresora

							#region Impresion Formato
							if(bImpresionDivision)
							{
								ImpresionPV(sImpresora);
								ImpresionPV2(sImpresora);
								ImpresionPV3(sImpresora);
								ImpresionPV4(sImpresora);
							}
							else
							{
								ImpresionSinDivision(sImpresora);
								ImpresionSinDivision1(sImpresora);
							}
							#endregion Impresion Formato
						
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set Impreso = 1 Where idCompra = {0}", idCompra));
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Habitaciones Set Ocupado = 10 Where idHabitacion = {0}", idHabitacion));
						}
					}
				}
			}
		}

		private void ImpresionSinDivision(string sImpresora)
		{
			#region Factura sin Division Impresion Directa
			#region Impresion Greta
			//			if (MenuLatinium.stIdDB == 29 || MenuLatinium.stIdDB == 36)
			//			{
			//				string sReporte = "FacturaLottePV.rpt";
			//				ReportDocument Report1 = new ReportDocument();
			//				Report1.FileName = MenuLatinium.stDirReportes + sReporte;
			//													
			//				ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
			//			{
			//				string strBaseDato = MenuLatinium.stDirNumero;
			//				string strUsuario = MenuLatinium.stDirConta;
			//				string strServidor = MenuLatinium.stDirServidor;
			//				string strImpresora = sImpresora; //MenuLatinium.ImpresoraNombre;
			//					
			//				crConnectionInfo.ServerName = strServidor;
			//				crConnectionInfo.DatabaseName = strBaseDato;
			//				crConnectionInfo.UserID = "infoelect";
			//				crConnectionInfo.Password = "Bl45o6$9";
			//			}										
			//								
			//				TableLogOnInfo crTableLogOnInfo; 
			//				Database crDatabase = Report1.Database; 
			//				Tables crTables = crDatabase.Tables; 
			//				Table crTable;
			//					
			//				for (int i=0; i< crTables.Count; i++)
			//				{
			//				{
			//					crTable = crTables [i]; 
			//					crTableLogOnInfo = crTable.LogOnInfo; 
			//					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
			//					crTable.ApplyLogOnInfo(crTableLogOnInfo);
			//				}
			//				}
			//										
			//				Report1.PrintOptions.PrinterName = sImpresora; //MenuLatinium.ImpresoraNombre;
			//				Report1.SetParameterValue("@idCompra", idCompra);
			//					
			//				Report1.PrintToPrinter(1, false, 0, 0);
			//			}
			#endregion Impresion Greta
			#region Impresion Pro
			//			if (MenuLatinium.stIdDB == 34)
			//			{
			//				string sReporte = "FacturaLottePVPro.rpt";
			//				ReportDocument Report1 = new ReportDocument();
			//				Report1.FileName = MenuLatinium.stDirReportes + sReporte;
			//													
			//				ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
			//			{
			//				string strBaseDato = MenuLatinium.stDirNumero;
			//				string strUsuario = MenuLatinium.stDirConta;
			//				string strServidor = MenuLatinium.stDirServidor;
			//				string strImpresora = sImpresora; //MenuLatinium.ImpresoraNombre;
			//					
			//				crConnectionInfo.ServerName = strServidor;
			//				crConnectionInfo.DatabaseName = strBaseDato;
			//				crConnectionInfo.UserID = "infoelect";
			//				crConnectionInfo.Password = "Bl45o6$9";
			//			}										
			//								
			//				TableLogOnInfo crTableLogOnInfo; 
			//				Database crDatabase = Report1.Database; 
			//				Tables crTables = crDatabase.Tables; 
			//				Table crTable;
			//					
			//				for (int i=0; i< crTables.Count; i++)
			//				{
			//				{
			//					crTable = crTables [i]; 
			//					crTableLogOnInfo = crTable.LogOnInfo; 
			//					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
			//					crTable.ApplyLogOnInfo(crTableLogOnInfo);
			//				}
			//				}
			//										
			//				Report1.PrintOptions.PrinterName = sImpresora; //MenuLatinium.ImpresoraNombre;
			//				Report1.SetParameterValue("@idCompra", idCompra);
			//					
			//				Report1.PrintToPrinter(1, false, 0, 0);
			//			}

			if (MenuLatinium.stIdDB == 34)
			{
				ParameterFields paramFields = new ParameterFields();

				ParameterField paramFieldIdCompra = new ParameterField();
				ParameterDiscreteValue discreteValIdCompra = new ParameterDiscreteValue();
				paramFieldIdCompra.ParameterFieldName = "@idCompra";
				discreteValIdCompra.Value = idCompra; 
				paramFieldIdCompra.CurrentValues.Add(discreteValIdCompra);
				paramFields.Add(paramFieldIdCompra);

				string sReporte = "FacturaLottePVPro.rpt";

				ReportDocument Report1 = new ReportDocument();
				Report1.Load(MenuLatinium.stDirReportes + sReporte);

				ConnectionInfo crConnectionInfo = new ConnectionInfo();
				crConnectionInfo.ServerName = MenuLatinium.stDirServidor;
				crConnectionInfo.DatabaseName = MenuLatinium.stDirNumero;
				crConnectionInfo.UserID = "infoelect";
				crConnectionInfo.Password = "Bl45o6$9";

				Tables crTables = Report1.Database.Tables;
				for (int i = 0; i < crTables.Count; i++)
				{
					Table crTable = crTables[i];
					try
					{
						TableLogOnInfo crTableLogOnInfo = crTable.LogOnInfo;
						crTableLogOnInfo.ConnectionInfo = crConnectionInfo;
						crTable.ApplyLogOnInfo(crTableLogOnInfo);
					}
					catch (Exception exTabla)
					{
						throw new Exception(string.Format("Error aplicando conexión a tabla {0}: {1}", crTable.Name, exTabla.Message), exTabla);
					}
				}

				if (sImpresora != null && sImpresora != "")
				{
					Report1.PrintOptions.PrinterName = sImpresora;
				}
				else
				{
					throw new Exception("Nombre de impresora vacío o inválido.");
				}
				Report1.SetParameterValue("@idCompra", idCompra);

				Report1.PrintToPrinter(2, false, 0, 0);
			}
				
			#endregion Impresion Pro
			#endregion Factura sin Division IMpresion Directa

			#region Factura sin Division Impresion Pantalla
			#region Greta
			if(MenuLatinium.stIdDB == 29)
			{
//				ParameterFields paramFields = new ParameterFields ();
//				ParameterField paramFieldIdDetCompra = new ParameterField ();
//				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
//				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
//				discreteValIdDetCompra.Value = idCompra; 
//				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
//				paramFields.Add (paramFieldIdDetCompra);
//				
//				PrintDocument pd = new PrintDocument();
//				pd.PrinterSettings.PrinterName = sImpresora;
//				Reporte Reporte = new Reporte("FacturaLottePV.rpt", "");
//				Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
//				Reporte.MdiParent = this.MdiParent;
//				Reporte.Titulo("Comprobante Venta");		 
//				Reporte.crVista.ParameterFieldInfo = paramFields;
//				if (bMenuCuadricula) Reporte.ShowDialog();
//				if (!bMenuCuadricula) Reporte.Show();
			}

			#region Marca como impreso
			string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + idCompra;
			Funcion.EjecutaSQL(cdsSeteoF, stGrabaImprime);
			#endregion Marca como impreso

			#endregion Greta
			#region Pro
			if(MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
			{
//				ParameterFields paramFields = new ParameterFields ();
//				ParameterField paramFieldIdDetCompra = new ParameterField ();
//				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
//				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
//				discreteValIdDetCompra.Value = idCompra; 
//				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
//				paramFields.Add (paramFieldIdDetCompra);
//				
//				PrintDocument pd = new PrintDocument();
//				pd.PrinterSettings.PrinterName = sImpresora;
//				Reporte Reporte = new Reporte("FacturaLottePVPro.rpt", "");
//				Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
//				Reporte.MdiParent = this.MdiParent;
//				Reporte.Titulo("Comprobante Venta");		 
//				Reporte.crVista.ParameterFieldInfo = paramFields;
//				if (bMenuCuadricula) Reporte.ShowDialog();
//				if (!bMenuCuadricula) Reporte.Show();


			}

			#region Marca como impreso
			string stGrabaImprimePro = "Update Compra Set Impreso = 1 Where idCompra = " + idCompra;
			Funcion.EjecutaSQL(cdsSeteoF, stGrabaImprimePro);
			#endregion Marca como impreso

			#endregion Pro
			#endregion Factura sin Division Impresion Pantalla
		}

		private void ImpresionSinDivision1(string sImpresora)
		{
			#region Factura sin Division Impresion Directa
			#region Impresion Greta
			//			if (MenuLatinium.stIdDB == 29 || MenuLatinium.stIdDB == 36)
			//			{
			//				string sReporte = "FacturaLottePV.rpt";
			//				ReportDocument Report1 = new ReportDocument();
			//				Report1.FileName = MenuLatinium.stDirReportes + sReporte;
			//													
			//				ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
			//			{
			//				string strBaseDato = MenuLatinium.stDirNumero;
			//				string strUsuario = MenuLatinium.stDirConta;
			//				string strServidor = MenuLatinium.stDirServidor;
			//				string strImpresora = sImpresora; //MenuLatinium.ImpresoraNombre;
			//					
			//				crConnectionInfo.ServerName = strServidor;
			//				crConnectionInfo.DatabaseName = strBaseDato;
			//				crConnectionInfo.UserID = "infoelect";
			//				crConnectionInfo.Password = "Bl45o6$9";
			//			}										
			//								
			//				TableLogOnInfo crTableLogOnInfo; 
			//				Database crDatabase = Report1.Database; 
			//				Tables crTables = crDatabase.Tables; 
			//				Table crTable;
			//					
			//				for (int i=0; i< crTables.Count; i++)
			//				{
			//				{
			//					crTable = crTables [i]; 
			//					crTableLogOnInfo = crTable.LogOnInfo; 
			//					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
			//					crTable.ApplyLogOnInfo(crTableLogOnInfo);
			//				}
			//				}
			//										
			//				Report1.PrintOptions.PrinterName = sImpresora; //MenuLatinium.ImpresoraNombre;
			//				Report1.SetParameterValue("@idCompra", idCompra);
			//					
			//				Report1.PrintToPrinter(1, false, 0, 0);
			//			}
			#endregion Impresion Greta

			#region Impresion Pro
			//			if (MenuLatinium.stIdDB == 34)
			//			{
			//				string sReporte = "FacturaLottePVPro.rpt";
			//				ReportDocument Report1 = new ReportDocument();
			//				Report1.FileName = MenuLatinium.stDirReportes + sReporte;
			//													
			//				ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
			//			{
			//				string strBaseDato = MenuLatinium.stDirNumero;
			//				string strUsuario = MenuLatinium.stDirConta;
			//				string strServidor = MenuLatinium.stDirServidor;
			//				string strImpresora = sImpresora; //MenuLatinium.ImpresoraNombre;
			//					
			//				crConnectionInfo.ServerName = strServidor;
			//				crConnectionInfo.DatabaseName = strBaseDato;
			//				crConnectionInfo.UserID = "infoelect";
			//				crConnectionInfo.Password = "Bl45o6$9";
			//			}										
			//								
			//				TableLogOnInfo crTableLogOnInfo; 
			//				Database crDatabase = Report1.Database; 
			//				Tables crTables = crDatabase.Tables; 
			//				Table crTable;
			//					
			//				for (int i=0; i< crTables.Count; i++)
			//				{
			//				{
			//					crTable = crTables [i]; 
			//					crTableLogOnInfo = crTable.LogOnInfo; 
			//					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
			//					crTable.ApplyLogOnInfo(crTableLogOnInfo);
			//				}
			//				}
			//										
			//				Report1.PrintOptions.PrinterName = sImpresora; //MenuLatinium.ImpresoraNombre;
			//				Report1.SetParameterValue("@idCompra", idCompra);
			//					
			//				Report1.PrintToPrinter(1, false, 0, 0);
			//			}
			#endregion Impresion Pro
			#endregion Factura sin Division Impresion Directa

			#region Factura sin Division Impresion Pantalla
			#region Greta
			if(MenuLatinium.stIdDB == 29)
			{
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = idCompra; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				PrintDocument pd = new PrintDocument();
				pd.PrinterSettings.PrinterName = sImpresora;
				Reporte Reporte = new Reporte("FacturaLottePV.rpt", "");
				Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Comprobante Venta");		 
				Reporte.crVista.ParameterFieldInfo = paramFields;
				if (bMenuCuadricula) Reporte.ShowDialog();
				if (!bMenuCuadricula) Reporte.Show();
			}
			#endregion Greta
			#region Pro
			if(MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
			{
//				ParameterFields paramFields = new ParameterFields ();
//				ParameterField paramFieldIdDetCompra = new ParameterField ();
//				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
//				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
//				discreteValIdDetCompra.Value = idCompra; 
//				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
//				paramFields.Add (paramFieldIdDetCompra);
//				
//				PrintDocument pd = new PrintDocument();
//				pd.PrinterSettings.PrinterName = sImpresora;
//				Reporte Reporte = new Reporte("FacturaLottePVPro.rpt", "");
//				Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
//				Reporte.MdiParent = this.MdiParent;
//				Reporte.Titulo("Comprobante Venta");		 
//				Reporte.crVista.ParameterFieldInfo = paramFields;
//				if (bMenuCuadricula) Reporte.ShowDialog();
//				if (!bMenuCuadricula) Reporte.Show();
			}
			#endregion Pro
			#endregion Factura sin Division Impresion Pantalla
		}

		private void ImpresionPV(string sImpresora)
		{
			#region Impresion Pantalla
			//					ParameterFields paramFields = new ParameterFields ();
			//				
			//					ParameterField pfRet = new ParameterField ();
			//					ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
			//					pfRet.ParameterFieldName = "@idCompra";
			//					idCompraRet.Value = iCompraOriginal;
			//					pfRet.CurrentValues.Add (idCompraRet);
			//					paramFields.Add (pfRet);
			//				
			//					Reporte miRepor = new Reporte("FacturaLottePV.rpt", "");
			//					miRepor.MdiParent = this.MdiParent;
			//					miRepor.Titulo("");
			//					miRepor.iNumeroCopias = 2;
			//					miRepor.crVista.ParameterFieldInfo = paramFields;
			//					if (bMenuCuadricula) miRepor.ShowDialog();
			//
			//					ParameterFields paramFields1 = new ParameterFields ();
			//
			//					ParameterField pfRet1 = new ParameterField ();
			//					ParameterDiscreteValue idCompraRet1 = new ParameterDiscreteValue ();
			//					pfRet1.ParameterFieldName = "@idCompra";
			//					idCompraRet1.Value = iCompraDividida;
			//					pfRet1.CurrentValues.Add (idCompraRet1);
			//					paramFields1.Add (pfRet1);
			//				
			//					Reporte miReporDivision = new Reporte("FacturaLottePV.rpt", "");
			//					miReporDivision.MdiParent = this.MdiParent;
			//					miReporDivision.Titulo("");
			//					miRepor.iNumeroCopias = 2;
			//					miReporDivision.crVista.ParameterFieldInfo = paramFields1;
			//					if (bMenuCuadricula) miReporDivision.ShowDialog();
			#endregion Impresion Pantalla
			
			if (MenuLatinium.stIdDB == 29)
			{
				#region Factura 1 Impresion Directa
				//				string sReporte = "FacturaLottePV.rpt";
				//				ReportDocument Report1 = new ReportDocument();
				//				Report1.FileName = MenuLatinium.stDirReportes + sReporte;
				//													
				//				ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
				//			{
				//				string strBaseDato = MenuLatinium.stDirNumero;
				//				string strUsuario = MenuLatinium.stDirConta;
				//				string strServidor = MenuLatinium.stDirServidor;
				//				string strImpresora = sImpresora; //MenuLatinium.ImpresoraNombre;
				//					
				//				crConnectionInfo.ServerName = strServidor;
				//				crConnectionInfo.DatabaseName = strBaseDato;
				//				crConnectionInfo.UserID = "infoelect";
				//				crConnectionInfo.Password = "Bl45o6$9";
				//			}										
				//								
				//				TableLogOnInfo crTableLogOnInfo; 
				//				Database crDatabase = Report1.Database; 
				//				Tables crTables = crDatabase.Tables; 
				//				Table crTable;
				//					
				//				for (int i=0; i< crTables.Count; i++)
				//				{
				//				{
				//					crTable = crTables [i]; 
				//					crTableLogOnInfo = crTable.LogOnInfo; 
				//					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
				//					crTable.ApplyLogOnInfo(crTableLogOnInfo);
				//				}
				//				}
				//										
				//				Report1.PrintOptions.PrinterName = sImpresora; //MenuLatinium.ImpresoraNombre;
				//				Report1.SetParameterValue("@idCompra", iCompraOriginal);
				//					
				//				Report1.PrintToPrinter(1, false, 0, 0);
				#endregion Factura 1 Impresion Directa

				#region Factura 1 Impresion Pantalla
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = iCompraOriginal; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				PrintDocument pd = new PrintDocument();
				pd.PrinterSettings.PrinterName = sImpresora;
				Reporte Reporte = new Reporte("FacturaLottePV.rpt", "");
				Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Comprobante Venta");		 
				Reporte.crVista.ParameterFieldInfo = paramFields;
				if (bMenuCuadricula) Reporte.ShowDialog();
				if (!bMenuCuadricula) Reporte.Show();
				#endregion Factura 1 Impresion Pantalla

				#region Marca como impreso
				string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + iCompraOriginal;
				Funcion.EjecutaSQL(cdsSeteoF, stGrabaImprime);
				#endregion Marca como impreso
			}

			try
			{
				// Tu código original aquí
				if (MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
				{

					#region Factura 1 Impresion Directa
					//				string sReporte = "FacturaLottePV.rpt";
					//				ReportDocument Report1 = new ReportDocument();
					//				Report1.FileName = MenuLatinium.stDirReportes + sReporte;
					//													
					//				ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
					//			{
					//				string strBaseDato = MenuLatinium.stDirNumero;
					//				string strUsuario = MenuLatinium.stDirConta;
					//				string strServidor = MenuLatinium.stDirServidor;
					//				string strImpresora = sImpresora; //MenuLatinium.ImpresoraNombre;
					//					
					//				crConnectionInfo.ServerName = strServidor;
					//				crConnectionInfo.DatabaseName = strBaseDato;
					//				crConnectionInfo.UserID = "infoelect";
					//				crConnectionInfo.Password = "Bl45o6$9";
					//			}										
					//								
					//				TableLogOnInfo crTableLogOnInfo; 
					//				Database crDatabase = Report1.Database; 
					//				Tables crTables = crDatabase.Tables; 
					//				Table crTable;
					//					
					//				for (int i=0; i< crTables.Count; i++)
					//				{
					//				{
					//					crTable = crTables [i]; 
					//					crTableLogOnInfo = crTable.LogOnInfo; 
					//					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
					//					crTable.ApplyLogOnInfo(crTableLogOnInfo);
					//				}
					//				}
					//										
					//				Report1.PrintOptions.PrinterName = sImpresora; //MenuLatinium.ImpresoraNombre;
					//				Report1.SetParameterValue("@idCompra", iCompraOriginal);
					//					
					//				Report1.PrintToPrinter(1, false, 0, 0);
					#endregion Factura 1 Impresion Directa

					#region Factura 1 Impresion Pantalla
//					ParameterFields paramFields = new ParameterFields();
//					ParameterField paramFieldIdDetCompra = new ParameterField();
//					ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue();
//					paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
//					discreteValIdDetCompra.Value = iCompraOriginal;
//					paramFieldIdDetCompra.CurrentValues.Add(discreteValIdDetCompra);
//					paramFields.Add(paramFieldIdDetCompra);
//        
//					PrintDocument pd = new PrintDocument();
//					pd.PrinterSettings.PrinterName = sImpresora;
//					Reporte Reporte = new Reporte("FacturaLottePVPro.rpt", "");
//					Reporte.oRpt.PrintOptions.PrinterName = sImpresora;
//					Reporte.MdiParent = this.MdiParent;
//					Reporte.Titulo("Comprobante Venta");
//					Reporte.crVista.ParameterFieldInfo = paramFields;
//        
//					if (bMenuCuadricula) 
//						Reporte.ShowDialog();
//					if (!bMenuCuadricula) 
//						Reporte.Show();
					#endregion Factura 1 Impresion Pantalla

					#region Marca como impreso
					string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + iCompraOriginal;
					Funcion.EjecutaSQL(cdsSeteoF, stGrabaImprime);
					#endregion Marca como impreso
				}
			}
			catch (SqlException sqlEx)
			{
				// Aquí puedes capturar excepciones específicas de SQL, como problemas con la base de datos
				Console.WriteLine("Error en la base de datos: " + sqlEx.Message);
				// También podrías mostrar un mensaje de error al usuario o loguear el error en un archivo
			}
			finally
			{
				// Código que se ejecuta siempre, sin importar si hubo un error o no
				// Por ejemplo, puedes cerrar conexiones o liberar recursos aquí
				Console.WriteLine("Proceso finalizado.");
			}

		}

		private void ImpresionPV2(string sImpresora)
		{
			if (MenuLatinium.stIdDB == 29)
			{
				#region Factura 1 reimprimir Impresion Directa
				//				string sReporte2 = "FacturaLottePV.rpt";
				//				ReportDocument Report2 = new ReportDocument();
				//				Report2.FileName = MenuLatinium.stDirReportes + sReporte2;
				//													
				//				ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
				//			{
				//				string strBaseDato = MenuLatinium.stDirNumero;
				//				string strUsuario = MenuLatinium.stDirConta;
				//				string strServidor = MenuLatinium.stDirServidor;
				//				string strImpresora = sImpresora; //MenuLatinium.ImpresoraNombre;
				//					
				//				crConnectionInfo.ServerName = strServidor;
				//				crConnectionInfo.DatabaseName = strBaseDato;
				//				crConnectionInfo.UserID = "infoelect";
				//				crConnectionInfo.Password = "Bl45o6$9";
				//			}										
				//								
				//				TableLogOnInfo crTableLogOnInfo; 
				//				Database crDatabase = Report2.Database; 
				//				Tables crTables = crDatabase.Tables; 
				//				Table crTable;
				//					
				//				for (int i=0; i< crTables.Count; i++)
				//				{
				//				{
				//					crTable = crTables [i]; 
				//					crTableLogOnInfo = crTable.LogOnInfo; 
				//					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
				//					crTable.ApplyLogOnInfo(crTableLogOnInfo);
				//				}
				//				}
				//										
				//				Report2.PrintOptions.PrinterName = sImpresora; //MenuLatinium.ImpresoraNombre;
				//				Report2.SetParameterValue("@idCompra", iCompraOriginal);
				//					
				//				Report2.PrintToPrinter(1, false, 0, 0);
				#endregion Factura 1 reimprimir	

				#region Factura 1 reimprimir Impresion Pantalla
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = iCompraOriginal; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				PrintDocument pd = new PrintDocument();
				pd.PrinterSettings.PrinterName = sImpresora;
				Reporte Reporte = new Reporte("FacturaLottePV.rpt", "");
				Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Comprobante Venta");		 
				Reporte.crVista.ParameterFieldInfo = paramFields;
				if (bMenuCuadricula) Reporte.ShowDialog();
				if (!bMenuCuadricula) Reporte.Show();
				#endregion Factura 1 reimprimir Impresion Pantalla
			}
			if (MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
			{
				#region Factura 1 reimprimir Impresion Directa
				//				string sReporte2 = "FacturaLottePVPro.rpt";
				//				ReportDocument Report2 = new ReportDocument();
				//				Report2.FileName = MenuLatinium.stDirReportes + sReporte2;
				//													
				//				ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
				//			{
				//				string strBaseDato = MenuLatinium.stDirNumero;
				//				string strUsuario = MenuLatinium.stDirConta;
				//				string strServidor = MenuLatinium.stDirServidor;
				//				string strImpresora = sImpresora; //MenuLatinium.ImpresoraNombre;
				//					
				//				crConnectionInfo.ServerName = strServidor;
				//				crConnectionInfo.DatabaseName = strBaseDato;
				//				crConnectionInfo.UserID = "infoelect";
				//				crConnectionInfo.Password = "Bl45o6$9";
				//			}										
				//								
				//				TableLogOnInfo crTableLogOnInfo; 
				//				Database crDatabase = Report2.Database; 
				//				Tables crTables = crDatabase.Tables; 
				//				Table crTable;
				//					
				//				for (int i=0; i< crTables.Count; i++)
				//				{
				//				{
				//					crTable = crTables [i]; 
				//					crTableLogOnInfo = crTable.LogOnInfo; 
				//					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
				//					crTable.ApplyLogOnInfo(crTableLogOnInfo);
				//				}
				//				}
				//										
				//				Report2.PrintOptions.PrinterName = sImpresora; //MenuLatinium.ImpresoraNombre;
				//				Report2.SetParameterValue("@idCompra", iCompraOriginal);
				//					
				//				Report2.PrintToPrinter(1, false, 0, 0);
				#endregion Factura 1 reimprimir	Impresion Directa

				#region Factura 1 reimprimir Impresion Pantalla
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = iCompraOriginal; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				PrintDocument pd = new PrintDocument();
				pd.PrinterSettings.PrinterName = sImpresora;
				Reporte Reporte = new Reporte("FacturaLottePVPro.rpt", "");
				Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Comprobante Venta");		 
				Reporte.crVista.ParameterFieldInfo = paramFields;
				if (bMenuCuadricula) Reporte.ShowDialog();
				if (!bMenuCuadricula) Reporte.Show();
				#endregion Factura 1 reimprimir Impresion Pantalla
			}
		}

		private void ImpresionPV3(string sImpresora)
		{
			if (MenuLatinium.stIdDB == 29)
			{
				#region Factura 2 Impresion Directa
				//				string sReporte3 = "FacturaLottePV.rpt";
				//				ReportDocument Report3 = new ReportDocument();
				//				Report3.FileName = MenuLatinium.stDirReportes + sReporte3;
				//													
				//				ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
				//			{
				//				string strBaseDato = MenuLatinium.stDirNumero;
				//				string strUsuario = MenuLatinium.stDirConta;
				//				string strServidor = MenuLatinium.stDirServidor;
				//				string strImpresora = sImpresora; //MenuLatinium.ImpresoraNombre;
				//					
				//				crConnectionInfo.ServerName = strServidor;
				//				crConnectionInfo.DatabaseName = strBaseDato;
				//				crConnectionInfo.UserID = "infoelect";
				//				crConnectionInfo.Password = "Bl45o6$9";
				//			}										
				//								
				//				TableLogOnInfo crTableLogOnInfo; 
				//				Database crDatabase = Report3.Database; 
				//				Tables crTables = crDatabase.Tables; 
				//				Table crTable;
				//					
				//				for (int i=0; i< crTables.Count; i++)
				//				{
				//				{
				//					crTable = crTables [i]; 
				//					crTableLogOnInfo = crTable.LogOnInfo; 
				//					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
				//					crTable.ApplyLogOnInfo(crTableLogOnInfo);
				//				}
				//				}
				//										
				//				Report3.PrintOptions.PrinterName = sImpresora; //MenuLatinium.ImpresoraNombre;
				//				Report3.SetParameterValue("@idCompra", iCompraDividida);
				//
				//				Report3.PrintToPrinter(1, false, 0, 0);
				#endregion Factura 2 Impresion Directa			

				#region Factura 2 Impresion Pantalla
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = iCompraDividida; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				PrintDocument pd = new PrintDocument();
				pd.PrinterSettings.PrinterName = sImpresora;
				Reporte Reporte = new Reporte("FacturaLottePV.rpt", "");
				Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Comprobante Venta");		 
				Reporte.crVista.ParameterFieldInfo = paramFields;
				if (bMenuCuadricula) Reporte.ShowDialog();
				if (!bMenuCuadricula) Reporte.Show();
				#endregion Factura 2 Impresion Pantalla
			}
			if (MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
			{
				#region Factura 2 Impresion Directa
				//				string sReporte3 = "FacturaLottePVPro.rpt";
				//				ReportDocument Report3 = new ReportDocument();
				//				Report3.FileName = MenuLatinium.stDirReportes + sReporte3;
				//													
				//				ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
				//			{
				//				string strBaseDato = MenuLatinium.stDirNumero;
				//				string strUsuario = MenuLatinium.stDirConta;
				//				string strServidor = MenuLatinium.stDirServidor;
				//				string strImpresora = sImpresora; //MenuLatinium.ImpresoraNombre;
				//					
				//				crConnectionInfo.ServerName = strServidor;
				//				crConnectionInfo.DatabaseName = strBaseDato;
				//				crConnectionInfo.UserID = "infoelect";
				//				crConnectionInfo.Password = "Bl45o6$9";
				//			}										
				//								
				//				TableLogOnInfo crTableLogOnInfo; 
				//				Database crDatabase = Report3.Database; 
				//				Tables crTables = crDatabase.Tables; 
				//				Table crTable;
				//					
				//				for (int i=0; i< crTables.Count; i++)
				//				{
				//				{
				//					crTable = crTables [i]; 
				//					crTableLogOnInfo = crTable.LogOnInfo; 
				//					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
				//					crTable.ApplyLogOnInfo(crTableLogOnInfo);
				//				}
				//				}
				//										
				//				Report3.PrintOptions.PrinterName = sImpresora; //MenuLatinium.ImpresoraNombre;
				//				Report3.SetParameterValue("@idCompra", iCompraDividida);
				//
				//				Report3.PrintToPrinter(1, false, 0, 0);
				#endregion Factura 2 Impresion Directa

				#region Factura 2 Impresion Pantalla
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = iCompraDividida; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				PrintDocument pd = new PrintDocument();
				pd.PrinterSettings.PrinterName = sImpresora;
				Reporte Reporte = new Reporte("FacturaLottePVPro.rpt", "");
				Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Comprobante Venta");		 
				Reporte.crVista.ParameterFieldInfo = paramFields;
				if (bMenuCuadricula) Reporte.ShowDialog();
				if (!bMenuCuadricula) Reporte.Show();
				#endregion Factura 2 Impresion Pantalla
			}
		}

		private void ImpresionPV4(string sImpresora)
		{
			if (MenuLatinium.stIdDB == 29)
			{
				#region Factura 2 reimprimir Impresion Directa
				//				string sReporte4 = "FacturaLottePV.rpt";
				//				ReportDocument Report4 = new ReportDocument();
				//				Report4.FileName = MenuLatinium.stDirReportes + sReporte4;
				//													
				//				ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
				//			{
				//				string strBaseDato = MenuLatinium.stDirNumero;
				//				string strUsuario = MenuLatinium.stDirConta;
				//				string strServidor = MenuLatinium.stDirServidor;
				//				string strImpresora = sImpresora; //MenuLatinium.ImpresoraNombre;
				//					
				//				crConnectionInfo.ServerName = strServidor;
				//				crConnectionInfo.DatabaseName = strBaseDato;
				//				crConnectionInfo.UserID = "infoelect";
				//				crConnectionInfo.Password = "Bl45o6$9";
				//			}										
				//								
				//				TableLogOnInfo crTableLogOnInfo; 
				//				Database crDatabase = Report4.Database; 
				//				Tables crTables = crDatabase.Tables; 
				//				Table crTable;
				//					
				//				for (int i=0; i< crTables.Count; i++)
				//				{
				//				{
				//					crTable = crTables [i]; 
				//					crTableLogOnInfo = crTable.LogOnInfo; 
				//					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
				//					crTable.ApplyLogOnInfo(crTableLogOnInfo);
				//				}
				//				}
				//										
				//				Report4.PrintOptions.PrinterName = sImpresora; //MenuLatinium.ImpresoraNombre;
				//				Report4.SetParameterValue("@idCompra", iCompraDividida);
				//
				//				Report4.PrintToPrinter(1, false, 0, 0);
				#endregion Factura 2 reimprimir Impresion Directa

				#region Factura 2 reimprimir Impresion Pantalla
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = iCompraDividida; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				PrintDocument pd = new PrintDocument();
				pd.PrinterSettings.PrinterName = sImpresora;
				Reporte Reporte = new Reporte("FacturaLottePV.rpt", "");
				Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Comprobante Venta");		 
				Reporte.crVista.ParameterFieldInfo = paramFields;
				if (bMenuCuadricula) Reporte.ShowDialog();
				if (!bMenuCuadricula) Reporte.Show();
				#endregion Factura 2 reimprimir Impresion Pantalla
			}
			if (MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
			{
				#region Factura 2 reimprimir Impresion Directa
				//				string sReporte4 = "FacturaLottePVPro.rpt";
				//				ReportDocument Report4 = new ReportDocument();
				//				Report4.FileName = MenuLatinium.stDirReportes + sReporte4;
				//													
				//				ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
				//			{
				//				string strBaseDato = MenuLatinium.stDirNumero;
				//				string strUsuario = MenuLatinium.stDirConta;
				//				string strServidor = MenuLatinium.stDirServidor;
				//				string strImpresora = sImpresora; //MenuLatinium.ImpresoraNombre;
				//					
				//				crConnectionInfo.ServerName = strServidor;
				//				crConnectionInfo.DatabaseName = strBaseDato;
				//				crConnectionInfo.UserID = "infoelect";
				//				crConnectionInfo.Password = "Bl45o6$9";
				//			}										
				//								
				//				TableLogOnInfo crTableLogOnInfo; 
				//				Database crDatabase = Report4.Database; 
				//				Tables crTables = crDatabase.Tables; 
				//				Table crTable;
				//					
				//				for (int i=0; i< crTables.Count; i++)
				//				{
				//				{
				//					crTable = crTables [i]; 
				//					crTableLogOnInfo = crTable.LogOnInfo; 
				//					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
				//					crTable.ApplyLogOnInfo(crTableLogOnInfo);
				//				}
				//				}
				//										
				//				Report4.PrintOptions.PrinterName = sImpresora; //MenuLatinium.ImpresoraNombre;
				//				Report4.SetParameterValue("@idCompra", iCompraDividida);
				//
				//				Report4.PrintToPrinter(1, false, 0, 0);
				#endregion Factura 2 reimprimir Impresion Directa

				#region Factura 2 reimprimir Impresion Pantalla
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = iCompraDividida; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				PrintDocument pd = new PrintDocument();
				pd.PrinterSettings.PrinterName = sImpresora;
				Reporte Reporte = new Reporte("FacturaLottePVPro.rpt", "");
				Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Comprobante Venta");		 
				Reporte.crVista.ParameterFieldInfo = paramFields;
				if (bMenuCuadricula) Reporte.ShowDialog();
				if (!bMenuCuadricula) Reporte.Show();
				#endregion Factura 2 reimprimir Impresion Pantalla
			}
		}

		private void ImpresionCupon(string sImpresora, int iAnticipo)
		{
			#region Impresion Cupon
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idAnticipo";
			discreteValIdDetCompra.Value = iAnticipo; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("AnticipoMitadPrecio.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Comprobante Venta");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula) Reporte.ShowDialog();
			if (!bMenuCuadricula) Reporte.Show();
			#endregion Impresion Cupon
		}

		private void ImpresionCupon2(string sImpresora, int iAnticipo)
		{
			#region Impresion Cupon
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idAnticipo";
			discreteValIdDetCompra.Value = iAnticipo; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("AnticipoMitadPrecio.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Comprobante Venta");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula) Reporte.ShowDialog();
			if (!bMenuCuadricula) Reporte.Show();
			#endregion Impresion Cupon
		}

		private void ImprimeComanda(int iComandaImpr, string sImpresora)
		{
			#region Impresion Comanda Caja
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idComanda";
			discreteValIdDetCompra.Value = iComandaImpr; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("ImpresionComanda.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Impresion Comanda Caja");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula) Reporte.ShowDialog();
			if (!bMenuCuadricula) Reporte.Show();
			#endregion Impresion Comanda Caja			
		}

		private void ImprimeComandaCliente(int iComandaImpr, string sImpresora)
		{
			#region Impresion Comanda Cliente
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idComanda";
			discreteValIdDetCompra.Value = iComandaImpr; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("ImpresionComanda.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Impresion Comanda Cocina");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula) Reporte.ShowDialog();
			if (!bMenuCuadricula) Reporte.Show();
			#endregion Impresion Comanda Cliente	
		}

		private void ImprimeComandaCancelada(int iComandaImpr, string sImpresora)
		{
			#region Impresion Comanda Cancelada
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idDetCompra";
			discreteValIdDetCompra.Value = iComandaImpr; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("ImpresionComandaCancelada.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Impresion Cancelacion Comanda Caja");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula) Reporte.ShowDialog();
			if (!bMenuCuadricula) Reporte.Show();
			#endregion Impresion Comanda Cancelada	
		}
		private void ImprimeComandaCanceladaCliente(int iComandaImpr, string sImpresora)
		{
			#region Impresion Comanda Cancelada
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idDetCompra";
			discreteValIdDetCompra.Value = iComandaImpr; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("ImpresionComandaCancelada.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Impresion Cancelacion Comanda Cocina");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula) Reporte.ShowDialog();
			if (!bMenuCuadricula) Reporte.Show();
			#endregion Impresion Comanda Cancelada	
		}

		private void ImprimeComandaBtnOrdenCaja(int iComandaImpr, string sImpresora)
		{
			#region Impresion Comanda Caja
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
			discreteValIdDetCompra.Value = iComandaImpr; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("ImpresionComandabtnOrden.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Impresion Comanda Caja");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula)
			{
				Reporte.ShowDialog();  
				Reporte.Dispose();
			}
			else
			{
				Reporte.Show();
			}
			#endregion Impresion Comanda Caja
		}

		private void ImprimeComandaBtnOrdenCliente(int iComandaImpr, string sImpresora)
		{
			#region Impresion Comanda Cliente
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
			discreteValIdDetCompra.Value = iComandaImpr; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("ImpresionComandabtnOrden.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Impresion Comanda Cocina");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
//			if (bMenuCuadricula) Reporte.ShowDialog();
//			if (!bMenuCuadricula) Reporte.Show();
			if (bMenuCuadricula)
			{
				Reporte.ShowDialog(); 
				Reporte.Dispose();
			}
			else
			{
				Reporte.Show();
			}
			#endregion Impresion Comanda Cliente
		}
		private void ImprimeComandaAutDescCaja(int iComandaImpr, string sImpresora)
		{
			#region Impresion Comanda Autoconsumo | Descuento Empleado Caja
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
			discreteValIdDetCompra.Value = iComandaImpr; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("ImpresionComandaAutDesc.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Impresion Comanda Caja");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula) Reporte.ShowDialog();
			if (!bMenuCuadricula) Reporte.Show();
			#endregion Impresion Comanda Autoconsumo | Descuento Empleado Caja
		}

		private void ImprimeComandaAutDescCliente(int iComandaImpr, string sImpresora)
		{
			#region Impresion Comanda Autoconsumo | Descuento Empleado Cliente
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
			discreteValIdDetCompra.Value = iComandaImpr; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("ImpresionComandaAutDesc.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Impresion Comanda Cocina");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula) Reporte.ShowDialog();
			if (!bMenuCuadricula) Reporte.Show();
			#endregion Impresion Comanda Autoconsumo | Descuento Empleado Cliente
		}


		private void ValidaAnticiposClientes()
		{
			#region Anticipos
			bUsoAnticipoCliente = false;
			dAnticipoCliente = 0;			
			
			dAnticipoCliente = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ValidaSaldoAnticipos {0}, 1", idCliente));
				
			if (dAnticipoCliente > 0)
			{
				if(DialogResult.Yes == MessageBox.Show(string.Format("El cliente tiene un saldo a favor de {0} dólares, ¿Desea cruzarlo?", dAnticipoCliente), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					bUsoAnticipoCliente = true;
				else
				{
					bUsoAnticipoCliente = false;
					dAnticipoCliente = 0.00m;
				}
			}
			#endregion Anticipos
		}

		int idAnticipoRP = 0;

		private void ValidaPromoReturnPass()
		{
			#region Anticipos
			dValorPromActiva = 0.00m;
			bUsoPromActiva = false;		
			idAnticipoRP = 0;

			if(idCompra > 0) idAnticipoRP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT CrediDigital FROM COMPRA WHERE IDCOMPRA = {0}", idCompra));
 
			if (idAnticipoRP > 0)
			{
				dValorPromActiva = 19.00m;
				bUsoPromActiva = true;

				return;
			}

			decimal dPrecioHabitacion = Funcion.decEscalarSQL(cdsSeteoF, string.Format("SELECT PrecioNormal FROM Habitaciones WHERE idHabitacion = {0}", (int)this.cmbHabitaciones.Value));
      
			if (dPrecioHabitacion == 35.00m)
			{			
				string sSQL = string.Format("Exec ValidaSaldoReturnPass {0}", idCliente);
				SqlDataReader dReader = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
							
				dReader.Read();
				if (dReader.HasRows)
				{				
					idAnticipoRP = dReader.GetInt32(0);
					dValorPromActiva = dReader.GetDecimal(1);
				}
				dReader.Close();		
	
				if (dValorPromActiva > 0)
				{
					if(DialogResult.Yes == MessageBox.Show("El Cliente Compro La Promocion RETURN PASS, ¿Desea Aplicarla?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					{
						bUsoPromActiva = true;

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
						{								
							if (Convert.ToInt32(dr.Cells["idArticulo"].Value) == 3422)
							{
								dr.Cells["DescuentoPorc"].Value = 45.69;

								this.Total();
							}
						}

						Funcion.EjecutaSQL(cdsSeteoF, string.Format("UPDATE COMPRA SET CrediDigital = {1} WHERE IDCOMPRA = {0}", idCompra, idAnticipoRP));
					}
					else
					{
						bUsoPromActiva = false;
						dValorPromActiva = 0.00m;
					}
				}
			}
			#endregion Anticipos
		}

		public string invertirCadena(String cadena)
		{
			String cadenaInvertida = "";
			for (int x = cadena.Length - 1; x >= 0; x--)
				cadenaInvertida = cadenaInvertida + cadena.Substring(x, 1);

			return cadenaInvertida;
		}

		public string Modulo11(string number)
		{
			int pivote = 2;
			int longitudCadena = number.Length;
			int cantidadTotal = 0;
			int b = 1;

			for (int i = 0; i < longitudCadena; i++)
			{
				if (pivote == 8) pivote = 2;

				int temporal = int.Parse(number.Substring(i, b));

				temporal *= pivote;

				pivote++;

				cantidadTotal += temporal;
			}

			cantidadTotal = 11 - cantidadTotal % 11;

			if (cantidadTotal == 11) cantidadTotal = 0;
			else if (cantidadTotal == 10) cantidadTotal = 1;

			return cantidadTotal.ToString();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			this.Total();

			bool bFiscal = true;
						
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;

			if ((int)this.cmbFormaPago.Value == 1 || (int)this.cmbFormaPago.Value == 3 || (int)this.cmbFormaPago.Value == 16 || (int)this.cmbFormaPago.Value == 8)
				if (!Validacion.vbComboVacio(this.cmbHabitaciones, "Seleccione una habitación")) return;

			if (!Validacion.vbComboVacio(this.cmbGuardia, "Seleccione un guardia")) return;
			
			if (!Validacion.vbComboVacio(this.cmbCobrador, "Seleccione el cobrador")) return;

			if (!Validacion.vbComboVacio(this.cmbCajero, "Seleccione el cajero")) return;

			if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione el tipo de identificación del cliente")) return;
									
			if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;
						
			if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, (int)this.cmbFormaPago.Value, cdsSeteoF)) return;

			if (!Validacion.vbComboVacio(this.cmbFormaPago, "Seleccione la forma de pago")) return;
			
			#region Valida Empleado Activo Autoconsumo
			if (!bMenuCuadricula)
			{
				if ((int)this.cmbFormaPago.Value == 30)
				{
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) from Nomina Where NIdentificacion = '{0}' And Estado <> 2", this.txtRuc.Text.ToString())) == 0)
					{
						MessageBox.Show(string.Format("Para '{0}' debe ser un empleado activo", this.cmbFormaPago.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
			}
			#endregion Valida Empleado Activo Autoconsumo

			if (!Validacion.ValidaTotales((decimal)this.txtTotal.Value)) return;

			if (!Validacion.vbGridVacio(this.uGridDetalle, "Ingrese los items de la factura", "Codigo")) return;
									
			#region Valida registro de servicio de hospedaje pasado los 15 minutos
			if (!bFacturaExtra)
			{
				if (bMenuCuadricula && !bServicioPagado)
				{
					int iMinutos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(MINUTE, '{0}', '{1}')", 
						Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"), Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd HH:mm")));

					if (iMinutos > 15)
					{
						int iServHosp = 0;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
							if ((int)dr.Cells["idArticulo"].Value == 3422) iServHosp++;
				
						if (iServHosp == 0)
						{
							MessageBox.Show("Pasado 15 minutos debe registrar el servicio de habitación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtBuscar.Focus();
							return;
						}
					}
				}
			}
			#endregion Valida registro de servicio de hospedaje pasado los 15 minutos
						
			#region Valida Valor Maximo Consumidor Final
			bool bDivideFactura = false;

			if((int)this.cmbTipoRuc.Value == 11)
			{
				if (Convert.ToDecimal(this.txtTotal.Value) > 50.00m)
				{
					if (!bMenuCuadricula)
					{
						MessageBox.Show("La factura supera los $50.00 \n\nIngrese datos del cliente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbTipoRuc.Focus();
						return;
					}
					else if (bMenuCuadricula)
					{
						if (Convert.ToDecimal(this.txtTotal.Value) > 100.00m)
						{
							MessageBox.Show("La factura supera los $100.00\n\nNO SE PUEDE DIVIDIR\n\nIngrese datos del cliente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.cmbTipoRuc.Focus();
							return;
						}
						else if (DialogResult.Yes == MessageBox.Show("La factura supera los $50\n\n¿Desea dividirla para no registrar los datos del cliente?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
						{
							bDivideFactura = true;		
						}
					}
				}
			}
			#endregion Valida Valor Maximo Consumidor Final

			#region Variables
			iEstado = 10;
			int idTipoRuc = 0;
			if (this.cmbTipoRuc.ActiveRow != null) idTipoRuc = (int)this.cmbTipoRuc.Value;			
			int idCobrador = 0;
			if (this.cmbCobrador.ActiveRow != null) idCobrador = (int)this.cmbCobrador.Value;
			int idCajero = 0;
			if (this.cmbCajero.ActiveRow != null) idCajero = (int)this.cmbCajero.Value;
			int idFormaPago = 0;
			if (this.cmbFormaPago.ActiveRow != null) idFormaPago = (int)this.cmbFormaPago.Value;

			DateTime dtFechaSalida = DateTime.Today;
			bool bFechaSalida = false;
			int idTurnoSalida = 0;
			if (this.dtFechaSalida.Value != System.DBNull.Value)
			{
				bFechaSalida = true;
				dtFechaSalida = Convert.ToDateTime(this.dtFechaSalida.Value);
			}
			if (this.optTurnoSalida.CheckedIndex > -1) idTurnoSalida = (int)this.optTurnoSalida.Value;
			if (bFacturaExtra) idCompra = 0;			
			#endregion Variables

			#region Check Cupon
			if ((bool)this.chkCupon.Checked)
			{
				if (this.txtCupon.Text.Length == 0)
				{
					MessageBox.Show("Ingrese el Codigo del Cupon.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
					this.txtCupon.Focus();
					return;
				}				
						
				#region Variables
				int iValCompraCupon = 0;
				int iValCodigoCupon = 0;
				#endregion Variables

				#region Validaciones
				iValCompraCupon = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesAnticipoCompra Where idCompraDivision = {0}", idCompra));
				iValCodigoCupon = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesAnticipoCompra Where Codigo = '{0}' And Estado = 0", this.txtCupon.Text.ToString()));				
				
				if (iValCompraCupon >= 1)
				{
					MessageBox.Show("Codigo cupon registrado en otra factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCupon.Focus();
					return;
				}
				if (iValCodigoCupon == 0)
				{
					MessageBox.Show("El codigo del cupon ya fue registrado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCupon.Focus();
					return;
				}				
				#endregion Validaciones

				#region Valores
				bUsoCupon = true;
				dValorCupon = Math.Round(Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select PrecioNormal From Habitaciones Where idHabitacion = '{0}'", (int)this.cmbHabitaciones.Value)), 2) / 2;
				#endregion Valores
			}
			#endregion Check Cupon
						
			decimal dCobroAnticipado = 0;
			
			if (bCobroAnticipado) 
				dCobroAnticipado = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.SaldoCobroAnticipado({0})", idCompra));

			using (VentaFPagoLot miCobro = new VentaFPagoLot(idCompra, (int)this.cmbFormaPago.Value, (decimal)this.txtTotal.Value, "", false, 0, 0, false, 0, bUsoAnticipoCliente, dAnticipoCliente, false, 0, 0, 1, idCliente, 0, 0, 0, 0, dSaldo, vValFPago, iCountFP, idOriginal, iTipoCLiente, this.chkAbono.Checked, bUsoGiftCard, dGiftCard, bUsoCupon, dValorCupon, bCobroAnticipado, dCobroAnticipado, bUsoPromActiva, dValorPromActiva, idAnticipoRP))
			{
				if (DialogResult.OK == miCobro.ShowDialog())
				{
					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
					{
						oConexion.Open();

						SqlTransaction oTransaction = null;
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandType = CommandType.Text;
						oCmdActualiza.CommandTimeout = 120;
						
						try
						{
							oTransaction = oConexion.BeginTransaction();
							oCmdActualiza.Transaction = oTransaction;

							#region Numeracion
							if (bNuevo && MenuLatinium.IdUsuario != 279)
							{
								string sSQLNumero = string.Format("Exec NumeracionLocales 1, {0}, 1", (int)this.cmbBodega.Value);
								oCmdActualiza.CommandText = sSQLNumero;
								this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
							}
							#endregion Numeracion

							#region Encabezado
							string sSQL = string.Format("Exec GuardaMaestroFacturas {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', {8}, '{9}', {10}, '{11}', {12}, {13}, {14}, '{15}', '{16}', {17}, '{18}', {19}, {20}, {21}, {22}, {23}, '{24}', {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, '{34}', '{35}', {36}, {37}, {38}",
								idCompra, idReserva, 2, iEstado, (int)this.cmbBodega.Value, idProyecto, (int)this.cmbHabitaciones.Value, //6
								Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbGuardia.Value, this.txtNumero.Text.ToString(), (int)this.txtTiempo.Value, Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"), (int)this.optTurnoIngreso.Value, //12
								idCliente, idTipoRuc, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(), bFechaSalida, dtFechaSalida.ToString("yyyyMMdd HH:mm"), idTurnoSalida, //19
								idCobrador, idCajero, (int)this.txtUnidad.Value, idFormaPago, this.txtNotas.Text.ToString(), //24
								0.00m, 0.00m, Convert.ToDecimal(this.txtIva.Value), Convert.ToDecimal(this.txtDescIvaTotal.Value), Convert.ToDecimal(this.txtIvaTotal.Value), Convert.ToDecimal(this.txtTotal.Value), false, (int)this.txtPersonas.Value, bFiscal, this.txtEncuesta.Text.ToString(), this.txtEncuesta2.Text.ToString(), iTarifaExpress, 
								(int)this.txtTiempoAdicional.Value, bToqueQueda); //31
							oCmdActualiza.CommandText = sSQL;
							idCompra = (int)oCmdActualiza.ExecuteScalar();
							#endregion Encabezado

							#region Detalle
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
							{								
								int iRegistroTarjetaVIP = 0;

								string sSQLDetalle = string.Format("Exec GuardaDetalleFacturas {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15},{16}",
									(int)dr.Cells["idDetCompra"].Value, idCompra, //1
									(int)dr.Cells["idArticulo"].Value, dr.Cells["Codigo"].Value.ToString(), dr.Cells["Articulo"].Value.ToString(), (int)dr.Cells["Cantidad"].Value, //5
									Convert.ToDecimal(dr.Cells["Precio"].Value), Convert.ToDecimal(dr.Cells["Impuesto"].Value), Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), //8
									Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value), Convert.ToDecimal(dr.Cells["Subtotal"].Value),//10
									(int)dr.Cells["Tipo"].Value, (bool)dr.Cells["Principal"].Value, (bool)dr.Cells["Bloqueado"].Value, (int)dr.Cells["idCombo"].Value, iRegistroTarjetaVIP,(bool)dr.Cells["Entregado"].Value); //15
								oCmdActualiza.CommandText = sSQLDetalle;
								oCmdActualiza.ExecuteNonQuery();								
							}
							#endregion Detalle	

							#region Guarda Cobro de la factura	
							if (miCobro.ultraGrid1.Rows.Count > 0)
							{
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in miCobro.ultraGrid1.Rows.All)
								{
									int idDetFormaPago = 0;
									int idBanco = 0;
									int idTarjeta = 0;
									int idCuenta = 0;
									int idPlan = 0;
									int idPlazo = 0;
									decimal dComTarjeta = 0.00m;
									int idDocumento = 0;
									string DesPrenda = "";
									DateTime dtFecCaducidad = DateTime.Today;
							
									if (dr.Cells["idDetFormaPago"].Value != System.DBNull.Value) idDetFormaPago = (int)dr.Cells["idDetFormaPago"].Value;
									if (dr.Cells["idBanco"].Value != System.DBNull.Value) idBanco = (int)dr.Cells["idBanco"].Value;
									if (dr.Cells["idTarjeta"].Value != System.DBNull.Value) idTarjeta = (int)dr.Cells["idTarjeta"].Value;
									if (dr.Cells["idCuenta"].Value != System.DBNull.Value) idCuenta = (int)dr.Cells["idCuenta"].Value;
									if (dr.Cells["idPlan"].Value != System.DBNull.Value) idPlan = (int)dr.Cells["idPlan"].Value;							
									if (dr.Cells["idPlazo"].Value != System.DBNull.Value) idPlazo = (int)dr.Cells["idPlazo"].Value;
									//if ((int)dr.Cells["idFormaPago"].Value == 3) dtFecCaducidad = (DateTime)dr.Cells["FechaCaducidad"].Value;
									if (dr.Cells["ValorUsoTarjeta"].Value != System.DBNull.Value) dComTarjeta = (decimal)dr.Cells["ValorUsoTarjeta"].Value;
									if (dr.Cells["idDocumento"].Value != System.DBNull.Value) idDocumento = (int)dr.Cells["idDocumento"].Value;
									if (dr.Cells["DescripcionPrenda"].Value != System.DBNull.Value) DesPrenda = dr.Cells["DescripcionPrenda"].Value.ToString();
								
									string sSQLCobro = string.Format("Exec GuardaDetFormaPagoLotte {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, {18}, '{19}',{20}", 
										idDetFormaPago, 
										idCompra, 1, (int)dr.Cells["idFormaPago"].Value, (decimal)dr.Cells["Valor"].Value,
										idBanco, idTarjeta, dr.Cells["CuentaCorriente"].Value.ToString(), dr.Cells["Numero"].Value.ToString(),
										dtFecCaducidad.ToString("yyyyMMdd"), dr.Cells["NumVoucher"].Value.ToString(), dr.Cells["Autorizacion"].Value.ToString(),
										idCuenta, dr.Cells["Lote"].Value.ToString(), dr.Cells["Referencia"].Value.ToString(), idPlan, 0, 0, 0, DesPrenda, 0);
									oCmdActualiza.CommandText = sSQLCobro;
									oCmdActualiza.ExecuteNonQuery();
									
									#region Cobro Prenda
									if ((int)dr.Cells["idFormaPago"].Value == 16)
									{
										string sSQLEP = string.Format("Exec GrabaEstadoPrenda {0}", idCompra);
										oCmdActualiza.CommandText = sSQLEP;
										oCmdActualiza.ExecuteNonQuery();
									}
									#endregion Cobro Prenda									
								}								
							}
							#endregion Guarda Cobro de la factura

							#region Autorizacion Factura
							if (bNuevo)
							{
								if (Funcion.bEscalarSQL(cdsSeteoF, "Select FacturaElectronica From SeteoF", false))
								{
									oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 1", idCompra);;
									string sClaveAcceso = oCmdActualiza.ExecuteScalar().ToString();

									string cadInvertida = invertirCadena(sClaveAcceso.Substring(0, 48));
									string ValM11 = Modulo11(cadInvertida).ToString().Trim();
									string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

									oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, idCompra);
									oCmdActualiza.ExecuteNonQuery();
								}
							}							
							#endregion Autorizacion Factura

							#region Cobros Anticipados Anticipos Promocionales
							if (bCobroAnticipado)
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Anticipos Set Estado = 3, Saldo = 0 Where Tipo = 15 And idOrigen = {0}", idCompra));

							if (idAnticipoRP > 0)
							{
								sSQL = string.Format("Insert Into AnticiposCompras (idCompra, idAnticipo, Valor, Estado) Values ({0}, {1}, {2}, 0) Exec ActualizaSaldoFacturas {1}, 24, 6", 
									idCompra, idAnticipoRP, dValorPromActiva);
								oCmdActualiza.CommandText = sSQL;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Cobros Anticipados Anticipos Promocionales

							#region Anticipos Clientes
							if ((int)this.cmbTipoRuc.Value != 11)
							{
								if(bUsoAnticipoCliente)
								{
									#region Actualiza saldo Anticipos Abonos
									sSQL = string.Format("Exec ActualizarAnticiposAbonosClientes {0}, {1}, {2}", idCliente, (decimal)this.txtTotal.Value, idCompra);
									oCmdActualiza.CommandText = sSQL;
									oCmdActualiza.ExecuteNonQuery();
																
									string sSQLMensajefact = string.Format("Exec MensajeAnticiposCruzados {0}", idCompra);
									oCmdActualiza.CommandText = sSQLMensajefact;
									string MensajeFact = oCmdActualiza.ExecuteScalar().ToString();

									string sSQLUpMsjCompra = string.Format("Update Compra Set Notas = '{0}' Where idCompra = {1}", MensajeFact, idCompra);
									oCmdActualiza.CommandText = sSQLUpMsjCompra;
									oCmdActualiza.ExecuteNonQuery();									
									
									MessageBox.Show(MensajeFact, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
									#endregion Actualiza saldo Anticipos Abonos
								}
							}
							#endregion Anticipos Clientes

							#region Comandas Autoconsumo | Descuento Empleado
							if (!bMenuCuadricula)
							{
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.uGridDetalle.Rows.All)
								{
									if((int)drr.Cells["idArticulo"].Value != 3422)
									{
										#region Valida Comanda Existencia
										string sSQLVALCO = string.Format("Exec ExistenciaComanda {0}", (int)drr.Cells["idArticulo"].Value);
										oCmdActualiza.CommandText = sSQLVALCO;
										iExisComandaAuDesc = (int)oCmdActualiza.ExecuteScalar();
										#endregion Valida Comanda Existencia

										#region Inserta Comanda
										if(iExisComandaAuDesc == 1)
										{
											string sSQLCOAD = string.Format("Exec InsertaComandasAutDesc {0}, {1}, {2}, {3}, {4}",
												idCompra, (int)this.cmbBodega.Value, (int)drr.Cells["idArticulo"].Value, (int)this.cmbHabitaciones.Value, (int)drr.Cells["Cantidad"].Value);
											oCmdActualiza.CommandText = sSQLCOAD;
											idComandaAutDesc = (int)oCmdActualiza.ExecuteScalar();
										}
										#endregion Inserta Comanda
									}
								}
							}
							#endregion Comandas Autoconsumo | Descuento Empleado

							#region Refacturacion - Factura Extra 
							if(bFacturaExtra)
							{
								string sSQLFE = string.Format("Exec GrabaFacturaExtra {0}, {1}", idReserva, idCompra);
								oCmdActualiza.CommandText = sSQLFE;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Refacturacion - Factura Extra

							#region Asiento Automatico
							string sSQLAsiento = string.Format("Exec CreaAsientoDeVenta {0}", idCompra);
							oCmdActualiza.CommandText = sSQLAsiento;							
							idAsiento = (int)oCmdActualiza.ExecuteScalar();
							#endregion Asiento Automatico
							
							#region Adiciona Hora Cuadricula
							if(iTipoModificacion != 0)
							{
								string sSQLACH = string.Format("Exec ModificaHoraFinalizaCuadricula {0}, {1}, {2}, {3}, {4}", (int)this.cmbBodega.Value, (int)this.cmbHabitaciones.Value,
									idCompra, (int)this.txtTiempoAdicional.Value, iTipoModificacion);
								oCmdActualiza.CommandText = sSQLACH;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Adiciona Hora Cuadricula
							             
							#region Ingresos Habitacion
							if (bMenuCuadricula)
							{
								if (bNuevo && (int)this.txtPersonas.Value > 2)
								{			
									#region Genera Ingresos Mas De Dos Personas - Actualiza Estado Ocupado (2) 
									int i = 1;
									while (i <= iIngresos)
									{
										int iTurnoIngreso = (int)this.optTurnoIngreso.Value;
										DateTime dtTurnoIngreso = Convert.ToDateTime(this.dtFechaIngreso.Value);
						
										string sSQLNI = string.Format("Exec GuardaIngresos {0}, {1}, {2}, '{3}', {4}, {5}", 
											idIngreso, idCompra, iTurnoIngreso, dtTurnoIngreso.ToString("yyyyMMdd HH:mm"), (int)this.txtTiempo.Value, 0);
										oCmdActualiza.CommandText = sSQLNI;
										oCmdActualiza.ExecuteNonQuery();

										i++;
									}					
					
									string sSQLEstado3 = string.Format("Exec ActualizaEstadoHabitacion {0}, {1}, {2}, {3}, {4}, 1", 
										idHabitacion, 0, (int)this.txtTiempo.Value, 2, bNuevo);
									oCmdActualiza.CommandText = sSQLEstado3;
									oCmdActualiza.ExecuteNonQuery();								
									#endregion Genera Ingresos Mas De Dos Personas - Actualiza Estado Ocupado (2) 
								}
								else
								{
									#region Actualiza Estado Limpieza (3)
									string sSQLEstado4 = string.Format("Exec ActualizaEstadoHabitacion {0}, {1}, {2}, {3}, 0, 0", idHabitacion, 0, 5, 3);
									oCmdActualiza.CommandText = sSQLEstado4;
									oCmdActualiza.ExecuteNonQuery();

									if (idReserva > 0)
									{
										oCmdActualiza.CommandText = string.Format("Update ReservaHabitacion Set Estado = 0, FechaFinaliza = GETDATE() Where idReserva = {0}", idReserva);
										oCmdActualiza.ExecuteNonQuery();
									}
									#endregion Actualiza Estado Limpieza (3)
								}								
							}
							#endregion Ingresos Habitacion
							
							#region Factura Mayor a 50
							if (bDivideFactura)
							{
								int iFPMayorEfectivo = 0;
								int iFPMayorTC = 0;
								int iFPMayorAnticipo = 0;
								int iFPMayorPrenda = 0;
								decimal dTotalMayorFactura = 0.00m;
					
								dTotalMayorFactura = (decimal)this.txtTotal.Value;

								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in miCobro.ultraGrid1.Rows.All)
								{
									if ((int)drr.Cells["idFormaPago"].Value == 1)	 iFPMayorEfectivo = 1;
									if ((int)drr.Cells["idFormaPago"].Value == 3)	 iFPMayorTC = 1;
									if ((int)drr.Cells["idFormaPago"].Value == 8)  iFPMayorAnticipo = 1;
									if ((int)drr.Cells["idFormaPago"].Value == 16) iFPMayorPrenda = 1;
								}								

								string sValNumCompra = "";
								int iValNuevaCompra = 0;									
								decimal dTotalFacturaMayor = 0.00m;
								decimal dValTotalHospedaje = 0.00m;
								decimal dValTotalProductos = 0.00m;
								dTotalFacturaMayor = (decimal)this.txtTotal.Value;

								if((int)iFPMayorEfectivo == 1 || iFPMayorTC == 1 || iFPMayorAnticipo == 1 || iFPMayorPrenda == 1)
								{										
									if(dTotalFacturaMayor >= 50.00m && dTotalFacturaMayor <= 100.00m)
									{
										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
										{	
											if((int)dr.Cells["idArticulo"].Value == 3422)
											{
												dValTotalHospedaje = dValTotalHospedaje + Convert.ToDecimal(dr.Cells["Subtotal"].Value);
											}

											if((int)dr.Cells["idArticulo"].Value != 3422)
											{
												dValTotalProductos = dValTotalProductos + Convert.ToDecimal(dr.Cells["Subtotal"].Value);
											}
										}

										dValTotalHospedaje = dValTotalHospedaje * dPorcentajeIVADec;
										dValTotalProductos = dValTotalProductos * dPorcentajeIVADec;
												
										#region Numeracion
										string sSQLNumero = string.Format("Exec NumeracionLocales 1, {0}, 1", (int)this.cmbBodega.Value);
										oCmdActualiza.CommandText = sSQLNumero;
										sValNumCompra = oCmdActualiza.ExecuteScalar().ToString();
										#endregion Numeracion

										#region Factura
										string sSQLTransforma = string.Format("Exec FacturaMontoMayor {0}, '{1}'", idCompra, sValNumCompra);
										oCmdActualiza.CommandText = sSQLTransforma;
										iValNuevaCompra = (int)oCmdActualiza.ExecuteScalar();
										#endregion Factura

										#region Kardex
										string sSQLDetalleKardex = string.Format("Exec MovimientoKardexFacturaDividida {0}, {1}, '{2}', {3}, '{4}', {5}",
											iValNuevaCompra, 1, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value, 
											sValNumCompra, idCliente);
										oCmdActualiza.CommandText = sSQLDetalleKardex;
										oCmdActualiza.ExecuteNonQuery();
										#endregion Kardex

										#region Registro Nueva Factura
										string sSQLRegistro = string.Format("Exec FacturasDivididas {0}, '{1}', {2}", iValNuevaCompra, sValNumCompra, idCompra);
										oCmdActualiza.CommandText = sSQLRegistro;
										oCmdActualiza.ExecuteNonQuery();
										#endregion Registro Nueva Factura

										#region Autorizacion Factura
										if (Funcion.bEscalarSQL(cdsSeteoF, "Select FacturaElectronica From SeteoF", false))
										{
											oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 1", iValNuevaCompra);;
											string sClaveAcceso = oCmdActualiza.ExecuteScalar().ToString();

											string cadInvertida = invertirCadena(sClaveAcceso.Substring(0, 48));
											string ValM11 = Modulo11(cadInvertida).ToString().Trim();
											string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

											oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, iValNuevaCompra);
											oCmdActualiza.ExecuteNonQuery();
										}
										#endregion Autorizacion Factura

										#region Tiempo Factura
										if (idCompra > 0)
										{
											oCmdActualiza.CommandText = string.Format("Update FacturaEdicion Set Hospedaje = 1, FacturaDividida = 1 Where idCompra = {0}", idCompra);
											oCmdActualiza.ExecuteNonQuery();
										}
										if (iValNuevaCompra > 0)
										{														
											string sSQLFactura1 = string.Format("Exec GuardaFacturaTiempo {0}", iValNuevaCompra);
											oCmdActualiza.CommandText = sSQLFactura1;
											oCmdActualiza.ExecuteNonQuery();

											oCmdActualiza.CommandText = string.Format("Update FacturaEdicion Set FacturaDividida = 1 Where idCompra = {0}", iValNuevaCompra);
											oCmdActualiza.ExecuteNonQuery();
										}
										#endregion Tiempo Factura

										#region Cobro Prenda
										if (iFPMayorPrenda == 1)
										{
											string sSQLEP = string.Format("Exec GrabaEstadoPrenda {0}", iValNuevaCompra);
											oCmdActualiza.CommandText = sSQLEP;
											oCmdActualiza.ExecuteNonQuery();
										}
										#endregion Cobro Prenda

										#region Impresion
										bImpresionDivision = true;
										iCompraOriginal = idCompra;
										iCompraDividida = iValNuevaCompra;
										#endregion Impresion
									}
								}
							}															
							// InsertaEdicionFactura(idCompra);
							#endregion Factura Mayor a 50

							oTransaction.Commit();

							oConexion.Close();
														
							bEdicion = false;
							bNuevo = false;

							#region Controles
							this.cmbBodega.Enabled = false;
							this.cmbHabitaciones.Enabled = false;
							this.dtFecha.Enabled = false;			
							this.cmbGuardia.Enabled = false;
							this.txtNumero.Enabled = false;
							this.btnCliente.Enabled = false;
							this.txtTiempo.Enabled = false;
							this.txtTiempoAdicional.Enabled = false;
							this.dtFechaIngreso.Enabled = false;
							this.optTurnoIngreso.Enabled = false;
							this.cmbTipoRuc.Enabled = false;
							this.txtRuc.Enabled = false;
							this.txtNombre.Enabled = false;
							this.cmbCobrador.Enabled = false;
							this.txtEncuesta.Enabled = false;
							this.txtEncuesta2.Enabled = false;
							this.cmbCajero.Enabled = false;
							this.txtUnidad.Enabled = false;
							this.dtFechaSalida.Enabled = false;
							this.optTurnoSalida.Enabled = false;					
							this.cmbFormaPago.Enabled = false;
							this.btnanticiposPendientes.Enabled = false;

							this.txtBuscar.Enabled = false;
							this.btnVer.Enabled = false;
			
							FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);
							FuncionesProcedimientos.EstadoGrids(false, this.uGridArticulos);
							this.uGridArticulos.DataSource = ultraDataSource3;
						
							this.txtNotas.Enabled = false;

							if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
							if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
							if (miAcceso.Editar) this.btnEditar.Enabled = false;
							if (bImpreso) if (miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
							if (!bImpreso) if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
							this.btnGuardar.Enabled = false;
							this.btnTemporal.Enabled = false;
							if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
							if (miAcceso.Anular) this.btnAnular.Enabled = true;
							this.btnCancelar.Enabled = true;			
							this.btnAnular.Text = "Anular";							
							iTipoModificacion = 0;
							bToqueQueda = false;
							bValidaPromocionMensaje = false;
							iTipoCLiente = 0;
							bUsoGiftCard = false;
							dGiftCard = 0.00m;
							idHistorialTarjetasVIP = 0;
							bUsoCupon = false;
							dValorCupon = 0.00m;
							iActivaTajetaVIP = 0;
							iTotalPuntosTarjetaVIP = 0;
							bPromoVideo = false;
							bRFVideoPromocion = false;
							bPromoHoras = false;
							#endregion Controles

							#region Impresion
							if (ImprimeRefacturacion)
							{
								using (frmMnsajeNumLotte Msje = new frmMnsajeNumLotte(this.txtNumero.Text.ToString(), false))
									if (DialogResult.OK == Msje.ShowDialog())
									{
										//ImprimirRefactura();
										#region Nombre Impresora
										//string sImpresora = "";
										//sImpresora = "SLK-TS100";

										string sImpresora = "";
										sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",(int) this.cmbBodega.Value));		
			              
										//string sImpresoraCocina = "";
										//sImpresoraCocina = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 2 ",(int) this.cmbBodega.Value));		
			
										#endregion Nombre Impresora

										#region Impresion Formato
										if(bImpresionDivision)
										{
											ImpresionPV(sImpresora);
											ImpresionPV2(sImpresora);
											ImpresionPV3(sImpresora);
											ImpresionPV4(sImpresora);
										}
										else
										{
											ImpresionSinDivision(sImpresora);
											ImpresionSinDivision1(sImpresora);
										}
										#endregion Impresion Formato
									}

								this.btnCancelar_Click(sender, e); 

								this.Close();
							}
							if (ImprimeConDEs && !bMenuCuadricula)
							{
								#region Valida iCompra Comanda
								int iValComandaOrdCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec BuscaComandaAutDesc {0}", idComandaAutDesc));
								#endregion Valida iCompra Comanda

								#region Imprime Comanda
								if(iValComandaOrdCompra > 0)
								{																				
									//string sImpresora = "SLK-TS100"; //Cambiar el nombre de la impresora Cocina  "SLK-TS100"  "RICOH Aficio MP 3350 PCL 6"
									//string sImpresoraCocina = "SLK-TS100_cocina";
									string sImpresora = "";
									sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",(int) this.cmbBodega.Value));		
			              
									string sImpresoraCocina = "";
									sImpresoraCocina = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 2 ",(int) this.cmbBodega.Value));		
			
									ImprimeComandaAutDescCaja(iValComandaOrdCompra, sImpresora);
									ImprimeComandaAutDescCliente(iValComandaOrdCompra, sImpresoraCocina);
									iExisComandaAuDesc = 0;
								}
								#endregion Imprime Comanda

								using (frmMnsajeNumLotte Msje = new frmMnsajeNumLotte(this.txtNumero.Text.ToString(), false))
									if (DialogResult.OK == Msje.ShowDialog())
									{
										//Imprimir(0);
										#region Nombre Impresora
										string sImpresora = "";
										sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",(int) this.cmbBodega.Value));		
			              
										//string sImpresoraCocina = "";
										//sImpresoraCocina = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 2 ",(int) this.cmbBodega.Value));		
			
										//string sImpresora = "";
										//sImpresora = "SLK-TS100";
										#endregion Nombre Impresora

										#region Impresion Formato
										if(bImpresionDivision)
										{
											ImpresionPV(sImpresora);
											ImpresionPV2(sImpresora);
											ImpresionPV3(sImpresora);
											ImpresionPV4(sImpresora);
										}
										else
										{
											ImpresionSinDivision(sImpresora);
											ImpresionSinDivision1(sImpresora);
										}
										#endregion Impresion Formato
									}
								this.Close();
							}

							if(bImprimirModificacion)
							{
								using (frmMnsajeNumLotte Msje = new frmMnsajeNumLotte(this.txtNumero.Text.ToString(), false))
									if (DialogResult.OK == Msje.ShowDialog())
									{
										//Imprimir(0);
										#region Nombre Impresora
										string sImpresora = "";
										sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",(int) this.cmbBodega.Value));		
			              
										//								string sImpresoraCocina = "";
										//								sImpresoraCocina = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 2 ",(int) this.cmbBodega.Value));		
										//			
										//								string sImpresora = "";
										//								sImpresora = "SLK-TS100";
										#endregion Nombre Impresora

										#region Impresion Formato
										if(bImpresionDivision)
										{
											ImpresionPV(sImpresora);
											ImpresionPV2(sImpresora);
											ImpresionPV3(sImpresora);
											ImpresionPV4(sImpresora);
										}
										else
										{
											ImpresionSinDivision(sImpresora);
											ImpresionSinDivision1(sImpresora);
										}
										#endregion Impresion Formato
									}
								this.btnCancelar_Click(sender, e); 
								this.Close();
							}

							if(bFacturaExtra)
							{
								using (frmMnsajeNumLotte Msje = new frmMnsajeNumLotte(this.txtNumero.Text.ToString(), false))
									if (DialogResult.OK == Msje.ShowDialog())
									{
										//Imprimir(0);
										#region Nombre Impresora
										string sImpresora = "";
										sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",(int) this.cmbBodega.Value));		
			              
										//string sImpresoraCocina = "";
										//sImpresoraCocina = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 2 ",(int) this.cmbBodega.Value));		
			
										//string sImpresora = "";
										//sImpresora = "SLK-TS100";  // sImpresora = "SLK-TS100";  "RICOH Aficio MP 3350 PCL 6"
										#endregion Nombre Impresora

										#region Impresion Formato
										if(bImpresionDivision)
										{
											ImpresionPV(sImpresora);
											ImpresionPV2(sImpresora);
											ImpresionPV3(sImpresora);
											ImpresionPV4(sImpresora);
										}
										else
										{
											ImpresionSinDivision(sImpresora);
											ImpresionSinDivision1(sImpresora);
										}
										#endregion Impresion Formato
									}
								this.Close();
							}					
							#endregion Impresion
					
							if (bMenuCuadricula) 
								this.Close();
						}
						catch (Exception ex)
						{
							oTransaction.Rollback();
							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
						}
						finally
						{
							oConexion.Close();
						}					
					}
				}
			}
		}

		private void InsertaEdicionFactura(int iValCompra)
		{
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec GuardaFacturaTiempo {0}", iValCompra));
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Crear Asientos", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if (bEdicion)
			{
				MessageBox.Show("No puede crear un asiento mientras edita la factura");				
			}
			else if (bAnulado) 
			{
				MessageBox.Show("No puede crear un asiento de una transacción Anulada");				
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
			{				
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Crear Asiento de esta Factura, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}			
			else
			{					
				string stExisteAsiento = string.Format("Exec AsientoMaestroExiste {0}", (int)idCompra);
				string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento, true);
				if (stMensaje.Length > 0)
				{
					MessageBox.Show(stMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				}
				else
				{
					stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento + ", 2", true);
					if (stMensaje.Length > 0)
					{
						if (DialogResult.Yes == MessageBox.Show(stMensaje, "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{						
							stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento + ", 3", true);

							#region Crea Asiento
							this.Cursor = Cursors.WaitCursor;
							string sSQLAsiento = string.Format("Exec CreaAsientoDeVenta {0}", (int)idCompra);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);

							MessageBox.Show("Asiento Generado con éxito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
							idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idAsiento From Compra Where idCompra = {0}", (int)idCompra));
							Cursor = Cursors.Default;							
							#endregion Crea Asiento
						}
					}				
					else if (DialogResult.Yes == MessageBox.Show("¿Desea crear Asiento?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
					{	
						#region Crea Asiento
						this.Cursor = Cursors.WaitCursor;
						string sSQLAsiento = string.Format("Exec CreaAsientoDeVenta {0}", (int)idCompra);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);

						MessageBox.Show("Asiento Generado con éxito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
						idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idAsiento From Compra Where idCompra = {0}", (int)idCompra));
						Cursor = Cursors.Default;
						#endregion Crea Asiento
					}
				}
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
			{
				string sMensaje = "ANULAR";
				if (bAnulado) sMensaje = "ELIMINAR";
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede '{1}' esta Factura, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper(), sMensaje), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				return;
			}

			this.dtFecha.Value = Funcion.dtEscalarSQL(cdsSeteoF, "Select dbo.FacturacionRetornaFecha()");

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFechaSalida, cdsSeteoF, true);

			this.optTurnoSalida.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTurno '{0}'", Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd HH:mm")));
			
			#region Nuevo Registro	
			if (bNuevo)
			{					
				iEstado = 10;
				
				if (!Validacion.vbNumeracion(bNumeracionAutomatica, idCompra, 1, (int)this.cmbBodega.Value, this.txtNumero, cdsSeteoF)) return;
										
				using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
				{
					if (DialogResult.OK == miNota.ShowDialog())
					{
						#region Encabezado
						string sSQL = string.Format("Exec GuardaMaestroFacturas {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', {8}, '{9}', {10}, '{11}', {12}, {13}, {14}, '{15}', '{16}', {17}, '{18}', {19}, {20}, {21}, {22}, {23}, '{24}', {25}, {26}, {27}, {28}, {29}, {30}, {31}",
							idCompra, idReserva, 2, iEstado, (int)this.cmbBodega.Value, idProyecto, (int)this.cmbHabitaciones.Value, //6
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 0, this.txtNumero.Text.ToString(), (int)this.txtTiempo.Value, Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"), (int)this.optTurnoIngreso.Value, //12
							idCliente, 0, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(), false, Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd HH:mm"), (int)this.optTurnoSalida.Value, //19
							0, 0, (int)this.txtUnidad.Value, 0, this.txtNotas.Text.ToString(), //24
							0.00m, 0.00m, Convert.ToDecimal(this.txtIva.Value), Convert.ToDecimal(this.txtDescIvaTotal.Value), Convert.ToDecimal(this.txtIvaTotal.Value), Convert.ToDecimal(this.txtTotal.Value), true); //31						
						idCompra = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
						#endregion Encabezado

						string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago 0, {0}, 1, 1, 0, 0, 0, '', '', '', '', '', 0, '', '', 0, 0, 0, 0", idCompra);							
						Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro, true);
												
						#region Impresión
						if (bNuevo)
						{
							frmMensajeNumeración miMsje = new frmMensajeNumeración (this.txtNumero.Text.ToString(), bAnulado, 2, false);
						}
						#endregion Impresión									

						bEdicion = false;
						bNuevo = false;

						if (bMenuCuadricula) this.Close();

						#region Controles
						this.cmbBodega.Enabled = false;
						this.cmbHabitaciones.Enabled = false;
						this.dtFecha.Enabled = false;			
						this.cmbGuardia.Enabled = false;
						this.txtNumero.Enabled = false;
						this.btnCliente.Enabled = false;
						this.txtTiempo.Enabled = false;
						this.dtFechaIngreso.Enabled = false;
						this.optTurnoIngreso.Enabled = false;
						this.cmbTipoRuc.Enabled = false;
						this.txtRuc.Enabled = false;
						this.txtNombre.Enabled = false;
						this.cmbCobrador.Enabled = false;
						this.txtEncuesta.Enabled = false;
						this.txtEncuesta2.Enabled = false;
						this.cmbCajero.Enabled = false;
						this.txtUnidad.Enabled = false;
						this.dtFechaSalida.Enabled = false;
						this.optTurnoSalida.Enabled = false;					
						this.cmbFormaPago.Enabled = false;

						this.txtBuscar.Enabled = false;
						this.btnVer.Enabled = false;		
			
						FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);
						FuncionesProcedimientos.EstadoGrids(false, this.uGridArticulos);	
						
						if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
						if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
						this.btnEditar.Enabled = false;
						this.btnImprimir.Enabled = false;
						this.btnGuardar.Enabled = false;
						this.btnTemporal.Enabled = false;
						this.btnAsiento.Enabled = false;
						this.btnAnular.Enabled = false;
						this.btnCancelar.Enabled = false;			
						this.btnAnular.Text = "Anular";
						#endregion Controles
					}
				}											
				return;
			}
			#endregion Nuevo Registro
				
			if (!bMenuCuadricula) this.Consulta(idCompra);

			#region Elimina Registro
			if (bAnulado)
			{						
				if (!miAcceso.Eliminar)
				{
					MessageBox.Show("El Documento esta Anulado, No tiene acceso a Eliminar", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);													
					return;
				}
//				if (DialogResult.Yes == MessageBox.Show("Esta seguro de ELIMINAR este Documento \n\nEste proceso no puede ser revertido", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
//				{
//					string sSQLBorrar = string.Format("Exec BorrarCompraVenta {0}, {1}, {2}", idCompra, idAsiento, 1);
//					Funcion.EjecutaSQL(cdsSeteoF, sSQLBorrar, true);
//
//					MessageBox.Show("Documento Eliminado del sistema", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//
//					this.btnCancelar_Click(sender, e);							
//				}

				return;						
			}
			#endregion Elimina Registro

			#region Valida Anulacion
			if (bAnulado)
			{
				MessageBox.Show("El Documento esta Anulado", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}
			if (!miAcceso.BAnulaTiempo)
			{
				if (!miAcceso.Anular)
				{
					MessageBox.Show("No tiene acceso a Anular " + Text.Trim(), "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
					return;
				}		
			}
			if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Bloquear, 0) From Compra Where idCompra = {0}", idCompra), true))
			{
				MessageBox.Show("No puede Anular la Factura esta Bloqueada", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				return;
			}
			string sNumeroNC = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec VentasValidaNotaDeCredito {0}", idCompra));
			if (sNumeroNC.Length > 0)
			{
				MessageBox.Show(string.Format("Esta Factura esta Asociada con la Nota de Credito N. '{0}'", sNumeroNC), "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From RetencionesVenta Where idCompra = {0}", idCompra)) > 0)
			{
				MessageBox.Show("No puede Anular la Factura esta Asociada a una Retencion que a Afecta a Multiples Facturas, \n\nAnule la Retencion o Consulte al Administrador", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			if (((int)this.cmbFormaPago.Value == 6 || (int)this.cmbFormaPago.Value == 9) && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec VentaValidaPagos {0}", idCompra)) > 0)
			{
				MessageBox.Show("Esta Factura ya tiene Pagos Registrados", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idFE_EstadoComprobante From Compra Where idCompra = {0}", idCompra), true) == 2)
			{
				if (!miAcceso.BAnulFacEnviada)
				{
					MessageBox.Show("No puede Anular la Factura esta autorizada por el S.R.I.", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}

			if (!FuncionesProcedimientos.CVValidaCapturadoDepositado (idCompra, 1, cdsSeteoF)) return;			
			if (Funcion.ValidaBloqueo(idCompra, 1, cdsSeteoF)) return;
			if (!FuncionesContabilidad.bEstadoCaja((int)this.cmbBodega.Value, (DateTime)this.dtFecha.Value, idCompra, 1, cdsSeteoF)) return;
		
			DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
			DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
			if ((miAcceso.BAnulaSoloDelDia && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Estado, 0) From Compra Where idCompra = {0}", idCompra), true) != 6) && dtFechaFac.ToShortDateString() != dtFechaHoy.ToShortDateString())
			{
				MessageBox.Show("Solo puede Anular Facturas de fecha " + dtFechaHoy.ToString("dd/MMM/yyyy"), "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);					
				return;
			}

			#region Valida registro de servicio de hospedaje pasado los 15 minutos
			if (bMenuCuadricula && !bServicioPagado)
			{
				int iMinutos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(MINUTE, '{0}', '{1}')", 
					Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"), Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd HH:mm")));

				if (iMinutos > 10)
				{
					MessageBox.Show(string.Format("Pasado los {0} minutos desde el ingreso no puede anular la factura", iMinutos), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtBuscar.Focus();
					return;				
				}
			}
			#endregion Valida registro de servicio de hospedaje pasado los 15 minutos
			#endregion Valida Anulacion

			#region Anulacion	
			if (Facturacion.AnulacionCompraVenta (idCompra, 1, cdsSeteoF))
			{
				#region Actualiza Estado Limpieza (3)
				if (bMenuCuadricula)
				{
					string sSQLEstado = string.Format("Exec ActualizaEstadoHabitacion {0}, {1}, {2}, {3}", idHabitacion, 0, 5, 3);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLEstado);
					bEdicion = false;

					this.Close();
				}
				#endregion Actualiza Estado Limpieza (3)
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set Numero = Case Estado When 6 Then '' When 10 Then '{0}' End, FechaTentativaLlegada = '{1}', idTipoCredito = {2} Where idCompra = {3}",
					this.txtNumero.Text.ToString(), Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd"), (int)this.optTurnoSalida.Value, idCompra));
        
				if (idCuponWeb > 0) 
					Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 0, 0, cdsSeteoF);

				this.Consulta(idCompra);
			}
			#endregion Anulacion
		}


		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (idCuponWeb > 0 && iEstado == 6 && bCuponActivo) 
				Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 0, 0, cdsSeteoF);

			if (bMenuCuadricula)
			{
				bEdicion = false;
				bNuevo = false;

				this.Close();
				return;
			}

			bEdicion = false;
			bNuevo = false;

			#region Controles
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbHabitaciones.Value = System.DBNull.Value;
			this.dtFecha.Value = System.DBNull.Value;
			this.cmbGuardia.Value = System.DBNull.Value;
			this.txtNumero.Text = "";
			this.txtTiempo.Value = 3;
			this.dtFechaIngreso.Value = System.DBNull.Value;
			this.optTurnoIngreso.Value = 0;			
			this.cmbTipoRuc.Value = System.DBNull.Value;
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.txtUnidad.Value = 0;
			this.dtFechaSalida.Value = System.DBNull.Value;
			this.optTurnoSalida.Value = 0;
			this.cmbCobrador.Value = System.DBNull.Value;
			this.cmbCajero.Value = System.DBNull.Value;
			this.cmbFormaPago.Value = System.DBNull.Value;
			this.txtPersonas.Value = 2;
			this.lblEstado.Text = "";
			this.lblIngresos.Text = "";
			this.txtBuscar.Text = "";
			this.txtEncuesta.Text = "";
			this.txtEncuesta2.Text = "";
			this.txtCupon.Text = "";
			this.txtTarjetaVIP.Text = "";
			this.txtTiempoAdicional.Enabled = false;
			this.txtTiempoAdicional.Value = 0;
		
			
			this.uGridDetalle.DataSource = this.ultraDataSource2;//Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleFactura 0");			
			this.uGridArticulos.DataSource = this.ultraDataSource3;

			this.txtNotas.Text = "";
			this.txtIva.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;			
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;
			
			this.cmbBodega.Enabled = false;
			this.cmbHabitaciones.Enabled = false;
			this.dtFecha.Enabled = false;			
			this.cmbGuardia.Enabled = false;
			this.txtNumero.Enabled = false;
			this.btnCliente.Enabled = false;
			this.txtTiempo.Enabled = false;
			this.dtFechaIngreso.Enabled = false;
			this.optTurnoIngreso.Enabled = false;
			this.cmbTipoRuc.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtNombre.Enabled = false;
			this.cmbCobrador.Enabled = false;
			this.txtEncuesta.Enabled = false;
			this.txtEncuesta2.Enabled = false;
			this.cmbCajero.Enabled = false;
			this.txtUnidad.Enabled = false;
			this.dtFechaSalida.Enabled = false;
			this.optTurnoSalida.Enabled = false;					
			this.cmbFormaPago.Enabled = false;
			this.txtPersonas.Enabled = false;
			this.txtBuscar.Enabled = false;
			this.btnVer.Enabled = false;
			this.txtNotas.Enabled = false;
			this.txtCupon.Enabled = false;
			this.txtTarjetaVIP.Enabled = false;
			dtfechaFact = false;
			
			FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridArticulos);	
						
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnTemporal.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnRefacturacion.Enabled = false;
			this.btnAnular.Text = "Anular";
			this.btnClaveDescuento.Enabled = false;
			this.btnDescEmpAut.Enabled = false;
			this.btnMitadPrecio.Enabled = false;
			this.btnCupon.Enabled = false;
			this.btnExtras.Enabled = false;
			this.btnOrdenCocina.Enabled = false;
			this.chkAbono.Enabled = false;
			this.chkTarjetaVIP.Enabled = false;
			this.chkTarjetaVIP.Checked = false;
			this.btnanticiposPendientes.Enabled = false;
			bBtnDesFactura = false;
			vValidaFPAgo = false;
			vValidaRucEmpleado = false;
			vValFPago = 0;
			ImprimeConDEs = false;
			ValNuevaFactura = false;
			iTarifaExpress = 0;
			
			bCuponMP = false;

			bUsoCupon = false;
			dValorCupon = 0.00m;

			bPromoVideo = false;
			bRFVideoPromocion = false;
			bPromoHoras = false;

			if (idCuponWeb > 0) 
				Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 0, idCompra, cdsSeteoF);
			#endregion Controles
			
			#region Variables
			idCompra = 0;
			idReserva = 0;
			iEstado = 0;
			idAsiento = 0;
			idCliente = 0;
			idProyecto = 0;

			bNumeracionAutomatica = false;
			bImpreso = false;			
			bEliminaAlValidar = false;			
			dAnticipo = 0.00m;
			dAnticipoCliente = 0.00m;
			bUsoAnticipo = false;
			bUsoAnticipoCliente = false;
			bAnulado = false;			
			idBloqueaTransacciones = 0;
			idCuponWeb = 0;
			iOrigen = 0;
			bNuevoIngreso = false;
			bCuponActivo = false;
			bUsoGiftCard = false;
			
			dGiftCard = 0;
			iCompraOriginal = 0;
			iCompraDividida = 0;
			bImpresionDivision = false;
			bFacturaTiempo = false;
			bActivaBoton = false;
			bImprimirModificacion = false;
			bFacturaExtra = false;
			iCompraFacturaExtra = 0;
			iActivoMitadPrecion = 0;
			iActivaChkCupon = 0;
			iActivoCuponBodega = 0;
			ValTiempoAdicionalPromo = 0;
			iTipoModificacion = 0;
			bToqueQueda = false;
			bValidaPromocionMensaje = false;
			iTipoCLiente = 0;
			idHistorialTarjetasVIP = 0;
			iExisComanda = 0;
			idComanda = 0;
			iValComandaExis = 0;
			iValComandaMins = 0;
			iComandaImp = 0;
			idComandaOrden = 0;
			idComandaAutDesc = 0;
			iExisComandaAuDesc = 0;
			iTotalPuntosTarjetaVIP = 0;
			#endregion Variables

			this.dtFecha.CalendarInfo.MinDate = Funcion.dtEscalarSQL(cdsSeteoF, "Select CONVERT(Date, FechaInicio) From SeteoF");
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void dtFechaIngreso_ValueChanged(object sender, System.EventArgs e)
		{
			if (dtFechaIngreso.Value != System.DBNull.Value)
				FuncionesProcedimientos.ListaPersonalFacturas(this.cmbGuardia, (DateTime)this.dtFechaIngreso.Value, 11, cdsSeteoF);	
		}

		bool bCobroAnticipado = false;

		private void btnTemporal_Click(object sender, System.EventArgs e)
		{
			this.Total();

			if (!Validacion.vbComboVacio(this.cmbGuardia, "Seleccione un Guardia")) return;

			if (bCobroAnticipado && bMenuCuadricula)
			{
				if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione el tipo de identificacion del cliente")) return;
				if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;

				if (idCliente == 0)
				{
					MessageBox.Show("El Cliente No Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			if (!Validacion.vbComboVacio(this.cmbFormaPago, "Seleccione una Forma De Pago")) return;

			if((int)this.cmbTipoRuc.Value == 11)
			{
				if (Convert.ToDecimal(this.txtTotal.Value) > 50.00m)
				{
					if (bMenuCuadricula)
					{
						MessageBox.Show("La factura supera los $50.00 \n\nIngrese datos del cliente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbTipoRuc.Focus();
						return;
					}
				}
			}

			bool bComanda = false;
			bool bExisteTipo10 = false;

			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in this.uGridDetalle.Rows.All)
			{
				if (r.Cells["Tipo"].Value == System.DBNull.Value) continue;

				int tipo = Convert.ToInt32(r.Cells["Tipo"].Value);

				if (tipo == 10 || tipo == 3 || tipo == 4)
				{
					bComanda = true;
					if (tipo == 10) bExisteTipo10 = true;				
				}
			}
			bGrabaComoTemporal = true;
			try
			{
				this.GrabaFactura();
			}
			finally
			{
				bGrabaComoTemporal = false;
			}

			if (idCompra <= 0) return;

			if (bMenuCuadricula && bComanda)
			{
				string sNumComanda = "";
				if (bExisteTipo10) 
				{
					string sSQL = string.Format("Exec NumeracionLocales 87, {0}, 1", (int)this.cmbBodega.Value);
					sNumComanda = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
				}
				bool bCrearOrdenUnaVez = true;

				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				{
					if (dr.Cells["Tipo"].Value == System.DBNull.Value) continue;

					int tipo = Convert.ToInt32(dr.Cells["Tipo"].Value);

					if (tipo == 10 || tipo == 3 || tipo == 4)
					{
						int btnOrden = bCrearOrdenUnaVez ? 1 : 0;

						InsertaComanda((int)dr.Cells["idArticulo"].Value,btnOrden,2,(decimal)dr.Cells["Precio"].Value,(decimal)dr.Cells["Impuesto"].Value,(decimal)dr.Cells["DescuentoArt"].Value,(decimal)dr.Cells["DescuentoPorc"].Value,(decimal)dr.Cells["Subtotal"].Value,(int)dr.Cells["Cantidad"].Value, sNumComanda);

						dr.Cells["Tipo"].Value = 0;
						bCrearOrdenUnaVez = false;
					}
				}
				if (idComandaOrden > 0)
				{
					int iValComandaOrd = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec BuscaComandaBtnOrden {0}", idComandaOrden));

					string sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Exec ConsultarBodegaProservicios {0}, 1 ", (int)this.cmbBodega.Value));
					string sImpresoraCocina = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Exec ConsultarBodegaProservicios {0}, 2 ", (int)this.cmbBodega.Value));

					ImprimeComandaBtnOrdenCaja(iValComandaOrd, sImpresora);
				
					ImprimeComandaBtnOrdenCliente(iValComandaOrd, sImpresoraCocina);

					iComandaImp = 0;
					iExisComanda = 0;

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaEstadoComandaBtnOrden {0}", idComandaOrden));
					idComandaOrden = 0;
				}
			}
		}
		private void ImpresionAutomaticaAnticipos(int idAnticipo, int iTipo, C1.Data.C1DataSet ds)
		{
			string sReporte = "";
			#region Nombre Impresora
                            
			string sImpresora = "";
			sImpresora = Funcion.sEscalarSQL(ds, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",iBodega));        
    
			#endregion Nombre Impresora

			ParameterFields paramFields = new ParameterFields();

			ParameterField paramFieldIdAnticipo = new ParameterField();
			ParameterDiscreteValue discreteValIdAnticipo = new ParameterDiscreteValue();
			paramFieldIdAnticipo.ParameterFieldName = "@idAnticipo";
			discreteValIdAnticipo.Value = idAnticipo; 
			paramFieldIdAnticipo.CurrentValues.Add(discreteValIdAnticipo);
			paramFields.Add(paramFieldIdAnticipo);

			sReporte = Funcion.sEscalarSQL(ds, string.Format("Select Reporte From TiposDeAnticipos Where Tipo = {0}", iTipo));

			ReportDocument Report1 = new ReportDocument();
			Report1.Load(MenuLatinium.stDirReportes + sReporte);

			ConnectionInfo crConnectionInfo = new ConnectionInfo();
			crConnectionInfo.ServerName = MenuLatinium.stDirServidor;
			crConnectionInfo.DatabaseName = MenuLatinium.stDirNumero;
			crConnectionInfo.UserID = "infoelect";
			crConnectionInfo.Password = "Bl45o6$9";

			Tables crTables = Report1.Database.Tables;
			for (int i = 0; i < crTables.Count; i++)
			{
				Table crTable = crTables[i];
				try
				{
					TableLogOnInfo crTableLogOnInfo = crTable.LogOnInfo;
					crTableLogOnInfo.ConnectionInfo = crConnectionInfo;
					crTable.ApplyLogOnInfo(crTableLogOnInfo);
				}
				catch (Exception exTabla)
				{
					throw new Exception(string.Format("Error aplicando conexión a tabla {0}: {1}", crTable.Name, exTabla.Message), exTabla);
				}
			}

			if (sImpresora != null && sImpresora != "")
			{
				Report1.PrintOptions.PrinterName = sImpresora;
			}
			else
			{
				throw new Exception("Nombre de impresora vacío o inválido.");
			}
			Report1.SetParameterValue("@idAnticipo", idAnticipo);

			Report1.PrintToPrinter(1, false, 0, 0);

			Funcion.EjecutaSQL(ds, string.Format("Update Anticipos Set Impreso = 1 Where idAnticipo = {0}", idAnticipo));
		}

		private void GrabaFactura()//(Infragistics.Win.UltraWinGrid.UltraGrid uGridFP)
		{
			int iTarifa = iTarifaExpress;

			#region Numeracion
			if (miAcceso.CambiarNumeracion)
			{
				this.txtNumero.Enabled = true;
				if (!Validacion.vbNumeracion(bNumeracionAutomatica, idCompra, 1, (int)this.cmbBodega.Value, this.txtNumero, cdsSeteoF)) return;
			}
			else
			{
				if (bNuevo)
				{
					if (this.txtNumero.Text == "")
					{
						string sSQLNumero = string.Format("Exec NumeracionLocales 1, {0}, 1", (int)this.cmbBodega.Value);
						this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, sSQLNumero);						
					}
				}
			}
			#endregion Numeracion

			if (bServicioPagado)
			{
				if (this.uGridDetalle.Rows.Count == 0)
				{
					MessageBox.Show("Si va a registrar una nueva factura de servicio multiple de habitacion debe registrar los productos a facturar", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}

			#region Variables
			iEstado = 6;
			
			int idTipoRuc = 0;
			if (this.cmbTipoRuc.ActiveRow != null) idTipoRuc = (int)this.cmbTipoRuc.Value;			
			int idCobrador = 0;
			if (this.cmbCobrador.ActiveRow != null) idCobrador = (int)this.cmbCobrador.Value;
			int idCajero = 0;
			if (this.cmbCajero.ActiveRow != null) idCajero = (int)this.cmbCajero.Value;
			int idFormaPago = 0;
			if (this.cmbFormaPago.ActiveRow != null) idFormaPago = (int)this.cmbFormaPago.Value;

			DateTime dtFechaSalida = DateTime.Today;
			bool bFechaSalida = false;
			int idTurnoSalida = 0;
			if (this.dtFechaSalida.Value != System.DBNull.Value)
			{
				bFechaSalida = true;
				dtFechaSalida = Convert.ToDateTime(this.dtFechaSalida.Value);
			}
			if ((int)this.optTurnoSalida.CheckedIndex > 0) idTurnoSalida = (int)this.optTurnoSalida.Value;
			#endregion Variables

			#region Validacion Cupon
			if(this.chkCupon.Checked == true)
			{
				MessageBox.Show("El registro del Cupon se lo debe realizar al CONTABILIZAR la Factura", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCupon.Text = "";
				this.txtCupon.Enabled = true;
				this.chkCupon.Enabled = true;
				this.txtCupon.Focus();
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				{
					int iArtHospedaje = 0;
					iArtHospedaje = (int)dr.Cells["idArticulo"].Value;

					if (iArtHospedaje == 3422)
					{
						dr.Cells["DescuentoPorc"].Value = 0.00;
						this.Total();
					}
				}
				return;
			}
			#endregion Validacion Cupon

			Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 2, idCompra, cdsSeteoF);
			idCuponWeb = 0;

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;
						
				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Encabezado
					string sSQL = string.Format("Exec GuardaMaestroFacturas {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', {8}, '{9}', {10}, '{11}', {12}, {13}, {14}, '{15}', '{16}', {17}, '{18}', {19}, {20}, {21}, {22}, {23}, '{24}', {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, '{34}', '{35}', {36}, {37}, {38}",
						idCompra, idReserva, 2, iEstado, (int)this.cmbBodega.Value, idProyecto, (int)this.cmbHabitaciones.Value, //6
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbGuardia.Value, this.txtNumero.Text.ToString(), (int)this.txtTiempo.Value, Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"), (int)this.optTurnoIngreso.Value, //12
						idCliente, idTipoRuc, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(), bFechaSalida, dtFechaSalida.ToString("yyyyMMdd HH:mm"), idTurnoSalida, //19
						idCobrador, idCajero, (int)this.txtUnidad.Value, idFormaPago, this.txtNotas.Text.ToString(), //24
						0.00m, 0.00m, Convert.ToDecimal(this.txtIva.Value), Convert.ToDecimal(this.txtDescIvaTotal.Value), Convert.ToDecimal(this.txtIvaTotal.Value), Convert.ToDecimal(this.txtTotal.Value), false, (int)this.txtPersonas.Value, false, "", "", iTarifaExpress, 
						(int)this.txtTiempoAdicional.Value, bToqueQueda); //31
					oCmdActualiza.CommandText = sSQL;
					idCompra = (int)oCmdActualiza.ExecuteScalar();
					#endregion Encabezado

					#region Detalle
					int iRegistroTarjetaVIP = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					{
						int tipo = 0;
						if (dr.Cells["Tipo"].Value != System.DBNull.Value) tipo = Convert.ToInt32(dr.Cells["Tipo"].Value);
						if (tipo == 10 || tipo == 3 || tipo == 4)
							continue;

						#region Registra Puntos Tarjeta VIP
						if(chkTarjetaVIP.Checked)
						{
							string sSQLConsumoVIP = string.Format("Exec RegistraPuntosConsumoTarjetasVIP {0}, {1}, '{2}', {3}, {4}, {5}",
								idCompra, (int)dr.Cells["idDetCompra"].Value, this.txtTarjetaVIP.Text, idCliente, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Cantidad"].Value);
							oCmdActualiza.CommandText = sSQLConsumoVIP;
							iRegistroTarjetaVIP = (int)oCmdActualiza.ExecuteScalar();
						}
						#endregion Registra Puntos Tarjeta VIP

						string sSQLDetalle = string.Format("Exec GuardaDetalleFacturas {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15},{16}",
							(int)dr.Cells["idDetCompra"].Value, idCompra, //1
							(int)dr.Cells["idArticulo"].Value, dr.Cells["Codigo"].Value.ToString(), dr.Cells["Articulo"].Value.ToString(), (int)dr.Cells["Cantidad"].Value, //5
							Convert.ToDecimal(dr.Cells["Precio"].Value), Convert.ToDecimal(dr.Cells["Impuesto"].Value), Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), //8
							Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value), Convert.ToDecimal(dr.Cells["Subtotal"].Value),//10
							(int)dr.Cells["Tipo"].Value, (bool)dr.Cells["Principal"].Value, (bool)dr.Cells["Bloqueado"].Value, (int)dr.Cells["idCombo"].Value, iRegistroTarjetaVIP,(bool)dr.Cells["Entregado"].Value); //15
						oCmdActualiza.CommandText = sSQLDetalle;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Detalle	
					
					#region Genera Ingreso - Actualiza Estado Ocupado (2)
					if (!bServicioPagado)
					{
						if (bNuevo)
						{
							int iTurnoIngreso = (int)this.optTurnoIngreso.Value;
							DateTime dtTurnoIngreso = Convert.ToDateTime(this.dtFechaIngreso.Value);
						
							string sSQLNI = string.Format("Exec GuardaIngresos {0}, {1}, {2}, '{3}', {4}, {5}", 
								idIngreso, idCompra, iTurnoIngreso, dtTurnoIngreso.ToString("yyyyMMdd HH:mm"), (int)this.txtTiempo.Value, 0);
							oCmdActualiza.CommandText = sSQLNI;
							oCmdActualiza.ExecuteNonQuery();

							int iOcupado = Funcion.iEscalarSQL(cdsSeteoF,string.Format("SELECT ISNULL(Ocupado,0) FROM Habitaciones WHERE  idHabitacion = {0} AND Bodega = {1}",(int)this.cmbHabitaciones.Value,(int)this.cmbBodega.Value));

							//int iEstadoHabitacion = (bGrabaComoTemporal ? 6 : 2); 

							if (iOcupado == 1)
								iEstadoHabitacion = 2;
							else
								iEstadoHabitacion = 6;

							string sSQLEstado = string.Format("Exec ActualizaEstadoHabitacion {0}, {1}, {2}, {3}, {4}, {5}", 
								idHabitacion, idCompra, (int)this.txtTiempo.Value, iEstadoHabitacion, bNuevo, bServicioPagado);
							oCmdActualiza.CommandText = sSQLEstado;
							oCmdActualiza.ExecuteNonQuery();
						}
					}
					else
					{
						oCmdActualiza.CommandText = string.Format("Update Habitaciones Set idCompra = {0} Where idHabitacion = {1}", 
							idCompra, idHabitacion);
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Genera Ingreso - Actualiza Estado Ocupado (2)					

					#region Adiciona Hora Cuadricula
					if(iTipoModificacion != 0)
					{
						string sSQLACH = string.Format("Exec ModificaHoraFinalizaCuadricula {0}, {1}, {2}, {3}, {4}", (int)this.cmbBodega.Value, (int)this.cmbHabitaciones.Value,
							idCompra, (int)this.txtTiempoAdicional.Value, iTipoModificacion);
						oCmdActualiza.CommandText = sSQLACH;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Adiciona Hora Cuadricula

					#region Promocion Temporal Video
					string sSQLRPV = string.Format("Exec GuardaPromocionVideo {0}, {1}, {2}", idCompra, (int)this.cmbHabitaciones.Value, this.chkVideo.Checked);
					oCmdActualiza.CommandText = sSQLRPV;
					oCmdActualiza.ExecuteNonQuery();
					#endregion Promocion Temporal Video

					if (bUsoPromActiva)
					{
						oCmdActualiza.CommandText = string.Format("UPDATE COMPRA SET CrediDigital = {1} WHERE IDCOMPRA = {0}", idCompra, idAnticipoRP);
						oCmdActualiza.ExecuteNonQuery();
					}

					oTransaction.Commit();
		
					bEdicion = false;
					bNuevo = false;

					if (bMenuCuadricula) 
					{
						if (bNuevo)
							MessageBox.Show(string.Format("Se Genero La Factura Numero : '{0}'.", this.txtNumero.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

						this.Close();
					}
					else
					{
						#region Controles
						this.cmbBodega.Enabled = false;
						this.cmbHabitaciones.Enabled = false;
						this.dtFecha.Enabled = false;			
						this.cmbGuardia.Enabled = false;
						this.txtNumero.Enabled = false;
						this.btnCliente.Enabled = false;
						this.txtTiempo.Enabled = false;
						this.dtFechaIngreso.Enabled = false;
						this.txtEncuesta.Enabled = false;
						this.txtEncuesta2.Enabled = false;
						this.optTurnoIngreso.Enabled = false;
						this.cmbTipoRuc.Enabled = false;
						this.txtRuc.Enabled = false;
						this.txtNombre.Enabled = false;
						this.cmbCobrador.Enabled = false;
						this.cmbCajero.Enabled = false;
						this.txtUnidad.Enabled = false;
						this.dtFechaSalida.Enabled = false;
						this.optTurnoSalida.Enabled = false;					
						this.cmbFormaPago.Enabled = false;

						this.txtBuscar.Enabled = false;
						this.btnVer.Enabled = false;		
			
						FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);
						FuncionesProcedimientos.EstadoGrids(false, this.uGridArticulos);	
						
						if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
						if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
						this.btnEditar.Enabled = false;
						this.btnImprimir.Enabled = false;
						this.btnGuardar.Enabled = false;
						this.btnTemporal.Enabled = false;
						this.btnAsiento.Enabled = false;
						this.btnAnular.Enabled = false;
						this.btnCancelar.Enabled = false;			
						this.btnAnular.Text = "Anular";
						this.chkAbono.Enabled = false;
						iTipoModificacion = 0;
						bToqueQueda = false;
						bValidaPromocionMensaje = false;
						iTipoCLiente = 0;
						bUsoGiftCard = false;
						dGiftCard = 0.00m;
						idHistorialTarjetasVIP = 0;
						bUsoCupon = false;
						dValorCupon = 0.00m;
						iTotalPuntosTarjetaVIP = 0;
						bPromoVideo = false;
						bRFVideoPromocion = false;
						bPromoHoras = false;
						#endregion Controles		
					}
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				}
				finally
				{
					oConexion.Close();
				}
			}
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (!bNumeracionAutomatica)
				{
					char Pad = '0';
					this.txtNumero.Text = this.txtNumero.Text.ToString().PadLeft(9, Pad);
				}
			}
		}

		private void cmbGuardia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNumero.Focus();
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
				if (this.txtTiempo.Enabled) this.txtTiempo.Focus(); else this.cmbTipoRuc.Focus();			
		}

		private void cmbTipoRuc_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbTipoRuc.ActiveRow != null) 
				{
					idCliente = 0;
					this.txtRuc.Text = "";
					this.txtNombre.Text = "";
					this.txtRuc.Enabled = true;

					#region Cupones
					this.txtCupon.Text = "";
					this.txtCupon.Enabled = false;
					this.chkCupon.Checked = false;
					this.chkCupon.Enabled = false;
					#endregion Cupones

					#region Abono Cliente
					this.chkAbono.Checked = false;
					this.chkAbono.Enabled = false;
					#endregion Abono Cliente

					#region Tarjetas VIP
					this.chkTarjetaVIP.Enabled = false;
					this.chkTarjetaVIP.Checked = false;
					this.txtTarjetaVIP.Enabled = false;
					this.txtTarjetaVIP.Text = "";
					this.txtPuntosTarjetaVIP.Value = 0;
					this.btnInfoTarjetasVIP.Enabled = false;
					this.btnCanjePuntos.Enabled = false;					
					#endregion Tarjetas VIP

					if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 6) this.txtRuc.MaxLength = 13;
					if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;

					#region Consumidor Final
					if ((int)this.cmbTipoRuc.Value == 11) 
					{	
						iTipoCLiente = 1;

						this.txtRuc.Enabled = false;
						this.txtRuc.MaxLength = 13;
											
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select cr.idCliente, cl.Ruc, cl.Nombre From Cliente cl Inner Join ClienteTipoRuc cr On cr.idCliente = cl.idCliente Where cr.idTipoRuc = 11", true);
						dr.Read();
						if (dr.HasRows)
						{
							idCliente = dr.GetInt32(0);
							this.txtRuc.Value = dr.GetString(1);
							this.txtNombre.Value = dr.GetString(2);
						}
						dr.Close();						
					}
					#endregion Consumidor Final	
          
					ValidaAnticiposClientes();

					int iValCheckLocal = 0;
					if(this.cmbBodega.Value == System.DBNull.Value) 
					{
						iValCheckLocal = 0;
					}
					else
					{
						iValCheckLocal = (int)this.cmbBodega.Value;
					}					

					if(idCompra > 0)
					{
						#region Cupones
						if(Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ChkActivo From CuponesEmpresasActivo Where Bodega = {0}", iValCheckLocal), true))
						{
							iActivaChkCupon = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaAceptaCupones {0}, {1}", (int)this.cmbBodega.Value, (int)this.cmbHabitaciones.Value));
							if(iActivaChkCupon == 1 && (int)this.cmbTipoRuc.Value != 11) this.chkCupon.Enabled = true;
						}
						#endregion Cupones

						#region Tarjetas VIP
						iActivaTajetaVIP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Activo From ActivaTarjetasVIP Where Bodega = {0}", (int)this.cmbBodega.Value));
						if (iActivaTajetaVIP == 1 && (int)this.cmbTipoRuc.Value != 11) 
						{
							this.chkTarjetaVIP.Enabled = true;
							this.btnInfoTarjetasVIP.Enabled = true;
						}
						#endregion Tarjetas VIP
					}

					#region Video Promocion
					if(!bPromoHoras)
					{
						if(bMenuCuadricula || bRFVideoPromocion)
						{
							if ((int)this.cmbTipoRuc.Value != 11) 
							{
								bool bPromocionVideo = Funcion.bEscalarSQL(cdsSeteoF, "Select Estado From PromocionesTemporales Where promocion = 'PROMOCION VIDEO KIKE JAV'", true);
								if((int)this.cmbBodega.Value == 4 && bPromocionVideo)
								{
									this.chkVideo.Enabled = true;
								}
							}
						}
					}
					#endregion Video Promocion
				}
			}
		}

		private void txtTiempo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (bMenuCuadricula) this.cmbTipoRuc.Focus(); else this.dtFechaIngreso.Focus();
		}

		private void cmbTipoRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if (this.txtRuc.Enabled) this.txtRuc.Focus(); else if (bMenuCuadricula) this.cmbCobrador.Focus(); else this.dtFechaSalida.Focus();
			}
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (bMenuCuadricula)
				{
					if (this.cmbCobrador.Enabled) this.cmbCobrador.Focus(); else this.txtUnidad.Focus();
				}
				else this.dtFechaSalida.Focus();

				#region Valida Socio
				if (vValidaRucEmpleado)
				{
					int ValCliente = 0;
					if (this.txtRuc.Text.Length == 0)
					{
						MessageBox.Show("Ingrese Numero identificacion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
					if (this.txtRuc.Text.Length < 9)
					{
						MessageBox.Show("Ingrese Numero identificacion correcto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					if (this.txtRuc.Text.Length > 9)
					{					
						#region Validacion
								
						string sSQLValCli = string.Format("Select COUNT(1) From Cliente Where idGrupoCliente = 9 And Proveedor = 0 And Ruc = '{0}'", this.txtRuc.Text);		
						ValCliente = Funcion.iEscalarSQL(cdsSeteoF, sSQLValCli);

						#endregion Validacion

						if (ValCliente == 1)
						{
							CargaDatosSocios();	
						}
						else
						{
							MessageBox.Show("Validar informacion del empleado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtNombre.Value = "";
							this.txtRuc.Focus();
							return;
						}				
					}					
				}
				#endregion Valida Socio
			}
		}

		private void cmbCobrador_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCajero.Focus();
		}

		private void cmbCajero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtUnidad.Focus();
		}

		private void txtUnidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbFormaPago.Focus(); 
		}

		private void cmbFormaPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtPersonas.Focus();
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnTemporal.Focus();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{	
			if (ValNuevaFactura)
			{
				this.uGridArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec FacturaLoteListaInventario '{0}'", this.txtBuscar.Text.ToString()));
				if (this.uGridArticulos.Rows.Count > 0) 
				{
					this.uGridArticulos.Focus();
				}
				else 
				{
					this.txtBuscar.Focus();
				}
			}
			else if(bActivaBoton || bFacturaExtra)
			{
				this.uGridArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec FacturaLoteListaInventario '{0}'", this.txtBuscar.Text.ToString()));
				if (this.uGridArticulos.Rows.Count > 0) 
				{
					this.uGridArticulos.Focus();
				}
				else 
				{
					this.txtBuscar.Focus();
				}
			}		
			else
			{
				if (!bMenuCuadricula && bEdicion)
				{
					if (!Validacion.vbComboVacio(this.cmbHabitaciones, "Seleccione una habitación")) return;
	
					int iCantidad = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDet in this.uGridDetalle.Rows.All)					
						if ((int)drDet.Cells["idArticulo"].Value == 3422) iCantidad++;
				
					if ((int)this.cmbFormaPago.Value == 30) this.cmbFormaPago.Enabled = false;
					if (iCantidad == 0)
						this.CargaArticulos(3422, true);
				}

				this.uGridArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec FacturaLoteListaInventario '{0}'", this.txtBuscar.Text.ToString()));

				if (this.uGridArticulos.Rows.Count > 0) this.uGridArticulos.Focus(); else this.txtBuscar.Focus();
			}
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void txtTiempo_Click(object sender, System.EventArgs e)
		{
			this.txtTiempo.SelectAll();
		}

		private void txtTiempo_Enter(object sender, System.EventArgs e)
		{
			this.txtTiempo.SelectAll();
		}

		private void txtUnidad_Click(object sender, System.EventArgs e)
		{
			this.txtUnidad.SelectAll();
		}

		private void txtUnidad_Enter(object sender, System.EventArgs e)
		{
			this.txtUnidad.SelectAll();
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.cmbTipoRuc.ActiveRow != null)
				if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 5) Funcion.CamposNumericos(sender, e);
		}

		bool bCargandoHabitaciones = false;

		private void cmbBodega_ValueChanged(object sender, EventArgs e)
		{
			if (!bFormCargado) return;

			if (this.cmbBodega.ActiveRow != null && this.cmbBodega.Value != null)
			{
				int idBodegaSeleccionada = Convert.ToInt32(this.cmbBodega.Value);

				bCargandoHabitaciones = true;

				this.cmbHabitaciones.DataSource = Funcion.dvProcedimiento(
					cdsSeteoF,
					string.Format("SELECT idHabitacion, Nombre FROM Habitaciones WHERE Bodega = {0} ORDER BY NumeroHabitacion", idBodegaSeleccionada)
					);

				bCargandoHabitaciones = false;

				bCobroAnticipado = Funcion.bEscalarSQL(
					cdsSeteoF,
					string.Format("SELECT CobroAnticipado FROM Bodega WHERE Bodega = {0}", idBodegaSeleccionada),
					true
					);

				DateTime fechaSeleccionada = this.dtFecha.Value != null ? Convert.ToDateTime(this.dtFecha.Value) : DateTime.Today;

				if (fechaSeleccionada.Date == DateTime.Today)
				{
					DateTime ahora = DateTime.Now;
					TimeSpan horaActual = ahora.TimeOfDay;
					int turnoActual = (horaActual >= TimeSpan.FromHours(7) && horaActual < TimeSpan.FromHours(18)) ? 1 : 2;
					DateTime fechaConsulta = ahora.Date;

					if (turnoActual == 2 && horaActual < TimeSpan.FromHours(7))
						fechaConsulta = fechaConsulta.AddDays(-1);
					// Viejo responsables caja
//					string sSQLTurno = string.Format(
//						"SELECT ISNULL(IdGuardia, 0), ISNULL(IdCobrador, 0), ISNULL(IdCajero, 0) " +
//						"FROM Caja " +
//						"WHERE CONVERT(DATE, FechaSistema) = '{0:yyyy-MM-dd}' " +
//						"AND TurnoCaja = {1} " +
//						"AND Bodega = {2}",
//						fechaConsulta, turnoActual, idBodegaSeleccionada
//						);

					// Nuevo Responsables Caja
					string sSQLTurno = string.Format(
						@"SELECT 
									ISNULL((SELECT TOP 1 idPersonal FROM ResponsablesCaja WHERE idCaja = c.idCaja AND idCargo = 11 AND Estado = 1), 0) AS IdGuardia,
									ISNULL((SELECT TOP 1 idPersonal FROM ResponsablesCaja WHERE idCaja = c.idCaja AND idCargo = 6 AND Estado = 1), 0) AS IdCobrador,
									ISNULL((SELECT TOP 1 idPersonal FROM ResponsablesCaja WHERE idCaja = c.idCaja AND idCargo = 4 AND Estado = 1), 0) AS IdCajero
							FROM (
									SELECT TOP 1 idCaja
									FROM Caja
									WHERE CONVERT(DATE, Fecha) = '{0:yyyy-MM-dd}'
										AND TurnoCaja = {1}
										AND Bodega = {2}
							) AS c",
						fechaConsulta, turnoActual, idBodegaSeleccionada
						);

					SqlDataReader drTurno = Funcion.rEscalarSQL(cdsSeteoF, sSQLTurno, true);

					bool responsablesValidos = false;

					if (drTurno.Read())
					{
						int idGuardia = drTurno.GetInt32(0);
						int idCobrador = drTurno.GetInt32(1);
						int idCajero = drTurno.GetInt32(2);

						if (idGuardia > 0 && idCobrador > 0 && idCajero > 0)
						{
							responsablesValidos = true;

							this.cmbGuardia.Value = idGuardia;
							this.cmbCobrador.Value = idCobrador;
							this.cmbCajero.Value = idCajero;

							this.cmbGuardia.Enabled = false;
							this.cmbCobrador.Enabled = false;
							this.cmbCajero.Enabled = false;
						}
					}
					drTurno.Close();

					if (!responsablesValidos)
					{
						MessageBox.Show("Aún no se han seleccionado responsables para la caja.", "Responsables faltantes", MessageBoxButtons.OK, MessageBoxIcon.Warning);

						this.cmbBodega.Value = idBodegaAnterior;

						this.cmbHabitaciones.DataSource = null;
						this.cmbGuardia.Value = null;
						this.cmbCobrador.Value = null;
						this.cmbCajero.Value = null;

						return; 
					}
				}

				idBodegaAnterior = idBodegaSeleccionada;

				FuncionesProcedimientos.ListaPersonalFacturas(this.cmbGuardia, fechaSeleccionada, 11, cdsSeteoF);
				FuncionesProcedimientos.ListaPersonalFacturas(this.cmbCobrador, fechaSeleccionada, 6, cdsSeteoF);
				FuncionesProcedimientos.ListaPersonalFacturas(this.cmbCajero, fechaSeleccionada, 4, cdsSeteoF);
			}
		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Cantidad", 0);
			FuncionesProcedimientos.FormatoGrid(e, "Precio", 6);						
			FuncionesProcedimientos.FormatoGrid(e, "Impuesto", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoPorc", 6);
		}

		private void uGridArticulos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnSeleccionar")
			{
				if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;
				if (!Validacion.vbComboVacio(this.cmbHabitaciones, "Seleccione una Habitación")) return;

				if (this.uGridDetalle.ActiveRow != null) uGridDetalle.ActiveRow.Selected = false;

				this.CargaArticulos((int)e.Cell.Row.Cells["idArticulo"].Value, false);
			}
		}

		private void frmFacturacionHabitaciones_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una Factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void frmFacturacionHabitaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (!bMenuCuadricula)
				{
					if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
					if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
					if (idCompra > 0)
					{
						if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
						if (e.KeyCode == Keys.I) if (bImpreso) 
																		 {
																			 if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
																		 }
																		 else if (miAcceso.ReImprimir) if (this.btnImprimir.Enabled) this.btnImprimir.Enabled = true;
						if (e.KeyCode == Keys.F5)
						{
							miAcceso = new Acceso(cdsSeteoF, "0703");

							this.Consulta(idCompra);					
						}
					}
				}

				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				//				if (e.KeyCode == Keys.Escape)
				//					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				//						this.btnCancelar_Click(sender, e);
				//
				//				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void uGridDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalle);	
		}

		private void txtUnidad_Validated(object sender, System.EventArgs e)
		{
			if (this.txtUnidad.Value == System.DBNull.Value)
				this.txtUnidad.Value = 0;
		}

		private void dtFechaSalida_ValueChanged(object sender, System.EventArgs e)
		{
			if (dtFechaSalida.Value != System.DBNull.Value)
			{
				FuncionesProcedimientos.ListaPersonalFacturas(this.cmbCobrador, (DateTime)this.dtFechaSalida.Value, 6, cdsSeteoF);	
				FuncionesProcedimientos.ListaPersonalFacturas(this.cmbCajero, (DateTime)this.dtFechaSalida.Value, 4, cdsSeteoF);	

				if (!bMenuCuadricula && bEdicion)
					this.optTurnoSalida.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTurno '{0}'", Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd HH:mm")));
			}
		}

		private void btnAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				if (miAcceso.CrearAsiento)
				{
					if (idAsiento == 0) MessageBox.Show("No Existe un Asiento Creado de esta Transacción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					else
					{					
						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0} And Borrar = 0) Set @Res = 1 Else Set @Res = 0 Select @Res", idAsiento), true))
						{
							Cursor = Cursors.WaitCursor;
							frmAsientos Asientos = new frmAsientos(idAsiento);
							Asientos.MdiParent = MdiParent;
							Asientos.Show();
							Cursor = Cursors.Default;
						}
						else if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0} And Borrar = 1) Set @Res = 1 Else Set @Res = 0 Select @Res", idAsiento), true))
						{
							MessageBox.Show("El Asiento de esta Factura esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						}
						else if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0}) Set @Res = 1 Else Set @Res = 0 Select @Res", idAsiento), true))
						{
							MessageBox.Show("El Asiento de esta Factura fue Eliminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						}
					}
				}
			}
		}

		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (!bActualiza) return;

			bActualiza = false;	
//				if (e.Cell.Column.ToString() == "Cantidad") 
//				{
////					if ((int)e.Cell.Row.Cells["idArticulo"].Value == 3422 && (int)e.Cell.Row.Cells["Cantidad"].Value != iIngresos)
////					{
////						e.Cell.Row.Cells["Cantidad"].Value = iIngresos;
////					}
//					this.Total(); 
//					bActualiza = true;				
			//	}
			if (e.Cell.Column.ToString() == "Cantidad") {this.Total(); bActualiza = true;}
			if (e.Cell.Column.ToString() == "Precio") {this.Total(); bActualiza = true;}
			if (e.Cell.Column.ToString() == "Impuesto") {this.Total(); bActualiza = true;}
			if (e.Cell.Column.ToString() == "DescuentoPorc") {this.Total();bActualiza = true;}

			bActualiza = true;
		}

		private void uGridDetalle_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			this.Total();
			
			if (bNuevo)
			{
				//				this.btnGuardar.Enabled = false;
				//				this.btnTemporal.Enabled = true;
				this.txtPersonas.Value = 2;
				iIngresos = 1;
			}
		}

		private void uGridDetalle_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			
		}

		private void uGridDetalle_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if(!bFacturaExtra)
				{
					if (iEstado == 10 && bImpreso && (int)this.cmbTipoRuc.Value != 11)
					{
						MessageBox.Show("No puede eliminar filas de facturas impresas con datos de clientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
						return;
					}
				}
				
				for (int i = 0; i < e.Rows.Length; i++)
				{	
					if (idReserva > 0 && (int)e.Rows[i].Cells["idArticulo"].Value == 3422)
					{
						MessageBox.Show("No puede eliminar el SERVICIO DE HOSPEDAJE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
						return;
					}

					if ((int)e.Rows[i].Cells["idArticulo"].Value == 3422)
					{
						MessageBox.Show("No puede eliminar el SERVICIO DE HOSPEDAJE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
						return;
					}

					if ((int)e.Rows[i].Cells["idArticulo"].Value == 4539)
					{
						MessageBox.Show("No puede eliminar el SERVICIO ADICIONAL HORAS", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
						return;
					}

					if (bValRefacturacion)
					{
						int iComDetCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From DetCompra Where idcomanda > 0 And idDetCompra = {0}", (int)e.Rows[i].Cells["idDetCompra"].Value));

						if ((int)e.Rows[i].Cells["idArticulo"].Value == 3422)
						{
							MessageBox.Show("No puede eliminar el SERVICIO DE HABITACIÓN DE RESERVACIONES", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cancel = true;
							return;
						}

						if(iComDetCompra > 0)
						{
							MessageBox.Show("No puede eliminar la COMANDA.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cancel = true;
							return;
						}
					}

					if ((int)this.cmbFormaPago.Value == 29 || (int)this.cmbFormaPago.Value == 30)
					{
						e.DisplayPromptMsg = false;
						return;
					}

					#region Comandas
					if ((int)e.Rows[i].Cells["idDetCompra"].Value > 0)
					{
						iValComandaExis = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaComandaDetCompra {0}", (int)e.Rows[i].Cells["idDetCompra"].Value));
						iValComandaMins = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaMinutosComandas {0}", (int)e.Rows[i].Cells["idDetCompra"].Value));
						
						if(iValComandaExis == 1)
						{
							if(iValComandaMins == 0)
							{
								if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Eliminar la Comanda?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
								{
									//string sImpresora = "SLK-TS100"; //Cambiar el nombre de la impresora Cocina
									//string sImpresoraCocina = "SLK-TS100_cocina";

									string sImpresora = "";
									sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec ConsultarBodegaProservicios {0}, 1 ",(int) this.cmbBodega.Value));		
			              
									string sImpresoraCocina = "";
									sImpresoraCocina = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 2 ",(int) this.cmbBodega.Value));		
												
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ModificaEstadoComandaEliminada {0}", (int)e.Rows[i].Cells["idDetCompra"].Value));									
									ImprimeComandaCancelada((int)e.Rows[i].Cells["idDetCompra"].Value, sImpresora);
									ImprimeComandaCanceladaCliente((int)e.Rows[i].Cells["idDetCompra"].Value, sImpresoraCocina);
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)e.Rows[i].Cells["idDetCompra"].Value));									
									
									e.DisplayPromptMsg = false;
									this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleFacturas {0}", idCompra));
									e.DisplayPromptMsg = false;
									Total();
								}
							}
							else
							{
								MessageBox.Show("Pasado los 5 minutos, no puede eliminar la Comanda.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								e.Cancel = true;
								return;
							}
						}
					}					
					#endregion Comandas

					if(iValComandaExis == 0)
					{
						if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Eliminar las filas seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
						{						
							if ((int)e.Rows[i].Cells["idDetCompra"].Value > 0)
							{	
								if (idCuponWeb > 0)
									if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CompraPromocion Where idPromocion = {0} And Cupon = 1", (int)e.Rows[i].Cells["idDetCompra"].Value)) > 0)
										Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 0, 0, cdsSeteoF);

								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec GrabaRegistroAnulacion {0}", (int)e.Rows[i].Cells["idDetCompra"].Value));
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)e.Rows[i].Cells["idDetCompra"].Value));
								e.DisplayPromptMsg = false;
								Total();														
							}
							else e.DisplayPromptMsg = false;
						}
					}
					else e.Cancel = true;
				}
			}
			else e.Cancel = true;
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			//			if (this.dtFecha.Value != System.DBNull.Value)
			//			{
			//				dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
			//				dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
			//			}
		}		

		private void CargaCliente()
		{
			string[] sParametros = new string[11];

			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 1, this.txtRuc.Text.ToString(), cdsSeteoF);
			
			if (sParametros[0] != null) idCliente = int.Parse(sParametros[0]);
			if (sParametros[1] != null) this.cmbTipoRuc.Value = int.Parse(sParametros[1]);
			if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];			
		}

		private void CargaClienteEmpresa(int iEmpresa)
		{
			if (iEmpresa == 29)
			{
				string sSQL = string.Format("Select idCliente, idTipoRuc, Ruc, Nombre From Cliente Where Proveedor = 0 And Ruc = '1792853818001'", this.txtRuc.Text.ToString());
				SqlDataReader drConsulta = Funcion.rEscalarSQL(cdsSeteo, sSQL, true);
				drConsulta.Read();
				if (drConsulta.HasRows)
				{	
					this.cmbTipoRuc.Value = drConsulta.GetInt32(1);
					idCliente = drConsulta.GetInt32(0);
					this.txtRuc.Text = drConsulta.GetString(2);
					this.txtNombre.Text = drConsulta.GetString(3);
				}
				drConsulta.Close();
				this.cmbTipoRuc.Enabled = false;
				this.txtRuc.Enabled = false;
			}

			if (iEmpresa == 34)
			{				
				string sSQL = string.Format("Select idCliente, idTipoRuc, Ruc, Nombre From Cliente Where Proveedor = 0 And Ruc = '1793101135001'", this.txtRuc.Text.ToString());
				SqlDataReader drConsulta = Funcion.rEscalarSQL(cdsSeteo, sSQL, true);
				drConsulta.Read();
				if (drConsulta.HasRows)
				{	
					this.cmbTipoRuc.Value = drConsulta.GetInt32(1);
					idCliente = drConsulta.GetInt32(0);
					this.txtRuc.Text = drConsulta.GetString(2);
					this.txtNombre.Text = drConsulta.GetString(3);
				}
				drConsulta.Close();
				this.cmbTipoRuc.Enabled = false;
				this.txtRuc.Enabled = false;
			}

			if (iEmpresa == 36)
			{
				string sSQL = string.Format("Select idCliente, idTipoRuc, Ruc, Nombre From Cliente Where Proveedor = 0 And Ruc = '1793101135001'", this.txtRuc.Text.ToString());
				SqlDataReader drConsulta = Funcion.rEscalarSQL(cdsSeteo, sSQL, true);
				drConsulta.Read();
				if (drConsulta.HasRows)
				{	
					this.cmbTipoRuc.Value = drConsulta.GetInt32(1);
					idCliente = drConsulta.GetInt32(0);
					this.txtRuc.Text = drConsulta.GetString(2);
					this.txtNombre.Text = drConsulta.GetString(3);
				}
				drConsulta.Close();
				this.cmbTipoRuc.Enabled = false;
				this.txtRuc.Enabled = false;
			}
		}

		private void CargaDatosSocios()
		{
			string sSQL = string.Format("Select idCliente, idTipoRuc, Ruc, Nombre From Cliente Where Ruc = '{0}' And Proveedor = 0 And idGrupoCliente = 9", this.txtRuc.Text.ToString());
			SqlDataReader drConsulta = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drConsulta.Read();
			if (drConsulta.HasRows)
			{	
				idCliente = drConsulta.GetInt32(0);
				this.cmbTipoRuc.Value = drConsulta.GetInt32(1);
				this.txtRuc.Text = drConsulta.GetString(2);
				this.txtNombre.Text = drConsulta.GetString(3);
			}
			drConsulta.Close();
		}

		private void btnCliente_Click(object sender, System.EventArgs e)
		{
			if (this.txtRuc.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el Numero de Identificacion del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
			}
			else
			{
				FuncionesProcedimientos.AbreFacturacionClientes((int)this.cmbTipoRuc.Value, this.txtRuc.Text.ToString(), true,  miAcceso.EditarCuentaNombre);

				CargaCliente();
			}
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbTipoRuc.ActiveRow != null)
				{					
					#region Valida Socio
//					if (vValidaRucEmpleado)
//					{
//						int ValCliente = 0;
//						if (this.txtRuc.Text.Length == 0)
//						{
//							MessageBox.Show("Ingrese Numero identificacion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							return;
//						}
//						if (this.txtRuc.Text.Length < 9)
//						{
//							MessageBox.Show("Ingrese Numero identificacion correcto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							return;
//						}
//
//						if (this.txtRuc.Text.Length > 9)
//						{					
//							#region Validacion								
//							string sSQLValCli = string.Format("Select COUNT(1) From Cliente Where idGrupoCliente = 9 And Proveedor = 0  And Ruc = '{0}'", this.txtRuc.Text);		
//							ValCliente = Funcion.iEscalarSQL(cdsSeteoF, sSQLValCli);
//							#endregion Validacion
//
//							if (ValCliente == 1)
//							{
//								CargaDatosSocios();	
//							}
//							else
//							{
//								MessageBox.Show("Validar informacion del empleado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//								this.txtNombre.Value = "";
//								this.txtRuc.Focus();
//								return;
//							}				
//						}					
//					}
					#endregion Valida Socio

					#region Busca Cliente 
//					if (!vValidaRucEmpleado)
//					{
						idCliente = 0;
						this.txtNombre.Text = "";

						CargaCliente();

						if (idCliente == 0)
						{
							FuncionesProcedimientos.AbreFacturacionClientes((int)this.cmbTipoRuc.Value, this.txtRuc.Text.ToString(), true, miAcceso.EditarCuentaNombre);

							CargaCliente();
						}
//					}
					#endregion Busca Cliente

					//if(idCompra > 0)
					ValidaAnticiposClientes();

					ValidaPromoReturnPass();
				}
			}
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbTipoRuc.ActiveRow != null)
				{
					if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF))
					{
						e.Cancel = true;
						return;
					}
				}
			}
			int idFormaPago = 0;
			if (this.cmbFormaPago.Value != null && this.cmbFormaPago.Value != DBNull.Value)
			{
				idFormaPago = Convert.ToInt32(this.cmbFormaPago.Value);
			}
			if (idFormaPago == 30)
			{
				if (this.txtRuc.Value == null || this.txtRuc.Value.ToString().Trim() == "")
				{
					MessageBox.Show("Ingrese la cédula o RUC del cliente.","Point Technology",MessageBoxButtons.OK,MessageBoxIcon.Information);
					e.Cancel = true;
					return;
				}
				if (this.dtFecha.Value == null || this.dtFecha.Value == DBNull.Value)
				{
					MessageBox.Show("Ingrese la fecha de la factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					e.Cancel = true;
					return;
				}
				DateTime fecha = (DateTime)this.dtFecha.Value;
				string cedulaRuc = this.txtRuc.Value.ToString().Trim();

				string sSQL = string.Format( "Exec RetornaTotalEmpleado '{0}','{1}'", cedulaRuc, fecha.ToString("yyyyMMdd"));

				decimal totalEmpleado = Funcion.decEscalarSQL(cdsSeteoF, sSQL);

				if (totalEmpleado > 40m)
				{
					MessageBox.Show(
						"El total del EMPLEADO supera el límite permitido de 40.","Point Technology",MessageBoxButtons.OK,MessageBoxIcon.Information);
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
			}
		}

		private void txtTotal_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtFechaSalida_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbHabitaciones.Focus();
		}

		private void cmbHabitaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
				if (miAcceso.CambiarFecha) this.dtFecha.Focus(); 
				else if (bNuevo) this.cmbGuardia.Focus(); 
				else this.cmbTipoRuc.Focus();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.cmbGuardia.Enabled) this.cmbGuardia.Focus(); else this.cmbTipoRuc.Focus();
		}

		private void dtFechaIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTipoRuc.Focus();
		}

		private void cmbHabitaciones_ValueChanged(object sender, System.EventArgs e)
		{
			if (!bFormCargado) return;
			if (bNuevo && bEdicion)
			{
				CargarTiempoEstadiaDesdeHabitacion();
				
			}
			if (bCargandoHabitaciones) return;

			if (bNuevo && MenuLatinium.IdUsuario == 279)
			{
				if (DialogResult.Yes == MessageBox.Show("¿Desea Cargar el Servicio de hospedaje?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				{
					//LimpiarHabitacion();
					CargaHabitacion(0, 1);
					CargarTiempoEstadiaDesdeHabitacion();
				}
			}			
		}

		private void dtFechaIngreso_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbHabitaciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtPersonas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNotas.Focus();
		}

		private void txtPersonas_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.txtPersonas.Value == System.DBNull.Value) this.txtPersonas.Value = 2;

				iIngresos = Convert.ToInt32(Math.Round(Convert.ToDecimal(this.txtPersonas.Value)/2.0m));
									
				if (iIngresos > 1 && DialogResult.Yes == MessageBox.Show(string.Format("Si ingresan {0} Personas se registraran {1} servicios\n\n¿Esta seguro de continuar?", (int)this.txtPersonas.Value, iIngresos), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					#region Carga Habitacion - actualiza cantidades
					if (this.uGridDetalle.Rows.Count == 0)
						this.CargaHabitacion(0, iIngresos);
					else
					{
						int iContH = 0;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
							if (dr.Cells["idArticulo"].Value.ToString().Length > 0)
								if ((int)dr.Cells["idArticulo"].Value == 3422) iContH++;

						if (iContH == 0) this.CargaHabitacion(0, iIngresos);
						else
						{
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
								if (dr.Cells["idArticulo"].Value.ToString().Length > 0)
									if ((int)dr.Cells["idArticulo"].Value == 3422) dr.Cells["Cantidad"].Value = iIngresos;
						}
					}
					#endregion Carga Habitacion - actualiza cantidades

					FuncionesProcedimientos.RetornaFechaServidor(this.dtFechaSalida, cdsSeteoF, true);

					this.optTurnoSalida.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTurno '{0}'", Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd HH:mm")));
				
					this.btnGuardar.Enabled = true;
					this.btnTemporal.Enabled = false;
				}
				else
				{
					if (this.uGridDetalle.Rows.Count > 0)
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
							if (dr.Cells["idArticulo"].Value.ToString().Length > 0)
								if ((int)dr.Cells["idArticulo"].Value == 3422) dr.Cells["Cantidad"].Value = iIngresos;

					this.txtPersonas.Value = 2;
					this.btnGuardar.Enabled = false;
					this.btnTemporal.Enabled = true;

					iIngresos = 1;
				}
			}
		}

		private void txtPersonas_Click(object sender, System.EventArgs e)
		{
			this.txtPersonas.SelectAll();
		}

		private void txtPersonas_Enter(object sender, System.EventArgs e)
		{
			this.txtPersonas.SelectAll();
		}

		private void uGridArticulos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtEncuesta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos (sender, e);
		}

		private void txtEncuesta2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos (sender, e);
		}

		private void btnClaveDescuento_Click(object sender, System.EventArgs e)
		{
			#region Variables
			int RetiFormaPago = 0;
			int RetiArticulo = 0;
			int PDiArticulo = 0;
			int FHidentificacion = 0;
			decimal PDdPrecio = 0.00m;
			#endregion Variables

			#region Validacion
			if (this.cmbFormaPago.Value == DBNull.Value || (int)this.cmbFormaPago.Value == 0)
			{
				MessageBox.Show("Seleccione una Forma de Pago ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbFormaPago.Focus();
				return;
			}
			else
			{
				RetiFormaPago = (int)this.cmbFormaPago.Value;
			}
			#endregion Validacion

			#region Descuento Factura
			if (bBtnDesFactura)
			{
				if (RetiFormaPago != 29 || RetiFormaPago != 30)
				{
					decimal RetdDescuento = 0.00m;
					int ValiArticulo = 0;
					if (this.uGridDetalle.Rows.Count <= 0)
					{
						MessageBox.Show("Ingrese Productos a la Factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.btnVer.Focus();
						return;
					}
					else
					{
						using (frmPorcentajeLotte miDescuento = new frmPorcentajeLotte())
						{
							if (DialogResult.OK == miDescuento.ShowDialog())
							{
								RetdDescuento = miDescuento.dDescuento;
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
								{
									ValiArticulo = (int)dr.Cells["idArticulo"].Value;
									bValHoraDes = miDescuento.bHora;

									if (ValiArticulo == 3422)
									{
										dr.Cells["DescuentoPorc"].Value = RetdDescuento;
									}
									#region controles
									dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
									dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
									this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFPagoCarga 2");
									this.btnGuardar.Enabled = true;
									this.btnCancelar.Enabled = true;							
									this.btnSalir.Enabled = false;
									this.btnNuevo.Enabled = false;
									this.btnConsultar.Enabled = false;
									this.btnEditar.Enabled = false;
									this.btnImprimir.Enabled = false;
									this.btnAsiento.Enabled = false;
									this.btnAnular.Enabled = false;
									this.btnClaveDescuento.Enabled = false;
									this.btnDescEmpAut.Enabled = false;
									this.btnRefacturacion.Enabled = false;
									vValFPago = 2;
									this.btnTemporal.Enabled = true;
									if (bValHoraDes) this.txtTiempo.Enabled = true;
									#endregion controles
									Total();
									this.btnGuardar.Enabled = true;
								}
							}
							else
							{
								return;
							}
						}
					}
				}
			}
			#endregion Descuento Factura
		}

		private void ValDetalleFormaPago(int ValidCompraOriginal)
		{
			countFP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From DetalleFormaPago Where idCompra = {0} ", ValidCompraOriginal));
			if (countFP > 1)
			{
				iCountFP = 1;
				idOriginal = ValidCompraOriginal;
			}
		}
		DateTime dtFechaRef = DateTime.Today;
 
		private void btnRefacturacion_Click(object sender, System.EventArgs e)
		{		
			#region Variables
			int iCompraOrigen = 0;
			int iCompraDestino = 0;
			int iCompraNDC = 0;
			int iCompraDestinoNDC = 0;
			int iCompraOrigenNDC = 0;
			int iCompraGeneral = 0;
			int ValiCuponUtilizado = 0;
			#endregion Variables			

			int EstadoCaja = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec EstadoCajaRefacturacion '{0}', {1}, {2}", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm"), (int)this.optTurnoSalida.Value, (int)this.cmbBodega.Value));

			int RetiCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From Compra Where idcompra = {0}", idCompra));

			int ValidaTVIPUtilizados = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From TarjetasVIpCompras Where idTipoFactura = 1 And idCompra = {0}", idCompra));

			if(ValidaTVIPUtilizados > 0)
			{
				MessageBox.Show("No se puede refacturar, registra cruce con TARJETA VIP.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (RetiCompra == 1)
			{
				int iValMinutos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaMinutosFactura {0}", idCompra));
				int iFactDivision = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec EstadoFacturaDivision {0}", idCompra));
				int iValCompraCupon = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesAnticipoCompra Where idCompraDivision = {0}", idCompra));

				if (iValMinutos > 60)
				{
					MessageBox.Show("Pasado la Hora no se puede Refacturar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (iFactDivision == 1)
				{
					string sSQLConsulta = string.Format("Exec ConsultaCompraOrigenDestino {0}", idCompra);
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLConsulta, true);
					dr.Read();
					if (dr.HasRows)
					{
						iCompraOrigen = dr.GetInt32(0);
						iCompraDestino = dr.GetInt32(1);
					}
					dr.Close();
				}

				if (iValCompraCupon >= 1)
				{
					MessageBox.Show("No puede refacturar, la factura registra un CUPÓN.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCupon.Focus();
					return;
				}

				if ((int)this.cmbFormaPago.Value == 14)
				{
					MessageBox.Show("No puede Refacturar Cupon", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					bEdicion = false;
					return;	
				}

				if ((int)this.cmbFormaPago.Value == 16)
				{
					MessageBox.Show("No puede Refacturar Prenda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					bEdicion = false;
					return;	
				}

				if ((int)this.cmbFormaPago.Value == 29)
				{
					MessageBox.Show("No puede Refacturar Autoconsumo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					bEdicion = false;
					return;	
				}

				if ((int)this.cmbFormaPago.Value == 30)
				{
					MessageBox.Show("No puede Refacturar Descuento Empleado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					bEdicion = false;
					return;	
				}

				ValiCuponUtilizado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesEmpresas Where idCompraDestino = {0}", idCompra));
				if (ValiCuponUtilizado == 1)
				{
					MessageBox.Show("No puede Refacturar, la factura contiene un cupon.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					bEdicion = false;
					return;
				}

				if (bAnulado)
				{
					MessageBox.Show("El Documento esta Anulado", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (EstadoCaja == 2)
				{
					MessageBox.Show("No se puede Refacturar si la CAJA ya esta Cerrada", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			
				#region Crea Factura
				if (DialogResult.Yes == MessageBox.Show("Esta seguro de REFACTURAR este Documento \n\nEste proceso no puede ser revertido y no podra cancelar el proceso.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
					{
						oConexion.Open();

						SqlTransaction oTransaction = null;
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandType = CommandType.Text;
						oCmdActualiza.CommandTimeout = 120;

						try
						{	
							oTransaction = oConexion.BeginTransaction();
							oCmdActualiza.Transaction = oTransaction;

							#region Numeracion
							string sNumVenta = "";
							string sSQLNumero = string.Format("Exec NumeracionLocales 1, {0}, 1", (int)this.cmbBodega.Value);
							oCmdActualiza.CommandText = sSQLNumero;
							sNumVenta = oCmdActualiza.ExecuteScalar().ToString();				
							#endregion Numeracion

							#region Notas de Credito
							if((int)this.cmbTipoRuc.Value == 11)
							{
								if (iCompraOrigen > 0)
								{
									string sSQLAnulacion = string.Format("Exec AnulacionCompraVenta {0}, {1}, '{2}'", iCompraOrigen, 1, "REFACTURACIÓN");
									oCmdActualiza.CommandText = sSQLAnulacion;
									oCmdActualiza.ExecuteNonQuery();
								}
								if (iCompraDestino > 0)
								{
									string sSQLAnulacion = string.Format("Exec AnulacionCompraVenta {0}, {1}, '{2}'", iCompraDestino, 1, "REFACTURACIÓN");
									oCmdActualiza.CommandText = sSQLAnulacion;
									oCmdActualiza.ExecuteNonQuery();
								}
								if (iCompraOrigen == 0 && iCompraDestino == 0)
								{
									string sSQLAnulacion = string.Format("Exec AnulacionCompraVenta {0}, {1}, '{2}'", idCompra, 1, "REFACTURACIÓN");
									oCmdActualiza.CommandText = sSQLAnulacion;
									oCmdActualiza.ExecuteNonQuery();
								}
							}
							else
							{
								if (iCompraOrigen > 0)
								{
									#region Numeracion
									string sNumNDC = "";
									string sSQLNumeracion = string.Format("Exec NumeracionLocales 5, {0}, 1", (int)this.cmbBodega.Value);
									oCmdActualiza.CommandText = sSQLNumeracion;
									sNumNDC = oCmdActualiza.ExecuteScalar().ToString();
									#endregion Numeracion

									string sSQLTransforma = string.Format("Exec CreaNotaDeCreditoCambioProducto {0}, '{1}', {2}, '{3}', {4}, {5}",
										iCompraOrigen, sNumNDC, 5, "REFACTURACION", 1, (int)this.cmbBodega.Value);
									oCmdActualiza.CommandText = sSQLTransforma;
									iCompraNDC = (int)oCmdActualiza.ExecuteScalar();

									#region Kardex
									string sSQLDetalleKardex = string.Format("Exec MovimientoKardexFacturaDivididaNDC {0}, {1}, '{2}', {3}, '{4}', {5}",
										iCompraNDC, 5, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value, 
										sNumNDC, idCliente);
									oCmdActualiza.CommandText = sSQLDetalleKardex;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Kardex

									#region Autorizacion Nota de Credito
									oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 5", iCompraNDC);;
									string sClaveAccesoOR = oCmdActualiza.ExecuteScalar().ToString();

									string cadInvertidaOR = invertirCadena(sClaveAccesoOR.Substring(0, 48));
									string ValM11OR = Modulo11(cadInvertidaOR).ToString().Trim();
									string sClaveSRIOR = sClaveAccesoOR.ToString().Substring(0, 48) + ValM11OR;

									oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRIOR, iCompraNDC);
									oCmdActualiza.ExecuteNonQuery();
									#endregion Autorizacion Nota de Credito
								}	
							
								if (iCompraDestino > 0)
								{
									#region Numeracion
									string sNumNDC = "";
									string sSQLNumeracion = string.Format("Exec NumeracionLocales 5, {0}, 1", (int)this.cmbBodega.Value);
									oCmdActualiza.CommandText = sSQLNumeracion;
									sNumNDC = oCmdActualiza.ExecuteScalar().ToString();
									#endregion Numeracion

									string sSQLTransforma = string.Format("Exec CreaNotaDeCreditoCambioProducto {0}, '{1}', {2}, '{3}', {4}, {5}",
										iCompraDestino, sNumNDC, 5, "REFACTURACION", 1, (int)this.cmbBodega.Value);
									oCmdActualiza.CommandText = sSQLTransforma;
									iCompraDestinoNDC = (int)oCmdActualiza.ExecuteScalar();

									#region Autorizacion Notas de Credito
									oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 5", iCompraDestinoNDC);;
									string sClaveAccesoDE = oCmdActualiza.ExecuteScalar().ToString();

									string cadInvertidaDE = invertirCadena(sClaveAccesoDE.Substring(0, 48));
									string ValM11DE = Modulo11(cadInvertidaDE).ToString().Trim();
									string sClaveSRIOR = sClaveAccesoDE.ToString().Substring(0, 48) + ValM11DE;

									oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRIOR, iCompraDestinoNDC);
									oCmdActualiza.ExecuteNonQuery();
									#endregion Autorizacion Notas de Credito
								}
							
								if (iCompraOrigen == 0 && iCompraDestino == 0)
								{
									#region Numeracion
									string sNumNDC = "";
									string sSQLNumeracion = string.Format("Exec NumeracionLocales 5, {0}, 1", (int)this.cmbBodega.Value);
									oCmdActualiza.CommandText = sSQLNumeracion;
									sNumNDC = oCmdActualiza.ExecuteScalar().ToString();
									#endregion Numeracion

									string sSQLTransforma = string.Format("Exec CreaNotaDeCreditoCambioProducto {0}, '{1}', {2}, '{3}', {4}, {5}",
										idCompra, sNumNDC, 5, "REFACTURACION", 1, (int)this.cmbBodega.Value);
									oCmdActualiza.CommandText = sSQLTransforma;
									iCompraOrigenNDC = (int)oCmdActualiza.ExecuteScalar();

									#region Kardex
									string sSQLDetalleKardex = string.Format("Exec MovimientoKardexFacturaDivididaNDC {0}, {1}, '{2}', {3}, '{4}', {5}",
										iCompraOrigenNDC, 5, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value, 
										sNumNDC, idCliente);
									oCmdActualiza.CommandText = sSQLDetalleKardex;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Kardex

									#region Reverso Anticipos Abono Cliente
									string sSQLReversoAnticipoAbCl = string.Format("Exec ReversoAnticiposAbonoClientes {0}", idCompra);
									oCmdActualiza.CommandText = sSQLReversoAnticipoAbCl;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Reverso Anticipos Abono Cliente

									#region Autorizacion Notas de Credito
									oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 5", iCompraOrigenNDC);;
									string sClaveAccesoNC = oCmdActualiza.ExecuteScalar().ToString();

									string cadInvertidaNC = invertirCadena(sClaveAccesoNC.Substring(0, 48));
									string ValM11NC = Modulo11(cadInvertidaNC).ToString().Trim();
									string sClaveSRIOR = sClaveAccesoNC.ToString().Substring(0, 48) + ValM11NC;

									oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRIOR, iCompraOrigenNDC);
									oCmdActualiza.ExecuteNonQuery();
									#endregion Autorizacion Notas de Credito
								}
							}
							#endregion Notas de Credito

							#region Anulacion
							//							if (iCompraOrigen > 0)
							//							{		
							//								string sSQL = string.Format("Exec AnulacionCompraVenta {0}, {1}, '{2}'", iCompraOrigen, 1, "REFACTURACIÓN");
							//								oCmdActualiza.CommandText = sSQL;
							//								oCmdActualiza.ExecuteNonQuery();
							//							}
							//							
							//							if (iCompraDestino > 0)
							//							{
							//								string sSQL = string.Format("Exec AnulacionCompraVenta {0}, {1}, '{2}'", iCompraDestino, 1, "REFACTURACIÓN");
							//								oCmdActualiza.CommandText = sSQL;
							//								oCmdActualiza.ExecuteNonQuery();
							//							}
							//							
							//							if (iCompraOrigen == 0 && iCompraDestino == 0)
							//							{
							//								string sSQL = string.Format("Exec AnulacionCompraVenta {0}, {1}, '{2}'", idCompra, 1, "REFACTURACIÓN");
							//								oCmdActualiza.CommandText = sSQL;
							//								oCmdActualiza.ExecuteNonQuery();
							//							}
							#endregion Anulacion

							#region Unificacion Factura
							if (iCompraOrigen > 0 && iCompraDestino > 0)
							{
								string sSQLTransforma = string.Format("Exec CompraTransformaCopiaUnificacionMaestro {0}, {1}, {2}, {3}, {4}, {5}, '{6}', '{7}', '{8}'",
									iCompraOrigen, iCompraDestino, (int)this.cmbBodega.Value, false, false, false, "", sNumVenta, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm"));
								oCmdActualiza.CommandText = sSQLTransforma;
								iCompraGeneral = (int)oCmdActualiza.ExecuteScalar();

								string sSQLDetalle = string.Format("Exec CompraTransformaCopiaUnificacionDetalle {0}, {1}, {2}", iCompraOrigen, iCompraDestino, iCompraGeneral);
								oCmdActualiza.CommandText = sSQLDetalle;
								oCmdActualiza.ExecuteNonQuery();

								string sSQLDetalleValores = string.Format("Exec ActualizaValoresFactura {0}, {1}, {2}", iCompraOrigen, iCompraDestino, iCompraGeneral);
								oCmdActualiza.CommandText = sSQLDetalleValores;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Unificacion Factura						

							#region Nueva Venta
							if (iCompraOrigen == 0 && iCompraDestino == 0)
							{
								string sSQLTransforma = string.Format("Exec CompraTransformaCopia {0}, 1, {1}, {2}, {3}, {4}, '{5}', '{6}', '{7}'",
									idCompra, (int)this.cmbBodega.Value, false, false, false, "", sNumVenta, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm"));
								oCmdActualiza.CommandText = sSQLTransforma;
								idCompraDest = (int)oCmdActualiza.ExecuteScalar();
							}
							else
							{
								idCompraDest = iCompraGeneral;
							}
							#endregion Nueva Venta

							#region Autorizacion Factura
							if (Funcion.bEscalarSQL(cdsSeteoF, "Select FacturaElectronica From SeteoF", false))
							{
								oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 1", idCompraDest);;
								string sClaveAcceso = oCmdActualiza.ExecuteScalar().ToString();

								string cadInvertida = invertirCadena(sClaveAcceso.Substring(0, 48));
								string ValM11 = Modulo11(cadInvertida).ToString().Trim();
								string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

								oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, idCompraDest);
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Autorizacion Factura

							#region Promocion Temporal Video
							string sSQLRPV = string.Format("Exec GuardaPromocionVideo {0}, {1}, {2}", idCompraDest, (int)this.cmbHabitaciones.Value, false);
							oCmdActualiza.CommandText = sSQLRPV;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Promocion Temporal Video
							
							oTransaction.Commit();
							//oConexion.Close();
						}						
						catch (Exception ex)
						{
							oTransaction.Rollback();
							oConexion.Close();
							MessageBox.Show("Error", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);												
							return;
						}
						finally
						{
							oConexion.Close();

							ImprimeRefacturacion = true;
							ValDetalleFormaPago(idCompra);
							dtFechaRef = Convert.ToDateTime(this.dtFecha.Value);
							this.btnCancelar_Click(sender, e);
							dtfechaFact = true;
							bValRefacturacion = true;
							this.Consulta(idCompraDest);
							EdicionRefactura();						
						}
					}
				}
				#endregion Crea Factura
			}
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbCobrador_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbFormaPago_ValueChanged(object sender, System.EventArgs e)
		{
			if(vValidaFPAgo)
			{
				if (this.cmbFormaPago.Value != System.DBNull.Value)
				{
					if ((int)this.cmbFormaPago.Value == 29 || (int)this.cmbFormaPago.Value == 30) 
					{
						if (miAcceso.BDesEmpleado) 
						{
							this.btnDescEmpAut.Enabled = true;
						}
						else 
						{
							this.btnDescEmpAut.Enabled = false;
						}
					}

					#region Efectivo
					if((int)this.cmbFormaPago.Value == 1)
					{												
						if(idHistorialTarjetasVIP > 0) 
						{
							bUsoGiftCard = false;
							this.txtNotas.Text = "";
							idCliente = 0;
							idHistorialTarjetasVIP = 0;
							dGiftCard = 0;
							this.txtRuc.Text = "";
							this.txtNombre.Text = "";
						}

						if(Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select AbonoCliente From SeteoC"), true))
						{
							if(idCompra > 0)
							{
								this.chkAbono.Enabled = true;
							}
						}
					}
					#endregion Efectivo

					#region Tarjeta de Credito
					if ((int)this.cmbFormaPago.Value == 3) 
					{
						this.chkAbono.Enabled = false;
						this.chkAbono.Checked = false;

						if(idHistorialTarjetasVIP > 0) 
						{
							bUsoGiftCard = false;
							this.txtNotas.Text = "";
							idCliente = 0;
							idHistorialTarjetasVIP = 0;
							dGiftCard = 0;
							this.txtRuc.Text = "";
							this.txtNombre.Text = "";
						}
					}
					#endregion Tarjeta de Credito

					#region Prenda
					if((int)this.cmbFormaPago.Value == 16)
					{
						this.chkAbono.Enabled = false;
						this.chkAbono.Checked = false;

						if(idHistorialTarjetasVIP > 0) 
						{
							bUsoGiftCard = false;
							this.txtNotas.Text = "";
							idCliente = 0;
							idHistorialTarjetasVIP = 0;
							dGiftCard = 0;
							this.txtRuc.Text = "";
							this.txtNombre.Text = "";
						}
					}
					#endregion Prenda

					#region Autoconsumos
					if ((int)this.cmbFormaPago.Value == 29) 
					{
						if(idHistorialTarjetasVIP > 0) 
						{
							bUsoGiftCard = false;
							this.txtNotas.Text = "";
							idCliente = 0;
							idHistorialTarjetasVIP = 0;
							dGiftCard = 0;
							this.txtRuc.Text = "";
							this.txtNombre.Text = "";
						}

						bUsoGiftCard = false;
						this.txtNotas.Text = "";

						ImprimeConDEs = true; 
						if (miAcceso.BDesEmpleado) this.btnDescEmpAut.Enabled = true;

						CargaClienteEmpresa(MenuLatinium.stIdDB);

						if (this.txtRuc.TextLength == 0)
						{
							MessageBox.Show("Ingrese Datos del Empleado", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtRuc.Focus();
							return;
						}
					}
					#endregion Autoconsumos

					#region Descuento Empleado
					if ((int)this.cmbFormaPago.Value == 30) 
					{
						if(idHistorialTarjetasVIP > 0) 
						{
							bUsoGiftCard = false;
							this.txtNotas.Text = "";
							idCliente = 0;
							idHistorialTarjetasVIP = 0;
							dGiftCard = 0;
							this.txtRuc.Text = "";
							this.txtNombre.Text = "";
						}

						bUsoGiftCard = false;
						this.txtNotas.Text = "";

						ImprimeConDEs = true; 
						if (miAcceso.BDesEmpleado) this.btnDescEmpAut.Enabled = true;

						#region Control
						idCliente = 0;
						this.cmbTipoRuc.Value = 5;
						this.txtRuc.Text = "";
						this.txtNombre.Text = "";
						this.cmbTipoRuc.Enabled = true;
						this.txtRuc.Enabled = true;
						#endregion COntrol
					}
					#endregion Descuento Empleado					

					#region Anticipo
					//					if ((int)this.cmbFormaPago.Value != 8) this.btnGuardar.Enabled = true;
					//
					//					if ((int)this.cmbFormaPago.Value == 8)
					//					{
					//						this.chkAbono.Enabled = false;
					//						this.chkAbono.Checked = false;
					//						if (this.txtRuc.TextLength == 0)
					//						{
					//							MessageBox.Show("Ingrese Datos del Cliente", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//							this.txtRuc.Focus();
					//							return;
					//						}
					//
					//						iCliente = idCliente;
					//						bUsoAnticipo = true;
					//
					//						using (frmListaAnticipos FRLA = new frmListaAnticipos(iCliente, (int)cmbBodega.Value))  //error
					//						{				
					//							if (DialogResult.OK == FRLA.ShowDialog())
					//							{
					//								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drNC in FRLA.uGrid.Rows.All)
					//								{
					//									if ((bool)drNC.Cells["Sel"].Value)
					//									{									
					//										IdAnticipo = (int)drNC.Cells["idAnticipo"].Value;
					//										sDocumento = (string)drNC.Cells["Numero"].Value;
					//										dSaldo = (decimal)drNC.Cells["Saldo"].Value;
					//										sMensajeAnticipo = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec MnsjAnticipo {0}, {1}", IdAnticipo, iCliente));
					//										this.txtNotas.Text = sMensajeAnticipo;
					//										this.txtNotas.Enabled = false;
					//									}
					//								}
					//
					//								if(dSaldo == 0 || IdAnticipo == 0)
					//								{
					//									MessageBox.Show("El Saldo del anticipo esta en Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//									this.cmbFormaPago.Focus();
					//									this.btnGuardar.Enabled = false;
					//									this.cmbFormaPago.Value = 1;
					//									return;
					//								}
					//								this.btnGuardar.Enabled = true;
					//							}
					//							else
					//							{
					//								MessageBox.Show("No a seleccionado un anticipo o no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//								this.btnGuardar.Enabled = false;
					//								this.cmbFormaPago.Value = 1;
					//								this.cmbFormaPago.Focus();							
					//								return;
					//							}
					//						}
					//					}
					//					else
					//					{
					//						this.txtNotas.Text = "";
					//						this.txtNotas.Enabled = true;
					//						IdAnticipo = 0;
					//						iCliente = 0;
					//						dSaldo = 0;
					//						sDocumento = "";
					//					}
					#endregion Anticipo

					#region Cupon
					//					if ((int)this.cmbFormaPago.Value == 14)
					//					{
					//						this.chkAbono.Enabled = false;
					//						this.chkAbono.Checked = false;
					//
					//						int iValCuponBodega = 0;
					//						int iValDiaCupon = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDiaCupon"));						
					//
					//						if(iValDiaCupon == 0)
					//						{
					//							MessageBox.Show("El dia actual no acepta CUPONES", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//							this.cmbTipoRuc.Value = 5;
					//							this.txtRuc.Focus();
					//							this.cmbFormaPago.Value = 1;
					//							return;
					//						}
					//						if ((int)this.cmbTipoRuc.Value == 11)
					//						{
					//							MessageBox.Show("CUPON valido solo con datos Cliente", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);														
					//							this.cmbTipoRuc.Value = 5;
					//							this.txtRuc.Focus();
					//							this.cmbFormaPago.Value = 1;
					//							return;
					//						}
					//						if (this.txtRuc.TextLength == 0)
					//						{
					//							MessageBox.Show("Cedula/RUC/Pasaporte vacio, Ingrese Datos del Cliente", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//							this.txtRuc.Focus();
					//							return;
					//						}
					//						if (this.txtNombre.TextLength == 0)
					//						{
					//							MessageBox.Show("Nombre del cliente vacio, Ingrese Datos del Cliente", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//							this.txtRuc.Focus();
					//							return;
					//						}
					//						bUsoAnticipo = true;
					//						using (frmListaCupones FRLC = new frmListaCupones())
					//						{				
					//							if (DialogResult.OK == FRLC.ShowDialog())
					//							{								
					//								bCuponMP = true;
					//								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drLC in FRLC.uGrid.Rows.All)
					//								{
					//									if ((bool)drLC.Cells["Sel"].Value)
					//									{
					//										iAnticipoCupon = (int)drLC.Cells["idAnticipo"].Value;
					//										sDocumentoCupon = (string)drLC.Cells["Numero"].Value;
					//										sDocumento = (string)drLC.Cells["Numero"].Value;
					//										dSaldoCupon = (decimal)drLC.Cells["Saldo"].Value;
					//										dSaldo = (decimal)drLC.Cells["Saldo"].Value;
					//										iValCuponBodega = (int)drLC.Cells["Bodega"].Value;
					//										sMensajeAnticipo = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec MnsjCupon {0}", iAnticipoCupon));
					//										this.txtNotas.Text = sMensajeAnticipo;
					//										this.txtNotas.Enabled = false;
					//									}
					//								}
					//
					//								#region Validaciones
					//								int iValCLienteMP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ca.idCliente From CuponesAnticipoCompra ca Inner Join Cliente cl On ca.idCliente = cl.idCliente Where ca.idAnticipo = {0} ", iAnticipoCupon));
					//								int iValValorHab = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaValorHabitacion {0}, {1}", (int)this.cmbHabitaciones.Value, iAnticipoCupon));
					//
					//								if(iValValorHab == 1)
					//								{
					//									MessageBox.Show("CUPON no valido, el valor de hospedaje no es el mismo.", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//									this.btnGuardar.Enabled = false;
					//									bCuponMP = false;
					//									this.txtNotas.Text = "";
					//									this.txtNotas.Enabled = true;
					//									iAnticipoCupon = 0;
					//									dSaldoCupon = 0;
					//									sDocumentoCupon = "";
					//									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					//									{
					//										if((int)dr.Cells["idArticulo"].Value == 3422)
					//										{
					//											dr.Cells["DescuentoPorc"].Value = 0.00;
					//											this.Total();
					//										}
					//									}
					//									return;
					//								}
					//
					//								if(idCliente != iValCLienteMP)
					//								{
					//									MessageBox.Show("El CUPON no pertenece al mismo cliente.", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//									this.btnGuardar.Enabled = false;
					//									this.txtRuc.Focus();
					//									bCuponMP = false;
					//									this.cmbFormaPago.Value = 1;
					//									this.txtNotas.Text = "";
					//									this.txtNotas.Enabled = true;
					//									iAnticipoCupon = 0;
					//									dSaldoCupon = 0;
					//									sDocumentoCupon = "";
					//									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					//									{
					//										if((int)dr.Cells["idArticulo"].Value == 3422)
					//										{
					//											dr.Cells["DescuentoPorc"].Value = 0.00;
					//											this.Total();
					//										}
					//									}									
					//									return;
					//								}
					//
					//								if(iValCuponBodega != (int)this.cmbBodega.Value)
					//								{
					//									MessageBox.Show("El CUPON no pertenece al mismo establecimiento.", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//									this.btnGuardar.Enabled = false;
					//									bCuponMP = false;
					//									this.txtNotas.Text = "";
					//									this.txtNotas.Enabled = true;
					//									iAnticipoCupon = 0;
					//									dSaldoCupon = 0;
					//									sDocumentoCupon = "";
					//									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					//									{
					//										if((int)dr.Cells["idArticulo"].Value == 3422)
					//										{
					//											dr.Cells["DescuentoPorc"].Value = 0.00;
					//											this.Total();
					//										}
					//									}
					//									return;
					//								}
					//								#endregion Validaciones
					//
					//								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					//								{
					//									if((int)dr.Cells["idArticulo"].Value == 3422)
					//									{
					//										decimal dDescuento = 50;
					//										dr.Cells["DescuentoPorc"].Value = dDescuento;
					//										this.Total();
					//									}
					//									else
					//									{
					//										dr.Cells["DescuentoPorc"].Value = 0.00;
					//									}
					//								}
					//								if(dSaldoCupon == 0 || iAnticipoCupon == 0)
					//								{
					//									MessageBox.Show("El Saldo del Cupon esta en Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//									this.cmbFormaPago.Focus();
					//									this.btnGuardar.Enabled = false;
					//									this.cmbFormaPago.Value = 1;
					//									return;
					//								}
					//								this.btnGuardar.Enabled = true;
					//							}
					//							else
					//							{
					//								MessageBox.Show("No a seleccionado un Cupon o no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//								this.btnGuardar.Enabled = false;
					//								bCuponMP = false;
					//								this.cmbFormaPago.Value = 1;
					//								this.cmbFormaPago.Focus();							
					//								return;
					//							}
					//						}
					//					}
					//					else
					//					{
					//						if((int)this.cmbFormaPago.Value == 14)
					//						{
					//							bCuponMP = false;
					//							this.txtNotas.Text = "";
					//							this.txtNotas.Enabled = true;
					//							iAnticipoCupon = 0;
					//							dSaldoCupon = 0;
					//							sDocumentoCupon = "";
					//							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					//							{
					//								if((int)dr.Cells["idArticulo"].Value == 3422)
					//								{
					//									dr.Cells["DescuentoPorc"].Value = 0.00;
					//									this.Total();
					//								}
					//							}
					//						}
					//						else
					//						{
					//							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					//							{
					//								if((int)dr.Cells["idArticulo"].Value == 3422)
					//								{
					//									dr.Cells["DescuentoPorc"].Value = 0.00;
					//									this.Total();
					//								}
					//							}
					//						}
					//					}
					#endregion Cupon

					#region Gift Card
					if((int)this.cmbFormaPago.Value == 22)
					{
						this.chkAbono.Enabled = false;
						this.chkAbono.Checked = false;

						#region Validaciones
						#endregion Validaciones

						using (frmCodigoTarjetasVIPlottes miCodigo = new frmCodigoTarjetasVIPlottes())
						{
							if (DialogResult.OK == miCodigo.ShowDialog())
							{
								bUsoGiftCard = true;
								idHistorialTarjetasVIP = miCodigo.idHistorialTarjetasVIP;
								dGiftCard = miCodigo.dSaldo;
								this.cmbTipoRuc.Value = miCodigo.idTipoRuc;
								this.txtNombre.Text = miCodigo.sNombreCliente;
								this.txtRuc.Text = miCodigo.sRucCliente;
								idCliente = miCodigo.iCliente;
								this.txtNotas.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec MensajeGiftCardCompra {0}", idHistorialTarjetasVIP));
							}							
						}
					}
					#endregion Gift Card
				}
			}

			if (bRFAnticipo) // Refacturacion Anticipo
			{
				#region Anticipo
				if (this.cmbFormaPago.Value != System.DBNull.Value)
				{
					this.chkAbono.Enabled = false;

					if ((int)this.cmbFormaPago.Value == 8)
					{
						if (this.txtRuc.TextLength == 0)
						{
							MessageBox.Show("Ingrese Datos del Cliente", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtRuc.Focus();
							return;
						}

						iCliente = idCliente;
						bUsoAnticipo = true;

						using (frmListaAnticipos FRLA = new frmListaAnticipos(iCliente, (int)cmbBodega.Value))  //error
						{				
							if (DialogResult.OK == FRLA.ShowDialog())
							{
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drNC in FRLA.uGrid.Rows.All)
								{
									if ((bool)drNC.Cells["Sel"].Value)
									{									
										IdAnticipo = (int)drNC.Cells["idAnticipo"].Value;
										sDocumento = (string)drNC.Cells["Numero"].Value;
										dSaldo = (decimal)drNC.Cells["Saldo"].Value;
										sMensajeAnticipo = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec MnsjAnticipo {0}, {1}", IdAnticipo, iCliente));
										this.txtNotas.Text = sMensajeAnticipo;
										this.txtNotas.Enabled = false;
									}
								}

								if(dSaldo == 0 || IdAnticipo == 0)
								{
									MessageBox.Show("El Saldo del anticipo esta en Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									this.cmbFormaPago.Focus();
									this.btnGuardar.Enabled = false;
									this.cmbFormaPago.Value = 1;
									return;
								}
								this.btnGuardar.Enabled = true;
							}
							else
							{
								MessageBox.Show("No a seleccionado un anticipo o no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
								this.cmbFormaPago.Value = 1;
								this.cmbFormaPago.Focus();							
								return;
							}
						}
					}
					else
					{
						if((int)this.cmbFormaPago.Value == 8)
						{
							this.txtNotas.Text = "";
							this.txtNotas.Enabled = true;
							IdAnticipo = 0;
							iCliente = 0;
							dSaldo = 0;
							sDocumento = "";
						}
					}
				}
				#endregion Anticipo
			}
		}

		private void btnDescEmpAut_Click(object sender, System.EventArgs e)
		{
			#region Variables
			int RetiFormaPago = 0;
			int RetiArticulo = 0;
			int PDiArticulo = 0;
			int FHidentificacion = 0;
			int iCountDepartamento = 0;
			decimal PDdPrecio = 0.00m;
			#endregion Variables

			#region Validacion
			if (this.cmbFormaPago.Value == DBNull.Value || (int)this.cmbFormaPago.Value == 0)
			{
				MessageBox.Show("Seleccione una Forma de Pago ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbFormaPago.Focus();
				return;
			}
			else
			{
				RetiFormaPago = (int)this.cmbFormaPago.Value;
			}
			#endregion Validacion

			#region Descuento Empleado
			if (bBtnDesEmpleado)
			{
				if (RetiFormaPago == 30 || RetiFormaPago == 29)
				{
					int iUsuarioCodigo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idGrupo From Usuario Where idUsuario = {0}", MenuLatinium.IdUsuario));

					if ((int)this.cmbTipoRuc.Value == 11)
					{
						MessageBox.Show("No Aplica a Consumidor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbTipoRuc.Focus();
						return;
					}
					if (this.txtRuc.Text.Length == 0)
					{
						MessageBox.Show("Ingrese Numero de identificación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtRuc.Focus();
						return;
					}
					if (this.txtRuc.Text.Length > 0)
					{
						if (this.uGridDetalle.Rows.Count <= 0)
						{
							MessageBox.Show("Ingrese Productos a la Factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.btnVer.Focus();
							return;
						}
						else 
						{
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
							{
								RetiArticulo = (int)dr.Cells["idArticulo"].Value;

								string sSQL = string.Format("Exec ValidaProductosDescuento {0}, {1}, {2}", RetiFormaPago, RetiArticulo, iUsuarioCodigo);
								SqlDataReader drr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
								drr.Read();
								if (drr.HasRows)
								{				
									PDiArticulo = drr.GetInt32(0);
									PDdPrecio = drr.GetDecimal(1);
								}
								drr.Close();

								if (PDiArticulo == RetiArticulo)
								{
									dr.Cells["Precio"].Value = PDdPrecio;
								}
								else
								{
									dr.Delete();
								}
							}
							Total();
							this.btnGuardar.Enabled = true;
						}
					}
				}
			}
			#endregion Descuento Empleado
		}

		private void ValidaTiempoAnulacion()
		{
			int ValiCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From Compra Where idcompra = {0}", idCompra));
			if (ValiCompra == 1)
			{
				int iValMinutos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(MINUTE, '{0}', GETDATE()) From Compra Where idCompra = {1}", 
					Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"), idCompra));

				if (iValMinutos > 15)
				{
					this.btnAnular.Enabled = false;
				}
				else
				{
					this.btnAnular.Enabled = true;
				}
			}
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_InitializeEditorButtonCheckState(object sender, Infragistics.Win.UltraWinEditors.InitializeCheckStateEventArgs e)
		{
		
		}

		private void mnuImprimirFactura_Click(object sender, System.EventArgs e)
		{
			Imprimir(0);
		}

		private void mnuImprimirRIDE_Click(object sender, System.EventArgs e)
		{
			if (bAnulado)	
			{
				MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From Compra Where idCompra = {0}", idCompra)) != 10)
			{
				MessageBox.Show("El documento no esta Contabilizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idFe_EstadoComprobante From Compra Where idCompra = {0}", idCompra)) != 2)
			{
				MessageBox.Show("El documento no ha sido autorizado por el S.R.I.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (MenuLatinium.stIdDB == 29)
			{
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = idCompra; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				Reporte Reporte = new Reporte("RIDEFacturaGreta.rpt", "");
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("RIDE");
				Reporte.crVista.ParameterFieldInfo = paramFields;
				Reporte.Show();
			}

			if (MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
			{
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = idCompra; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				Reporte Reporte = new Reporte("RIDEFacturaPro.rpt", "");
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("RIDE");
				Reporte.crVista.ParameterFieldInfo = paramFields;
				Reporte.Show();
			}
		}

		private void mnuImprimirComprobanteVenta_Click(object sender, System.EventArgs e)
		{
			if (bAnulado)	
			{
				MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From Compra Where idCompra = {0}", idCompra)) != 10)
			{
				MessageBox.Show("El documento no esta Contabilizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			// Greta
			if (MenuLatinium.stIdDB == 29)
			{
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = idCompra; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				Reporte Reporte = new Reporte("FacturaLottePV.rpt", "");
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Comprobante Venta");
				Reporte.crVista.ParameterFieldInfo = paramFields;
				Reporte.Show();
			}
			// Pro
			if (MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
			{
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = idCompra; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				Reporte Reporte = new Reporte("FacturaLottePVPro.rpt", "");
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Comprobante Venta");
				Reporte.crVista.ParameterFieldInfo = paramFields;
				Reporte.Show();
			}
		}

		private void btnCupon_Click(object sender, System.EventArgs e)
		{
			#region Variables
			int iValCuponCreado = 0;
			int iValNomCliente = 0;
			#endregion Variables

			#region Validaciones
			iValCuponCreado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesAnticipoCompra Where idCompra = {0}", idCompra));
			iValNomCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesAnticipoCompra ca Inner Join Cliente cl On ca.idCliente = cl.idCliente Where cl.Nombre = '{0}' And ca.Estado = 0", this.txtNombre.Text.ToString()));

			if (iValCuponCreado == 1)
			{
				MessageBox.Show("La factura Actual ya registra un cupón", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (idCompra == 0)
			{
				MessageBox.Show("La Factura debe estar en TEMPORAL", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}			

			if (this.txtRuc.Text.Length == 0)
			{
				MessageBox.Show("No se puede generar un cupon, identificacion esta vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}

			if(this.cmbTipoRuc.Value == System.DBNull.Value || (int)this.cmbTipoRuc.Value == 0)
			{
				MessageBox.Show("No se puede generar un cupon, tipo de identificacion vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipoRuc.Focus();
				return;
			}

			if (this.txtNombre.Text.Length == 0)
			{
				MessageBox.Show("No se puede generar un cupon, Nombre del cliente esta vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}			

			if(iValNomCliente >= 1)
			{
				MessageBox.Show("El Cliente registra un cupon pendiente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validaciones

			if(iActivoMitadPrecion == 1)
			{
				if((int)this.cmbTipoRuc.Value != 11)
				{
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de aplicar un CUPON \n\nEste proceso no puede ser revertido y no podra cancelar el proceso.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
						{
							#region Variables
							string sCodAnticipo = "";
							string cCodAnt = "";
							int iArtHospedaje = 0;
							int iAnticipoMP = 0;
							int iOrigenAnticipo = 2;
							decimal dSubtotalHospedaje = 0.00m;
							decimal dSubTotalHosDesc = 0.00m;
							#endregion Variables

							#region Valor Hospedaje Descuento
							dSubTotalHosDesc = Math.Round(Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select PrecioNormal From Habitaciones Where idHabitacion = '{0}'", (int)this.cmbHabitaciones.Value)), 2) / 2;
							#endregion Valor Hospedaje Descuento

							iArtHospedaje = (int)dr.Cells["idArticulo"].Value;
							if (iArtHospedaje == 3422)
							{
								using (VentaFPagoLot miCobro = new VentaFPagoLot (0, 1, dSubTotalHosDesc, 13, 4))
								{
									if (DialogResult.OK == miCobro.ShowDialog())
									{
										using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))					
										{
											oConexion.Open();

											SqlTransaction oTransaction = null;
											SqlCommand oCmdActualiza = oConexion.CreateCommand();
											oCmdActualiza.Connection = oConexion;
											oCmdActualiza.CommandType = CommandType.Text;
											oCmdActualiza.CommandTimeout = 120;
						
											try
											{							
												oTransaction = oConexion.BeginTransaction();
												oCmdActualiza.Transaction = oTransaction;

												#region Numeracion Anticipo
												string sSQLNumero = string.Format("Exec NumeracionLocales 24, {0}, 0", (int)this.cmbBodega.Value);
												oCmdActualiza.CommandText = sSQLNumero;
												cCodAnt = oCmdActualiza.ExecuteScalar().ToString();
												#endregion NUmeracion Anticipo

												#region Numeracion Cupon
												string sSQLNumeroAnt = string.Format("Exec AnticipoGeneraCodigo '{0}'", this.txtRuc.Text.ToString());
												oCmdActualiza.CommandText = sSQLNumeroAnt;
												sCodAnticipo = oCmdActualiza.ExecuteScalar().ToString();
												#endregion Numeracion Cupon

												#region Guarda Anticipo
												DateTime dtFecha = (DateTime) this.dtFecha.Value;

												string sSQLGraba = string.Format("Exec GuardaAnticipos {0}, {1}, '{2}', {3}, '{4}', {5}, {6}, {7}, '{8}', {9}, {10}, {11}, '{12}', {13}, '{14}', '{15}', {16}", 
													0, 0, dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, cCodAnt, idCliente, 1, 0, 
													"CUPON MITAD DE PRECIO", 3, Convert.ToDecimal(dSubTotalHosDesc), 13, "",
													0, iArtHospedaje, dtFecha.ToString("yyyyMMdd"), (int)this.optTurnoSalida.Value);
												oCmdActualiza.CommandText = sSQLGraba;
												iAnticipoMP = (int)oCmdActualiza.ExecuteScalar();
												#endregion Guarda Anticipo

												#region Guarda Cobro
												if (miCobro.ultraGrid1.Rows.Count > 0)
												{
													foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in miCobro.ultraGrid1.Rows.All)
													{
														int idDetFormaPago = 0;
														int idBanco = 0;
														int idTarjeta = 0;
														int idCuenta = 0;
														int idPlan = 0;
														int idPlazo = 0;
														decimal dComTarjeta = 0.00m;
														DateTime dtFecCaducidad = DateTime.Today;
							
														if (drr.Cells["idDetFormaPago"].Value != System.DBNull.Value) idDetFormaPago = (int)drr.Cells["idDetFormaPago"].Value;
														if (drr.Cells["idBanco"].Value != System.DBNull.Value) idBanco = (int)drr.Cells["idBanco"].Value;
														if (drr.Cells["idTarjeta"].Value != System.DBNull.Value) idTarjeta = (int)drr.Cells["idTarjeta"].Value;
														if (drr.Cells["idCuenta"].Value != System.DBNull.Value) idCuenta = (int)drr.Cells["idCuenta"].Value;
														if (drr.Cells["idPlan"].Value != System.DBNull.Value) idPlan = (int)drr.Cells["idPlan"].Value;							
														if (drr.Cells["idPlazo"].Value != System.DBNull.Value) idPlazo = (int)drr.Cells["idPlazo"].Value;
														if ((int)drr.Cells["idFormaPago"].Value == 3 || (int)drr.Cells["idFormaPago"].Value == 11) if (drr.Cells["FechaCaducidad"].Value != System.DBNull.Value) dtFecCaducidad = (DateTime)drr.Cells["FechaCaducidad"].Value;
														if (drr.Cells["ValorUsoTarjeta"].Value != System.DBNull.Value) dComTarjeta = (decimal )drr.Cells["ValorUsoTarjeta"].Value;

														string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, 0", 
															idDetFormaPago, 
															iAnticipoMP, iOrigenAnticipo,
															(int)drr.Cells["idFormaPago"].Value, (decimal )drr.Cells["Valor"].Value,
															idBanco, idTarjeta,
															drr.Cells["CuentaCorriente"].Value.ToString(), drr.Cells["Numero"].Value.ToString(),
															dtFecCaducidad.ToString("yyyyMMdd"),								
															drr.Cells["NumVoucher"].Value.ToString(), drr.Cells["Autorizacion"].Value.ToString(),
															idCuenta, 
															drr.Cells["Lote"].Value.ToString(), drr.Cells["Referencia"].Value.ToString(),
															idPlan, idPlazo, dComTarjeta);
														oCmdActualiza.CommandText = sSQLCobro;
														oCmdActualiza.ExecuteNonQuery();
													}
												}
												#endregion Guarda Cobro

												#region Saldo
												string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 24, 1", iAnticipoMP);
												oCmdActualiza.CommandText = sSQLSaldo;
												oCmdActualiza.ExecuteNonQuery();
												#endregion Saldo
												
												#region Historial Anticipos
												string sSQLHistorial = string.Format("Exec HistorialAnticiposMitadPrecio '{0}', {1}, {2}, {3}, {4}, {5}, '{6}', {7}, {8}, {9}", 
													sCodAnticipo, idCompra, (int)this.cmbBodega.Value, idCliente, iAnticipoMP, (int)this.cmbHabitaciones.Value,dtFecha.ToString("yyyyMMdd"),
													(int)this.optTurnoSalida.Value, dSubtotalHospedaje, Convert.ToDecimal(dSubTotalHosDesc));
												oCmdActualiza.CommandText = sSQLHistorial;
												oCmdActualiza.ExecuteNonQuery();
												#endregion Historial Anticipos

												oTransaction.Commit();
							
												oConexion.Close();
												
												#region Controles
												this.btnMitadPrecio.Enabled = false;
												#endregion Controles

												#region Impresion
												string sImpresora = "";
												sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",(int) this.cmbBodega.Value));		
			              
												impresionCuponesDesign(sImpresora, iAnticipoMP);
												impresionCuponesDesign2(sImpresora, iAnticipoMP);
												#endregion Impresion
												
											}
											catch (Exception ex)
											{
												oTransaction.Rollback();

												MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
											}
											finally
											{
												oConexion.Close();
											}
										}
									}
								}
							}
						}						
					}
				}
				else
				{
					MessageBox.Show("Cupon disponible solo para clientes con DATOS.", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtRuc.Focus();
					return;
				}
			}
		}

		private void btnExtras_Click(object sender, System.EventArgs e)
		{
			#region Variables Validacion
			int RetiCompra = 0;
			int iValMinutos = 0;
			int EstadoCaja = 0;
			int iValHospedaje = 0;
			#endregion Variables Validacion

			RetiCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From Compra Where idcompra = {0}", idCompra));

			if (RetiCompra == 1)
			{
				iValMinutos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaMinutosFactura {0}", idCompra));

				iValHospedaje = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From DetCompra Where idArticulo = 3422 And idCompra = {0}", idCompra));				
				
				EstadoCaja = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec EstadoCajaRefacturacion '{0}', {1}, {2}", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm"), (int)this.optTurnoSalida.Value, (int)this.cmbBodega.Value));

				if (iValMinutos > 60)
				{
					MessageBox.Show("Pasado la Hora no puede Facturar Extras", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (iValHospedaje == 0)
				{
					MessageBox.Show("No puede Facturar Extras, La factura no cuenta con Servicio de Hospedaje", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (EstadoCaja == 2)
				{
					MessageBox.Show("No puede Facturar Extras si la CAJA ya esta Cerrada", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (bAnulado)
				{
					MessageBox.Show("El Documento esta Anulado", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}				

				#region Factura Extra
				if (DialogResult.Yes == MessageBox.Show("Esta seguro de Crear una Factura Extra. \n\nEste proceso no puede ser revertido y no podra cancelar la acción.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{					
					int iTiempoEstadia = 0;					

					this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFPagoCarga 5");
					vValFPago = 5;

					FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
					FuncionesProcedimientos.EstadoGrids(true, this.uGridArticulos);

					string sSQL = string.Format("Exec ConsultaCompraIndividualFacturaExtra {0}", idCompra);
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
					dr.Read();
					if (dr.HasRows)
					{
						iCompraFacturaExtra = dr.GetInt32(0);
						this.cmbBodega.Value = dr.GetInt32(1);
						this.cmbHabitaciones.Value = dr.GetInt32(2);
						this.dtFecha.Value = dr.GetDateTime(3);
						if (dr.GetValue(4) != System.DBNull.Value) this.txtEncuesta.Text = dr.GetString(4);
						if (dr.GetValue(5) != System.DBNull.Value) this.txtEncuesta2.Text = dr.GetString(5);
						this.txtPersonas.Value = dr.GetInt32(6);
						this.cmbFormaPago.Value = dr.GetInt32(7);
						this.cmbTipoRuc.Value = dr.GetInt32(8);
						this.txtRuc.Text = dr.GetString(9);
						this.txtNombre.Text = dr.GetString(10);
						idCliente = dr.GetInt32(11);
						this.dtFechaIngreso.Value = dr.GetDateTime(12);
						this.dtFechaIngreso.Value = dr.GetDateTime(13);
						this.optTurnoIngreso.Value = dr.GetInt32(14);
						this.optTurnoIngreso.Value = dr.GetInt32(15);
						this.txtUnidad.Value = dr.GetInt32(16);
					}
					dr.Close();

					this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleFacturas {0}", 0));

					dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
					dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

					iTiempoEstadia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaPromocionTiempoEstadia({0}, {1}, {2}, {3})", 
						(int)this.cmbBodega.Value, 3422, 2, (int)this.cmbHabitaciones.Value)); 					

					if (iTiempoEstadia == 0)
					{
						iTiempoEstadia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec FacturacionRetornaHoras {0}", (int)this.cmbHabitaciones.Value));

						if (iTiempoEstadia == 0) iTiempoEstadia = 5;
					}
					this.txtTiempo.MaxValue = iTiempoEstadia;
					this.txtTiempo.Value = iTiempoEstadia;

					#region Saldo Anticipos
					this.ValidaAnticiposClientes();
					#endregion Saldo Anticipos

					#region Valores
					idReserva = iCompraFacturaExtra;
					this.txtNumero.Text = "";
					this.txtNotas.Text = "";
					this.txtIva.Value = 0.00m;
					this.txtDescIvaTotal.Value = 0.00m;
					this.txtIvaTotal.Value = 0.00m;
					this.txtTotal.Value = 0.00m;
					#endregion Valores

					#region Control
					bEdicion = true;
					bNuevo = true;
					bFacturaExtra = true;

					this.btnGuardar.Enabled = true;					
					this.btnVer.Enabled = true;
					
					//this.cmbGuardia.Enabled = true;
					//this.cmbCobrador.Enabled = true;
					this.cmbFormaPago.Enabled = true;
					//this.cmbCajero.Enabled = true;
					this.cmbTipoRuc.Enabled = true;
					this.txtRuc.Enabled = true;
					this.txtBuscar.Enabled = true;
					this.txtNotas.Enabled = true;
					this.btnanticiposPendientes.Enabled = true;
					this.btnOrdenCocina.Enabled = true;

					this.btnCancelar.Enabled = false;
					this.btnNuevo.Enabled = false;
					this.btnEditar.Enabled = false;
					this.btnConsultar.Enabled = false;
					this.btnImprimir.Enabled = false;
					this.btnTemporal.Enabled = false;
					this.btnAsiento.Enabled = false;
					this.btnAnular.Enabled = false;
					this.btnSalir.Enabled = false;
					this.btnRefacturacion.Enabled = false;
					this.btnMitadPrecio.Enabled = false;
					this.btnDescEmpAut.Enabled = false;
					this.btnClaveDescuento.Enabled = false;
					this.btnCliente.Enabled = false;
					this.btnExtras.Enabled = false;

					this.cmbBodega.Enabled = false;
					this.cmbHabitaciones.Enabled = false;
					this.dtFecha.Enabled = false;
					this.txtEncuesta.Enabled = false;
					this.txtEncuesta2.Enabled = false;
					this.txtPersonas.Enabled = false;
					this.txtUnidad.Enabled = false;
					this.txtTiempo.Enabled = false;
					this.dtFechaIngreso.Enabled = false;
					this.dtFechaSalida.Enabled = false;
					this.optTurnoIngreso.Enabled = false;
					this.optTurnoSalida.Enabled = false;

					this.txtTiempoAdicional.Enabled = false;

					bFacturaTiempo = false;
					#endregion Control
				}
				#endregion Factura Extra
			}
		}

		private void btnCupon_Click_1(object sender, System.EventArgs e)
		{
			#region Variables
			int iValCuponCreado = 0;
			int iValNomCliente = 0;
			int iValIdentCliente = 0;
			int iValCodCliente = 0;
			#endregion Variables

			iValCuponCreado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesEmpresas Where idCompraEmision = {0}", idCompra));

			if (iValCuponCreado == 1)
			{
				MessageBox.Show("La factura Actual ya registra un cupón", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (idCompra == 0)
			{
				MessageBox.Show("No se puede generar un cupon, La Factura debe estar en TEMPORAL", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.cmbTipoRuc.Value == 11)
			{
				MessageBox.Show("No se puede generar cupones a consumidor final.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipoRuc.Focus();
				return;
			}

			if (this.txtRuc.Text.Length == 0)
			{
				MessageBox.Show("No se puede generar un cupon, identificacion esta vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}

			if(this.cmbTipoRuc.Value == System.DBNull.Value || (int)this.cmbTipoRuc.Value == 0)
			{
				MessageBox.Show("No se puede generar un cupon, tipo de identificacion vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipoRuc.Focus();
				return;
			}

			if (this.txtNombre.Text.Length == 0)
			{
				MessageBox.Show("No se puede generar un cupon, Nombre del cliente esta vacio.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}

			//			if((int)this.cmbBodega.Value == 2 || (int)this.cmbBodega.Value == 4)
			//			{
			//				MessageBox.Show("Solo se puede generar cupones desde el establecimiento TANTRA LOFT.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				return;
			//			}
			#region Validacion Datos Ingresados

			#region Validacion
			iValNomCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesEmpresas Where NobreCliente = '{0}' And Estado = 0", this.txtNombre.Text.ToString()));
			iValIdentCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesEmpresas Where IdentificacionCliente = '{0}' And Estado = 0", this.txtRuc.Text.ToString()));
			iValCodCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesEmpresas Where Codigo = '{0}'", this.txtCupon.Text.ToString()));

			if(iValNomCliente >= 1)
			{
				MessageBox.Show("El nombre del Cliente registra un cupon pendiente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if(iValIdentCliente == 1)
			{
				MessageBox.Show("La identificacion Cliente registra un cupon pendiente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if(iValCodCliente == 1)
			{
				MessageBox.Show("El codigo del cupon del cupón ya existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion		

			#endregion Validacion Datos Ingresados
			
			if(iActivoCuponBodega == 1)
			{
				if((int)this.cmbTipoRuc.Value != 11)
				{
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de generar un CUPON \n\nEste proceso no puede ser revertido y no podra cancelar el proceso.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
						{
							#region Variables
							string sCodCupon = "";
							int iArtHospedaje = 0;
							int iCuponOE = 0;
							decimal dSubtotalHospedaje = 0.00m;
							decimal dSubTotalHosRound = 0.00m;
							#endregion Variables

							#region Valor Hospedaje Descuento
							dSubTotalHosRound = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select PrecioNormal From Habitaciones Where idHabitacion = '{0}'", (int)this.cmbHabitaciones.Value));
							#endregion Valor Hospedaje Descuento

							iArtHospedaje = (int)dr.Cells["idArticulo"].Value;

							if (iArtHospedaje == 3422)
							{
								using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))					
								{
									oConexion.Open();

									SqlTransaction oTransaction = null;
									SqlCommand oCmdActualiza = oConexion.CreateCommand();
									oCmdActualiza.Connection = oConexion;
									oCmdActualiza.CommandType = CommandType.Text;
									oCmdActualiza.CommandTimeout = 120;
						
									try
									{							
										oTransaction = oConexion.BeginTransaction();
										oCmdActualiza.Transaction = oTransaction;

										#region Numeracion Cupon
										string sSQLNumeroAnt = string.Format("Exec CuponesEmpresasGeneraCodigo '{0}', {1}", this.txtRuc.Text.ToString(), (int)this.cmbBodega.Value);
										oCmdActualiza.CommandText = sSQLNumeroAnt;
										sCodCupon = oCmdActualiza.ExecuteScalar().ToString();
										#endregion Numeracion Cupon

										#region Historial Cupones
										string sSQLHistorial = string.Format("Exec InsertaCuponesEmpresas {0}, '{1}', {2}, {3}, {4}, {5}, {6}", 
											idCompra, sCodCupon, idCliente, (int)this.optTurnoSalida.Value, (int)this.cmbBodega.Value, (int)this.cmbHabitaciones.Value, dSubTotalHosRound);
										oCmdActualiza.CommandText = sSQLHistorial;
										iCuponOE = (int)oCmdActualiza.ExecuteScalar();
										#endregion Historial Cupones

										oTransaction.Commit();
							
										oConexion.Close();
												
										#region Controles
										this.btnCupon.Enabled = false;
										#endregion Controles

										#region Impresion
										string sImpresora = "";
										sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",(int) this.cmbBodega.Value));		
			              
										string sImpresoraCocina = "";
										sImpresoraCocina = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 2 ",(int) this.cmbBodega.Value));		
			
										//string sImpresora = "SLK-TS100"; //Cambiar el nombre de la impresora   "SLK-TS100"  "RICOH Aficio MP 3350 PCL 6"
										impresionCuponesDesign(sImpresora, iCuponOE);
										impresionCuponesDesign2(sImpresora, iCuponOE);										
										#endregion Impresion												
									}
									catch (Exception ex)
									{
										oTransaction.Rollback();

										MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
									}
									finally
									{
										oConexion.Close();
									}
								}
							} 
						}
					}
				}
			}
		}

		private void impresionCuponesPlatinum(string sImpresora, int iCuponEmpresas)
		{
			#region Impresion Cupon Platinum
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idCuponesEmpresas";
			discreteValIdDetCompra.Value = iCuponEmpresas; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("CuponesEntreEmpresaPlatinum.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Cupones");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula) Reporte.ShowDialog();
			if (!bMenuCuadricula) Reporte.Show();
			#endregion Impresion Cupon Platinum
		}

		private void impresionCuponesPlatinum2(string sImpresora, int iCuponEmpresas)
		{
			#region Impresion Cupon Platinum
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idCuponesEmpresas";
			discreteValIdDetCompra.Value = iCuponEmpresas; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("CuponesEntreEmpresaPlatinum.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Cupones");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula) Reporte.ShowDialog();
			if (!bMenuCuadricula) Reporte.Show();
			#endregion Impresion Cupon Platinum
		}

		private void impresionCuponesDesign(string sImpresora, int iCuponEmpresas)
		{
			#region Impresion Cupon Design
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idCuponesEmpresas";
			discreteValIdDetCompra.Value = iCuponEmpresas; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("CuponesEntreEmpresaDesign.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Cupones");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula) Reporte.ShowDialog();
			if (!bMenuCuadricula) Reporte.Show();
			#endregion Impresion Cupon Design
		}

		private void impresionCuponesDesign2(string sImpresora, int iCuponEmpresas)
		{
			#region Impresion Cupon Design
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idCuponesEmpresas";
			discreteValIdDetCompra.Value = iCuponEmpresas; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("CuponesEntreEmpresaDesign.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Cupones");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			if (bMenuCuadricula) Reporte.ShowDialog();
			if (!bMenuCuadricula) Reporte.Show();
			#endregion Impresion Cupon Design
		}

		private void uGridDetalle_BeforeRowUpdate(object sender, Infragistics.Win.UltraWinGrid.CancelableRowEventArgs e)
		{
		
		}

		private void chkCupon_CheckedChanged(object sender, System.EventArgs e)
		{
			if ((bool)this.chkCupon.Checked)
			{			
				#region Validacion Dia Cupones
				int iValDiaCupon = 0;

				iValDiaCupon = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDiaCuponesEmpresas"));						

				if(iValDiaCupon == 0)
				{
					MessageBox.Show("El dia actual no acepta CUPONES", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCupon.Enabled = false;
					this.chkCupon.Checked = false;					
					this.txtCupon.Text = "";
					this.txtCupon.Focus();
					return;
				} 					
				#endregion Validacion Dia Cupones

				this.txtCupon.Enabled = true;
			}
			else 
			{
				this.txtCupon.Enabled = false;
				this.txtCupon.Text = "";
			}
		}

		private void txtCupon_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if(idCompra > 0)
				{
					if(this.txtCupon.TextLength > 0)
					{
						this.ValidaCodigoCupon();
					}
				}
			}
		}

		private void txtCupon_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				this.ValidaCodigoCupon();
			}
		}

		private void ValidaCodigoCupon()
		{
			#region Variables
			int iValCliente = 0;
			int iValTiempo = 0;
			int iValDiaCupon = 0;
			#endregion Variable	
			
			iValDiaCupon = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDiaCupon"));						

			if(iValDiaCupon == 0)
			{
				MessageBox.Show("El dia actual no acepta CUPONES", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
								
			#region Validacion Credenciales Cliente
					
			if (this.txtRuc.Text.Length == 0)
			{
				MessageBox.Show("No se puede validar el cupon, identificacion esta vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCupon.Text = "";
				this.txtRuc.Focus();
				return;
			}

			if(this.cmbTipoRuc.Value == System.DBNull.Value || (int)this.cmbTipoRuc.Value == 0)
			{
				MessageBox.Show("No se puede validar el cupon, tipo de identificacion vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCupon.Text = "";
				this.cmbTipoRuc.Focus();
				return;
			}

			if (this.txtNombre.Text.Length == 0)
			{
				MessageBox.Show("No se puede validar el cupon, Nombre del cliente esta vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCupon.Text = "";
				this.txtRuc.Focus();
				return;
			}
			#endregion Validacion Credenciales Cliente

			#region Validacion
			//				iValCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaCuponClientePlatinum '{0}', '{1}', {2}", this.txtRuc.Text.ToString(), this.txtCupon.Text.ToString(), (int)this.cmbHabitaciones.Value));
			//				iValTiempo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaTiempoCuponesPlatinum '{0}', '{1}'", this.txtRuc.Text.ToString(), this.txtCupon.Text.ToString()));
			//
			//				if(iValTiempo > 30) //Dias en Vencer
			//				{
			//					MessageBox.Show("El tiempo del cupon a Expirado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtCupon.Focus();
			//					return;
			//				}
			//				
			//				if(iValCliente == 0)
			//				{
			//					MessageBox.Show("No se puede validar el cupon, revisar la Cedula, el Nombre, el valor de la habitacion o si el cupon ya se utilizo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtCupon.Text = "";
			//					this.txtCupon.Focus();
			//					return;
			//				}
			#endregion Validacion

			#region Proceso
			if (bEdicion)
			{
				if (DialogResult.Yes == MessageBox.Show("Esta seguro de Aplicar el Cupon a la factura?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					{
						int iArtHospedaje = 0;
						iArtHospedaje = (int)dr.Cells["idArticulo"].Value;

						if (iArtHospedaje == 3422)
						{
							decimal dDescuento = 50;
							dr.Cells["DescuentoPorc"].Value = dDescuento;
							this.Total();
							this.txtCupon.Enabled = false;
							this.chkCupon.Enabled = false;
						}
						else
						{
							dr.Cells["DescuentoPorc"].Value = 0.00;
							this.Total();
						}
					}
				}									
			}
			#endregion Proceso
		}

		private void txtCupon_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if(idCompra > 0)
				{
					if(this.txtCupon.TextLength > 0)
					{
						this.ValidaCodigoCupon();
					}
				}
			}
		}

		private void txtCupon_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void optTurnoSalida_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnTiempoAdicional_Click(object sender, System.EventArgs e)
		{			
			if (bMenuCuadricula)
			{
				if (bEdicion)
				{
					if ((int)this.txtTiempoAdicional.Value > 0)
					{
						if((int)txtTiempoAdicional.Value < ValTiempoAdicionalPromo)
						{
							MessageBox.Show("El tiempo Servicio Adicional Horas no puede tener menor valor al actual.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtTiempoAdicional.Value = ValTiempoAdicionalPromo;
							return;
						}

						if (iIngresos > 1) return;

						if ((int)this.txtTiempoAdicional.Value > (int)this.txtTiempo.Value - 1) return;

						int idArticuloHA = 4539;
						
						int iCantidad = 0;

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDet in this.uGridDetalle.Rows.All)
							if (drDet.Cells["idArticulo"].Value.ToString().Length > 0)
								if ((int)drDet.Cells["idArticulo"].Value == idArticuloHA) iCantidad = iCantidad + (int)drDet.Cells["Cantidad"].Value;
					
						if (iCantidad == 0)
						{
							iTipoModificacion = 1;

							CargaHabitacionTiempoAdicional(idArticuloHA);							
						}
						else
						{
							iTipoModificacion = 2;

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)							
								if((int)dr.Cells["idArticulo"].Value == idArticuloHA)								
									dr.Cells["Cantidad"].Value = (int)this.txtTiempoAdicional.Value;
						}
					}
				}
			}
		}

		private void CargaArticuloHabitacion(int iModoActualizacion)
		{
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ModificaHoraFinalizaCuadricula {0}, {1}, {2}, {3}, {4}", (int)this.cmbBodega.Value, (int)this.cmbHabitaciones.Value,
				idCompra, (int)this.txtTiempoAdicional.Value, iTipoModificacion));
		}

		private void txtTiempoAdicional_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void btnanticiposPendientes_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.txtRuc.TextLength == 0)
			{
				MessageBox.Show("Ingrese Datos del Cliente", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}

			if((int)this.cmbTipoRuc.Value == 11)
			{
				MessageBox.Show("No valido a CONSUMIDOR FINAL", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}

			if(idCliente == 0)
			{
				MessageBox.Show("Cliente no ingresado.", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
			#endregion Validacion

			using (frmAnticipoAbonoCliente AnticiposPendientes = new frmAnticipoAbonoCliente(idCliente, this.txtNombre.Text))
			{				
				if (DialogResult.OK == AnticiposPendientes.ShowDialog())
				{
				}
			}
		}

		private void InsertaComanda(int iArticulo, int BtnOrden, int OrPromocion, decimal OrRetPrecio, decimal OrRetImpuesto, decimal OrRetDescuentoArt, decimal OrRetDescuentoPorc, decimal OrRetSubTotal, int OrRetCantidad, string sNumComanda)
		{
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;
						
				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Comanda Promocion
					if(OrPromocion == 1)
					{
						string sSQL = string.Format("Exec InsertaComandas {0}, {1}, {2}, {3}", idCompra, (int)this.cmbBodega.Value, iArticulo, (int)this.cmbHabitaciones.Value);
						oCmdActualiza.CommandText = sSQL;
						idComanda = (int)oCmdActualiza.ExecuteScalar();	
					}					
					#endregion Comanda Promocion

					#region Sin Promocion
					if(OrPromocion == 2)
					{
						string sSQL = string.Format("Exec InsertaComandasSinPromocion {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, '{10}'", 
							idCompra, (int)this.cmbBodega.Value, iArticulo, (int)this.cmbHabitaciones.Value, OrRetPrecio, OrRetImpuesto, OrRetDescuentoArt, OrRetDescuentoPorc, OrRetSubTotal, OrRetCantidad, sNumComanda);
						oCmdActualiza.CommandText = sSQL;
						idComanda = (int)oCmdActualiza.ExecuteScalar();	
					}		
					#endregion Sin Promocion

					#region Boton Orden Comanda
					if(BtnOrden == 1)
					{
						string sSQLBtn = string.Format("Exec InsertaComandaOrdenBoton {0}, {1}", idComanda, idCompra);
						oCmdActualiza.CommandText = sSQLBtn;
						idComandaOrden = (int)oCmdActualiza.ExecuteScalar();
					}
					#endregion Boton Orden Comanda

					iComandaImp = idComanda;

					oTransaction.Commit();

					#region Controles
					idComanda = 0;
					iValComandaExis = 0;
					iValComandaMins = 0;
					#endregion Controles
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				}
				finally
				{
					oConexion.Close();
				}
			}
		}

		private void ValidaTemporalArticulo(int iArticulo)
		{
			iExisComanda = Funcion.iEscalarSQL(cdsSeteo, string.Format("Exec ExistenciaComanda {0}", iArticulo));
		}

		private void btnOrdenCocina_Click(object sender, System.EventArgs e)
		{
			#region Peticion Cuadricula
			if(bMenuCuadricula)
			{
				using (frmOrdeneCocina OrdCocina = new frmOrdeneCocina())
				{
					if (DialogResult.OK == OrdCocina.ShowDialog())
					{
						if (OrdCocina.uGridDetalle.Rows.Count > 0)
						{
							int iFila = uGridDetalle.Rows.Count - 1;

							foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in OrdCocina.uGridDetalle.Rows.All)
							{
								int iRetArticulo = (int)dr.Cells["idArticulo"].Value;
								string sRetCodArticulo = dr.Cells["Codigo"].Value.ToString();
								string sRetArticulo = dr.Cells["Articulo"].Value.ToString();
								int iCantidad = (int)dr.Cells["Cantidad"].Value;

								decimal dRetPrecio = Convert.ToDecimal(dr.Cells["Precio"].Value);
								decimal dRetImpuesto = Convert.ToDecimal(dr.Cells["Impuesto"].Value);
								decimal dRetDescuentoPorc = 0.00m;
								decimal dRetDescuentoArt = ((dRetPrecio / Funcion.decEscalarSQL(cdsSeteoF,"Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)")) * dRetDescuentoPorc) / 100;

								decimal dRetSubTotal = Convert.ToDecimal(dr.Cells["Subtotal"].Value);
								#region agregafila
								iFila++;
								uGridDetalle.DisplayLayout.Bands[0].AddNew();

								if (bFacturaExtra) uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;

								uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = iRetArticulo;
								uGridDetalle.Rows[iFila].Cells["Codigo"].Value = sRetCodArticulo;
								uGridDetalle.Rows[iFila].Cells["Articulo"].Value = sRetArticulo;
								uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = iCantidad;
								uGridDetalle.Rows[iFila].Cells["Precio"].Value = dRetPrecio;
								uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = dRetImpuesto;
								uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = dRetDescuentoArt;
								uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = dRetDescuentoPorc;
								uGridDetalle.Rows[iFila].Cells["SubTotal"].Value = dRetSubTotal;

								uGridDetalle.Rows[iFila].Cells["Principal"].Value = 0;
								uGridDetalle.Rows[iFila].Cells["Tipo"].Value = 10;  
								uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = false;
								uGridDetalle.Rows[iFila].Cells["idCombo"].Value = 0;
								uGridDetalle.Rows[iFila].Cells["Entregado"].Value = false;


								#endregion agregafila
							}
							Total();
						}
					}
				}
			}
			#endregion Peticion Cuadricula				

			#region Peticion Facturacion
			if(!bMenuCuadricula)
			{
				#region Validaciones
				if (this.txtRuc.Text.Length == 0)
				{
					MessageBox.Show("No se puede generar una Comanda, identificacion esta vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					return;
				}

				if(this.cmbTipoRuc.Value == System.DBNull.Value || (int)this.cmbTipoRuc.Value == 0)
				{
					MessageBox.Show("No se puede generar una Comanda, tipo de identificacion vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTipoRuc.Focus();
					return;
				}

				if (this.txtNombre.Text.Length == 0)
				{
					MessageBox.Show("No se puede generar una Comanda, Nombre del cliente esta vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					return;
				}

				if (!Validacion.vbComboVacio(this.cmbBodega, "No se puede generar una Comanda, seleccione un local")) return;
				if (!Validacion.vbComboVacio(this.cmbHabitaciones, "No se puede generar una Comanda, seleccione una habitación")) return;
				if (!Validacion.vbComboVacio(this.cmbGuardia, "No se puede generar una Comanda, seleccione un Guardia")) return;
				if (!Validacion.vbComboVacio(this.cmbTipoRuc, "No se puede generar una Comanda, seleccione el tipo de identificación")) return;									
				if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;			
				if (!Validacion.vbComboVacio(this.cmbFormaPago, "No se puede generar una Comanda, seleccione la forma de pago")) return;
				if (!Validacion.vbComboVacio(this.cmbCobrador, "No se puede generar una Comanda, seleccione el cobrador")) return;
				if (!Validacion.vbComboVacio(this.cmbCajero, "No se puede generar una Comanda, seleccione el cajero")) return;

				#endregion Validaciones

				using (frmOrdeneCocina OrdCocina = new frmOrdeneCocina())
				{
					if (DialogResult.OK == OrdCocina.ShowDialog())
					{
						if (OrdCocina.uGridDetalle.Rows.Count > 0)
						{								
							int iFila = uGridDetalle.Rows.Count - 1;

							if (DialogResult.Yes == MessageBox.Show("Esta seguro Ordenar la Comanda.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
							{
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in OrdCocina.uGridDetalle.Rows.All)
								{
									#region Variables
									int iRetArticulo = 0;
									string sRetCodArticulo = "";
									string sRetArticulo = "";
									int iCantidad = 0;
									decimal dRetPrecio = 0.000000m;
									decimal dRetImpuesto = 0.00m;
									decimal dRetDescuentoPorc = 0.00m;
									decimal dRetDescuentoArt = 0.00m;
									decimal dRetSubTotal = 0.00m;
									#endregion Variables

									#region Valores
									iRetArticulo = (int)dr.Cells["idArticulo"].Value;
									sRetCodArticulo = dr.Cells["Codigo"].Value.ToString();
									sRetArticulo = dr.Cells["Articulo"].Value.ToString();
									iCantidad = (int)dr.Cells["Cantidad"].Value;
									dRetPrecio = Convert.ToDecimal(dr.Cells["Precio"].Value);
									dRetImpuesto = Convert.ToDecimal(dr.Cells["Impuesto"].Value);
									dRetDescuentoArt = ((Convert.ToDecimal(dr.Cells["Precio"].Value)/Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)"))) * dRetDescuentoPorc)/100;
									dRetSubTotal = Convert.ToDecimal(dr.Cells["Subtotal"].Value);
									#endregion Valores

									#region Valores Grid
									iFila++;
									uGridDetalle.DisplayLayout.Bands[0].AddNew();
									if(bFacturaExtra) uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;
									uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = iRetArticulo;											
									uGridDetalle.Rows[iFila].Cells["Codigo"].Value = sRetCodArticulo;
									uGridDetalle.Rows[iFila].Cells["Articulo"].Value = sRetArticulo;									
									uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = iCantidad;	
									uGridDetalle.Rows[iFila].Cells["Precio"].Value = dRetPrecio;
									uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = dRetImpuesto;
									uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = dRetDescuentoArt;
									uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = dRetDescuentoPorc;
									uGridDetalle.Rows[iFila].Cells["SubTotal"].Value = dRetSubTotal;
									uGridDetalle.Rows[iFila].Cells["Principal"].Value = 0;
									uGridDetalle.Rows[iFila].Cells["Tipo"].Value = 0;																		
									uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = false;
									uGridDetalle.Rows[iFila].Cells["idCombo"].Value = 0;
									uGridDetalle.Rows[iFila].Cells["Entregado"].Value = false;
									#endregion Valores Grid
							
									Total();
								}
							}
						}
					}
				}
			}
			#endregion Peticion Facturacion
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void chkTarjetaVIP_CheckedChanged(object sender, System.EventArgs e)
		{
			if ((bool)this.chkTarjetaVIP.Checked)
			{			
				this.txtTarjetaVIP.Enabled = true;
			}
			else 
			{
				this.txtTarjetaVIP.Enabled = false;
				this.txtTarjetaVIP.Text = "";
				this.txtPuntosTarjetaVIP.Value = 0;
				this.txtPuntosTarjetaVIP.Visible = false;
				this.lblPuntosTarjetaVIP.Visible = false;
				this.btnCanjePuntos.Enabled = false;
			}
		}

		private void ValidaCodigoTarjetaVIP()
		{
			#region Variables
			int iValCliente = 0;
			int iValTiempo = 0;			
			#endregion Variable	
											
			#region Validacion Credenciales Cliente
			if (this.txtRuc.Text.Length == 0)
			{
				MessageBox.Show("No se puede validar la TARJETA VIP, identificacion esta vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTarjetaVIP.Text = "";
				this.txtRuc.Focus();
				return;
			}

			if(this.cmbTipoRuc.Value == System.DBNull.Value || (int)this.cmbTipoRuc.Value == 0)
			{
				MessageBox.Show("No se puede validar la TARJETA VIP, tipo de identificacion vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTarjetaVIP.Text = "";
				this.cmbTipoRuc.Focus();
				return;
			}

			if (this.txtNombre.Text.Length == 0)
			{
				MessageBox.Show("No se puede validar la TARJETA VIP, Nombre del cliente esta vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTarjetaVIP.Text = "";
				this.txtRuc.Focus();
				return;
			}
			#endregion Validacion Credenciales Cliente

			#region Validacion
			iValCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaTarjetaVIP '{0}', '{1}'", this.txtRuc.Text.ToString(), this.txtTarjetaVIP.Text.ToString()));
			iValTiempo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaTiempoTarjetasVIP '{0}', '{1}'", this.txtRuc.Text.ToString(), this.txtTarjetaVIP.Text.ToString()));

			if(iValTiempo > 365) //Dias en Vencer (1 año)
			{
				MessageBox.Show("El tiempo de la TARJETA VIP a Expirado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCupon.Focus();
				return;
			}
				
			if(iValCliente == 0)
			{
				MessageBox.Show("No se puede validar, revisar la Cedula, el Nombre, el codigo, o la fecha vencimiento de la TARJETA VIP.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTarjetaVIP.Text = "";
				this.txtTarjetaVIP.Focus();
				return;
			}
			#endregion Validacion

			#region Proceso
			if (bEdicion)
			{
				this.txtPuntosTarjetaVIP.Visible = true;
				this.lblPuntosTarjetaVIP.Visible = true;
				this.btnCanjePuntos.Enabled = true;

				iTotalPuntosTarjetaVIP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CalcularPuntosTarjetaVIP '{0}'", this.txtTarjetaVIP.Text.ToString()));
				this.txtPuntosTarjetaVIP.Value = iTotalPuntosTarjetaVIP;
			}
			#endregion Proceso
		}

		private void txtTarjetaVIP_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				this.ValidaCodigoTarjetaVIP();
			}
		}

		private void btnInfoTarjetasVIP_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.txtRuc.TextLength == 0)
			{
				MessageBox.Show("Ingrese Datos del Cliente", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}

			if((int)this.cmbTipoRuc.Value == 11)
			{
				MessageBox.Show("No valido a CONSUMIDOR FINAL", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}

			if(idCliente == 0)
			{
				MessageBox.Show("Cliente no ingresado.", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
			#endregion Validacion

			using (frmConsultaTarjetasVIP ConsultaTajertaVIP = new frmConsultaTarjetasVIP(idCliente, this.txtNombre.Text))
			{				
				if (DialogResult.OK == ConsultaTajertaVIP.ShowDialog())
				{
				}
			}
		}

		private void txtCupon_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private void InsertaPuntosCanjeados(int RACidCompra, int RACidArticulo, string RACCodArticulo, string RACArticulo, int RACCantidad, decimal RACPrecio, decimal RACImpuesto, decimal RACDescuentoArt, decimal RACDescuentoPorc, decimal RACSubTotal, string RACCodTarjeta)
		{
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;
						
				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					int iRegistroTarjetaVIP = 0;

					#region Registra Puntos Tarjeta VIP
					if(chkTarjetaVIP.Checked)
					{
						//						string sSQLConsumoVIP = string.Format("Exec RegistraPuntosConsumoTarjetasVIP {0}, {1}, '{2}', {3}, {4}, {5}",
						//							idCompra, 0, this.txtTarjetaVIP.Text, idCliente, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Cantidad"].Value);
						//						oCmdActualiza.CommandText = sSQLConsumoVIP;
						//						iRegistroTarjetaVIP = (int)oCmdActualiza.ExecuteScalar();
					}
					#endregion Registra Puntos Tarjeta VIP

					#region Detalle
					string sSQLDetalle = string.Format("Exec GuardaDetalleFacturas {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}",
						0, RACidCompra, RACidArticulo, RACCodArticulo, RACArticulo, RACCantidad, //5
						RACPrecio, RACImpuesto, RACDescuentoArt, RACDescuentoPorc, RACSubTotal,//10
						0, true, false, 0, iRegistroTarjetaVIP); //14
					oCmdActualiza.CommandText = sSQLDetalle;
					oCmdActualiza.ExecuteNonQuery();
					#endregion Detalle

					oTransaction.Commit();
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				}
				finally
				{
					oConexion.Close();
				}
			}
		}

		private void btnCanjePuntos_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.txtRuc.TextLength == 0)
			{
				MessageBox.Show("Ingrese Datos del Cliente", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}

			if((int)this.cmbTipoRuc.Value == 11)
			{
				MessageBox.Show("No valido a CONSUMIDOR FINAL", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}

			if(idCliente == 0)
			{
				MessageBox.Show("Cliente no ingresado.", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}

			if(this.txtTarjetaVIP.TextLength == 0)
			{
				MessageBox.Show("Ingrese Codigo de la Tarjeta VIP", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if((int)this.txtPuntosTarjetaVIP.Value <= 0)
			{
				MessageBox.Show("No puede canjear puntos", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			#endregion Validacion

			#region Peticion Cuadricula
			if(bMenuCuadricula)
			{
				if(idCompra == 0)
				{
					MessageBox.Show("Para realizar el canjeo de puntos, la Factura Debe estar en Temporal", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				else
				{
					using (frmCanjeoPuntos CanjePuntos = new frmCanjeoPuntos(idCliente, this.txtNombre.Text, (int)this.txtPuntosTarjetaVIP.Value))
					{
						if (DialogResult.OK == CanjePuntos.ShowDialog())
						{
							int iFila = uGridDetalle.Rows.Count - 1;

							if (DialogResult.Yes == MessageBox.Show("Esta seguro realizar el canjeo de puntos con los producto seleccionados \n\nEl proceso no puede ser reversible.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
							{
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drPuntosCanje in CanjePuntos.uGridPuntosCanje.Rows.All)
								{
									if ((bool)drPuntosCanje.Cells["Sel"].Value)
									{
										#region Variables
										int iRetArticulo = 0;
										string sRetCodArticulo = "";
										string sRetArticulo = "";
										int iCantidad = 0;
										int iTotalPuntos = 0;
										decimal dRetPrecio = 0.000000m;
										decimal dRetImpuesto = 0.00m;
										decimal dRetDescuentoPorc = 0.00m;
										decimal dRetDescuentoArt = 0.00m;
										decimal dRetSubTotal = 0.00m;
										#endregion Variables

										#region Valores
										iRetArticulo = (int)drPuntosCanje.Cells["idArticulo"].Value;
										sRetCodArticulo = drPuntosCanje.Cells["Codigo"].Value.ToString();
										sRetArticulo = drPuntosCanje.Cells["Articulo"].Value.ToString();
										iCantidad = (int)drPuntosCanje.Cells["Cantidad"].Value;
										dRetPrecio = Convert.ToDecimal(drPuntosCanje.Cells["Precio"].Value);
										dRetImpuesto = Convert.ToDecimal(drPuntosCanje.Cells["Impuesto"].Value);
										dRetDescuentoArt = ((Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value)/Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)"))) * dRetDescuentoPorc)/100;
										dRetDescuentoPorc = Convert.ToDecimal(drPuntosCanje.Cells["DescuentoPorc"].Value);
										dRetSubTotal = Convert.ToDecimal(drPuntosCanje.Cells["Subtotal"].Value);
										iTotalPuntos = (int)drPuntosCanje.Cells["TotalPuntos"].Value;
										#endregion Valores

										#region Valores Grid
										iFila++;
										uGridDetalle.DisplayLayout.Bands[0].AddNew();									
										uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = iRetArticulo;											
										uGridDetalle.Rows[iFila].Cells["Codigo"].Value = sRetCodArticulo;
										uGridDetalle.Rows[iFila].Cells["Articulo"].Value = sRetArticulo;									
										uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = iCantidad;	
										uGridDetalle.Rows[iFila].Cells["Precio"].Value = dRetPrecio;
										uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = dRetImpuesto;
										uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = dRetDescuentoArt;
										uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = dRetDescuentoPorc;
										uGridDetalle.Rows[iFila].Cells["SubTotal"].Value = dRetSubTotal;
										uGridDetalle.Rows[iFila].Cells["Principal"].Value = 0;
										uGridDetalle.Rows[iFila].Cells["Tipo"].Value = 0;																		
										uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = false;
										uGridDetalle.Rows[iFila].Cells["idCombo"].Value = 0;
										uGridDetalle.Rows[iFila].Cells["Entregado"].Value = false;
										#endregion Valores Grid

										#region Registra Articulos Canje
										InsertaPuntosCanjeados(idCompra, iRetArticulo, sRetCodArticulo, sRetArticulo, iCantidad, dRetPrecio, dRetImpuesto, dRetDescuentoArt, dRetDescuentoPorc, dRetSubTotal, this.txtTarjetaVIP.Text);
										#endregion Registra Articulos Canje
									}
								}
							}
						}
					}
				}
			}
			#endregion Peticion Cuadricula
		}

		private void chkVideo_CheckedChanged(object sender, System.EventArgs e)
		{
			if ((bool)this.chkVideo.Checked)
			{
				#region Validacion
				if (this.txtRuc.Text.Length == 0)
				{
					MessageBox.Show("identificacion esta vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.chkVideo.Checked = false;
					bPromoVideo = false;
					return;
				}

				if(this.cmbTipoRuc.Value == System.DBNull.Value || (int)this.cmbTipoRuc.Value == 0)
				{
					MessageBox.Show("tipo de identificacion vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.chkVideo.Checked = false;
					bPromoVideo = false;
					return;
				}

				if (this.txtNombre.Text.Length == 0)
				{
					MessageBox.Show("Nombre del cliente esta vacia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.chkVideo.Checked = false;
					bPromoVideo = false;
					return;
				}

				if(!bRFVideoPromocion && bMenuCuadricula)
				{
					int iValClienteVideo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("ConsultaPromocionVideoCliente '{0}'", this.txtRuc.Text.ToString()));
					if(iValClienteVideo > 1)
					{
						MessageBox.Show("El cliente ya registra la promoción del 10%.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.chkVideo.Checked = false;
						bPromoVideo = false;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
						{
							int iArtHospedaje = 0;
							iArtHospedaje = (int)dr.Cells["idArticulo"].Value;

							if (iArtHospedaje == 3422)
							{
								decimal dDescuento = 0;
								dr.Cells["DescuentoPorc"].Value = dDescuento;
								this.Total();
							}
						}
						return;
					}
				}				
				#endregion Validacion

				bPromoVideo = true;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				{
					int iArtHospedaje = 0;
					iArtHospedaje = (int)dr.Cells["idArticulo"].Value;

					if (iArtHospedaje == 3422)
					{
						decimal dDescuento = 10;
						dr.Cells["DescuentoPorc"].Value = dDescuento;
						this.Total();						
					}
					else
					{
						dr.Cells["DescuentoPorc"].Value = 0.00;
						this.Total();
					}
				}
			}
			else
			{
				bPromoVideo = false;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				{
					int iArtHospedaje = 0;
					iArtHospedaje = (int)dr.Cells["idArticulo"].Value;

					if (iArtHospedaje == 3422)
					{
						decimal dDescuento = 0;
						dr.Cells["DescuentoPorc"].Value = dDescuento;
						this.Total();						
					}
					else
					{
						dr.Cells["DescuentoPorc"].Value = 0.00;
						this.Total();
					}
				}
			}
		}

		private void BtnReImprimirComanda_Click(object sender, System.EventArgs e)
		{
			int bodega = (int) this.cmbBodega.Value;
			using (frmReImprimirComandas REIC = new frmReImprimirComandas(idCompra,bodega))
			{
				REIC.ShowDialog();
			}
		}
		private void CargarTiempoEstadiaDesdeHabitacion()
		{
			if (this.cmbHabitaciones.Value != null)
			{
				bool bActivoFact = false;
				int iTiempoModEstadia = 0;
				int iTiempoEstadia = 0;

				bActivoFact = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select top 1 Activo From TiempoModificacionEstadia Where Bodega = {0} Order by FechaSistema desc",
					(int)this.cmbBodega.Value), true);

				iTiempoModEstadia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select top 1 idTiempoModificacionEstadia From TiempoModificacionEstadia Where Bodega = {0} Order by FechaSistema desc",
					(int)this.cmbBodega.Value));

				if (bActivoFact && iTiempoModEstadia > 0)
				{
					iTiempoEstadia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("ValidaTiempoEstadia {0}", iTiempoModEstadia));
				}

				if (iTarifaExpress == 2)
				{
					iTiempoEstadia = 2;
				}

				if (iTiempoEstadia == 0)
				{
					iTiempoEstadia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaPromocionTiempoEstadia({0}, {1}, {2}, {3})",
						(int)this.cmbBodega.Value, 3422, 2, (int)this.cmbHabitaciones.Value));

					if (iTiempoEstadia == 0)
					{
						iTiempoEstadia =
							Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec FacturacionRetornaHoras {0}", (int)this.cmbHabitaciones.Value));

						if (iTiempoEstadia == 0) iTiempoEstadia = 5;
					}
				}

				this.txtTiempo.MaxValue = iTiempoEstadia;
				this.txtTiempo.Value = iTiempoEstadia;
			}
		}

		private void btnReturnPass_Click(object sender, System.EventArgs e)
		{
			string sRUC = "";
			int idTipoRuc = 5;
			if ((int)this.cmbTipoRuc.Value != 11) 
			{
				sRUC = this.txtRuc.Text.ToString();
				idTipoRuc = (int)this.cmbTipoRuc.Value;
			}

			using (frmAnticiposClientes Anticipos = new frmAnticiposClientes(true, sRUC, (int)this.cmbBodega.Value, 
							 19.00m, 5, 0, idTipoRuc,(int)this.cmbFormaPago.Value))
			{
				if (DialogResult.OK == Anticipos.ShowDialog())
				{	
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("UPDATE Anticipos SET idOrigen = {0} WHERE idAnticipo = {1}", idCompra, (int)Anticipos.txtIdAnticipo.Value));

					ImpresionAutomaticaAnticipos((int)Anticipos.txtIdAnticipo.Value, 5, cdsSeteoF);
				}
			}
		}
		private void AbrirAnticipoRegistroCobro()
		{
			if (bAnticipoYaAbierto) return;
			bAnticipoYaAbierto = true;

			if (bServicioPagado) return;

			if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione el tipo de identificacion del cliente")) return;
			if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;

			if (idCliente == 0)
			{
				MessageBox.Show("El Cliente No Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			decimal dTotal = 0.00m;
			if (this.txtTotal.Value != null) dTotal = Convert.ToDecimal(this.txtTotal.Value);

			using (frmAnticiposClientes Anticipos = new frmAnticiposClientes(true,this.txtRuc.Text.Trim(),iBodega,dTotal,15,idCompraV,(int)this.cmbTipoRuc.Value,(int)this.cmbFormaPago.Value))
			{
				if (DialogResult.OK == Anticipos.ShowDialog())
				{
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("UPDATE Anticipos SET idOrigen = {0} WHERE idAnticipo = {1}",idCompraV, (int)Anticipos.txtIdAnticipo.Value));

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("UPDATE Habitaciones SET Ocupado = CASE WHEN Ocupado IN (6) THEN 2 ELSE Ocupado END WHERE idHabitacion = {0}",idHabitacion));
					
					ImpresionAutomaticaAnticipos((int)Anticipos.txtIdAnticipo.Value, 15, cdsSeteoF);
					
					this.btnCancelar_Click(this, EventArgs.Empty);

					this.Close();
				}
				else
				{
					bAnticipoYaAbierto = false;
				}
			}
		}

		private void LimpiarHabitacion()
		{
			for (int i = uGridDetalle.Rows.Count - 1; i >= 0; i--)
			{
				Infragistics.Win.UltraWinGrid.UltraGridRow row = uGridDetalle.Rows[i];

				if (row.Cells["Tipo"].Value != DBNull.Value)
				{
					int tipo = Convert.ToInt32(row.Cells["Tipo"].Value);

					if (tipo == 0 || tipo == 2)
					{
						row.Delete(false);
					}
				}
			}
		}
		private void AbrirVerificacionPago()
		{
			btnGuardar_Click(this, EventArgs.Empty);
		}
	}
}
