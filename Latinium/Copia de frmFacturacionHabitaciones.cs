using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
//using CrystalDecisions;
//using CrystalDecisions.CrystalReports.Engine;
//using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

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
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
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
		private System.Windows.Forms.Button btnIngresos;
		private System.Windows.Forms.Label lblIngresos;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPersonas;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

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

		int idHabitacion = 0;
		int iBodega = 0;
		bool bMenuCuadricula = false;
		int idCompraV = 0;
		bool bServicioPagado = false;

		public frmFacturacionHabitaciones(int IdHabitacion, int IBodega, bool BMenuCuadricula, int IdCompraV, bool BServicioPagado)
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Turno");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFacturacionHabitaciones));
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeleccionar", 0);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio2");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			this.cmbCajero = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.txtTiempo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
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
			this.btnIngresos = new System.Windows.Forms.Button();
			this.lblIngresos = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtPersonas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
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
			this.cmbCajero.Location = new System.Drawing.Point(416, 112);
			this.cmbCajero.Name = "cmbCajero";
			this.cmbCajero.Size = new System.Drawing.Size(240, 22);
			this.cmbCajero.TabIndex = 267;
			this.cmbCajero.ValueMember = "idPersonal";
			this.cmbCajero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCajero_KeyDown);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label19.Location = new System.Drawing.Point(360, 115);
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
			this.txtTiempo.Location = new System.Drawing.Point(784, 47);
			this.txtTiempo.MaskInput = "nnnnnnnnn";
			this.txtTiempo.MaxValue = 8;
			this.txtTiempo.MinValue = 3;
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
			this.label14.Location = new System.Drawing.Point(728, 50);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(41, 17);
			this.label14.TabIndex = 258;
			this.label14.Text = "Tiempo";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource2;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance3;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 353;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance4;
			ultraGridColumn11.Format = "#,##0";
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.PromptChar = ' ';
			ultraGridColumn11.Width = 63;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance5;
			ultraGridColumn12.Format = "#,##0.000000";
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn12.PromptChar = ' ';
			ultraGridColumn12.Width = 98;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance6;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Width = 80;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance7;
			ultraGridColumn14.Format = "#,##0.000000";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance8;
			ultraGridColumn15.Format = "#,##0.000000";
			ultraGridColumn15.Header.Caption = "% Desc";
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridColumn15.Width = 88;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance9;
			ultraGridColumn16.Format = "#,##0.000000";
			ultraGridColumn16.Header.Caption = "Subtotal";
			ultraGridColumn16.Header.VisiblePosition = 11;
			ultraGridColumn16.PromptChar = ' ';
			ultraGridColumn16.Width = 113;
			ultraGridColumn17.Header.VisiblePosition = 12;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 14;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 15;
			ultraGridColumn20.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn20});
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance12;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 9F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance13;
			appearance14.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance14.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.uGridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 184);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(816, 180);
			this.uGridDetalle.TabIndex = 252;
			this.uGridDetalle.AfterRowsDeleted += new System.EventHandler(this.uGridDetalle_AfterRowsDeleted);
			this.uGridDetalle.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDetalle_BeforeCellUpdate);
			this.uGridDetalle.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalle_BeforeRowsDeleted);
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
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
																																 ultraDataColumn16});
			// 
			// btnCliente
			// 
			this.btnCliente.Enabled = false;
			this.btnCliente.Image = ((System.Drawing.Image)(resources.GetObject("btnCliente.Image")));
			this.btnCliente.Location = new System.Drawing.Point(832, 78);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(26, 26);
			this.btnCliente.TabIndex = 250;
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
			// 
			// cmbCobrador
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCobrador.Appearance = appearance17;
			this.cmbCobrador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 335;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbCobrador.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCobrador.DisplayMember = "Nombre";
			this.cmbCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCobrador.Enabled = false;
			this.cmbCobrador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador.Location = new System.Drawing.Point(96, 112);
			this.cmbCobrador.Name = "cmbCobrador";
			this.cmbCobrador.Size = new System.Drawing.Size(240, 22);
			this.cmbCobrador.TabIndex = 265;
			this.cmbCobrador.ValueMember = "idPersonal";
			this.cmbCobrador.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCobrador_KeyDown);
			// 
			// cmbGuardia
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGuardia.Appearance = appearance18;
			this.cmbGuardia.CausesValidation = false;
			this.cmbGuardia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGuardia.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 335;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbGuardia.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbGuardia.DisplayMember = "Nombre";
			this.cmbGuardia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGuardia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGuardia.Enabled = false;
			this.cmbGuardia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGuardia.Location = new System.Drawing.Point(256, 47);
			this.cmbGuardia.Name = "cmbGuardia";
			this.cmbGuardia.Size = new System.Drawing.Size(240, 22);
			this.cmbGuardia.TabIndex = 264;
			this.cmbGuardia.ValueMember = "idPersonal";
			this.cmbGuardia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbGuardia_KeyDown);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.Location = new System.Drawing.Point(8, 115);
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
			this.label17.Location = new System.Drawing.Point(192, 50);
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
			this.label15.Location = new System.Drawing.Point(728, 115);
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
			this.label10.Location = new System.Drawing.Point(8, 50);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(34, 17);
			this.label10.TabIndex = 255;
			this.label10.Text = "Fecha";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFormaPago
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance19;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 152;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(96, 144);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(150, 22);
			this.cmbFormaPago.TabIndex = 253;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbFormaPago_KeyDown);
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
			this.lblHabitacion.Location = new System.Drawing.Point(8, 8);
			this.lblHabitacion.Name = "lblHabitacion";
			this.lblHabitacion.Size = new System.Drawing.Size(816, 30);
			this.lblHabitacion.TabIndex = 249;
			this.lblHabitacion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblHabitacion.Visible = false;
			// 
			// uGridArticulos
			// 
			this.uGridArticulos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridArticulos.DataSource = this.ultraDataSource3;
			appearance20.BackColor = System.Drawing.Color.White;
			this.uGridArticulos.DisplayLayout.Appearance = appearance20;
			this.uGridArticulos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 3;
			ultraGridColumn31.Width = 300;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance21;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.Caption = "Precio";
			ultraGridColumn32.Header.VisiblePosition = 4;
			ultraGridColumn32.Width = 80;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.Caption = "Reserva";
			ultraGridColumn33.Header.VisiblePosition = 5;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 80;
			ultraGridColumn34.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance22.Image = ((object)(resources.GetObject("appearance22.Image")));
			ultraGridColumn34.CellButtonAppearance = appearance22;
			ultraGridColumn34.Header.Caption = "...";
			ultraGridColumn34.Header.VisiblePosition = 2;
			ultraGridColumn34.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn34.Width = 20;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			this.uGridArticulos.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridArticulos.DisplayLayout.Override.ActiveRowAppearance = appearance23;
			this.uGridArticulos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridArticulos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridArticulos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance24.BackColor = System.Drawing.Color.Transparent;
			this.uGridArticulos.DisplayLayout.Override.CardAreaAppearance = appearance24;
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridArticulos.DisplayLayout.Override.CellAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 9F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridArticulos.DisplayLayout.Override.HeaderAppearance = appearance26;
			this.uGridArticulos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance27.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance27.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridArticulos.DisplayLayout.Override.RowAlternateAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridArticulos.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridArticulos.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.uGridArticulos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridArticulos.Location = new System.Drawing.Point(832, 216);
			this.uGridArticulos.Name = "uGridArticulos";
			this.uGridArticulos.Size = new System.Drawing.Size(421, 272);
			this.uGridArticulos.TabIndex = 247;
			this.uGridArticulos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridArticulos_ClickCellButton);
			this.uGridArticulos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridArticulos_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// dtFecha
			// 
			appearance30.FontData.Name = "Tahoma";
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance30;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 48);
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
			this.lblNumero.Location = new System.Drawing.Point(512, 50);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 17);
			this.lblNumero.TabIndex = 269;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance31;
			this.txtNumero.CausesValidation = false;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(576, 47);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 22);
			this.txtNumero.TabIndex = 268;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(512, 83);
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
			this.lblSubTotal.Location = new System.Drawing.Point(560, 368);
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
			this.lblIva.Location = new System.Drawing.Point(560, 416);
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
			this.lblDescuento.Location = new System.Drawing.Point(560, 392);
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
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 83);
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
			this.lblCliente.Location = new System.Drawing.Point(264, 83);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(86, 17);
			this.lblCliente.TabIndex = 422;
			this.lblCliente.Text = "N. Identificación";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance32;
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
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance33;
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
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance34;
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
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance35;
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
			appearance36.FontData.Name = "Tahoma";
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance36;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Width = 233;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn35,
																										 ultraGridColumn36});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(96, 80);
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
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance37;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtRuc.Location = new System.Drawing.Point(368, 80);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(128, 22);
			this.txtRuc.TabIndex = 419;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			// 
			// txtNombre
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance38;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNombre.Location = new System.Drawing.Point(576, 80);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(248, 22);
			this.txtNombre.TabIndex = 420;
			// 
			// txtBuscar
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance39;
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
			this.lblEstado.Location = new System.Drawing.Point(416, 144);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 446;
			// 
			// txtUnidad
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUnidad.Appearance = appearance40;
			this.txtUnidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUnidad.Enabled = false;
			this.txtUnidad.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtUnidad.FormatString = "#0";
			this.txtUnidad.Location = new System.Drawing.Point(784, 112);
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
			this.btnAsiento.Location = new System.Drawing.Point(474, 504);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
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
			this.btnNuevo.Location = new System.Drawing.Point(8, 504);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
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
			this.btnConsultar.Location = new System.Drawing.Point(82, 504);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
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
			this.btnEditar.Location = new System.Drawing.Point(162, 504);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
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
			this.btnImprimir.Location = new System.Drawing.Point(240, 504);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
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
			this.btnGuardar.Location = new System.Drawing.Point(319, 504);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
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
			this.btnAnular.Location = new System.Drawing.Point(556, 504);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
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
			this.btnCancelar.Location = new System.Drawing.Point(634, 504);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
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
			this.btnSalir.Location = new System.Drawing.Point(708, 504);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 457;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Location = new System.Drawing.Point(0, 488);
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
			this.lblLocal.Location = new System.Drawing.Point(8, 15);
			this.lblLocal.Name = "lblLocal";
			this.lblLocal.Size = new System.Drawing.Size(31, 16);
			this.lblLocal.TabIndex = 696;
			this.lblLocal.Text = "Local";
			this.lblLocal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblLocal.Visible = false;
			// 
			// cmbBodega
			// 
			appearance41.FontData.Name = "Tahoma";
			this.cmbBodega.Appearance = appearance41;
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn37.Header.VisiblePosition = 2;
			ultraGridColumn38.Header.VisiblePosition = 1;
			ultraGridColumn38.Width = 180;
			ultraGridColumn39.Header.VisiblePosition = 0;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 12);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(250, 22);
			this.cmbBodega.TabIndex = 695;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.Visible = false;
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// lblHabitaciones
			// 
			this.lblHabitaciones.AutoSize = true;
			this.lblHabitaciones.Location = new System.Drawing.Point(336, 15);
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
			appearance42.FontData.Name = "Tahoma";
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaIngreso.Appearance = appearance42;
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
			appearance43.FontData.Name = "Tahoma";
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaSalida.Appearance = appearance43;
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
			appearance44.ForeColor = System.Drawing.Color.Black;
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTurnoIngreso.Appearance = appearance44;
			this.optTurnoIngreso.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTurnoIngreso.Enabled = false;
			this.optTurnoIngreso.ItemAppearance = appearance45;
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
			appearance46.ForeColor = System.Drawing.Color.Black;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTurnoSalida.Appearance = appearance46;
			this.optTurnoSalida.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTurnoSalida.Enabled = false;
			this.optTurnoSalida.ItemAppearance = appearance47;
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
			// 
			// ultraDataSource1
			// 
			ultraDataColumn22.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22,
																																 ultraDataColumn23});
			// 
			// cmbHabitaciones
			// 
			appearance48.FontData.Name = "Tahoma";
			this.cmbHabitaciones.Appearance = appearance48;
			this.cmbHabitaciones.CausesValidation = false;
			this.cmbHabitaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbHabitaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbHabitaciones.DataSource = this.ultraDataSource1;
			ultraGridColumn40.Header.VisiblePosition = 0;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 1;
			ultraGridColumn41.Width = 250;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn40,
																										 ultraGridColumn41});
			this.cmbHabitaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbHabitaciones.DisplayMember = "Nombre";
			this.cmbHabitaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbHabitaciones.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbHabitaciones.Enabled = false;
			this.cmbHabitaciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbHabitaciones.Location = new System.Drawing.Point(408, 12);
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
			this.btnTemporal.Location = new System.Drawing.Point(397, 504);
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
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance49;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(64, 392);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(400, 66);
			this.txtNotas.TabIndex = 707;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(8, 417);
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
			// btnIngresos
			// 
			this.btnIngresos.Enabled = false;
			this.btnIngresos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnIngresos.Image = ((System.Drawing.Image)(resources.GetObject("btnIngresos.Image")));
			this.btnIngresos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnIngresos.Location = new System.Drawing.Point(64, 464);
			this.btnIngresos.Name = "btnIngresos";
			this.btnIngresos.Size = new System.Drawing.Size(104, 24);
			this.btnIngresos.TabIndex = 710;
			this.btnIngresos.Text = "Nuevo Ingreso";
			this.btnIngresos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnIngresos.Visible = false;
			this.btnIngresos.Click += new System.EventHandler(this.btnIngresos_Click);
			// 
			// lblIngresos
			// 
			this.lblIngresos.AutoSize = true;
			this.lblIngresos.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIngresos.ForeColor = System.Drawing.Color.Firebrick;
			this.lblIngresos.Location = new System.Drawing.Point(272, 144);
			this.lblIngresos.Name = "lblIngresos";
			this.lblIngresos.Size = new System.Drawing.Size(0, 23);
			this.lblIngresos.TabIndex = 711;
			this.lblIngresos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(728, 147);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(49, 17);
			this.label2.TabIndex = 712;
			this.label2.Text = "Personas";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPersonas
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPersonas.Appearance = appearance50;
			this.txtPersonas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPersonas.Enabled = false;
			this.txtPersonas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPersonas.FormatString = "#0";
			this.txtPersonas.Location = new System.Drawing.Point(784, 144);
			this.txtPersonas.MaskInput = "nnnnnnn";
			this.txtPersonas.MaxValue = 4;
			this.txtPersonas.MinValue = 2;
			this.txtPersonas.Name = "txtPersonas";
			this.txtPersonas.PromptChar = ' ';
			this.txtPersonas.Size = new System.Drawing.Size(40, 22);
			this.txtPersonas.TabIndex = 713;
			this.txtPersonas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPersonas_KeyDown);
			this.txtPersonas.Click += new System.EventHandler(this.txtPersonas_Click);
			this.txtPersonas.Validated += new System.EventHandler(this.txtPersonas_Validated);
			this.txtPersonas.Enter += new System.EventHandler(this.txtPersonas_Enter);
			// 
			// frmFacturacionHabitaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1256, 534);
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
			this.Controls.Add(this.txtPersonas);
			this.Controls.Add(this.lblIngresos);
			this.Controls.Add(this.btnIngresos);
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
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
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
				
		bool bNumeracionAutomatica = false;
		bool bImpreso = false;		
		bool bAnulado = false;
		bool bEliminaAlValidar = false;			
		decimal dAnticipo = 0.00m;
		bool bUsoAnticipo = false;	
		private bool bActualiza = true;
		int iOrigen = 0;
		DateTime dtFechaPromocion = DateTime.Today;
		decimal dPorcentajeIVA = 0.00m;
		decimal dPorcentajeIVADec = 0.00m;
		int idIngreso = 0;
		bool bNuevoIngreso = false;
		int iIngresos = 0;
		bool bCuponActivo = false;
		
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void Consulta(int idCompraC)
		{
			#region Maestro Detalle
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
				this.dtFecha.Value = dr.GetDateTime(10);
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
				this.cmbFormaPago.Value = dr.GetInt32(25);
				this.txtNotas.Text = dr.GetString(26);
				this.txtIva.Value = dr.GetDecimal(29);
				this.txtDescIvaTotal.Value = dr.GetDecimal(30);
				this.txtIvaTotal.Value = dr.GetDecimal(31);
				this.txtTotal.Value = dr.GetDecimal(32);
				idAsiento = dr.GetInt32(33);
				iOrigen = dr.GetInt32(34); 
				this.txtPersonas.Value = dr.GetInt32(35);
			}
			dr.Close();

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));
			this.cmbBodega.Value = iBodega;
			this.cmbHabitaciones.Value = idHabitacion;

			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleFacturas {0}", idCompra));

			iIngresos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Ingresos Where idCompra = {0}", idCompra));

			this.lblIngresos.Text = "INGRESOS : " + iIngresos.ToString();

			string[] sPrmCW = new string[2];
			sPrmCW = Facturacion.sEditarCuponWeb(idCompra, cdsSeteoF);
			idCuponWeb = int.Parse(sPrmCW[0]);

			if (!bMenuCuadricula)
			{
				#region Habilita Controles - permisos de usuario
				if (!bAnulado)
				{
					if (iEstado != 6)
					{
						if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
						{
							if (miAcceso.Editar) this.btnEditar.Enabled = true;// && ((int)this.cmbFormaPago.Value != 9 || (int)this.cmbFormaPago.Value != 6)) 
							if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
							if (miAcceso.BAnulaSoloDelDia && !bAnulado)
							{
								if ((int)this.cmbFormaPago.Value != 9)
								{
									if ((DateTime)this.dtFecha.Value == DateTime.Today)
										this.btnAnular.Enabled = true;
								}
							}
							if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;						
							if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
						}
					}
					else
					{
						if (miAcceso.Editar && ((int)this.cmbFormaPago.Value != 9 || (int)this.cmbFormaPago.Value != 6)) this.btnEditar.Enabled = true;
						if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
						if (miAcceso.BAnulaSoloDelDia && !bAnulado)
						{
							if ((int)this.cmbFormaPago.Value != 9)
							{
								if ((DateTime)this.dtFecha.Value == DateTime.Today) this.btnAnular.Enabled = true;
							}
						}
						if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;						
						//if (miAcceso.CrearAsiento && !bAnulado) this.btnAsiento.Enabled = true;
					}

					#region Impresion
					if (iEstado != 6)
					{
						if (!bImpreso) if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
						
						if (bImpreso) if (miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
					}									
					#endregion Impresion
			
				}
				else if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;
				#endregion Habilita Controles - permisos de usuario

				this.btnCancelar.Enabled = true;
				if (bAnulado) this.btnAnular.Text = "Eliminar";
			}
			#endregion Maestro Detalle			
		}

		private void frmFacturacionHabitaciones_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0703");
			
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			
			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga 1");
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFPagoCarga");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));

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
								
				string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
				iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);							
				
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				this.btnSalir.Enabled = true;

				if (idCompraV > 0) this.Consulta(idCompraV);
			}

			if (bMenuCuadricula)
			{
				this.StartPosition = FormStartPosition.CenterScreen;
				this.ControlBox = false;
				this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;// BorderStyle.FixedSingle;
				
				this.lblHabitacion.Visible = true;
				this.lblHabitacion.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec LocalHabitacionTematica {0}", idHabitacion));
								
				string sSQL = string.Format("Select ISNULL(Ocupado, 0) From Habitaciones Where idHabitacion = {0}", idHabitacion);
				int iEstadoHabitacion = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (iEstadoHabitacion == 1)	this.btnNuevo_Click(sender, e);

				if (iEstadoHabitacion == 2 && idCompraV > 0)
				{
					this.Consulta(idCompraV);

					this.btnEditar_Click(sender, e);
				}
				else this.btnNuevo_Click(sender, e);
			}
		}

		private void Total()
		{		
			#region Creacion de Variables
			decimal dDescuentoLineal = 0.00m;			
			#endregion Creacion de Variables
			
			this.txtIva.Value = 0.00m;			
			this.txtDescIvaTotal.Value = 0.00m;			
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;			

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
			{
				dDescuentoLineal = 0.00m;				

				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value && dr.Cells["DescuentoPorc"].Value != DBNull.Value)
				{
					dr.Cells["SubTotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 6);
					dDescuentoLineal = (Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value))/100.00m;

					dr.Cells["DescuentoArt"].Value = dDescuentoLineal;

					this.txtIva.Value = Convert.ToDecimal(this.txtIva.Value) + Math.Round(Convert.ToDecimal(dr.Cells["Subtotal"].Value), 6);
					this.txtDescIvaTotal.Value = Convert.ToDecimal(this.txtDescIvaTotal.Value) + Math.Round(Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 6);
				}

				this.txtIvaTotal.Value = ((Convert.ToDecimal(this.txtIva.Value) - Convert.ToDecimal(this.txtDescIvaTotal.Value)) * dPorcentajeIVA)/100.00m;
			
				this.txtTotal.Value = Math.Round(Convert.ToDecimal(this.txtIva.Value), 6)					
					- Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value), 6)
					+ Math.Round(Convert.ToDecimal(this.txtIvaTotal.Value), 6);
			}
			#endregion Desplazamiento de grilla			
		}

		private void ValidaCargaReservas()
		{
			idReserva = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Numero_Reserva {0}", idHabitacion)); 

			if (idReserva > 0 && DialogResult.Yes == MessageBox.Show("La Habitación esta reservada\n\n¿Desea aplicar la reservación?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DetalleFacturaReservaciones {0}, {1}, {2}, {3}", idReserva, 3422, idHabitacion, 0));
		}

		private void CargaHabitacion(int idPromocionD, int iCantidad)
		{
			int iFila = uGridDetalle.Rows.Count - 1;
			
			string sSQL = string.Format("Exec DetalleFacturaReservaciones {0}, {1}, {2}, {3}", 
				0, 3422, (int)this.cmbHabitaciones.Value, idPromocionD);
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
				#endregion Asigna valores en el grid
			} 
			drC.Close();
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
						//						if (bMenuCuadricula)
						//						{
						#region Valida Promociones Tarjetas VIP
						DataSet oDS = new DataSet();					
						cdsSeteoF.Schema.Connections[0].Open();
						C1.Data.SchemaObjects.Connection miConeccion = cdsSeteoF.Schema.Connections[0];
						miConeccion = cdsSeteoF.Schema.Connections[0];
						string sSQL = string.Format("Exec ValidaPromocionesTarjetaVIP {0}, {1}, {2}, {3}, '{4}'",
							(int)this.cmbBodega.Value, 3422, 2, (int)this.cmbHabitaciones.Value, Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"));
						SqlDataAdapter oDA = new SqlDataAdapter(sSQL, (SqlConnection) miConeccion.DbConnection);
						oDA.Fill(oDS);
					
						if (oDS.Tables[0].Rows.Count > 0)
						{
							using (frmValidaNumeroTarjetasVIP VTI = new frmValidaNumeroTarjetasVIP(sSQL, idCompra))//frmValidaTarjetaIngresos VTI = new frmValidaTarjetaIngresos(true, sSQL))
							{
								if (DialogResult.OK == VTI.ShowDialog())
								{
									#region Recorre Grid Promociones Tarjeta VIP
									int i = 0;
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drVTI in VTI.uGridDetalle.Rows.All)
									{							
										if ((bool)drVTI.Cells["Sel"].Value)
										{
											idPromocionD = 0;
											bDescuento = false;
											bAdicional = false;
											bOpcional = false;

											idPromocionD = (int)drVTI.Cells["idPromocion"].Value;
											bDescuento = (bool)drVTI.Cells["SumaDescuento"].Value;
											bAdicional = (bool)drVTI.Cells["Adicional"].Value;
											bOpcional = (bool)drVTI.Cells["Opcional"].Value;

											#region Promociones
											if (bDescuento)
											{
												if (i == 0) CargaHabitacion(idPromocionD, 1);
											}

											if (bAdicional)
											{
												if (i == 0) CargaHabitacion(0, 1);
				
												this.CargaArticulosPromocionAdicional(idPromocionD);												
											}

											if (bOpcional)
											{
												if (i == 0) CargaHabitacion(0, 1);

												foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
												{	
													if ((bool)dr.Cells["Principal"].Value) 
														CargaArticulosPromocionOpcional(idPromocionD);
												}
											}
											#endregion Promociones

											i++;
										}
									}
									#endregion Recorre Grid Promociones Tarjeta VIP

									bValidaNumeroIngresos = true;
									bPromocionVIP = true;
								}									
							}
						}
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
										//											else 
										//											{
										//												if (DialogResult.Yes == MessageBox.Show("¿Desea cargar el servicio de hospedaje?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2)) 									
										//													CargaHabitacion(idPromocionD, 1);
										//											}
										#endregion Carga Promociones
									}
									//										else
									//										{
									//											if (DialogResult.Yes == MessageBox.Show("¿Desea cargar el servicio de hospedaje?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2)) 									
									//											{		
									//												bValidaNumeroIngresos = false;
									//												CargaHabitacion(idPromocionD, 1);
									//											}
									//											else bValidaNumeroIngresos = false;
									//										}
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
							if (DialogResult.Yes == MessageBox.Show("¿Desea cargar el servicio de hospedaje?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2)) 									
							{		
								//bValidaNumeroIngresos = false;
								CargaHabitacion(idPromocionD, (iIngresos - iCantidad));
							}								
							//							}								
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

				#region Valida si hay promociones cargadas
				//				int iCont = 0;
				//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				//					if (dr.Cells["Tipo"].Value.ToString().Length > 0)
				//						if ((int)dr.Cells["Tipo"].Value > 1) iCont++;
				#endregion Valida si hay promociones cargadas

				//				if (iCont == 0)
				//				{
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
					#endregion Informacion Productos sin promocion
				}	
				else if (oDS.Tables[0].Rows.Count > 0)
				{
					#region Lista de Promociones
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
							#endregion Informacion Productos cuando no acepta la promocion
						}
					}
					#endregion Lista de Promociones
				}
				//				}
				//				else
				//				{
				//					#region Informacion Productos cuando ya existe una promocion cargada
				//					uGridDetalle.Rows[iFila].Cells["idDetCompra"].Value = 0;
				//					uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = (int)this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value;
				//					uGridDetalle.Rows[iFila].Cells["Codigo"].Value = this.uGridArticulos.ActiveRow.Cells["Codigo"].Value.ToString();				
				//					uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = 1;
				//					uGridDetalle.Rows[iFila].Cells["Precio"].Value = Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value)/Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)"));
				//					uGridDetalle.Rows[iFila].Cells["Impuesto"].Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente(GETDATE(), 1)"));
				//					uGridDetalle.Rows[iFila].Cells["DescuentoArt"].Value = ((Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value)/Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)"))) * dDescuentaPorc)/100;
				//					uGridDetalle.Rows[iFila].Cells["DescuentoPorc"].Value = dDescuentaPorc;
				//					uGridDetalle.Rows[iFila].Cells["SubTotal"].Value = Convert.ToDecimal(this.uGridArticulos.ActiveRow.Cells["Precio1"].Value);
				//					uGridDetalle.Rows[iFila].Cells["Principal"].Value = true;
				//					uGridDetalle.Rows[iFila].Cells["Tipo"].Value = 0;				
				//					uGridDetalle.Rows[iFila].Cells["Bloqueado"].Value = false;
				//					uGridDetalle.Rows[iFila].Cells["idCombo"].Value = 0;
				//					uGridDetalle.Rows[iFila].Cells["Articulo"].Value = this.uGridArticulos.ActiveRow.Cells["Articulo"].Value.ToString();
				//					#endregion Informacion Productos cuando ya existe una promocion cargada
				//				}
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
								uGridDetalle.DisplayLayout.Bands[0].AddNew();	
								iFila++;
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
				iFila++;
				uGridDetalle.DisplayLayout.Bands[0].AddNew();	

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
				#endregion Asigna valores en el grid
			} 
			drC.Close();					
		}
		
		private void btnNuevo_Click(object sender, System.EventArgs e)
		{	
			if (!bMenuCuadricula) this.btnCancelar_Click(sender, e);

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFechaIngreso, cdsSeteoF, true);

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
												
				this.btnCancelar.Enabled = true;			
				
				dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
				dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

				ValidaCargaReservas();

				this.btnVer_Click(sender, e);

				if (!bServicioPagado) this.txtPersonas.Enabled = true;

				this.btnTemporal.Enabled = true;

				this.cmbGuardia.TabIndex = 0;				
			}
			#endregion Peticion desde cuadricula

			#region Peticion desde menu
			if (!bMenuCuadricula)
			{				
				this.dtFecha.Value = DateTime.Today;				
									
				#region Habilita controles 
				if (iBodega == 0) 
				{
					this.cmbBodega.Value = 1; 
					this.cmbBodega.Enabled = true;
				}
				else this.cmbBodega.Value = iBodega;
			
				this.cmbHabitaciones.Enabled = true;
				if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;				
				this.dtFechaIngreso.Enabled = true;
				this.dtFechaSalida.Enabled = true;
				this.cmbTipoRuc.Value = 5;					
				this.txtUnidad.Enabled = true;
				this.cmbFormaPago.Enabled = true;
				
				this.txtNotas.Enabled = true;

				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;				
				this.btnGuardar.Enabled = true;
				this.btnTemporal.Enabled = false;			
				this.btnAsiento.Enabled = false;				
				#endregion Habilita controles 
				
				dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
				dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

				FuncionesProcedimientos.RetornaFechaServidor(this.dtFechaSalida, cdsSeteoF, true);

				this.optTurnoSalida.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTurno '{0}'", Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd HH:mm")));
			}
			#endregion Peticion desde menu

			iIngresos = 1;

			this.cmbGuardia.Enabled = true;
			this.txtNumero.Enabled = true;				
			this.txtTiempo.Enabled = true;
			this.cmbTipoRuc.Enabled = true;
			this.txtRuc.Enabled = true;
			
			this.cmbCobrador.Enabled = true;
			this.cmbCajero.Enabled = true;
			this.txtUnidad.Enabled = true;

			this.cmbFormaPago.Enabled = true;
			this.txtPersonas.Enabled = true;
			this.txtNotas.Enabled = true;

			this.txtBuscar.Enabled = true;
			this.btnVer.Enabled = true;

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
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaFacturas Busqueda = new frmBuscaFacturas(1, iBodega, miAcceso.BLimiteBusquedaFechas))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
					this.Consulta((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			#region Peticion desde cuadricula
			if (bMenuCuadricula)
			{
				FuncionesProcedimientos.RetornaFechaServidor(this.dtFechaSalida, cdsSeteoF, true);

				this.optTurnoSalida.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTurno '{0}'", Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd HH:mm")));
								
				this.dtFecha.Value = Funcion.dtEscalarSQL(cdsSeteoF, "Select dbo.FacturacionRetornaFecha()");
				
				this.dtFechaPromocion = (DateTime)this.dtFechaIngreso.Value;

				this.cmbTipoRuc.Enabled = true;
				if ((int)this.cmbTipoRuc.Value != 11) this.txtRuc.Enabled = true;
				this.cmbCobrador.Enabled = true;
				this.cmbCajero.Enabled = true;				

				this.txtBuscar.Enabled = true;
				this.btnVer.Enabled = true;
				this.btnVer_Click(sender, e);

				FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
				FuncionesProcedimientos.EstadoGrids(true, this.uGridArticulos);

				this.btnIngresos.Enabled = true;
				this.btnCliente.Enabled = true;
				this.btnAnular.Enabled = true;
				this.btnTemporal.Enabled = true;

				dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
				dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

				this.CargaArticulos(3422, true);
			}
			#endregion Peticion desde cuadricula

			#region Peticion desde menu
			if (!bMenuCuadricula)
			{	
				if (iOrigen == 2 && iEstado == 6)
				{
					MessageBox.Show("La factura no ha sido contabilizada, no puede editar", "Información", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (miAcceso.CambiarNumeracion) this.txtNumero.Enabled = true;

				if (!bImpreso) 
				{
					if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
					
					this.btnCliente.Enabled = true;

					this.cmbTipoRuc.Enabled = true;
					this.txtRuc.Enabled = true;

					this.txtBuscar.Enabled = true;
					this.btnVer.Enabled = true;
					this.btnVer_Click(sender, e);
					
					FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
					FuncionesProcedimientos.EstadoGrids(true, this.uGridArticulos);
				}

				if (bImpreso) 
				{
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

						FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
						FuncionesProcedimientos.EstadoGrids(true, this.uGridArticulos);
					}
				}

				dPorcentajeIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
				dPorcentajeIVADec = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

				this.dtFechaIngreso.Enabled = true;
				this.dtFechaSalida.Enabled = true;
//				this.optTurnoIngreso.Enabled = true;
//				this.optTurnoSalida.Enabled = true;
				this.cmbCajero.Enabled = true;
				this.cmbCobrador.Enabled = true;

				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnTemporal.Enabled = false;
				this.btnAsiento.Enabled = false;
				this.btnAnular.Enabled = false;
			}
			#endregion Peticion desde menu

			this.txtUnidad.Enabled = true;
			this.cmbFormaPago.Enabled = true;
			this.txtNotas.Enabled = true;
      			
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			
			bNuevo = false;
			bEdicion = true;

			this.cmbTipoRuc.Focus();
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
				#endregion Impresion
				
				if (bMenuCuadricula) miRepor.ShowDialog();
				if (!bMenuCuadricula) miRepor.Show();
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

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Imprimir(0);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			this.Total();

			bool bFiscal = true;

			if (Convert.ToDecimal(this.txtTotal.Value) == 0.00m) bFiscal = false;

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;

			if (!Validacion.vbComboVacio(this.cmbHabitaciones, "Seleccione una habitación")) return;

			if (!Validacion.vbComboVacio(this.cmbGuardia, "Seleccione un guardia")) return;

			if (bFiscal)
				if (!Validacion.vbNumeracion(bNumeracionAutomatica, idCompra, 1, (int)this.cmbBodega.Value, this.txtNumero, cdsSeteoF)) return;
			
			if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione el tipo de identificación del cliente")) return;
									
			if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;
			
			if (!Validacion.vbComboVacio(this.cmbFormaPago, "Seleccione la forma de pago")) return;

			if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, (int)this.cmbFormaPago.Value, cdsSeteoF)) return;

			if (!Validacion.vbComboVacio(this.cmbCobrador, "Seleccione el cobrador")) return;

			if (!Validacion.vbComboVacio(this.cmbCajero, "Seleccione el cajero")) return;
			
			#region Valida registro de servicio de hospedaje pasado los 15 minutos
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
						MessageBox.Show("Pasado 10 minutos debe registrar el servicio de habitación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtBuscar.Focus();
						return;
					}
				}
			}
			#endregion Valida registro de servicio de hospedaje pasado los 15 minutos
			
			if (this.uGridDetalle.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los items de la factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBuscar.Focus();
				return;
			}

			if (bFiscal)
				if (!Validacion.ValidaTotales((decimal)this.txtTotal.Value)) return;

			if (!bFiscal)
				this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 1, {0}, 0", (int)this.cmbBodega.Value));

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
			#endregion Variables

			using (VentaFPagoLot miCobro = new VentaFPagoLot(idCompra, (int)this.cmbFormaPago.Value, (decimal)this.txtTotal.Value, this.txtNumero.Text.ToString(), false, 0, 0, false, 0, false, 0, false, 0, 0, 0, idCliente, 0, 0, 0, 0))
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
						oCmdActualiza.CommandTimeout = 0;
						
						try
						{
							oTransaction = oConexion.BeginTransaction();
							oCmdActualiza.Transaction = oTransaction;

							#region Encabezado
							string sSQL = string.Format("Exec GuardaMaestroFacturas {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', {8}, '{9}', {10}, '{11}', {12}, {13}, {14}, '{15}', '{16}', {17}, '{18}', {19}, {20}, {21}, {22}, {23}, '{24}', {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}",
								idCompra, idReserva, 2, iEstado, (int)this.cmbBodega.Value, idProyecto, (int)this.cmbHabitaciones.Value, //6
								Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbGuardia.Value, this.txtNumero.Text.ToString(), (int)this.txtTiempo.Value, Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"), (int)this.optTurnoIngreso.Value, //12
								idCliente, idTipoRuc, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(), bFechaSalida, dtFechaSalida.ToString("yyyyMMdd HH:mm"), idTurnoSalida, //19
								idCobrador, idCajero, (int)this.txtUnidad.Value, idFormaPago, this.txtNotas.Text.ToString(), //24
								0.00m, 0.00m, Convert.ToDecimal(this.txtIva.Value), Convert.ToDecimal(this.txtDescIvaTotal.Value), Convert.ToDecimal(this.txtIvaTotal.Value), Convert.ToDecimal(this.txtTotal.Value), false, (int)this.txtPersonas.Value, bFiscal); //30
							oCmdActualiza.CommandText = sSQL;
							idCompra = (int)oCmdActualiza.ExecuteScalar();
							#endregion Encabezado

							#region Detalle
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
							{							
								string sSQLDetalle = string.Format("Exec GuardaDetalleFacturas {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}",
									(int)dr.Cells["idDetCompra"].Value, idCompra, //1
									(int)dr.Cells["idArticulo"].Value, dr.Cells["Codigo"].Value.ToString(), dr.Cells["Articulo"].Value.ToString(), (int)dr.Cells["Cantidad"].Value, //5
									Convert.ToDecimal(dr.Cells["Precio"].Value), Convert.ToDecimal(dr.Cells["Impuesto"].Value), Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), //8
									Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value), Convert.ToDecimal(dr.Cells["Subtotal"].Value),//10
									(int)dr.Cells["Tipo"].Value, (bool)dr.Cells["Principal"].Value, (bool)dr.Cells["Bloqueado"].Value, (int)dr.Cells["idCombo"].Value); //14
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
									if ((int)dr.Cells["idFormaPago"].Value == 3) dtFecCaducidad = (DateTime)dr.Cells["FechaCaducidad"].Value;
									if (dr.Cells["ValorUsoTarjeta"].Value != System.DBNull.Value) dComTarjeta = (decimal)dr.Cells["ValorUsoTarjeta"].Value;
									if (dr.Cells["idDocumento"].Value != System.DBNull.Value) idDocumento = (int)dr.Cells["idDocumento"].Value;
									if (dr.Cells["DescripcionPrenda"].Value != System.DBNull.Value) DesPrenda = dr.Cells["DescripcionPrenda"].Value.ToString();
								
									string sSQLCobro = string.Format("Exec GuardaDetFormaPagoLotte {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, {18}, '{19}',{20}", 
										idDetFormaPago, 
										idCompra, 1, (int)dr.Cells["idFormaPago"].Value, (decimal)dr.Cells["Valor"].Value,
										idBanco, idTarjeta, dr.Cells["CuentaCorriente"].Value.ToString(), dr.Cells["Numero"].Value.ToString(),
										dtFecCaducidad.ToString("yyyyMMdd"),	dr.Cells["NumVoucher"].Value.ToString(), dr.Cells["Autorizacion"].Value.ToString(),
										idCuenta, dr.Cells["Lote"].Value.ToString(), dr.Cells["Referencia"].Value.ToString(), 0, 0, 0, 0, DesPrenda,0);
									oCmdActualiza.CommandText = sSQLCobro;
									oCmdActualiza.ExecuteNonQuery();
								}								
							}
							#endregion Guarda Cobro de la factura
							
							#region Asiento Automatico
							string sSQLAsiento = string.Format("Exec CreaAsientoDeVenta {0}", idCompra);
							oCmdActualiza.CommandText = sSQLAsiento;
							oCmdActualiza.ExecuteNonQuery();

							string sSQLRetIdAsiento = string.Format("Select ISNULL(idAsiento, 0) From Compra Where idCompra = {0}", idCompra);
							oCmdActualiza.CommandText = sSQLRetIdAsiento; 
							idAsiento = (int)oCmdActualiza.ExecuteScalar();
							#endregion Asiento Automatico
							
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
					
									string sSQLEstado = string.Format("Exec ActualizaEstadoHabitacion {0}, {1}, {2}, {3}, {4}, 1", 
										idHabitacion, 0, (int)this.txtTiempo.Value, 2, bNuevo);
									oCmdActualiza.CommandText = sSQLEstado;
									oCmdActualiza.ExecuteNonQuery();								
									#endregion Genera Ingresos Mas De Dos Personas - Actualiza Estado Ocupado (2) 
								}
								else
								{
									#region Actualiza Estado Limpieza (3)
									string sSQLEstado = string.Format("Exec ActualizaEstadoHabitacion {0}, {1}, {2}, {3}, 0, 0", idHabitacion, 0, 5, 3);
									oCmdActualiza.CommandText = sSQLEstado;
									oCmdActualiza.ExecuteNonQuery();

									if (idReserva > 0)
									{
										oCmdActualiza.CommandText = string.Format("Update ReservaHabitacion Set Estado = 0, FechaFinaliza = GETDATE() Where idReserva = {0}", idReserva);
										oCmdActualiza.ExecuteNonQuery();
									}
									#endregion Actualiza Estado Limpieza (3)
								}								
							}

							oTransaction.Commit();		

							#region Acumula Puntos Tarjetas VIP
							if (bMenuCuadricula)
							{								
								using (frmValidaTarjetaIngresos VTI = new frmValidaTarjetaIngresos(false, ""))
								{
									if (DialogResult.OK == VTI.ShowDialog())
									{
										string sSQLVTI = string.Format("Insert Into TarjetasIngresosCompra (idTarjetasIngresos, idCompra) Values ({0}, {1})", VTI.idTarjetasIngresos, idCompra);
										Funcion.EjecutaSQL(cdsSeteoF, sSQLVTI);
									}
								}
							}					
							#endregion Acumula Puntos Tarjetas VIP

							Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 2, idCompra, cdsSeteoF);
							idCuponWeb = 0;

							#region Impresion
							if (bMenuCuadricula && bFiscal) 
							{
								using (frmMnsajeNumLotte Msje = new frmMnsajeNumLotte(this.txtNumero.Text.ToString(), false))							
									if (DialogResult.OK == Msje.ShowDialog())
										Imprimir(0);

								this.Close();
							}
							#endregion Impresion

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
							this.dtFechaIngreso.Enabled = false;
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
			}
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
				if (DialogResult.Yes == MessageBox.Show("Esta seguro de ELIMINAR este Documento \n\nEste proceso no puede ser revertido", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					string sSQLBorrar = string.Format("Exec BorrarCompraVenta {0}, {1}, {2}", idCompra, idAsiento, 1);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLBorrar, true);

					MessageBox.Show("Documento Eliminado del sistema", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.btnCancelar_Click(sender, e);							
				}

				return;						
			}
			#endregion Elimina Registro

			#region Valida Anulacion
			if (bAnulado)
			{
				MessageBox.Show("El Documento esta Anulado", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}
			if (!miAcceso.Anular)
			{
				MessageBox.Show("No tiene acceso a Anular " + Text.Trim(), "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
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

				if (iMinutos > 15)
				{
					MessageBox.Show("Pasado los 10 minutos desde el ingreso no puede anular la factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
			this.cmbCajero.Enabled = false;
			this.txtUnidad.Enabled = false;
			this.dtFechaSalida.Enabled = false;
			this.optTurnoSalida.Enabled = false;					
			this.cmbFormaPago.Enabled = false;
			this.txtPersonas.Enabled = false;
			this.txtBuscar.Enabled = false;
			this.btnVer.Enabled = false;
			this.txtNotas.Enabled = false;
			
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
			bUsoAnticipo = false;						
			bAnulado = false;			
			idBloqueaTransacciones = 0;
			idCuponWeb = 0;
			iOrigen = 0;
			bNuevoIngreso = false;
			bCuponActivo = false;
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

		private void btnTemporal_Click(object sender, System.EventArgs e)
		{
			this.Total();

			if (!Validacion.vbComboVacio(this.cmbGuardia, "Seleccione un Guardia")) return;

			if (!Validacion.vbNumeracion(bNumeracionAutomatica, idCompra, 1, (int)this.cmbBodega.Value, this.txtNumero, cdsSeteoF)) return;

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

			Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 2, idCompra, cdsSeteoF);
			idCuponWeb = 0;

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
						
				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Encabezado
					string sSQL = string.Format("Exec GuardaMaestroFacturas {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', {8}, '{9}', {10}, '{11}', {12}, {13}, {14}, '{15}', '{16}', {17}, '{18}', {19}, {20}, {21}, {22}, {23}, '{24}', {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, 1",
						idCompra, idReserva, 2, iEstado, (int)this.cmbBodega.Value, idProyecto, (int)this.cmbHabitaciones.Value, //6
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbGuardia.Value, this.txtNumero.Text.ToString(), (int)this.txtTiempo.Value, Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"), (int)this.optTurnoIngreso.Value, //12
						idCliente, idTipoRuc, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(), bFechaSalida, dtFechaSalida.ToString("yyyyMMdd HH:mm"), idTurnoSalida, //19
						idCobrador, idCajero, (int)this.txtUnidad.Value, idFormaPago, this.txtNotas.Text.ToString(), //24
						0.00m, 0.00m, Convert.ToDecimal(this.txtIva.Value), Convert.ToDecimal(this.txtDescIvaTotal.Value), Convert.ToDecimal(this.txtIvaTotal.Value), Convert.ToDecimal(this.txtTotal.Value), false, (int)this.txtPersonas.Value); //31
					oCmdActualiza.CommandText = sSQL;
					idCompra = (int)oCmdActualiza.ExecuteScalar();
					#endregion Encabezado

					#region Detalle
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					{							
						string sSQLDetalle = string.Format("Exec GuardaDetalleFacturas {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}",
							(int)dr.Cells["idDetCompra"].Value, idCompra, //1
							(int)dr.Cells["idArticulo"].Value, dr.Cells["Codigo"].Value.ToString(), dr.Cells["Articulo"].Value.ToString(), (int)dr.Cells["Cantidad"].Value, //5
							Convert.ToDecimal(dr.Cells["Precio"].Value), Convert.ToDecimal(dr.Cells["Impuesto"].Value), Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), //8
							Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value), Convert.ToDecimal(dr.Cells["Subtotal"].Value),//10
							(int)dr.Cells["Tipo"].Value, (bool)dr.Cells["Principal"].Value, (bool)dr.Cells["Bloqueado"].Value, (int)dr.Cells["idCombo"].Value); //14
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

							string sSQLEstado = string.Format("Exec ActualizaEstadoHabitacion {0}, {1}, {2}, {3}, {4}, {5}", 
								idHabitacion, idCompra, (int)this.txtTiempo.Value, 2, bNuevo, bServicioPagado);
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

					oTransaction.Commit();

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

					if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 6) this.txtRuc.MaxLength = 13;
					if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;

					#region Consumidor Final
					if ((int)this.cmbTipoRuc.Value == 11) 
					{						
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
			if (!bMenuCuadricula && bEdicion)
			{
				if (!Validacion.vbComboVacio(this.cmbHabitaciones, "Seleccione una habitación")) return;
	
				int iCantidad = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDet in this.uGridDetalle.Rows.All)					
					if ((int)drDet.Cells["idArticulo"].Value == 3422) iCantidad++;

				if (iCantidad == 0)
					this.CargaArticulos(3422, true);
			}

			this.uGridArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec FacturaLoteListaInventario '{0}'", this.txtBuscar.Text.ToString()));

			if (this.uGridArticulos.Rows.Count > 0) this.uGridArticulos.Focus(); else this.txtBuscar.Focus();
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

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null)			
				this.cmbHabitaciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idHabitacion, Nombre From Habitaciones Where Bodega = {0} Order by NumeroHabitacion", (int)this.cmbBodega.Value));
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
				if (this.uGridDetalle.ActiveRow != null) uGridDetalle.ActiveRow.Selected = false;

//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
//				{
//					if ((int)dr.Cells["idArticulo"].Value == (int)e.Cell.Row.Cells["idArticulo"].Value) 
//					{
//						MessageBox.Show("El articulo seleccionado ya esta cargado en la lista de consumos, modifique la Cantidad", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						uGridDetalle.ActiveRow = dr;
//						uGridDetalle.ActiveRow.Selected = true;
//						return;
//					}
//				}
				
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
				if (e.KeyCode == Keys.Escape)
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
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

			if (e.Cell.Column.ToString() == "Cantidad") 
			{
//				if ((int)e.Cell.Row.Cells["idArticulo"].Value == 3422 && (int)e.Cell.Row.Cells["Cantidad"].Value != iIngresos)
//				{
//					e.Cell.Row.Cells["Cantidad"].Value = iIngresos;
//				}
				this.Total(); 
				bActualiza = true;				
			}
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
				this.btnGuardar.Enabled = false;
				this.btnTemporal.Enabled = true;
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
				if (iEstado == 10 && bImpreso && (int)this.cmbTipoRuc.Value != 11)
				{
					MessageBox.Show("No puede eliminar filas de facturas impresas con datos de clientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
				
				for (int i = 0; i < e.Rows.Length; i++)
				{	
					if (idReserva > 0 && (int)e.Rows[i].Cells["idArticulo"].Value == 3422)
					{
						MessageBox.Show("No puede eliminar el servicio de habitación de reservaciones", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
						return;
					}

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

		private void btnCliente_Click(object sender, System.EventArgs e)
		{
			if (this.txtRuc.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el Numero de Identificacion del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
			}
			else
			{
				FuncionesProcedimientos.AbreFacturacionClientes((int)this.cmbTipoRuc.Value, this.txtRuc.Text.ToString(), true);

				CargaCliente();
			}
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbTipoRuc.ActiveRow != null)
				{
					#region Busca Cliente 
					idCliente = 0;
					this.txtNombre.Text = "";

					CargaCliente();

					if (idCliente == 0)
					{
						FuncionesProcedimientos.AbreFacturacionClientes((int)this.cmbTipoRuc.Value, this.txtRuc.Text.ToString(), true);

						CargaCliente();
					}
					#endregion Busca Cliente
				}
			}
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (this.cmbTipoRuc.ActiveRow != null)
					if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) e.Cancel = true;
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

		private void btnIngresos_Click(object sender, System.EventArgs e)
		{
//			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de generar nuevo ingreso?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
//			{
//				bNuevoIngreso = true;
//				MessageBox.Show("Nuevo ingreso generado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//				this.btnIngresos.Enabled = false;
//				this.btnGuardar.Enabled = false;
//				iIngresos++;
//			}
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
			if (bNuevo && bEdicion)
				if (this.cmbHabitaciones.ActiveRow != null)
					this.txtTiempo.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec FacturacionRetornaHoras {0}", (int)this.cmbHabitaciones.Value));
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

	}
}
