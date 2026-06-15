using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmInformeRetiroMercaderiaProductos.
	/// </summary>
	public class frmInformeRetiroMercaderiaProductos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label53;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnEstados;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstados;
		private System.Windows.Forms.Button btnTransferir;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private System.Windows.Forms.Label label36;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodegaDestino;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private System.Windows.Forms.Button btnCotizacion;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstadoAccesorios;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmInformeRetiroMercaderiaProductos()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmInformeRetiroMercaderiaProductos));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SelTF");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleIngresoRetiro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDeRetiro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRetiro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroNC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNDC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorNotaDeCredito");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloMod");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoNuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PVP");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoProducto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDeTransferencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Destino");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaActual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UbicacionActual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ult_Conteo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas_Pagadas");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas_Vencidas");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasAtraso");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalFactura");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalRecaudado");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoProducto");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoFactura");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalCosto");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NuevoPrecio");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Margen");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NuevoCosto");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioSugerido");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDTEstadoDT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoProdDT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoAccesorios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoAccesoriosM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasDT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoDT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoDTM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Garantia");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioRevisoDT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRevisoDT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransferencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompraTransf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDTEstadoCompras");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoProdCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasCompras");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoCompras");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoComprasM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioRevisoCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRevisoCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnRevisionDT", 0);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalFactura", 29, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalFactura", 29, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalRecaudado", 30, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalRecaudado", 30, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalCosto", 33, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalCosto", 33, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NuevoPrecio", 35, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NuevoPrecio", 35, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NuevoCosto", 37, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NuevoCosto", 37, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PrecioSugerido", 38, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PrecioSugerido", 38, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CostoProducto", 31, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CostoProducto", 31, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CostoFactura", 32, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CostoFactura", 32, true);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SelTF");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleIngresoRetiro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDeRetiro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaRetiro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroNC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaNDC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorNotaDeCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloMod");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoNuevo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PVP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoProducto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDeTransferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Destino");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BodegaActual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UbicacionActual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ult_Conteo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas_Pagadas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas_Vencidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasAtraso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalRecaudado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoProducto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NuevoPrecio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Margen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NuevoCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioSugerido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDTEstadoDT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoProdDT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoAccesorios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoAccesoriosM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasDT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoDT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoDTM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Garantia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioRevisoDT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaRevisoDT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTransferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompraTransf");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDTEstadoCompras");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoProdCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasCompras");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoCompras");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoComprasM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioRevisoCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaRevisoCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDTEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDTEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubgrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubgrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			this.lblContador = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label7 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnEstados = new System.Windows.Forms.Button();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btnTransferir = new System.Windows.Forms.Button();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label36 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbBodegaDestino = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label12 = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnCotizacion = new System.Windows.Forms.Button();
			this.optEstadoAccesorios = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoAccesorios)).BeginInit();
			this.SuspendLayout();
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(392, 104);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 691;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(200, 10);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(33, 17);
			this.label53.TabIndex = 688;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1112, 8);
			this.groupBox1.TabIndex = 689;
			this.groupBox1.TabStop = false;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 107);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(37, 17);
			this.label7.TabIndex = 687;
			this.label7.Text = "Buscar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(35, 17);
			this.lblBodega.TabIndex = 686;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBusqueda
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance1;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(64, 104);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(216, 22);
			this.txtBusqueda.TabIndex = 0;
			this.txtBusqueda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBusqueda_KeyDown);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(296, 103);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 1;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 2;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
			// 
			// dtHasta
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance3;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 3;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource1;
			appearance4.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance4;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 20;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 100;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Ingreso N.";
			ultraGridColumn4.Header.VisiblePosition = 7;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 90;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Fecha Retiro";
			ultraGridColumn5.Header.VisiblePosition = 9;
			ultraGridColumn5.Width = 80;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Ingreso N";
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Width = 90;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Numero N/C";
			ultraGridColumn7.Header.VisiblePosition = 11;
			ultraGridColumn7.Width = 90;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Fecha N/C";
			ultraGridColumn8.Header.VisiblePosition = 12;
			ultraGridColumn8.Width = 80;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance5;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Valor N/C";
			ultraGridColumn9.Header.VisiblePosition = 13;
			ultraGridColumn9.Width = 75;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 160;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Codigo Original";
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Width = 180;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "Codigo Nuevo";
			ultraGridColumn13.Header.VisiblePosition = 5;
			ultraGridColumn13.Width = 180;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Width = 200;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance6;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 34;
			ultraGridColumn15.Width = 80;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.Width = 150;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.Caption = "Estado";
			ultraGridColumn18.Header.VisiblePosition = 18;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 55;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.Caption = "N. Transferencia";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 21;
			ultraGridColumn21.Width = 150;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 20;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.Caption = "Ubicacion Actual";
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 150;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.Caption = "Ult. Conteo";
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.Caption = "N. Factura";
			ultraGridColumn25.Header.VisiblePosition = 14;
			ultraGridColumn25.Width = 80;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.Caption = "F. Factura";
			ultraGridColumn26.Header.VisiblePosition = 16;
			ultraGridColumn26.Width = 70;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance7;
			ultraGridColumn27.Format = "#,##0";
			ultraGridColumn27.Header.Caption = "C. Pagadas";
			ultraGridColumn27.Header.VisiblePosition = 24;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 90;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance8;
			ultraGridColumn28.Format = "#,##0";
			ultraGridColumn28.Header.Caption = "C Vencidas";
			ultraGridColumn28.Header.VisiblePosition = 25;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 90;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance9;
			ultraGridColumn29.Format = "#,##0";
			ultraGridColumn29.Header.Caption = "Dias de Atraso";
			ultraGridColumn29.Header.VisiblePosition = 26;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 90;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance10;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.Caption = "T. Factura";
			ultraGridColumn30.Header.VisiblePosition = 27;
			ultraGridColumn30.Width = 80;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance11;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.Caption = "Total Recaudado";
			ultraGridColumn31.Header.VisiblePosition = 28;
			ultraGridColumn31.Width = 100;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance12;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.Caption = "Costo Producto";
			ultraGridColumn32.Header.VisiblePosition = 29;
			ultraGridColumn32.Width = 90;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance13;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.Caption = "Costo Factura";
			ultraGridColumn33.Header.VisiblePosition = 30;
			ultraGridColumn33.Width = 90;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance14;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.Caption = "Total Costo";
			ultraGridColumn34.Header.VisiblePosition = 31;
			ultraGridColumn34.Width = 90;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance15;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 35;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance16;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.Caption = "Nuevo Precio";
			ultraGridColumn36.Header.VisiblePosition = 36;
			ultraGridColumn36.Width = 90;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance17;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.VisiblePosition = 37;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance18;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.Caption = "Nuevo Costo";
			ultraGridColumn38.Header.VisiblePosition = 32;
			ultraGridColumn38.Width = 90;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance19;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.Caption = "Precio Sugerido";
			ultraGridColumn39.Header.VisiblePosition = 33;
			ultraGridColumn39.Width = 90;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn40.Header.VisiblePosition = 38;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 140;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn41.Header.Caption = "Estado Prod DT";
			ultraGridColumn41.Header.VisiblePosition = 39;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 120;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn42.Header.VisiblePosition = 40;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn43.Header.Caption = "Accesorios";
			ultraGridColumn43.Header.VisiblePosition = 56;
			ultraGridColumn43.Width = 100;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn44.Header.Caption = "Notas";
			ultraGridColumn44.Header.VisiblePosition = 41;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 200;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn45.Header.Caption = "Estado";
			ultraGridColumn45.Header.VisiblePosition = 42;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn46.Header.Caption = "Estado";
			ultraGridColumn46.Header.VisiblePosition = 43;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 100;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance20;
			ultraGridColumn47.Header.VisiblePosition = 57;
			ultraGridColumn47.Width = 60;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn48.Header.Caption = "Revisado Por";
			ultraGridColumn48.Header.VisiblePosition = 44;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 100;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn49.Header.Caption = "Fecha Revisión";
			ultraGridColumn49.Header.VisiblePosition = 45;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 100;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn50.Header.VisiblePosition = 46;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn51.Header.VisiblePosition = 47;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn52.Header.VisiblePosition = 48;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 120;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn53.Header.Caption = "Estado Art Compras";
			ultraGridColumn53.Header.VisiblePosition = 49;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 120;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn54.Header.Caption = "Notas";
			ultraGridColumn54.Header.VisiblePosition = 50;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 200;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn55.Header.VisiblePosition = 51;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn56.Header.Caption = "Estado";
			ultraGridColumn56.Header.VisiblePosition = 53;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 100;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn57.Header.Caption = "Usuario Reviso";
			ultraGridColumn57.Header.VisiblePosition = 52;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 100;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn58.Header.Caption = "Fecha Reviso";
			ultraGridColumn58.Header.VisiblePosition = 54;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 100;
			ultraGridColumn59.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance21.Image = ((object)(resources.GetObject("appearance21.Image")));
			appearance21.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance21.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn59.CellButtonAppearance = appearance21;
			ultraGridColumn59.Header.Caption = "";
			ultraGridColumn59.Header.VisiblePosition = 58;
			ultraGridColumn59.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn59.Width = 30;
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
																										 ultraGridColumn59});
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance22;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance23;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance24;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance25;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance26;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance27;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance28;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance29;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance30.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance30;
			appearance31.ForeColor = System.Drawing.Color.Black;
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance32.FontData.BoldAsString = "True";
			appearance32.FontData.Name = "Arial";
			appearance32.FontData.SizeInPoints = 8F;
			appearance32.ForeColor = System.Drawing.Color.White;
			appearance32.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance32;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance34;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 152);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1096, 208);
			this.uGridInformacion.TabIndex = 681;
			this.uGridInformacion.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_CellChange);
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton);
			this.uGridInformacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInformacion_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn1.DefaultValue = false;
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(System.DateTime);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn35.DataType = typeof(System.Decimal);
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn37.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn40.DataType = typeof(int);
			ultraDataColumn42.DataType = typeof(int);
			ultraDataColumn42.DefaultValue = 0;
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(System.DateTime);
			ultraDataColumn50.DataType = typeof(int);
			ultraDataColumn51.DataType = typeof(int);
			ultraDataColumn52.DataType = typeof(int);
			ultraDataColumn55.DataType = typeof(int);
			ultraDataColumn58.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50,
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54,
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58});
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1016, 71);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(88, 24);
			this.btnExcel.TabIndex = 11;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnEstados
			// 
			this.btnEstados.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnEstados.Enabled = false;
			this.btnEstados.Image = ((System.Drawing.Image)(resources.GetObject("btnEstados.Image")));
			this.btnEstados.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEstados.Location = new System.Drawing.Point(1016, 39);
			this.btnEstados.Name = "btnEstados";
			this.btnEstados.Size = new System.Drawing.Size(88, 24);
			this.btnEstados.TabIndex = 10;
			this.btnEstados.Text = "Estados";
			this.btnEstados.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEstados.Click += new System.EventHandler(this.btnEstados_Click);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn59.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn59,
																																 ultraDataColumn60});
			// 
			// cmbEstados
			// 
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DataSource = this.ultraDataSource3;
			this.cmbEstados.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn60.Header.VisiblePosition = 0;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 78;
			ultraGridColumn61.Header.VisiblePosition = 1;
			ultraGridColumn61.Width = 98;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn60,
																										 ultraGridColumn61});
			ultraGridBand2.SummaryFooterCaption = "Grand Summaries";
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEstados.DisplayMember = "Descripcion";
			this.cmbEstados.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbEstados.Location = new System.Drawing.Point(656, 168);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(100, 104);
			this.cmbEstados.TabIndex = 693;
			this.cmbEstados.ValueMember = "idDTEstado";
			this.cmbEstados.Visible = false;
			// 
			// btnTransferir
			// 
			this.btnTransferir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnTransferir.Enabled = false;
			this.btnTransferir.Image = ((System.Drawing.Image)(resources.GetObject("btnTransferir.Image")));
			this.btnTransferir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTransferir.Location = new System.Drawing.Point(1016, 6);
			this.btnTransferir.Name = "btnTransferir";
			this.btnTransferir.Size = new System.Drawing.Size(88, 24);
			this.btnTransferir.TabIndex = 9;
			this.btnTransferir.Text = "Transferir";
			this.btnTransferir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnTransferir.Click += new System.EventHandler(this.btnTransferir_Click);
			// 
			// optEstados
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance35;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance36;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendiente";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Revisado";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Transferido";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(392, 38);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(312, 26);
			this.optEstados.TabIndex = 5;
			this.optEstados.Text = "Todos";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// optTipo
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipo.Appearance = appearance37;
			this.optTipo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTipo.ItemAppearance = appearance38;
			this.optTipo.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem5.DataValue = 0;
			valueListItem5.DisplayText = "Todos";
			valueListItem6.DataValue = 1;
			valueListItem6.DisplayText = "Departamento Técnico";
			valueListItem7.DataValue = 2;
			valueListItem7.DisplayText = "Compras";
			this.optTipo.Items.Add(valueListItem5);
			this.optTipo.Items.Add(valueListItem6);
			this.optTipo.Items.Add(valueListItem7);
			this.optTipo.ItemSpacingVertical = 10;
			this.optTipo.Location = new System.Drawing.Point(392, 5);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(312, 26);
			this.optTipo.TabIndex = 4;
			this.optTipo.Visible = false;
			this.optTipo.ValueChanged += new System.EventHandler(this.optTipo_ValueChanged);
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.Location = new System.Drawing.Point(8, 43);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(41, 17);
			this.label36.TabIndex = 697;
			this.label36.Text = "Bodega";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			appearance39.ForeColor = System.Drawing.Color.Black;
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance39;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn62.Header.VisiblePosition = 2;
			ultraGridColumn63.Header.VisiblePosition = 1;
			ultraGridColumn63.Width = 180;
			ultraGridColumn64.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(216, 22);
			this.cmbBodega.TabIndex = 6;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(720, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 17);
			this.label1.TabIndex = 699;
			this.label1.Text = "Destino";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodegaDestino
			// 
			this.cmbBodegaDestino.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance40.ForeColor = System.Drawing.Color.Black;
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodegaDestino.Appearance = appearance40;
			this.cmbBodegaDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodegaDestino.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn65.Header.VisiblePosition = 2;
			ultraGridColumn66.Header.VisiblePosition = 1;
			ultraGridColumn66.Width = 180;
			ultraGridColumn67.Header.VisiblePosition = 0;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67});
			this.cmbBodegaDestino.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbBodegaDestino.DisplayMember = "Nombre";
			this.cmbBodegaDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodegaDestino.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodegaDestino.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaDestino.Location = new System.Drawing.Point(784, 7);
			this.cmbBodegaDestino.MaxDropDownItems = 30;
			this.cmbBodegaDestino.Name = "cmbBodegaDestino";
			this.cmbBodegaDestino.Size = new System.Drawing.Size(216, 22);
			this.cmbBodegaDestino.TabIndex = 8;
			this.cmbBodegaDestino.ValueMember = "Bodega";
			this.cmbBodegaDestino.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodegaDestino_KeyDown);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(8, 75);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(53, 17);
			this.label12.TabIndex = 702;
			this.label12.Text = "SubGrupo";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn61.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn61,
																																 ultraDataColumn62});
			// 
			// cmbSubGrupo
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance41;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataSource = this.ultraDataSource2;
			ultraGridColumn68.Header.VisiblePosition = 0;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.Caption = "Subgrupo";
			ultraGridColumn69.Header.VisiblePosition = 1;
			ultraGridColumn69.Width = 300;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn68,
																										 ultraGridColumn69});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbSubGrupo.DisplayMember = "Subgrupo";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(64, 72);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(216, 22);
			this.cmbSubGrupo.TabIndex = 7;
			this.cmbSubGrupo.ValueMember = "idSubgrupo";
			this.cmbSubGrupo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSubGrupo_KeyDown);
			// 
			// btnCotizacion
			// 
			this.btnCotizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCotizacion.Enabled = false;
			this.btnCotizacion.Image = ((System.Drawing.Image)(resources.GetObject("btnCotizacion.Image")));
			this.btnCotizacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCotizacion.Location = new System.Drawing.Point(1016, 103);
			this.btnCotizacion.Name = "btnCotizacion";
			this.btnCotizacion.Size = new System.Drawing.Size(88, 24);
			this.btnCotizacion.TabIndex = 703;
			this.btnCotizacion.Text = "Cotización";
			this.btnCotizacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCotizacion.Click += new System.EventHandler(this.btnCotizacion_Click);
			// 
			// optEstadoAccesorios
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstadoAccesorios.Appearance = appearance42;
			this.optEstadoAccesorios.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstadoAccesorios.CheckedIndex = 0;
			this.optEstadoAccesorios.ItemAppearance = appearance43;
			this.optEstadoAccesorios.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem8.DataValue = 0;
			valueListItem8.DisplayText = "Todos";
			valueListItem9.DataValue = 1;
			valueListItem9.DisplayText = "Completo";
			valueListItem10.DataValue = 2;
			valueListItem10.DisplayText = "Incompleto";
			this.optEstadoAccesorios.Items.Add(valueListItem8);
			this.optEstadoAccesorios.Items.Add(valueListItem9);
			this.optEstadoAccesorios.Items.Add(valueListItem10);
			this.optEstadoAccesorios.ItemSpacingVertical = 10;
			this.optEstadoAccesorios.Location = new System.Drawing.Point(392, 70);
			this.optEstadoAccesorios.Name = "optEstadoAccesorios";
			this.optEstadoAccesorios.Size = new System.Drawing.Size(312, 26);
			this.optEstadoAccesorios.TabIndex = 704;
			this.optEstadoAccesorios.Text = "Todos";
			// 
			// frmInformeRetiroMercaderiaProductos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1112, 366);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label36);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.optEstadoAccesorios);
			this.Controls.Add(this.btnCotizacion);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.cmbBodegaDestino);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.btnTransferir);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.btnEstados);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmInformeRetiroMercaderiaProductos";
			this.Text = "Informe Retiro de Mercaderia por Articulos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmInformeRetiroMercaderiaProductos_KeyDown);
			this.Load += new System.EventHandler(this.frmInformeRetiroMercaderiaProductos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoAccesorios)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmInformeRetiroMercaderiaProductos_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;			
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Codigo, Nombre, Bodega From Bodega Where Bodega In (5, 20, 95)");
			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT idDTEstado, Descripcion FROM DTEstados");
			this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSubgrupo, SubGrupo From ArticuloSubGrupo Order By SubGrupo");

			miAcceso = new Acceso(cdsSeteoF, "1006");

			if (miAcceso.BTodos) 
			{
				this.optTipo.Visible = true;
				this.optTipo.Value = 0;
			}			
			else if (miAcceso.BDepartamentoTecnico) this.optTipo.Value = 1; 
			else if (miAcceso.BCompras) this.optTipo.Value = 2; 

			if (miAcceso.BEstados) this.btnEstados.Enabled = true;
			if (miAcceso.BExportar) this.btnExcel.Enabled = true;
			if (miAcceso.BTransferir) 
			{
				this.cmbBodegaDestino.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Codigo, Nombre, Bodega From Bodega Where Bodega In (36, 46, 49, 79, 81, 94)");
				this.cmbBodegaDestino.Value = 36;
				this.cmbBodega.Enabled = true;
				this.btnTransferir.Enabled = true;			
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{			
			if (!Validacion.vbComparaFechas((DateTime)this.dtDesde.Value, (DateTime)this.dtHasta.Value, this.dtDesde, "La fecha inicial no puede ser posterior a la fecha inicial", false)) return;
			
			#region Consulta
			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;
			int idSubgrupo = 0;
			if (this.cmbSubGrupo.ActiveRow != null) idSubgrupo = (int)this.cmbSubGrupo.Value;

			string sSQL = string.Format("Exec InformeRetiroDeMercaderiaPorArticulo '{0}', '{1}', '{2}', {3}, {4}, {5}, {6}, {7}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), this.txtBusqueda.Text.ToString(),
				(int)this.optTipo.Value, (int)this.optEstados.Value, iBodega, idSubgrupo, (int)this.optEstadoAccesorios.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInformacion);
			
			this.lblContador.Text = this.uGridInformacion.Rows.Count + " Articulos Encontrados";
			#endregion Consulta

			#region Totales
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
			{					
				//dr.Cells["TotalRecaudado"].Value = Convert.ToDecimal(dr.Cells["CuotaEntrada"].Value) + Convert.ToDecimal(dr.Cells["ValorCobrado"].Value);

				dr.Cells["TotalCosto"].Value = Convert.ToDecimal(dr.Cells["CostoFactura"].Value) + Convert.ToDecimal(dr.Cells["CostoProducto"].Value);

				if (Convert.ToDecimal(dr.Cells["NuevoPrecio"].Value) == 0.00m) dr.Cells["NuevoPrecio"].Value = Convert.ToDecimal(dr.Cells["TotalFactura"].Value) - Convert.ToDecimal(dr.Cells["TotalCosto"].Value);
				dr.Cells["NuevoCosto"].Value = Convert.ToDecimal(dr.Cells["TotalCosto"].Value) - Convert.ToDecimal(dr.Cells["TotalRecaudado"].Value);
				dr.Cells["PrecioSugerido"].Value = Convert.ToDecimal(dr.Cells["TotalFactura"].Value) * 0.50m;
			}
			#endregion Totales
		}

		private void btnEstados_Click(object sender, System.EventArgs e)
		{
			frmEstadoDTCompraRM Estados = new frmEstadoDTCompraRM();
			Estados.ShowDialog();
		}

		private void frmInformeRetiroMercaderiaProductos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnRevisionDT")
			{
				if (e.Cell.Row.Cells["idDetalleIngresoRetiro"].Value != System.DBNull.Value)
				{					
					#region Valida opcion DT o Compras
					if ((int)this.optTipo.Value == 0)
					{
						MessageBox.Show("Para ingresar un estado seleccione Departamento Técnico o Compras", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.optTipo.Focus();
						return;	
					}
					#endregion Valida opcion DT o Compras

					#region Departamento Tecnico
					if ((int)this.optTipo.Value == 1)
					{
						if ((int)e.Cell.Row.Cells["EstadoDT"].Value == 3)
						{
							MessageBox.Show("El articulo ya fue transferido a la Bodega de Obsoletos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);							
							return;	
						}

						#region Guarda Observaciones
						using (frmDTEstados EstadosDT = new frmDTEstados((int)e.Cell.Row.Cells["idDetalleIngresoRetiro"].Value, e.Cell.Row.Cells["EstadoProducto"].Value.ToString(), e.Cell.Row.Cells["Observaciones"].Value.ToString(), 
										 (int)e.Cell.Row.Cells["EstadoDT"].Value, (int)e.Cell.Row.Cells["idDTEstadoDT"].Value, (int)e.Cell.Row.Cells["EstadoAccesorios"].Value, e.Cell.Row.Cells["NotasDT"].Value.ToString(), 
										 (int)e.Cell.Row.Cells["Garantia"].Value, miAcceso.BModRevDT))
						{
							if (DialogResult.OK == EstadosDT.ShowDialog())
							{
								string sSQL = string.Format("Exec RMActualizaEstadosDTCompras {0}, {1}, {2}, '{3}', {4}, {5}, {6}, 0, 0, 0, {7}",
									(int)e.Cell.Row.Cells["idDetalleIngresoRetiro"].Value, false, (int)EstadosDT.cmbEstados.Value, EstadosDT.txtObservaciones.Text.ToString(), 2,
									(int)e.Cell.Row.Cells["idSerial"].Value, (int)EstadosDT.optEstado.Value, 0);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL);
								
								this.btnVer_Click(sender, e);
							}
						}
						#endregion Guarda Observaciones
					}
					#endregion Departamento Tecnico

					#region Compras
					if ((int)this.optTipo.Value == 2)
					{						
						if ((int)e.Cell.Row.Cells["EstadoCompras"].Value == 3)
						{
							MessageBox.Show("El articulo ya fue transferido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);							
							return;	
						}
												
						using (frmComprasEstado EstadosDT = new frmComprasEstado((int)e.Cell.Row.Cells["idDetalleIngresoRetiro"].Value, 
										 e.Cell.Row.Cells["EstadoProducto"].Value.ToString(), e.Cell.Row.Cells["Observaciones"].Value.ToString(), 
										 e.Cell.Row.Cells["EstadoProdDT"].Value.ToString(), (int)e.Cell.Row.Cells["idDTEstadoDT"].Value, e.Cell.Row.Cells["NotasDT"].Value.ToString(), 
										 (int)e.Cell.Row.Cells["EstadoCompras"].Value,
										 Convert.ToDecimal(e.Cell.Row.Cells["NuevoCosto"].Value), 
										 Convert.ToDecimal(e.Cell.Row.Cells["PVP"].Value), 
										 Convert.ToDecimal(e.Cell.Row.Cells["Descuento"].Value), 
										 Convert.ToDecimal(e.Cell.Row.Cells["NuevoPrecio"].Value), 
										 Convert.ToDecimal(e.Cell.Row.Cells["Margen"].Value), 
										 e.Cell.Row.Cells["NotasCompras"].Value.ToString(), miAcceso.BModRevCompras))
						{
							if (DialogResult.OK == EstadosDT.ShowDialog())
							{
								string sSQL = string.Format("Exec RMActualizaEstadosDTCompras {0}, {1}, {2}, '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10}",
									(int)e.Cell.Row.Cells["idDetalleIngresoRetiro"].Value, true, 0, EstadosDT.txtObservacionesCompras.Text.ToString(), 2,
									(int)e.Cell.Row.Cells["idSerial"].Value, 0,
									Convert.ToDecimal(EstadosDT.txtDescuento.Value), Convert.ToDecimal(EstadosDT.txtNuevoPrecio.Value), Convert.ToDecimal(EstadosDT.txtMargen.Value), 
									(int)EstadosDT.txtGarantia.Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL);
							
								this.btnVer_Click(sender, e);
							}
						}
					}
					#endregion Compras					
				}			
			}			
		}

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void OcultaColumnasDT(bool bEstado)
		{
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["EstadoProdDT"].Hidden = bEstado;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["NotasDT"].Hidden = bEstado;			
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["EstadoDTM"].Hidden = bEstado;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["UsuarioRevisoDT"].Hidden = bEstado;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["FechaRevisoDT"].Hidden = bEstado;
//			this.uGridInformacion.DisplayLayout.Bands[0].Columns["EstadoAccesoriosM"].Hidden = bEstado;
		}

		private void OcultaColumnasCompras(bool bEstado)
		{
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["EstadoProdCompra"].Hidden = bEstado;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["NotasCompras"].Hidden = bEstado;			
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["EstadoComprasM"].Hidden = bEstado;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["UsuarioRevisoCompra"].Hidden = bEstado;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["FechaRevisoCompra"].Hidden = bEstado;
		}

		private void optTipo_ValueChanged(object sender, System.EventArgs e)
		{						
			if ((int)this.optTipo.Value == 0)
			{
				OcultaColumnasDT(false);
				OcultaColumnasCompras(false);				
			}

			if ((int)this.optTipo.Value == 1)
			{
				OcultaColumnasDT(false);
				OcultaColumnasCompras(true);
				this.btnCotizacion.Enabled = true;
			}

			if ((int)this.optTipo.Value == 2)
			{
				OcultaColumnasDT(true);
				OcultaColumnasCompras(false);
			}

			this.btnVer_Click(sender, e);
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void txtBusqueda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void btnTransferir_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if ((int)this.optTipo.Value != 2)
			{			
				MessageBox.Show("Transferencia solo para compras", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if ((int)this.optEstados.Value != 2)
			{			
				MessageBox.Show("Para Transferir debe seleccionar el estado revisado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (this.txtBusqueda.Text.ToString().Trim().Length > 0)
			{			
				MessageBox.Show("Quite el filtro en la opcion de Busqueda y haga click en Ver", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.txtBusqueda.Text = "";
				return;
			}

			if (this.uGridInformacion.Rows.Count == 0)
			{			
				MessageBox.Show("No existen productos para transferir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
				if ((bool)dr.Cells["SelTF"].Value) iCont++;

			if (iCont == 0)
			{			
				MessageBox.Show("No existen productos seleccionados para transferir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una bodega")) return;
			if (!Validacion.vbComboVacio(this.cmbBodegaDestino, "Seleccione la bodega de destino")) return;
			#endregion Validacion
			
			this.Cursor = Cursors.WaitCursor;

			this.btnVer_Click(sender, e);

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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
					
					#region Variables
					DateTime dtFecha = DateTime.Today;
					string sNumero = "";
					int idRealiza = 0; 						
					int idCompraE = 0;
					int idCompraI = 0;								
					int iBodega = 0;
					int iBodegaDestino = 0;
					int idTransferencia = 0;
					#endregion Variables

					if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;			
					if (this.cmbBodegaDestino.ActiveRow != null) iBodegaDestino = (int)this.cmbBodegaDestino.Value;

					#region Guarda Transferencia
					string sSQLNumero = "Exec NumeracionLocales 42, 0, 0";
					oCmdActualiza.CommandText = sSQLNumero;
					sNumero = oCmdActualiza.ExecuteScalar().ToString();					

					string sSQLRealiza = string.Format("If Exists(Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0})) Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0}) Else Select 0", MenuLatinium.IdUsuario);
					oCmdActualiza.CommandText = sSQLRealiza;
					idRealiza = (int)oCmdActualiza.ExecuteScalar();

					string sSQLGrabaMaestro = string.Format("Exec TransferenciasGrabaMaestro {0}, '{1}', '{2}', {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, '{10}', '{11}', {12}, {13}", 
						0, sNumero, dtFecha.ToString("yyyyMMdd HH:mm"), iBodega, iBodegaDestino, 1, 
						"TRANSFERENCIA DE MERCADERIA DE REMATES GENERADA DESDE EL DEPARTAMENTO TECNICO A BODEGAS DE REMATES", "", 0, 0, "", "", idRealiza, 0);
					oCmdActualiza.CommandText = sSQLGrabaMaestro;
					idTransferencia = (int)oCmdActualiza.ExecuteScalar();

					string sSQLIE = string.Format("Exec CreaIEBodegaTFDTRemates {0}, {1}", idTransferencia, iBodegaDestino);
					oCmdActualiza.CommandText = sSQLIE;
					SqlDataReader drIE = oCmdActualiza.ExecuteReader();					
					drIE.Read();

					if (drIE.HasRows)
					{
						idCompraE = drIE.GetInt32(0);
						idCompraI = drIE.GetInt32(1);
					}
					drIE.Close();
					#endregion Guarda Transferencia

					#region Guarda Detalle de transferencia
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
					{
						if ((bool)dr.Cells["SelTF"].Value)
						{
							if ((int)dr.Cells["BodegaActual"].Value == iBodega)
							{
								string sSQLGrabaDetalle = string.Format("Exec TransferenciasGrabaDetalle {0}, {1}, {2}, {3}, {4}",
									0, idTransferencia, (int)dr.Cells["idArticuloMod"].Value, 1, false);
								oCmdActualiza.CommandText = sSQLGrabaDetalle;
								int idDetTransferencia = (int)oCmdActualiza.ExecuteScalar();		

								string sSQLActEstadoIdTf = string.Format("Update DTDetalleIngresoRetiros Set idTransferencia = {0}, idDetCompraTransf = {1}, EstadoDT = 3, EstadoCompras = 3 Where idDetalleIngresoRetiro = {2}", 
									idTransferencia, idDetTransferencia, (int)dr.Cells["idDetalleIngresoRetiro"].Value);
								oCmdActualiza.CommandText = sSQLActEstadoIdTf;
								oCmdActualiza.ExecuteNonQuery();
					
								string sSQLDIEB = string.Format("Exec CreaDetalleIEBodegaTFDTRemates {0}, {1}, {2}, {3}, {4}, {5}, {6}", idCompraE, idCompraI, 
									(int)dr.Cells["idArticuloMod"].Value, (int)dr.Cells["idSerial"].Value, idDetTransferencia, 0.00m, iBodega);							
								oCmdActualiza.CommandText = sSQLDIEB;
								oCmdActualiza.ExecuteNonQuery();
							}
						}
					}
					#endregion Guarda Detalle de transferencia

					#region Procesa transferencia
					string sSQLProc = string.Format("Exec ProcesarTransferencias {0}, {1}", idCompraE, 8);
					oCmdActualiza.CommandText = sSQLProc;
					oCmdActualiza.ExecuteNonQuery();
					
					string sSQLActCantEnvioTf = string.Format("Update DetCompraTransf Set Enviado = Cantidad Where idCompraTransf = {0}", idTransferencia);
					oCmdActualiza.CommandText = sSQLActCantEnvioTf;
					oCmdActualiza.ExecuteNonQuery();
					#endregion Procesa transferencia

					oTransaction.Commit();			

					#region Impresion
					ParameterFields paramFields = new ParameterFields ();
				
					ParameterField paramFieldCompra = new ParameterField ();
					ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
					paramFieldCompra.ParameterFieldName = "@idCompra";
					ValIdCompra.Value = int.Parse(idCompraE.ToString());
					paramFieldCompra.CurrentValues.Add (ValIdCompra);
					paramFields.Add (paramFieldCompra);
									
					Reporte Rep = new Reporte("RepTransfSalida.rpt", "");
					Rep.MdiParent = this.MdiParent;
					Rep.Titulo("Transferencia");
					Rep.crVista.ParameterFieldInfo = paramFields;
					Rep.Show();
					#endregion Impresion

					this.btnVer_Click(sender, e);

					this.Cursor = Cursors.Default;
				}
				catch (Exception ex)
				{
					try
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click(sender, e);
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbSubGrupo.Focus();
		}

		private void cmbSubGrupo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBusqueda.Focus();
		}

		private void cmbBodegaDestino_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnTransferir.Focus();
		}

		private void btnCotizacion_Click(object sender, System.EventArgs e)
		{
			frmRMRCotizacionAccesorios CA = new frmRMRCotizacionAccesorios();
			CA.ShowDialog();
		}

		private void uGridInformacion_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "SelTF")
			{
				if (e.Cell.Row.Cells["idDetalleIngresoRetiro"].Value != System.DBNull.Value)
				{
					this.uGridInformacion.UpdateData();
					
					string sSQL = string.Format("Exec ActualizaSelTfRemates {0}, {1}", 
						(int)e.Cell.Row.Cells["idDetalleIngresoRetiro"].Value, (bool)e.Cell.Row.Cells["SelTF"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					
				}
			}
		}
	}
}
