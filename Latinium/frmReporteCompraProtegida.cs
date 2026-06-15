using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using Infragistics.Shared;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteCompraProtegida.
	/// </summary>
	public class frmReporteCompraProtegida : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox gbBotones;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnDocumentoEnviado;
		/// <summary>
		/// Variable del diseñador requerida.
		#region Variables
		public static int iEstadoSeguroDocumentoEnviado = 5;
		public static int iEstadoSeguroSeguroLiquidado = 6;
		public static int iEstadoSeguroCerrado = 9;
		public static int iRobo = 1;
		public static int iDano = 2;
		public static int iCompraFormaPagoDepositoAseguradora = 25;
		public static int iCompraFormaPagoDepositoDeducible = 26;
		public static int iCompraFormaPagoDepositoTransferencia = 13;
		//Pagina Inicial
		public static int Pagina = 1;
		public static int TotalPagina = 1;
		public static int TamanoPagina = 20;

		#endregion Variables
		public static string FechaDesde = "05/07/2019";
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optProductos;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnFin;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPagina;
		private System.Windows.Forms.Button btnAnterior;
		private System.Windows.Forms.Button btnInicio;
		private System.Windows.Forms.Button btnSiguiente;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		/// </summary>
		private System.ComponentModel.Container components = null;


		public frmReporteCompraProtegida()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteCompraProtegida));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Linea");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio_Contado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Seguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSeguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaNotificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoNotificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProforma");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Asegurado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoCobertura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cobertura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ProformaST");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deducible");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCierre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipoD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipoP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DirectorioCP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SelDocumentoEnviado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorDocumentoEnviado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorDocumentoEnviadoCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoReemplazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraFacturaAnterior");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bTieneFacturaAnterior");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bTieneDesarrolloActual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSeguroLiquidado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bTieneDepositoAnterior");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seguro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Linea");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio_Contado");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Seguro");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSeguro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNotificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoNotificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Asegurado");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCobertura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobertura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProformaST");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deducible");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCierre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipoD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipoP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DirectorioCP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SelDocumentoEnviado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorDocumentoEnviado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorDocumentoEnviadoCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoReemplazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraFacturaAnterior");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloGE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bTieneFacturaAnterior");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bTieneDesarrolloActual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSeguroLiquidado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bTieneDepositoAnterior");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnGestion", 0);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnProforma", 1);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 2);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 17, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Costo", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Costo", 16, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Precio_Contado", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Precio_Contado", 15, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Seguro", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Seguro", 18, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Linea");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio_Contado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Seguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSeguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaNotificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoNotificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProforma");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Asegurado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoCobertura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cobertura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ProformaST");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deducible");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCierre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipoD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipoP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DirectorioCP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SelDocumentoEnviado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorDocumentoEnviado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorDocumentoEnviadoCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoReemplazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraFacturaAnterior");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn96 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn97 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn98 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn99 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bTieneFacturaAnterior");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn100 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bTieneDesarrolloActual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn101 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSeguroLiquidado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn102 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bTieneDepositoAnterior");
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.btnExcel = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnDocumentoEnviado = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.optProductos = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnFin = new System.Windows.Forms.Button();
			this.txtPagina = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnAnterior = new System.Windows.Forms.Button();
			this.btnInicio = new System.Windows.Forms.Button();
			this.btnSiguiente = new System.Windows.Forms.Button();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnSalir = new System.Windows.Forms.Button();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optProductos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtPagina)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
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
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(-10, 166);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1306, 9);
			this.gbBotones.TabIndex = 975;
			this.gbBotones.TabStop = false;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1181, 55);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(105, 28);
			this.btnExcel.TabIndex = 977;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.DateTime);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(bool);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(System.DateTime);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(bool);
			ultraDataColumn37.DefaultValue = false;
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn38.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn42.DataType = typeof(int);
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(bool);
			ultraDataColumn49.DataType = typeof(bool);
			ultraDataColumn50.DataType = typeof(System.DateTime);
			ultraDataColumn51.DataType = typeof(bool);
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
																																 ultraDataColumn51});
			// 
			// btnDocumentoEnviado
			// 
			this.btnDocumentoEnviado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnDocumentoEnviado.Image = ((System.Drawing.Image)(resources.GetObject("btnDocumentoEnviado.Image")));
			this.btnDocumentoEnviado.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDocumentoEnviado.Location = new System.Drawing.Point(1181, 18);
			this.btnDocumentoEnviado.Name = "btnDocumentoEnviado";
			this.btnDocumentoEnviado.Size = new System.Drawing.Size(105, 27);
			this.btnDocumentoEnviado.TabIndex = 1034;
			this.btnDocumentoEnviado.Text = "D. enviados";
			this.btnDocumentoEnviado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDocumentoEnviado.Click += new System.EventHandler(this.btnDocumentoEnviado_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.txtBusqueda);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.label53);
			this.groupBox1.Controls.Add(this.lblBodega);
			this.groupBox1.Controls.Add(this.optProductos);
			this.groupBox1.Controls.Add(this.lblContador);
			this.groupBox1.Controls.Add(this.btnVer);
			this.groupBox1.Controls.Add(this.dtDesde);
			this.groupBox1.Controls.Add(this.dtHasta);
			this.groupBox1.Location = new System.Drawing.Point(19, 9);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1133, 148);
			this.groupBox1.TabIndex = 1035;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Datos principales";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// txtBusqueda
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance1;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(221, 111);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(249, 24);
			this.txtBusqueda.TabIndex = 1001;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(10, 111);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(198, 18);
			this.label1.TabIndex = 1000;
			this.label1.Text = "Cédula/Ruc/Nombre/N° Factura:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(259, 28);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(44, 18);
			this.label53.TabIndex = 996;
			this.label53.Text = "Hasta:";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(10, 28);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(47, 18);
			this.lblBodega.TabIndex = 995;
			this.lblBodega.Text = "Desde:";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optProductos
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.optProductos.Appearance = appearance2;
			this.optProductos.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optProductos.ItemAppearance = appearance3;
			this.optProductos.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Sin reclamo";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Notificado";
			valueListItem4.DataValue = 5;
			valueListItem4.DisplayText = "Documentos enviado";
			valueListItem5.DataValue = 6;
			valueListItem5.DisplayText = "Seguro liquidado";
			valueListItem6.DataValue = 9;
			valueListItem6.DisplayText = "Cerrado";
			this.optProductos.Items.Add(valueListItem1);
			this.optProductos.Items.Add(valueListItem2);
			this.optProductos.Items.Add(valueListItem3);
			this.optProductos.Items.Add(valueListItem4);
			this.optProductos.Items.Add(valueListItem5);
			this.optProductos.Items.Add(valueListItem6);
			this.optProductos.ItemSpacingVertical = 10;
			this.optProductos.Location = new System.Drawing.Point(10, 65);
			this.optProductos.Name = "optProductos";
			this.optProductos.Size = new System.Drawing.Size(700, 27);
			this.optProductos.TabIndex = 999;
			this.optProductos.ValueChanged += new System.EventHandler(this.optProductos_ValueChanged);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(499, 28);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 24);
			this.lblContador.TabIndex = 998;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(490, 102);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(96, 27);
			this.btnVer.TabIndex = 997;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtDesde
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance4;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(86, 28);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(135, 24);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 993;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// dtHasta
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance5;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(326, 28);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(135, 24);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 994;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.btnFin);
			this.groupBox2.Controls.Add(this.txtPagina);
			this.groupBox2.Controls.Add(this.btnAnterior);
			this.groupBox2.Controls.Add(this.btnInicio);
			this.groupBox2.Controls.Add(this.btnSiguiente);
			this.groupBox2.Controls.Add(this.uGridDocumentos);
			this.groupBox2.Location = new System.Drawing.Point(19, 194);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1267, 341);
			this.groupBox2.TabIndex = 1036;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Detalle";
			// 
			// btnFin
			// 
			this.btnFin.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnFin.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFin.Image = ((System.Drawing.Image)(resources.GetObject("btnFin.Image")));
			this.btnFin.Location = new System.Drawing.Point(192, 305);
			this.btnFin.Name = "btnFin";
			this.btnFin.Size = new System.Drawing.Size(29, 18);
			this.btnFin.TabIndex = 1058;
			this.btnFin.Click += new System.EventHandler(this.btnFin_Click);
			// 
			// txtPagina
			// 
			this.txtPagina.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			this.txtPagina.Appearance = appearance6;
			this.txtPagina.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPagina.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPagina.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPagina.Enabled = false;
			this.txtPagina.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPagina.Location = new System.Drawing.Point(86, 305);
			this.txtPagina.Name = "txtPagina";
			this.txtPagina.ReadOnly = true;
			this.txtPagina.Size = new System.Drawing.Size(68, 23);
			this.txtPagina.TabIndex = 1056;
			// 
			// btnAnterior
			// 
			this.btnAnterior.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnterior.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAnterior.Image = ((System.Drawing.Image)(resources.GetObject("btnAnterior.Image")));
			this.btnAnterior.Location = new System.Drawing.Point(48, 305);
			this.btnAnterior.Name = "btnAnterior";
			this.btnAnterior.Size = new System.Drawing.Size(29, 18);
			this.btnAnterior.TabIndex = 1055;
			this.btnAnterior.Click += new System.EventHandler(this.btnAnterior_Click);
			// 
			// btnInicio
			// 
			this.btnInicio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnInicio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnInicio.Image = ((System.Drawing.Image)(resources.GetObject("btnInicio.Image")));
			this.btnInicio.Location = new System.Drawing.Point(19, 305);
			this.btnInicio.Name = "btnInicio";
			this.btnInicio.Size = new System.Drawing.Size(29, 18);
			this.btnInicio.TabIndex = 1054;
			this.btnInicio.Click += new System.EventHandler(this.btnInicio_Click);
			// 
			// btnSiguiente
			// 
			this.btnSiguiente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSiguiente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnSiguiente.Image = ((System.Drawing.Image)(resources.GetObject("btnSiguiente.Image")));
			this.btnSiguiente.Location = new System.Drawing.Point(163, 305);
			this.btnSiguiente.Name = "btnSiguiente";
			this.btnSiguiente.Size = new System.Drawing.Size(29, 18);
			this.btnSiguiente.TabIndex = 1057;
			this.btnSiguiente.Click += new System.EventHandler(this.btnSiguiente_Click);
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource1;
			appearance7.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance7;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 53;
			ultraGridColumn2.Width = 8;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance8;
			ultraGridColumn4.Header.Caption = "Cédula";
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 180;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Dirección";
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Width = 160;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.Caption = "F. Factura";
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Width = 80;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance9;
			ultraGridColumn8.Header.Caption = "Número";
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.Width = 90;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.Caption = "Local";
			ultraGridColumn9.Header.VisiblePosition = 10;
			ultraGridColumn9.Width = 120;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.Caption = "Forma de pago";
			ultraGridColumn10.Header.VisiblePosition = 11;
			ultraGridColumn10.Width = 140;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 12;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Código";
			ultraGridColumn12.Header.VisiblePosition = 13;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "Artículo";
			ultraGridColumn13.Header.VisiblePosition = 14;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance10;
			ultraGridColumn14.Header.Caption = "Línea";
			ultraGridColumn14.Header.VisiblePosition = 15;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance11;
			ultraGridColumn15.Header.VisiblePosition = 16;
			ultraGridColumn15.Width = 120;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance12;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Precio contado ($)";
			ultraGridColumn16.Header.VisiblePosition = 17;
			ultraGridColumn16.Width = 120;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance13;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Costo  ($)";
			ultraGridColumn17.Header.VisiblePosition = 18;
			ultraGridColumn17.Width = 90;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance14;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Valor  ($)";
			ultraGridColumn18.Header.VisiblePosition = 19;
			ultraGridColumn18.Width = 90;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance15;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Valor seguro  ($)";
			ultraGridColumn19.Header.VisiblePosition = 20;
			ultraGridColumn19.Width = 90;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn20.CellAppearance = appearance16;
			ultraGridColumn20.Header.VisiblePosition = 22;
			ultraGridColumn20.Width = 160;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn21.Header.Caption = "Estado seguro";
			ultraGridColumn21.Header.VisiblePosition = 23;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn22.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn22.Header.Caption = "F. Notificación";
			ultraGridColumn22.Header.VisiblePosition = 28;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn23.Header.Caption = "Tipo notificación";
			ultraGridColumn23.Header.VisiblePosition = 24;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellAppearance = appearance17;
			ultraGridColumn24.Header.VisiblePosition = 25;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn25.Header.VisiblePosition = 32;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance18;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.Caption = "Valor asegurado  ($)";
			ultraGridColumn26.Header.VisiblePosition = 21;
			ultraGridColumn26.Width = 118;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn27.Header.Caption = "Tipo cobertura";
			ultraGridColumn27.Header.VisiblePosition = 27;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn28.Header.VisiblePosition = 29;
			ultraGridColumn28.Width = 120;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance19;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.Caption = "Prof: serv. técnico";
			ultraGridColumn29.Header.VisiblePosition = 30;
			ultraGridColumn29.Width = 140;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn30.CellAppearance = appearance20;
			ultraGridColumn30.Header.VisiblePosition = 26;
			ultraGridColumn30.Width = 60;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn31.Header.VisiblePosition = 33;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn32.Header.VisiblePosition = 34;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn33.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn33.Header.Caption = "Fecha de cierre";
			ultraGridColumn33.Header.VisiblePosition = 31;
			ultraGridColumn33.Width = 155;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn34.Header.VisiblePosition = 35;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Header.VisiblePosition = 51;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 36;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 37;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.Caption = "Valor doc. enviado ($)";
			ultraGridColumn38.Header.VisiblePosition = 38;
			ultraGridColumn38.Width = 146;
			ultraGridColumn39.AutoEdit = false;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.Caption = "Valor de cobro";
			ultraGridColumn39.Header.VisiblePosition = 39;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.AutoEdit = false;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn40.Header.Caption = "Código reemplazo";
			ultraGridColumn40.Header.VisiblePosition = 48;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 8;
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
			ultraGridColumn49.Header.VisiblePosition = 49;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 50;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 52;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance21.Image = ((object)(resources.GetObject("appearance21.Image")));
			appearance21.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn52.CellButtonAppearance = appearance21;
			ultraGridColumn52.Header.Caption = "...";
			ultraGridColumn52.Header.VisiblePosition = 3;
			ultraGridColumn52.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn52.Width = 30;
			ultraGridColumn53.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance22.Image = ((object)(resources.GetObject("appearance22.Image")));
			appearance22.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance22.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn53.CellButtonAppearance = appearance22;
			ultraGridColumn53.Header.Caption = "Proforma";
			ultraGridColumn53.Header.VisiblePosition = 4;
			ultraGridColumn53.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn53.Width = 60;
			ultraGridColumn54.AutoEdit = false;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn54.Header.Caption = "N°";
			ultraGridColumn54.Header.VisiblePosition = 0;
			ultraGridColumn54.Width = 28;
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
																										 ultraGridColumn54});
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance23;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance24;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance25;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance26;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4});
			ultraGridBand1.SummaryFooterCaption = "Sumatoria:";
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Hidden = true;
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance27;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance28.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance28;
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.CellAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance30.FontData.BoldAsString = "True";
			appearance30.FontData.Name = "Arial";
			appearance30.FontData.SizeInPoints = 8F;
			appearance30.ForeColor = System.Drawing.Color.White;
			appearance30.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance30;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance31.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance31.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance33;
			this.uGridDocumentos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDocumentos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(19, 28);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(1229, 267);
			this.uGridDocumentos.TabIndex = 980;
			this.uGridDocumentos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_ClickCellButton);
			this.uGridDocumentos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDocumentos_InitializeLayout_1);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1181, 92);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(105, 27);
			this.btnSalir.TabIndex = 1037;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn52.DataType = typeof(int);
			ultraDataColumn53.DataType = typeof(int);
			ultraDataColumn54.DataType = typeof(int);
			ultraDataColumn58.DataType = typeof(System.DateTime);
			ultraDataColumn67.DataType = typeof(System.Decimal);
			ultraDataColumn68.DataType = typeof(System.Decimal);
			ultraDataColumn69.DataType = typeof(System.Decimal);
			ultraDataColumn70.DataType = typeof(System.Decimal);
			ultraDataColumn72.DataType = typeof(int);
			ultraDataColumn73.DataType = typeof(System.DateTime);
			ultraDataColumn74.DataType = typeof(int);
			ultraDataColumn76.DataType = typeof(int);
			ultraDataColumn77.DataType = typeof(System.Decimal);
			ultraDataColumn78.DataType = typeof(bool);
			ultraDataColumn80.DataType = typeof(System.Decimal);
			ultraDataColumn82.DataType = typeof(int);
			ultraDataColumn83.DataType = typeof(int);
			ultraDataColumn84.DataType = typeof(System.DateTime);
			ultraDataColumn85.DataType = typeof(int);
			ultraDataColumn86.DataType = typeof(int);
			ultraDataColumn88.DataType = typeof(bool);
			ultraDataColumn88.DefaultValue = false;
			ultraDataColumn89.DataType = typeof(System.Decimal);
			ultraDataColumn89.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn90.DataType = typeof(System.Decimal);
			ultraDataColumn92.DataType = typeof(int);
			ultraDataColumn93.DataType = typeof(int);
			ultraDataColumn94.DataType = typeof(int);
			ultraDataColumn95.DataType = typeof(System.Decimal);
			ultraDataColumn96.DataType = typeof(System.Decimal);
			ultraDataColumn97.DataType = typeof(int);
			ultraDataColumn98.DataType = typeof(int);
			ultraDataColumn99.DataType = typeof(bool);
			ultraDataColumn100.DataType = typeof(bool);
			ultraDataColumn101.DataType = typeof(System.DateTime);
			ultraDataColumn102.DataType = typeof(bool);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54,
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58,
																																 ultraDataColumn59,
																																 ultraDataColumn60,
																																 ultraDataColumn61,
																																 ultraDataColumn62,
																																 ultraDataColumn63,
																																 ultraDataColumn64,
																																 ultraDataColumn65,
																																 ultraDataColumn66,
																																 ultraDataColumn67,
																																 ultraDataColumn68,
																																 ultraDataColumn69,
																																 ultraDataColumn70,
																																 ultraDataColumn71,
																																 ultraDataColumn72,
																																 ultraDataColumn73,
																																 ultraDataColumn74,
																																 ultraDataColumn75,
																																 ultraDataColumn76,
																																 ultraDataColumn77,
																																 ultraDataColumn78,
																																 ultraDataColumn79,
																																 ultraDataColumn80,
																																 ultraDataColumn81,
																																 ultraDataColumn82,
																																 ultraDataColumn83,
																																 ultraDataColumn84,
																																 ultraDataColumn85,
																																 ultraDataColumn86,
																																 ultraDataColumn87,
																																 ultraDataColumn88,
																																 ultraDataColumn89,
																																 ultraDataColumn90,
																																 ultraDataColumn91,
																																 ultraDataColumn92,
																																 ultraDataColumn93,
																																 ultraDataColumn94,
																																 ultraDataColumn95,
																																 ultraDataColumn96,
																																 ultraDataColumn97,
																																 ultraDataColumn98,
																																 ultraDataColumn99,
																																 ultraDataColumn100,
																																 ultraDataColumn101,
																																 ultraDataColumn102});
			// 
			// frmReporteCompraProtegida
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(1305, 560);
			this.Controls.Add(this.btnDocumentoEnviado);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.gbBotones);
			this.Name = "frmReporteCompraProtegida";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Administración compra protegida";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteCompraProtegida_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteCompraProtegida_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optProductos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtPagina)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		string sDirectorio = @"\\192.168.1.10\Proyectos\Compra_Protegida"; //c:\Compra_Protegida\";

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmReporteCompraProtegida_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0231");

			if (!Funcion.Permiso("1510", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Administracion Compra Protegida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExcel.Enabled =  true;
			
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse(FechaDesde);
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse(FechaDesde);

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;			
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = DateTime.Parse(FechaDesde);//"01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;
		}

		private void Consulta()
		{
			//Consulta el tamaño antes de la consulta
			iRegistrosRequerimiento(Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 
				(int)this.optProductos.Value, this.txtBusqueda.Text.ToString());

			string sSQLPa = string.Format("Exec AdministracionCompraProtegida '{0}', '{1}', {2}, '{3}' , {4}, {5}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 
				(int)this.optProductos.Value, this.txtBusqueda.Text.ToString(),Pagina,TamanoPagina);
			FuncionesProcedimientos.dsGeneralModulo(sSQLPa, this.uGridDocumentos);

			//this.lblContador.Text = this.uGridDocumentos.Rows.Count + " REGISTROS ENCONTRADOS";

			UltraGridColumn column	= this.uGridDocumentos.DisplayLayout.Bands[0].Columns["btnGestion"];
			
			//Si es con estado 5 ya no me permita dar al mas, y en los cerrados
			if((int)this.optProductos.Value == iEstadoSeguroDocumentoEnviado || (int)this.optProductos.Value == iEstadoSeguroCerrado)
			{
				column.CellActivation = Activation.Disabled;
			}
			else
			{
				column.CellActivation = Activation.AllowEdit;
			}

			//Contador de filas
			this.uGridDocumentos = Funcion.ContadorFilas(this.uGridDocumentos, "Row");

		}

		public void iRegistrosRequerimiento(string dtDesde ,string dtHasta, int optProductos,string txtBusqueda)
		{
			int iRegistros = 0;
			string squery = "";	
			try
			{
				squery = String.Format("EXEC AdministracionCompraProtegidaCount '{0}', '{1}', {2}, '{3}'", 
					dtDesde,
					dtHasta,
					optProductos,
					txtBusqueda
					);
				iRegistros = Funcion.iEscalarSQL(cdsSeteoF, squery,false);// this.udgSeguimiento.Rows.Count;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
			string sRegistros = "";
	
			if(iRegistros == 1)
			{
				sRegistros = "1 REGISTRO ENCONTRADO";
			}
			else if(iRegistros == -1)
			{
				sRegistros = "0 REGISTROS ENCONTRADOS";
			}
			else
			{
				sRegistros = iRegistros.ToString() + " REGISTROS ENCONTRADOS";
			}
			this.lblContador.Text = sRegistros;
			CalculoPaginas(iRegistros);
		}

		private void CalculoPaginas(int iRegistros)
		{
			
			TotalPagina = Convert.ToInt32(iRegistros/TamanoPagina) ;
			decimal dTotalPagina = Convert.ToDecimal(iRegistros) / Convert.ToDecimal(TamanoPagina);

			if (dTotalPagina > TotalPagina)
			{
				TotalPagina = TotalPagina + 1;
			}
			//Validar si iRegistros es igual a 0, darle valor a 1 a total de la página
			if (iRegistros == 0)
			{
				TotalPagina = 1;
				Pagina = 1;
			}

			MostrarPaginaActual();
		}

		public void iRegistrosRequerimiento(int idResponsable,int idAsignado ,string dtDate,string dtHasta ,int idTipoRequerimiento,
			string Numero,bool bAsignacionesMias,int idEstado)
		{
			int iRegistros = 0;
			string squery = "";	
			try
			{
				squery = String.Format("EXEC RequerimientoConCount {0},{1},'{2}','{3}',{4},'{5}',{6},{7},'{8}'", 
					idResponsable,
					idAsignado,
					dtDate,
					dtHasta,
					idTipoRequerimiento,
					Numero,
					bAsignacionesMias,
					idEstado,
					MenuLatinium.stDirFacturacion
					);
				iRegistros = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, squery,false);// this.udgSeguimiento.Rows.Count;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
			string sRegistros = "";
	
			if(iRegistros == 1)
			{
				sRegistros = "1 REGISTRO";
			}
			else
			{
				sRegistros = iRegistros.ToString() + " REGISTROS";
			}
			this.lblContador.Text = sRegistros;
			CalculoPaginas(iRegistros);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{			
			#region Valida Fecha
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}
     			
			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}
			#endregion Valida Fecha

			if (this.optProductos.CheckedIndex == - 1)
			{
				MessageBox.Show("Seleccione una opción.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.optProductos.Focus();
				return;
			}

			SettearPaginacion();
			Consulta();	
		}

		private void SettearPaginacion()
		{
			Pagina = 1;
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			//Cargar todos los datos para la exportación
			ConsultarDatosCompletos();
			FuncionesProcedimientos.ExportaExcel(this.uGridDocumentos);
		}

		private void ConsultarDatosCompletos()
		{
			string sSQLPa = string.Format("Exec AdministracionCompraProtegida '{0}', '{1}', {2}, '{3}', {4}, {5}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), 
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 
				(int)this.optProductos.Value, 
				this.txtBusqueda.Text,
				1,
				0); //this.txtBusqueda.Text.ToString()
			FuncionesProcedimientos.dsGeneralModulo(sSQLPa, this.uGridDocumentos);
			//Label de registros
			this.lblContador.Text = this.uGridDocumentos.Rows.Count + " REGISTROS ENCONTRADOS";
			//COntador de filas
			this.uGridDocumentos = Funcion.ContadorFilas(uGridDocumentos, "Row");
		}

		private void frmReporteCompraProtegida_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridDocumentos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void Imprimir(int idProforma, string sCarpeta)
		{			
			ParameterFields paramFields = new ParameterFields();
			ParameterField paramFieldCedula = new ParameterField ();
			ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
			paramFieldCedula.ParameterFieldName = "@idcompra";
			ValCedula.Value = idProforma;
			paramFieldCedula.CurrentValues.Add (ValCedula);
			paramFields.Add (paramFieldCedula);
			Reporte Reporte = new Reporte("ReporteProformaSeguro.Rpt");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
			

			//Reporte miRep = new Reporte("ReporteProformaSeguro.rpt", "");			
												
			//			using (Reporte Reporte = new Reporte("ReporteProformaSeguro.Rpt"))			
			//			{
				
			//Reporte.ExportaPDF(sCarpeta + "proforma.pdf");				
			//			}
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

		public string invertirCadena(String cadena)
		{
			String cadenaInvertida = "";
			for (int x = cadena.Length - 1; x >= 0; x--)
				cadenaInvertida = cadenaInvertida + cadena.Substring(x, 1);

			return cadenaInvertida;
		}

		public int getIdSerialFrm(int getidCompra)
		{
			int getIdSerial = 0;
			using (frmReporteCompraProtegidaSerial frcps = new frmReporteCompraProtegidaSerial(getidCompra))
			{
				if (DialogResult.OK == frcps.ShowDialog())
				{	
					getIdSerial = frmReporteCompraProtegidaSerial.idSerial;
				}
				else
				{
					getIdSerial = 0;
				}
			}
			return getIdSerial;
		}

		private void uGridDocumentos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Botón gestión
			if (e.Cell.Column.ToString().Equals("btnGestion"))
			{
				#region Variables
				int iFacturaAseguradora = 0;
				int iCobroAseguradora = 0;
				#endregion Variables

				#region Notificacion
				if ((int)e.Cell.Row.Cells["EstadoSeguro"].Value == 1)
				{
					bool bNotificado = false;

					using (frmTipoReclamoCompraProtegida TRCP = new frmTipoReclamoCompraProtegida())
					{ 
						if (DialogResult.OK == TRCP.ShowDialog( ))
						{
							string sCarpeta = sDirectorio + @"\" + 
								e.Cell.Row.Cells["idDetCompra"].Value.ToString() + "_" +
								e.Cell.Row.Cells["Cedula"].Value.ToString() + "_" +
								e.Cell.Row.Cells["Numero"].Value.ToString() + @"\";

							DirectoryInfo DirInfo = Directory.CreateDirectory(sCarpeta);

							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetCompra Set EstadoSeguro = 3, FechaNotificacion = GETDATE(), FechaNotificacionCliente = '{3}', FechaSiniestro = '{4}', TipoNotificacion = {0}, DirectorioCP = '{2}' Where idDetCompra = {1}", 
								TRCP.iTipo, (int)e.Cell.Row.Cells["idDetCompra"].Value, sCarpeta, Convert.ToDateTime(TRCP.dtNotificacion.Value).ToString("yyyyMMdd"), Convert.ToDateTime(TRCP.dtSiniestro.Value).ToString("yyyyMMdd")));

							bNotificado = true;
						}					
					}

					#region Proforma
					if (bNotificado)
					{
						using (frmProformaSeguro PS = new frmProformaSeguro((int)e.Cell.Row.Cells["idProforma"].Value, (int)e.Cell.Row.Cells["idDetCompra"].Value))
						{ 
							if (DialogResult.OK == PS.ShowDialog())
							{
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetCompra Set idProforma = {0} Where idDetCompra = {1}", 
									(int)PS.txtIdCompra.Value, (int)e.Cell.Row.Cells["idDetCompra"].Value));

								MessageBox.Show("Se ha generado la notificacion de reclamo de compra protegida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

								Imprimir((int)PS.txtIdCompra.Value, "");
							}
						}
					}	
					#endregion Proforma							

					Consulta();

					return;
				}
				#endregion Notificacion

				#region Documentacion
				if ((int)e.Cell.Row.Cells["EstadoSeguro"].Value == 3)
				{
					if ((int)e.Cell.Row.Cells["idProforma"].Value == 0)
					{
						MessageBox.Show("Debe generar la proforma para continuar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						return;
					}

					using (frmDocumentacionCompraProtegida DCP = new frmDocumentacionCompraProtegida((int)e.Cell.Row.Cells["TipoNotificacion"].Value, (int)e.Cell.Row.Cells["idDetCompra"].Value, e.Cell.Row.Cells["DirectorioCP"].Value.ToString()))
					{ 
						if (DialogResult.OK == DCP.ShowDialog())
						{	
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec GuardaDocsCompraProtegida {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, '{13}', '{14}', '{15}', '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', '{22}', '{23}', {24}, {25}", 
								(int)e.Cell.Row.Cells["idDetCompra"].Value, DCP.iEstado, 
								DCP.chkFormulario.Checked, DCP.chkCertificado.Checked, DCP.chkCedula.Checked, DCP.chkFactura.Checked, DCP.chkProforma.Checked, DCP.chkDenuncia.Checked,
								DCP.chkCBIMEI.Checked, DCP.chkInformeTecnico.Checked, DCP.chkPST.Checked, DCP.chkCartaAseguradora.Checked, DCP.chkCartaBroker.Checked,
								DCP.sFormulario, DCP.sCertificado, DCP.sCedula, DCP.sFactura, DCP.sProforma,
								DCP.sDenuncia, DCP.sIMEI, DCP.sInformeTecnico, DCP.sPST, DCP.sAseguradora, DCP.sBroker, DCP.ChkCobroAseguradora.Checked, DCP.chkFacturaAseguradora.Checked));
		
							string sMensaje = "Documentacion Completa";//\n\nSe enviara al seguro via correo electronico";
							if (DCP.iEstado == 3) sMensaje = "Documentacion Incompleta";

							MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);							

							Consulta();

							return;
						}
					}					
				}
				#endregion Documentacion

				#region Seguro liquidado

				#region Facturación Deducible NO
				if ((int)e.Cell.Row.Cells["EstadoSeguro"].Value == iEstadoSeguroSeguroLiquidado && e.Cell.Row.Cells["Deducible"].Value.ToString() == "NO")
				{
					int idCompra = (int)e.Cell.Row.Cells["idCompra"].Value;
					int idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
					int idCompraNuevo = 0;
					int idProforma = (int)e.Cell.Row.Cells["idProforma"].Value;
					string Deducible = e.Cell.Row.Cells["Deducible"].Value.ToString();
					int TipoNotificacion = (int)e.Cell.Row.Cells["TipoNotificacion"].Value;
					decimal Valor_Asegurado = Convert.ToDecimal(e.Cell.Row.Cells["Valor_Asegurado"].Value);
					string sNumero = "";

					//Abrir el registro de depósito cuando el idCOmpraP, sea mayor a 0, ESTE CAMPO MUESTRA QUE EXISTE UNA FACTURA, YA GENERADA, 
					//lo que nos faltaría solo sería registrar el depósito, como un anticipo, en los casos que si tenga deducible, en los otros 
					//casos se cambiaría directamente al estado, cerrado, todos deberían ir a eset estado, una vez completado su proceso.

					// Tiene una factura anterior a la fecha 19/04/3023
					bool bTieneFacturaAnterior = (bool)e.Cell.Row.Cells["bTieneFacturaAnterior"].Value;
					// Mostrar el valor de la aseguradora, revisando si existe un deoósito anterior al 19/04/3023
					bool bTieneDepositoAnterior = (bool)e.Cell.Row.Cells["bTieneDepositoAnterior"].Value;
					//Control que determina que tiene desarrollo actual
					bool bTieneDesarrolloActual  = (bool)e.Cell.Row.Cells["bTieneDesarrolloActual"].Value;
					//Proceso Nuevo
					bool bProcesoActual = false;

					if(bTieneDesarrolloActual == false && bTieneDepositoAnterior == true)
					{
						//Actualización
						int idAnticipo  = (int)e.Cell.Row.Cells["idAnticipo"].Value;
						string squery = String.Format("Exec ReporteCompraProtegidaValorDocumentoEnviadoAct {0},{1}", idDetCompra, idAnticipo);
						Funcion.EjecutaSQL(cdsSeteoF, squery);
					}

					if(bTieneDepositoAnterior == true || bTieneDesarrolloActual == true )
					{
						bProcesoActual = true;
					}

					if(bProcesoActual)
					{
						#region POR ROBO
						if ((int)e.Cell.Row.Cells["TipoNotificacion"].Value == iRobo)
						{      
							//Generar factura, solo que en el procedimiento se debe excluir el aumento de la forma de pago del deducible.
							idCompraNuevo = CreaFacturaCompraProtegidaDeducible(idCompra, idDetCompra, idProforma,
								Valor_Asegurado, TipoNotificacion, Deducible, "", "");

							using (frmReporteCompraProtegidaSeguroLiquidadoFactura FRMRCPF = new frmReporteCompraProtegidaSeguroLiquidadoFactura(idCompraNuevo, iEstadoSeguroSeguroLiquidado, idDetCompra))
							{ 
								if (DialogResult.OK == FRMRCPF.ShowDialog())
								{
									//Facturación electrónica
									FacturacionElectronicaSeguroLiquidado(idCompraNuevo);
									//Guardar depósito y cambiar de estado a cerrado
									int getcmbCuentas = 0;
									try{getcmbCuentas = (int)FRMRCPF.cmbCuentas.Value;}
									catch(Exception ex){string sms = ex.Message;}
									GuardarDepositoSeguroLiquidado(idCompraNuevo, getcmbCuentas, FRMRCPF.txtCheque.Text.ToString(), 
										sNumero, FRMRCPF.txtNotasDeposito.Text.ToString(), idDetCompra, false, iCompraFormaPagoDepositoTransferencia);
									//Mensaje de éxito al guardar el seguro liquidado
									MensajeExitoSeguroLiquidado();
								}
							}
						}
						#endregion POR ROBO
						#region POR DAÑO
						if ((int)e.Cell.Row.Cells["TipoNotificacion"].Value == iDano)
						{
							#region Perdida total
							//Valor del rpoducto se lo saca del producto asegurado.
							decimal dValorFacturar = (decimal)e.Cell.Row.Cells["Valor_Asegurado"].Value;
							//Generar factura, solo que en el procedimiento se debe excluir el aumento de la forma de pago del deducible.
							idCompraNuevo = CreaFacturaCompraProtegidaDeducible(idCompra, idDetCompra, idProforma,
								dValorFacturar, TipoNotificacion, Deducible, "", "");

							using (frmReporteCompraProtegidaSeguroLiquidadoFactura FRMRCPF = new frmReporteCompraProtegidaSeguroLiquidadoFactura(idCompraNuevo, iEstadoSeguroSeguroLiquidado, idDetCompra))
							{ 
								if (DialogResult.OK == FRMRCPF.ShowDialog())
								{
									//Facturación electrónica
									FacturacionElectronicaSeguroLiquidado(idCompraNuevo);
									//Guardar depósito y cambiar de estado a cerrado
									int getcmbCuentas = 0;
									try{getcmbCuentas = (int)FRMRCPF.cmbCuentas.Value;}
									catch(Exception ex){string sms = ex.Message;}
									GuardarDepositoSeguroLiquidado(idCompraNuevo, getcmbCuentas, FRMRCPF.txtCheque.Text.ToString(), 
										sNumero, FRMRCPF.txtNotasDeposito.Text.ToString(), idDetCompra, false, iCompraFormaPagoDepositoTransferencia);
									//Mensaje de éxito al guardar el seguro liquidado
									MensajeExitoSeguroLiquidado();
								}
							}
							#endregion Perdida total
						}
						#endregion POR DAÑO
					}

					if(bTieneFacturaAnterior)
					{
						//Cerrar el estado
						string squery = string.Format("Update DetCompra Set EstadoSeguro = {0} , FechaCierre = GETDATE() Where idDetCompra = {1}", 
							iEstadoSeguroCerrado, (int)e.Cell.Row.Cells["idDetCompra"].Value); 
						Funcion.EjecutaSQL(cdsSeteoF, squery);
						//Éxito
						MessageBox.Show("El registro se cerró correctamente, ya que no existe un depósito que registrar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}

					Consulta();
					return;
				}
			
					
				#endregion Facturación Deducible NO
				
				#region Facturacion Deducible SI
				if ((int)e.Cell.Row.Cells["EstadoSeguro"].Value == iEstadoSeguroSeguroLiquidado && e.Cell.Row.Cells["Deducible"].Value.ToString() == "SI")
				{
					int idCompra = (int)e.Cell.Row.Cells["idCompra"].Value;
					int idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
					int idCompraNuevo = 0;
					int idProforma = (int)e.Cell.Row.Cells["idProforma"].Value;
					string Deducible = e.Cell.Row.Cells["Deducible"].Value.ToString();
					int TipoNotificacion = (int)e.Cell.Row.Cells["TipoNotificacion"].Value;
					decimal Valor_Asegurado = Convert.ToDecimal(e.Cell.Row.Cells["Valor_Asegurado"].Value);
					string sNumero = "";

					//Abrir el registro de depósito cuando el idCOmpraP, sea mayor a 0, ESTE CAMPO MUESTRA QUE EXISTE UNA FACTURA, YA GENERADA, 
					//lo que nos faltaría solo sería registrar el depósito, como un anticipo, en los casos que si tenga deducible, en los otros 
					//casos se cambiaría directamente al estado, cerrado, todos deberían ir a eset estado, una vez completado su proceso.

					// Tiene una factura anterior a la fecha 19/04/3023
					bool bTieneFacturaAnterior = (bool)e.Cell.Row.Cells["bTieneFacturaAnterior"].Value;
					// Mostrar el valor de la aseguradora, revisando si existe un deoósito anterior al 19/04/3023
					bool bTieneDepositoAnterior = (bool)e.Cell.Row.Cells["bTieneDepositoAnterior"].Value;
					//Control que determina que tiene desarrollo actual
					bool bTieneDesarrolloActual  = (bool)e.Cell.Row.Cells["bTieneDesarrolloActual"].Value;
					//Proceso Nuevo
					bool bProcesoActual = false;

					if(bTieneDesarrolloActual == false && bTieneDepositoAnterior == true)
					{
						//Actualización
						int idAnticipo  = (int)e.Cell.Row.Cells["idAnticipo"].Value;
						string squery = String.Format("Exec ReporteCompraProtegidaValorDocumentoEnviadoAct {0},{1}", idDetCompra, idAnticipo);
						Funcion.EjecutaSQL(cdsSeteoF, squery);
					}

					if(bTieneDepositoAnterior == true || bTieneDesarrolloActual == true )
					{
						bProcesoActual = true;
					}

					if(bProcesoActual)
					{
						#region POR ROBO
						if (TipoNotificacion == iRobo)
						{   
							using (frmTipoCoberturaCP TC = new frmTipoCoberturaCP())
							{ 
								if (DialogResult.OK == TC.ShowDialog( ))
								{
									#region Perdida total
									if(TC.bTipo == false)
									{
										decimal dValorFacturar = Convert.ToDecimal(TC.txtValorProforma.Value);
										//Generar factura
										idCompraNuevo = CreaFacturaCompraProtegidaDeducible(idCompra, idDetCompra, idProforma,
											dValorFacturar, TipoNotificacion, Deducible, "", "");

										using (frmReporteCompraProtegidaSeguroLiquidadoFactura FRMRCPF = new frmReporteCompraProtegidaSeguroLiquidadoFactura(idCompraNuevo, iEstadoSeguroSeguroLiquidado, idDetCompra))
										{ 
											if (DialogResult.OK == FRMRCPF.ShowDialog())
											{
												//Facturación electrónica
												FacturacionElectronicaSeguroLiquidado(idCompraNuevo);
												//Guardar depósito y cambiar de estado a cerrado
												int getcmbCuentas = 0;
												try{getcmbCuentas = (int)FRMRCPF.cmbCuentas.Value;}
												catch(Exception ex){string sms = ex.Message;}
												GuardarDepositoSeguroLiquidado(idCompraNuevo, getcmbCuentas, FRMRCPF.txtCheque.Text.ToString(), 
													sNumero, FRMRCPF.txtNotasDeposito.Text.ToString(), idDetCompra, true, iCompraFormaPagoDepositoTransferencia);
												//Mensaje de éxito al guardar el seguro liquidado
												MensajeExitoSeguroLiquidado();
											}
										}
									}
									else
									{
										MessageBox.Show("No existe reparación, en los casos de robo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
									}
									#endregion Perdida total
								}
							}
						}
						#endregion POR ROBO
						#region POR DAÑO
						if (TipoNotificacion == iDano)
						{
							using (frmTipoCoberturaCP TC = new frmTipoCoberturaCP())
							{ 
								if (DialogResult.OK == TC.ShowDialog( ))
								{
									string sNumeroD = "";
									string sNumeroP = "";

									#region Perdida total
									if(TC.bTipo == false)
									{
										decimal dValorFacturar = Convert.ToDecimal(TC.txtValorProforma.Value);
										//decimal dValorFacturar = (decimal)e.Cell.Row.Cells["Valor_Asegurado"].Value;
										//Generar factura
										idCompraNuevo = CreaFacturaCompraProtegidaDeducible((int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idDetCompra"].Value, (int)e.Cell.Row.Cells["idProforma"].Value,
											dValorFacturar, (int)e.Cell.Row.Cells["TipoNotificacion"].Value, 
											e.Cell.Row.Cells["Deducible"].Value.ToString(), "", "");

										using (frmReporteCompraProtegidaSeguroLiquidadoFactura FRMRCPF = new frmReporteCompraProtegidaSeguroLiquidadoFactura(idCompraNuevo, iEstadoSeguroSeguroLiquidado, (int)e.Cell.Row.Cells["idDetCompra"].Value))
										{ 
											if (DialogResult.OK == FRMRCPF.ShowDialog())
											{
												//Facturación electrónica
												FacturacionElectronicaSeguroLiquidado(idCompraNuevo);
												//Guardar depósito y cambiar de estado a cerrado
												int getcmbCuentas = 0;
												try{getcmbCuentas = (int)FRMRCPF.cmbCuentas.Value;}
												catch(Exception ex){string sms = ex.Message;}
												GuardarDepositoSeguroLiquidado(idCompraNuevo, getcmbCuentas, FRMRCPF.txtCheque.Text.ToString(), 
													sNumero, FRMRCPF.txtNotasDeposito.Text.ToString(), idDetCompra, true, iCompraFormaPagoDepositoTransferencia);
												//Mensaje de éxito al guardar el seguro liquidado
												MensajeExitoSeguroLiquidado();
											}
										}
									}
									#endregion Perdida total
									#region Reparación
									if(TC.bTipo == true)
									{
										#region Creación factura producto
										if (e.Cell.Row.Cells["Deducible"].Value.ToString() == "SI")         
											sNumeroD = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 1, 1, 1"));
																					                
										decimal dValorFacturar = Convert.ToDecimal(TC.txtValorProforma.Value);
											
										string sSQL = string.Format("Exec CreaFacturasCompraProtegida {0}, {1}, {2}, {3}, {4}, '{5}', '{6}', '{7}', {8}, 0, 0, 1", 
											(int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idDetCompra"].Value, (int)e.Cell.Row.Cells["idProforma"].Value,
											dValorFacturar, (int)e.Cell.Row.Cells["TipoNotificacion"].Value, 
											e.Cell.Row.Cells["Deducible"].Value.ToString(), sNumeroD, sNumeroP, TC.bTipo);
										idCompra = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
										#endregion Creación factura producto
										#region Formulario factura
										using (frmReporteCompraProtegidaSeguroLiquidadoFactura FRMRCPF2 = new frmReporteCompraProtegidaSeguroLiquidadoFactura(idCompra, iEstadoSeguroSeguroLiquidado, (int)e.Cell.Row.Cells["idDetCompra"].Value))
										{ 
											if (DialogResult.OK == FRMRCPF2.ShowDialog( ))
											{
												//Facturación electrónica
												FacturacionElectronicaSeguroLiquidado(idCompra);
												//Guardar depósito y cambiar de estado a cerrado
												int getcmbCuentas = 0;
												try{getcmbCuentas = (int)FRMRCPF2.cmbCuentas.Value;}
												catch(Exception ex){string sms = ex.Message;}
												GuardarDepositoSeguroLiquidado(idCompra, getcmbCuentas, FRMRCPF2.txtCheque.Text.ToString(), 
													sNumero, FRMRCPF2.txtNotasDeposito.Text.ToString(), idDetCompra, true, iCompraFormaPagoDepositoTransferencia);
												//Mensaje de éxito al guardar el seguro liquidado
												MensajeExitoSeguroLiquidado();							
											}
											#endregion Formulario factura
										}
									}
									#endregion Reparación
								}
							}
						}
						#endregion POR DAÑO
					}
					if(bTieneFacturaAnterior)
					{
						#region Factura anterior
						using (frmTipoCoberturaCP TC = new frmTipoCoberturaCP())
						{ 
							if (DialogResult.OK == TC.ShowDialog( ))
							{
								using (frmSegurosRegistroDeposito Deposito = new frmSegurosRegistroDeposito())
								{
									if (DialogResult.OK == Deposito.ShowDialog())
									{
										//Cuando si existe una factura de por medio, esto puede ser por que el desarrollo se empezo el 15042024
										//Incremental
										sNumero = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 24, 1, 0"));
										//Gurdar depósito
										string sSQLC = string.Format("Exec ReporteCompraDepositoClienteCrear {0}, 26, 24, {1}, '{2}', '{3}', '{4}', {5}, {6}", 
											idCompra, Deposito.cmbCuentas.Value, Deposito.txtCheque.Text.ToString(), 
											sNumero, Deposito.txtNotas.Text.ToString(), (int)e.Cell.Row.Cells["idDetCompra"].Value, Convert.ToDecimal(TC.txtValorProforma.Value));
										int idAnticipo = Funcion.iEscalarSQL(cdsSeteoF, sSQLC);
										//Cerrar el estado
										string squery = string.Format("Update DetCompra Set EstadoSeguro = {0}, idAnticipoD = {1} , FechaCierre = GETDATE() Where idDetCompra = {2}", 
											iEstadoSeguroCerrado , idAnticipo, (int)e.Cell.Row.Cells["idDetCompra"].Value); 
										Funcion.EjecutaSQL(cdsSeteoF, squery);
										//Éxito
										MessageBox.Show("Deposito registrado correctamente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
									}
								}
							}
						}
						#endregion Factura anterior
					}
					Consulta();
					return;
				}
				
				#endregion Facturacion Deducible SI

				#endregion Seguro liquidado
			}
			#endregion Botón gestión
			#region Ver Proforma
			if (e.Cell.Column.ToString().Equals("btnProforma"))
			{				
				if ((int)e.Cell.Row.Cells["EstadoSeguro"].Value > 1)
				{
					using (frmProformaSeguro PS = new frmProformaSeguro((int)e.Cell.Row.Cells["idProforma"].Value, (int)e.Cell.Row.Cells["idDetCompra"].Value))
					{ 
						if (DialogResult.OK == PS.ShowDialog())
						{
							Consulta();		
						}
					}
				}				
			}
			#endregion Ver Proforma
		}

		private void FacturacionElectronicaSeguroLiquidado(int idCompra)
		{
			#region Facturacion Electronica
			if (Funcion.bEscalarSQL(cdsSeteoF, "Select FacturaElectronica From SeteoF", false))
			{
				string sClaveAcceso = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec FE_ClaveDeAcceso {0}, 1", idCompra));
								
				string cadInvertida = invertirCadena(sClaveAcceso.Substring(0, 48));
				string ValM11 = Modulo11(cadInvertida).ToString().Trim();
				string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, idCompra));
			}
			#endregion Facturacion Electronica
		}

		private void GuardarDepositoSeguroLiquidado(int idCompra, int cmbCuentas, string txtCheque, string sNumero, string txtNotasDeposito, int idDetCompra, bool DepositoDeducible, int idFormaPago)
		{
//			if(DepositoDeducible)
//			{
//				string sSQLC = string.Format("Exec ReporteCompraProtegidaFormaPagoAct {0}, {1}, '{2}', '{3}', {4}", 
//					idCompra, cmbCuentas, txtCheque,
//					txtNotasDeposito, idFormaPago);
//				Funcion.EjecutaSQL(cdsSeteoF, sSQLC);
//			}
			string squery = String.Format("Update DetCompra Set EstadoSeguro = {0}, FechaCierre = GETDATE() Where idDetCompra = {1}", 
				iEstadoSeguroCerrado,
				idDetCompra);
			//Cambiar a estado cerrado
			Funcion.EjecutaSQL(cdsSeteoF, squery);
		}
		private void MensajeExitoSeguroLiquidado()
		{
			MessageBox.Show("Depoósito registrado correctamente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private int CreaFacturaCompraProtegidaDeducible(int idCompra, int idDetCompra, int idProforma,
		decimal Valor_Asegurado, int TipoNotificacion, 
		string deducible, string sNumeroD, string sNumeroP)
		{
			string sSQL = string.Format("Exec CreaFacturasCompraProtegida {0}, {1}, {2}, {3}, {4}, '{5}', '{6}', '{7}', 0, 0, 0, 0", 
				idCompra,idDetCompra, idProforma,
				Valor_Asegurado, TipoNotificacion, 
				deducible, sNumeroD, sNumeroP);
			idCompra = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			return idCompra;
		}


		private void optProductos_ValueChanged(object sender, System.EventArgs e)
		{
			SettearPaginacion();
			Consulta();
		}

		private void btnDocumentoEnviado_Click(object sender, System.EventArgs e)
		{
			using (frmReporteCompraProtegidaDocumentoEnviado frmrcpde = new frmReporteCompraProtegidaDocumentoEnviado())
			{		
				if (DialogResult.OK == frmrcpde.ShowDialog())
				{
				}
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnInicio_Click(object sender, System.EventArgs e)
		{
			Pagina = 1;
			MostrarPaginaActual();
			Consulta();
		}

		private void btnAnterior_Click(object sender, System.EventArgs e)
		{
			if(Pagina > 1)
			{
				Pagina = Pagina - 1;
			}
			MostrarPaginaActual();
			Consulta();
		}

		private void MostrarPaginaActual()
		{
			this.txtPagina.Text = Pagina.ToString() + "/" + TotalPagina.ToString();
		}

		private void btnSiguiente_Click(object sender, System.EventArgs e)
		{
			if(Pagina < TotalPagina)
			{
				Pagina = Pagina + 1;
			}
			MostrarPaginaActual();
			Consulta();
		}

		private void btnFin_Click(object sender, System.EventArgs e)
		{
			Pagina = TotalPagina;
			MostrarPaginaActual();
			Consulta();
		}

		private void uGridDocumentos_InitializeLayout_1(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}
	}
}
