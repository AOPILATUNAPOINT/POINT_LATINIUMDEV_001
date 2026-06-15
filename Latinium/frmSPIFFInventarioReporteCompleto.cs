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
	/// Descripción breve de frmSPIFFInventarioReporteCompleto.
	/// </summary>
	public class frmSPIFFInventarioReporteCompleto : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Button btnConsultar;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.ToolTip toolTip1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource UDSSPIFFInventarioReporteCon;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsSPIFFSubGrupoRangoCabList;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbSPIFFSubGrupoRangoCab;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsSPIFFRangoCon;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgSPIFFRangoCOn;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsSPIFFListaVendedorCon;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdSPIFFInventarioReporteCon;
		private System.ComponentModel.IContainer components;

		public frmSPIFFInventarioReporteCompleto()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSPIFFSubGrupoRangoCab");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSPIFFSubGrupoRangoCab");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSPIFFInventarioReporteCompleto));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subgrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Edad_Dias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rango_Dias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Stock");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaMovimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SPIFF");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Personal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasHastaVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SPIFFPorCobrar");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaId");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoItem");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSPIFFSubGrupoRangoCab");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotalFactura");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SPIFF", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SPIFF", 14, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SPIFFPorCobrar", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SPIFFPorCobrar", 21, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subgrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Edad_Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rango_Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Stock");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaMovimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SPIFF");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Personal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasHastaVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SPIFFPorCobrar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BodegaId");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoItem");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSPIFFSubGrupoRangoCab");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotalFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioPromocion");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSPIFFRango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtitulo", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 1);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSPIFFRango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtitulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seleccionar");
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cmbSPIFFSubGrupoRangoCab = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsSPIFFSubGrupoRangoCabList = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label53 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ugdSPIFFInventarioReporteCon = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.UDSSPIFFInventarioReporteCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsSPIFFListaVendedorCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.label12 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.udgSPIFFRangoCOn = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsSPIFFRangoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnSalir = new System.Windows.Forms.Button();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnExcel = new System.Windows.Forms.Button();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbSPIFFSubGrupoRangoCab)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFSubGrupoRangoCabList)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdSPIFFInventarioReporteCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.UDSSPIFFInventarioReporteCon)).BeginInit();
			this.groupBox5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFListaVendedorCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udgSPIFFRangoCOn)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFRangoCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.lblEstado);
			this.groupBox2.Controls.Add(this.cmbSPIFFSubGrupoRangoCab);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Controls.Add(this.lblBodega);
			this.groupBox2.Controls.Add(this.dtHasta);
			this.groupBox2.Controls.Add(this.cmbBodega);
			this.groupBox2.Controls.Add(this.label53);
			this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox2.Location = new System.Drawing.Point(19, 18);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(883, 102);
			this.groupBox2.TabIndex = 993;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(499, 55);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 32);
			this.lblEstado.TabIndex = 1044;
			// 
			// cmbSPIFFSubGrupoRangoCab
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSPIFFSubGrupoRangoCab.Appearance = appearance1;
			this.cmbSPIFFSubGrupoRangoCab.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSPIFFSubGrupoRangoCab.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSPIFFSubGrupoRangoCab.DataSource = this.udsSPIFFSubGrupoRangoCabList;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 331;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSPIFFSubGrupoRangoCab.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSPIFFSubGrupoRangoCab.DisplayMember = "Nombre";
			this.cmbSPIFFSubGrupoRangoCab.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSPIFFSubGrupoRangoCab.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSPIFFSubGrupoRangoCab.Location = new System.Drawing.Point(96, 28);
			this.cmbSPIFFSubGrupoRangoCab.Name = "cmbSPIFFSubGrupoRangoCab";
			this.cmbSPIFFSubGrupoRangoCab.Size = new System.Drawing.Size(384, 25);
			this.cmbSPIFFSubGrupoRangoCab.TabIndex = 1043;
			this.cmbSPIFFSubGrupoRangoCab.ValueMember = "idSPIFFSubGrupoRangoCab";
			this.cmbSPIFFSubGrupoRangoCab.ValueChanged += new System.EventHandler(this.ultraCombo2_ValueChanged);
			this.cmbSPIFFSubGrupoRangoCab.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraCombo2_InitializeLayout);
			// 
			// udsSPIFFSubGrupoRangoCabList
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsSPIFFSubGrupoRangoCabList.Band.Columns.AddRange(new object[] {
																																						 ultraDataColumn1,
																																						 ultraDataColumn2});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(10, 28);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(49, 19);
			this.label1.TabIndex = 1042;
			this.label1.Text = "SPIFF:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(499, 28);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(74, 19);
			this.lblBodega.TabIndex = 993;
			this.lblBodega.Text = "Cortado al:";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblBodega.Click += new System.EventHandler(this.lblBodega_Click);
			// 
			// dtHasta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance2;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(576, 28);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.ReadOnly = true;
			this.dtHasta.Size = new System.Drawing.Size(288, 25);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 989;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 180;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(96, 65);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(384, 25);
			this.cmbBodega.TabIndex = 1064;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_TextChanged);
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(10, 65);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(43, 19);
			this.label53.TabIndex = 1063;
			this.label53.Text = "Local:";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.ugdSPIFFInventarioReporteCon);
			this.groupBox1.Controls.Add(this.groupBox5);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(19, 129);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1027, 416);
			this.groupBox1.TabIndex = 994;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Campañas";
			// 
			// ugdSPIFFInventarioReporteCon
			// 
			this.ugdSPIFFInventarioReporteCon.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdSPIFFInventarioReporteCon.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdSPIFFInventarioReporteCon.DataSource = this.UDSSPIFFInventarioReporteCon;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ugdSPIFFInventarioReporteCon.DisplayLayout.Appearance = appearance3;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn9.Header.Caption = "Código";
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn10.Header.Caption = "Artículo";
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn11.Header.VisiblePosition = 13;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn15.Header.Caption = "Fecha edad";
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn16.Header.Caption = "Edad días";
			ultraGridColumn16.Header.VisiblePosition = 9;
			ultraGridColumn17.Header.Caption = "Rango días";
			ultraGridColumn17.Header.VisiblePosition = 10;
			ultraGridColumn18.Header.VisiblePosition = 11;
			ultraGridColumn19.Header.Caption = "Fecha movimiento";
			ultraGridColumn19.Header.VisiblePosition = 14;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance4;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellButtonAppearance = appearance5;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "SPIFF ($)";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn21.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.Caption = "Vendedor";
			ultraGridColumn21.Header.VisiblePosition = 15;
			ultraGridColumn21.Width = 200;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.Caption = "Número";
			ultraGridColumn23.Header.VisiblePosition = 16;
			ultraGridColumn24.Header.Caption = "Fecha factura";
			ultraGridColumn24.Header.VisiblePosition = 17;
			ultraGridColumn25.Header.VisiblePosition = 22;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 23;
			ultraGridColumn26.Hidden = true;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance6;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellButtonAppearance = appearance7;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.Caption = "SPIFF Vendido ($)";
			ultraGridColumn27.Header.VisiblePosition = 20;
			ultraGridColumn28.Header.VisiblePosition = 24;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 25;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 26;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 27;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 5;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance8;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellButtonAppearance = appearance9;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.Caption = "Subtotal Factura ($)";
			ultraGridColumn33.Header.VisiblePosition = 18;
			ultraGridColumn33.Width = 115;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance10;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellButtonAppearance = appearance11;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.Caption = "Costo ($)";
			ultraGridColumn34.Header.VisiblePosition = 28;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance12;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellButtonAppearance = appearance13;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.Caption = "Precio ($)";
			ultraGridColumn35.Header.VisiblePosition = 29;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn36.Header.Caption = "Tipo Promoción";
			ultraGridColumn36.Header.VisiblePosition = 30;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance14;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellButtonAppearance = appearance15;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.Caption = "Precio Promoción ($)";
			ultraGridColumn37.Header.VisiblePosition = 31;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn37});
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance16;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance17;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			this.ugdSPIFFInventarioReporteCon.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ugdSPIFFInventarioReporteCon.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugdSPIFFInventarioReporteCon.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdSPIFFInventarioReporteCon.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.ugdSPIFFInventarioReporteCon.DisplayLayout.Override.CardAreaAppearance = appearance18;
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdSPIFFInventarioReporteCon.DisplayLayout.Override.CellAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 8F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdSPIFFInventarioReporteCon.DisplayLayout.Override.HeaderAppearance = appearance20;
			this.ugdSPIFFInventarioReporteCon.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdSPIFFInventarioReporteCon.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdSPIFFInventarioReporteCon.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.ugdSPIFFInventarioReporteCon.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdSPIFFInventarioReporteCon.Location = new System.Drawing.Point(19, 175);
			this.ugdSPIFFInventarioReporteCon.Name = "ugdSPIFFInventarioReporteCon";
			this.ugdSPIFFInventarioReporteCon.Size = new System.Drawing.Size(989, 222);
			this.ugdSPIFFInventarioReporteCon.TabIndex = 1056;
			this.ugdSPIFFInventarioReporteCon.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ugdSPIFFInventarioReporteCon_InitializeRow);
			this.ugdSPIFFInventarioReporteCon.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdSPIFFInventarioReporteCon_InitializeLayout);
			// 
			// UDSSPIFFInventarioReporteCon
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.DateTime);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(System.DateTime);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(System.DateTime);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			this.UDSSPIFFInventarioReporteCon.Band.Columns.AddRange(new object[] {
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
																																						 ultraDataColumn34});
			// 
			// groupBox5
			// 
			this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox5.Controls.Add(this.lblVendedor);
			this.groupBox5.Controls.Add(this.cmbVendedor);
			this.groupBox5.Controls.Add(this.btnConsultar);
			this.groupBox5.Controls.Add(this.label12);
			this.groupBox5.Controls.Add(this.txtBuscar);
			this.groupBox5.Controls.Add(this.lblContador);
			this.groupBox5.Controls.Add(this.udgSPIFFRangoCOn);
			this.groupBox5.Location = new System.Drawing.Point(19, 28);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(989, 129);
			this.groupBox5.TabIndex = 1055;
			this.groupBox5.TabStop = false;
			this.groupBox5.Text = "Búsqueda";
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(19, 28);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(70, 19);
			this.lblVendedor.TabIndex = 1061;
			this.lblVendedor.Text = "Vendedor:";
			// 
			// cmbVendedor
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance23;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataSource = this.udsSPIFFListaVendedorCon;
			ultraGridColumn38.Header.VisiblePosition = 0;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 1;
			ultraGridColumn39.Width = 335;
			ultraGridColumn40.Header.Caption = "Código";
			ultraGridColumn40.Header.VisiblePosition = 2;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(86, 28);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(384, 26);
			this.cmbVendedor.TabIndex = 1060;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// udsSPIFFListaVendedorCon
			// 
			ultraDataColumn35.DataType = typeof(int);
			this.udsSPIFFListaVendedorCon.Band.Columns.AddRange(new object[] {
																																				 ultraDataColumn35,
																																				 ultraDataColumn36,
																																				 ultraDataColumn37});
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(874, 28);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(93, 27);
			this.btnConsultar.TabIndex = 1059;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(19, 65);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(53, 19);
			this.label12.TabIndex = 1058;
			this.label12.Text = "Buscar:";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance24;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(86, 65);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(384, 25);
			this.txtBuscar.TabIndex = 1057;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(19, 92);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 24);
			this.lblContador.TabIndex = 1056;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// udgSPIFFRangoCOn
			// 
			this.udgSPIFFRangoCOn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.udgSPIFFRangoCOn.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgSPIFFRangoCOn.DataSource = this.udsSPIFFRangoCon;
			appearance25.BackColor = System.Drawing.Color.White;
			this.udgSPIFFRangoCOn.DisplayLayout.Appearance = appearance25;
			this.udgSPIFFRangoCOn.DisplayLayout.AutoFitColumns = true;
			ultraGridBand5.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn41.Header.VisiblePosition = 2;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 85;
			ultraGridColumn42.Header.Caption = "Rango días";
			ultraGridColumn42.Header.VisiblePosition = 1;
			ultraGridColumn42.Width = 274;
			ultraGridColumn43.Header.Caption = "...";
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Width = 39;
			ultraGridColumn44.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn44.Header.Caption = "Estado";
			ultraGridColumn44.Header.VisiblePosition = 3;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn44.Width = 167;
			ultraGridColumn45.AutoEdit = false;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn45.Header.Caption = "N°";
			ultraGridColumn45.Header.VisiblePosition = 4;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 21;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45});
			ultraGridBand5.Header.Caption = "Detalle";
			this.udgSPIFFRangoCOn.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.udgSPIFFRangoCOn.DisplayLayout.GroupByBox.Hidden = true;
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgSPIFFRangoCOn.DisplayLayout.Override.ActiveRowAppearance = appearance26;
			this.udgSPIFFRangoCOn.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgSPIFFRangoCOn.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgSPIFFRangoCOn.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance27.BackColor = System.Drawing.Color.Transparent;
			this.udgSPIFFRangoCOn.DisplayLayout.Override.CardAreaAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance28.FontData.BoldAsString = "True";
			appearance28.FontData.Name = "Arial";
			appearance28.FontData.SizeInPoints = 10F;
			appearance28.ForeColor = System.Drawing.Color.White;
			appearance28.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgSPIFFRangoCOn.DisplayLayout.Override.HeaderAppearance = appearance28;
			this.udgSPIFFRangoCOn.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance29.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance29.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgSPIFFRangoCOn.DisplayLayout.Override.RowAlternateAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgSPIFFRangoCOn.DisplayLayout.Override.RowSelectorAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgSPIFFRangoCOn.DisplayLayout.Override.SelectedRowAppearance = appearance31;
			this.udgSPIFFRangoCOn.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgSPIFFRangoCOn.FlatMode = true;
			this.udgSPIFFRangoCOn.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.udgSPIFFRangoCOn.Location = new System.Drawing.Point(490, 30);
			this.udgSPIFFRangoCOn.Name = "udgSPIFFRangoCOn";
			this.udgSPIFFRangoCOn.Size = new System.Drawing.Size(355, 81);
			this.udgSPIFFRangoCOn.TabIndex = 1062;
			this.toolTip1.SetToolTip(this.udgSPIFFRangoCOn, "Nuevo");
			this.udgSPIFFRangoCOn.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgSPIFFRangoCOn_AfterCellUpdate);
			// 
			// udsSPIFFRangoCon
			// 
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn40.DataType = typeof(bool);
			this.udsSPIFFRangoCon.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40});
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
			this.btnSalir.Location = new System.Drawing.Point(950, 28);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(96, 26);
			this.btnSalir.TabIndex = 997;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
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
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnExcel.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(950, 74);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(96, 28);
			this.btnExcel.TabIndex = 998;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// frmSPIFFInventarioReporteCompleto
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(1065, 579);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox2);
			this.Name = "frmSPIFFInventarioReporteCompleto";
			this.Text = "Edad de inventario";
			this.Load += new System.EventHandler(this.frmSPIFFInventarioReporteCompleto_Load);
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbSPIFFSubGrupoRangoCab)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFSubGrupoRangoCabList)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdSPIFFInventarioReporteCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.UDSSPIFFInventarioReporteCon)).EndInit();
			this.groupBox5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFListaVendedorCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udgSPIFFRangoCOn)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFRangoCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void lblBodega_Click(object sender, System.EventArgs e)
		{
		
		}

		private Acceso miAcceso;
		public bool bMostrarTodosSPIFF = false;

		private void frmSPIFFInventarioReporteCompleto_Load(object sender, System.EventArgs e)
		{
			SeguridadAcceso();
			AccesoExcel();
			FechaCorte();
			BodegaVinculacion();
			SPIFFSubGrupoRangoCabList();
			SPIFFRangoCon();
		}

		private void AccesoExcel()
		{
			miAcceso = new Acceso(cdsSeteoF, "20054701");
			if (miAcceso.BExportar) this.btnExcel.Enabled =  true;
		}

		private void SPIFFListaVendedorCon()
		{
			string dtFecha = Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd");

			//Se va a poder revisar por todos los locales, para ello se debe quitar la validación de elección de locales y que se muestren todos
			int Bodega = 0;
			int cmbSPIFFSubGrupoRangoCab = 0;
			try
			{
				Bodega = (int)this.cmbBodega.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			try
			{
				cmbSPIFFSubGrupoRangoCab = (int)this.cmbSPIFFSubGrupoRangoCab.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

		  this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec SPIFFListaVendedorCon '{0}', {1}, {2}, {3}", 
			dtFecha, 
			Bodega, 
			MenuLatinium.iNivel,
			cmbSPIFFSubGrupoRangoCab
			));
		}

		private void SPIFFRangoCon()
		{
			string squery = "EXEC SPIFFRangoCon";
			this.udgSPIFFRangoCOn.DataSource = Funcion.dvProcedimiento(cdsSeteoF, squery);
		}

		private void SeguridadAcceso()
		{
			if (!Funcion.Permiso("2168", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar a " + this.Text, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void FechaCorte()
		{
			#region Fechas default
			DateTime dtInicio = DateTime.Today.AddMonths(-540);
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
			this.dtHasta.Value = DateTime.Today;
			#endregion Fechas default
		}

		private void BodegaVinculacion()
		{
			#region Bodega
			FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 1, DateTime.Today, false, cdsSeteoF);
			int IBodega = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec BodegaAsignadaPorUsuarioBuscaFacturas {0}", MenuLatinium.IdUsuario));
			if (IBodega > 0) 
			{
				this.cmbBodega.Value = IBodega;		
				bMostrarTodosSPIFF = false; 
			}
			else
			{
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 1, DateTime.Today, false, cdsSeteoF);
				bMostrarTodosSPIFF = true; 
				//this.cmbBodega.Enabled = true;
			}
			#endregion Bodega
		}

		//1725275059

		private void SPIFFSubGrupoRangoCabList()
		{
			string sSQLPa = string.Format("Exec SPIFFSubGrupoRangoCabReporteList {0}", bMostrarTodosSPIFF); //this.txtBusqueda.Text.ToString()
			this.cmbSPIFFSubGrupoRangoCab.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLPa);
		}

		private void SPIFFSubGrupoRangoCabModel(int idSPIFFSubGrupoRangoCab)
		{
			if(idSPIFFSubGrupoRangoCab > 0)
			{
				this.lblEstado.Text = Funcion.sEscalarSQL(cdsSeteoF, String.Format("Exec SPIFFSubGrupoRangoCabModel '{0}', {1}", "se.Descripcion", idSPIFFSubGrupoRangoCab));
			}
		}

		private void ultraCombo2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraCombo2_ValueChanged(object sender, System.EventArgs e)
		{
			int getidSPIFFSubGrupoRangoCab = 0;
			try{getidSPIFFSubGrupoRangoCab = (int)this.cmbSPIFFSubGrupoRangoCab.Value;}catch(Exception ex){string sms = ex.Message;}
			SPIFFSubGrupoRangoCabModel(getidSPIFFSubGrupoRangoCab);
			if(this.cmbSPIFFSubGrupoRangoCab.ActiveRow != null)
			{
				this.cmbBodega.Enabled = true;
				this.cmbVendedor.Enabled = true;
				this.txtBuscar.Enabled = true;
				//LISTADO TODOS LOS VENDEDORES DEL SPIFF
				SPIFFListaVendedorCon();

			}
			else
			{
				this.cmbBodega.Enabled = false;
				this.cmbVendedor.Enabled = false;
				this.txtBuscar.Enabled = false;
			}


		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			SPIFFInventarioReporteCon();
		}

		private void SPIFFInventarioReporteCon()
		{


			#region Validación
			if(!Validacion.vbComboVacio(this.cmbSPIFFSubGrupoRangoCab, "El campo SPIFF es obligatorio."))return;
			//if(!Validacion.vbComboVacio(this.cmbBodega, "El campo Bodega es obligatorio."))return;
			//Validar si ha seleccionado al menos un rango días
			int icontador = Funcion.iEscalarSQL(cdsSeteoF, "Exec SPIFFRangoSelCount", false);
			if(icontador == 0)
			{
				MessageBox.Show(string.Format("Proceso Finalizado, se debe seleccionar al menos un rango de días."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validación

			//Se va a poder revisar por todos los locales, para ello se debe quitar la validación de elección de locales y que se muestren todos
			int Bodega = 0;
			try
			{
				Bodega = (int)this.cmbBodega.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			int idPersonal = 0;
			try{idPersonal = (int)this.cmbVendedor.Value;}catch(Exception ex){string sms = ex.Message;}
			// Mostrar el modal de "procesando"
			MenuLatinium.MostrarModalProcesando();

			try
			{
				string squery = String.Format("EXEC SPIFFInventarioReporteCon_Comercial '{0}', {1}, {2}, {3}, {4}, {5}",
					this.txtBuscar.Text, 0, Bodega, 0, (int)this.cmbSPIFFSubGrupoRangoCab.Value, idPersonal
					);
				FuncionesProcedimientos.dsGeneralModulo(squery, this.ugdSPIFFInventarioReporteCon);
			}
			catch (Exception ex)
			{
				// Manejar cualquier excepción que pueda ocurrir durante la consulta
				MessageBox.Show("Ocurrió un error durante la consulta: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				// Ocultar el modal de "procesando" una vez que se haya completado la consulta
				MenuLatinium.OcultarModalProcesando();
			}


			//this.ugdSPIFFInventarioReporteCon.DataSource = Funcion.dvProcedimiento(cdsSeteoF, squery);
			//Label de registros
			this.lblContador.Text = this.ugdSPIFFInventarioReporteCon.Rows.Count + " REGISTROS ENCONTRADOS";

			// Asegúrate de que la columna 'Bodega' exista en la grilla
			if (ugdSPIFFInventarioReporteCon.DisplayLayout.Bands[0].Columns.Exists("Bodega"))
			{
				// Ordenación 
				// Accede a la banda de la grilla
				UltraGridBand band = ugdSPIFFInventarioReporteCon.DisplayLayout.Bands[0];

				// Accede a la columna que deseas ordenar
				UltraGridColumn column = band.Columns["Bodega"];

				// Establece la propiedad de ordenación
				column.SortIndicator = SortIndicator.Ascending;

				ugdSPIFFInventarioReporteCon.Refresh();
			}
			else
			{
				MessageBox.Show("La columna 'Bodega' no existe en la grilla.");
			}


			//COntador de filas
			//this.ugdSPIFFInventarioReporteCon = Funcion.ContadorFilas(ugdSPIFFInventarioReporteCon, "Row");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void ugdSPIFFInventarioReporteCon_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			decimal getSPIFF = 0;
			try{getSPIFF = Convert.ToDecimal(e.Row.Cells["SPIFFPorCobrar"].Value);}catch(Exception ex){string sms = ex.Message;}
			if (getSPIFF > 0)
			{
				e.Row.Appearance.BackColor = Color.Yellow;
				e.Row.Appearance.ForeColor = Color.Red;
			}
		}

		private void udgSPIFFRangoCOn_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Seleccionar")
			{
				int idSPIFFRango = (int)e.Cell.Row.Cells["idSPIFFRango"].Value;
				bool Seleccionar = (bool)e.Cell.Row.Cells["Seleccionar"].Value;
				SPIFFRangoSel(idSPIFFRango, Seleccionar);
			}
		}

		public void SPIFFRangoSel(int idSPIFFRango, bool Seleccionar)
		{
			#region Guardar datos
			//Actualiza valores en la compra
			string squery = string.Format("EXEC SPIFFRangoSel {0}, {1}", 
				idSPIFFRango, Seleccionar);
			Funcion.EjecutaSQL(cdsSeteoF, squery);
			#endregion Guardar datos
		}

		private void cmbBodega_TextChanged(object sender, System.EventArgs e)
		{
			SPIFFListaVendedorCon();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ugdSPIFFInventarioReporteCon);
		}

		private void ugdSPIFFInventarioReporteCon_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
