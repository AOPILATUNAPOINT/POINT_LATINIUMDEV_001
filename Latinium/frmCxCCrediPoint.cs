using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmCxCCrediPoint.
	/// </summary>
	public class frmCxCCrediPoint : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocales;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempoUso;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgDocumento;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCxCCrediPoint()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumDocAntiguo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Documento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaDePago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EntFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapitalPagado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapitalVencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapitalPorVencer");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PlazoOriginal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasPagadas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasVencidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasPorPagar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencido30Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencido60Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencido90Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencido90180Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencido180360Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencido360Dias");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumDocAntiguo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Documento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaDePago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EntFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasMora");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapitalPagado");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapitalVencido");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapitalPorVencer");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PlazoOriginal");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasPagadas");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasVencidas");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasPorPagar");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencido30Dias");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencido60Dias");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencido90Dias");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencido90180Dias");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencido180360Dias");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencido360Dias");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 13, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 14, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CapitalPagado", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CapitalPagado", 16, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CapitalVencido", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CapitalVencido", 17, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CapitalPorVencer", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CapitalPorVencer", 18, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vencido30Dias", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vencido30Dias", 23, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vencido60Dias", 24, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vencido60Dias", 24, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vencido90Dias", 25, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vencido90Dias", 25, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vencido90180Dias", 26, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vencido90180Dias", 26, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vencido180360Dias", 27, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vencido180360Dias", 27, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vencido360Dias", 28, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vencido360Dias", 28, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.optFiltro = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label1 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTiempoUso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.rdgDocumento = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoUso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgDocumento)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(System.Double);
			ultraDataColumn15.DataType = typeof(System.Double);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(System.Double);
			ultraDataColumn18.DataType = typeof(System.Double);
			ultraDataColumn19.DataType = typeof(System.Double);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(System.Double);
			ultraDataColumn25.DataType = typeof(System.Double);
			ultraDataColumn26.DataType = typeof(System.Double);
			ultraDataColumn27.DataType = typeof(System.Double);
			ultraDataColumn28.DataType = typeof(System.Double);
			ultraDataColumn29.DataType = typeof(System.Double);
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
																																 ultraDataColumn29});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 25;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 96;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 90;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Numero EM";
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 96;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Doc";
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 50;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "N. Lote";
			ultraGridColumn7.Header.VisiblePosition = 14;
			ultraGridColumn7.Width = 50;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Forma De Pago";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 120;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Ent Financiera";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 120;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Width = 65;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 9;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 52;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn12.Width = 110;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 11;
			ultraGridColumn13.Width = 250;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance2;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 12;
			ultraGridColumn14.Width = 75;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance3;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 13;
			ultraGridColumn15.Width = 75;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance4;
			ultraGridColumn16.Format = "";
			ultraGridColumn16.Header.Caption = "D Mora";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 75;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance5;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Cap Pagado";
			ultraGridColumn17.Header.VisiblePosition = 20;
			ultraGridColumn17.Width = 100;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance6;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Cap Vencido";
			ultraGridColumn18.Header.VisiblePosition = 21;
			ultraGridColumn18.Width = 100;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance7;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Cap Por Vencer";
			ultraGridColumn19.Header.VisiblePosition = 22;
			ultraGridColumn19.Width = 100;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance8;
			ultraGridColumn20.Format = "";
			ultraGridColumn20.Header.Caption = "Pl Orig";
			ultraGridColumn20.Header.VisiblePosition = 16;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance9;
			ultraGridColumn21.Header.Caption = "C Pagadas";
			ultraGridColumn21.Header.VisiblePosition = 17;
			ultraGridColumn21.Width = 80;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance10;
			ultraGridColumn22.Header.Caption = "C Vencidas";
			ultraGridColumn22.Header.VisiblePosition = 18;
			ultraGridColumn22.Width = 80;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance11;
			ultraGridColumn23.Header.Caption = "C Por Vencer";
			ultraGridColumn23.Header.VisiblePosition = 19;
			ultraGridColumn23.Width = 80;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance12;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "V 30";
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 90;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance13;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.Caption = "V 60";
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Width = 90;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance14;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.Caption = "V 90";
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 90;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance15;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.Caption = "V 90 - 180";
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Width = 90;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance16;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.Caption = "V 180-360";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Width = 90;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance17;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.Caption = "V 360";
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Width = 90;
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
																										 ultraGridColumn29});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance18;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance19;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance20;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance21;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance22;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance23;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance24;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance25;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance26;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance27;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance28;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance29;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance30.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance30;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance31.FontData.BoldAsString = "True";
			appearance31.FontData.Name = "Arial";
			appearance31.FontData.SizeInPoints = 8.5F;
			appearance31.ForeColor = System.Drawing.Color.White;
			appearance31.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance31;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance33;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(488, 152);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(792, 304);
			this.ultraGrid1.TabIndex = 663;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(bool);
			ultraDataColumn32.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// uGridLocales
			// 
			this.uGridLocales.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocales.DataSource = this.ultraDataSource2;
			appearance34.BackColor = System.Drawing.Color.White;
			this.uGridLocales.DisplayLayout.Appearance = appearance34;
			this.uGridLocales.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 0;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 100;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Width = 191;
			ultraGridColumn32.Header.Caption = "...";
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn32.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn32.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			ultraGridBand2.Header.Caption = "LOCALES";
			ultraGridBand2.HeaderVisible = true;
			this.uGridLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance35.ForeColor = System.Drawing.Color.Black;
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocales.DisplayLayout.Override.ActiveRowAppearance = appearance35;
			this.uGridLocales.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridLocales.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridLocales.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance36.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocales.DisplayLayout.Override.CardAreaAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance37.FontData.BoldAsString = "True";
			appearance37.FontData.Name = "Arial";
			appearance37.FontData.SizeInPoints = 8F;
			appearance37.ForeColor = System.Drawing.Color.White;
			appearance37.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocales.DisplayLayout.Override.HeaderAppearance = appearance37;
			appearance38.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance38.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowAlternateAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowSelectorAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.SelectedRowAppearance = appearance40;
			this.uGridLocales.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocales.Location = new System.Drawing.Point(8, 8);
			this.uGridLocales.Name = "uGridLocales";
			this.uGridLocales.Size = new System.Drawing.Size(232, 448);
			this.uGridLocales.TabIndex = 665;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource3;
			appearance41.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance41;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 68;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 186;
			ultraGridColumn35.Header.Caption = "...";
			ultraGridColumn35.Header.VisiblePosition = 2;
			ultraGridColumn35.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn35.Width = 25;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35});
			ultraGridBand3.Header.Caption = "ENTIDAD FINANCIERA";
			ultraGridBand3.HeaderVisible = true;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance42.ForeColor = System.Drawing.Color.Black;
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance42;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance43.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance44.FontData.BoldAsString = "True";
			appearance44.FontData.Name = "Arial";
			appearance44.FontData.SizeInPoints = 8F;
			appearance44.ForeColor = System.Drawing.Color.White;
			appearance44.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance44;
			appearance45.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance45.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance45;
			appearance46.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance46.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance46;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance47;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(248, 8);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(232, 448);
			this.ultraGrid2.TabIndex = 666;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(bool);
			ultraDataColumn35.DefaultValue = false;
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35});
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
			// 
			// optFiltro
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro.Appearance = appearance48;
			this.optFiltro.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro.ItemAppearance = appearance49;
			this.optFiltro.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendientes";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Vencidos";
			this.optFiltro.Items.Add(valueListItem1);
			this.optFiltro.Items.Add(valueListItem2);
			this.optFiltro.Items.Add(valueListItem3);
			this.optFiltro.ItemSpacingVertical = 10;
			this.optFiltro.Location = new System.Drawing.Point(1064, 8);
			this.optFiltro.Name = "optFiltro";
			this.optFiltro.Size = new System.Drawing.Size(88, 80);
			this.optFiltro.TabIndex = 667;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label1.Location = new System.Drawing.Point(488, 112);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 17);
			this.label1.TabIndex = 669;
			this.label1.Text = "Buscar";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance50;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(544, 112);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(240, 22);
			this.txtBuscar.TabIndex = 668;
			// 
			// txtTiempoUso
			// 
			this.txtTiempoUso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempoUso.Enabled = false;
			this.txtTiempoUso.Location = new System.Drawing.Point(64, 32);
			this.txtTiempoUso.MinValue = 0;
			this.txtTiempoUso.Name = "txtTiempoUso";
			this.txtTiempoUso.PromptChar = ' ';
			this.txtTiempoUso.Size = new System.Drawing.Size(72, 22);
			this.txtTiempoUso.TabIndex = 671;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(16, 32);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(35, 16);
			this.label2.TabIndex = 672;
			this.label2.Text = "Desde";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(152, 32);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(33, 16);
			this.label3.TabIndex = 674;
			this.label3.Text = "Hasta";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraNumericEditor1
			// 
			this.ultraNumericEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor1.Enabled = false;
			this.ultraNumericEditor1.Location = new System.Drawing.Point(200, 32);
			this.ultraNumericEditor1.MinValue = 0;
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(72, 22);
			this.ultraNumericEditor1.TabIndex = 673;
			// 
			// rdgDocumento
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgDocumento.Appearance = appearance51;
			this.rdgDocumento.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rdgDocumento.ItemAppearance = appearance52;
			this.rdgDocumento.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Todos";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Facturas";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Notas de Credito";
			valueListItem7.DataValue = 3;
			valueListItem7.DisplayText = "Notas de Debito";
			valueListItem8.DataValue = 4;
			valueListItem8.DisplayText = "Anticipos";
			this.rdgDocumento.Items.Add(valueListItem4);
			this.rdgDocumento.Items.Add(valueListItem5);
			this.rdgDocumento.Items.Add(valueListItem6);
			this.rdgDocumento.Items.Add(valueListItem7);
			this.rdgDocumento.Items.Add(valueListItem8);
			this.rdgDocumento.ItemSpacingVertical = 10;
			this.rdgDocumento.Location = new System.Drawing.Point(1160, 8);
			this.rdgDocumento.Name = "rdgDocumento";
			this.rdgDocumento.Size = new System.Drawing.Size(120, 128);
			this.rdgDocumento.TabIndex = 675;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtTiempoUso);
			this.groupBox1.Controls.Add(this.ultraNumericEditor1);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Location = new System.Drawing.Point(488, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(280, 64);
			this.groupBox1.TabIndex = 676;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Dias de Mora";
			// 
			// frmCxCCrediPoint
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1288, 462);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.rdgDocumento);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.optFiltro);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.uGridLocales);
			this.Name = "frmCxCCrediPoint";
			this.Text = "Cuentas Por Cobrar CREDIPOINT";
			this.Load += new System.EventHandler(this.frmCxCCrediPoint_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoUso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgDocumento)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCxCCrediPoint_Load(object sender, System.EventArgs e)
		{
			//this.uGridLocales.DataSource = Funcion.dvProcedimiento (, "Select Bodega, Nombre, CONVERT(Bit, 0) As Sel From Bodega Where Activo = 1 And Factura = 1 Order By Nombre")
		}
	}
}

