using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Xml.Schema;
using System.Globalization;
using C1.Data;
using System.IO;
using System.Text;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmDetalleATS.
	/// </summary>
	public class frmDetalleATS1 : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Button btnGenerar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinGrid.UltraDropDown cmbCreditoTributario;
		public Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoRUC;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbComprobante;
		public Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoProveedor;
		public Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoPago;
		public Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoFormaPago;
		public Infragistics.Win.UltraWinGrid.UltraDropDown cmbPais;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridCompras;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridVentas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btnGeneraXML;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Button btnTalonResumen;
		private System.Windows.Forms.SaveFileDialog saveFileDialog2;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridTransacciones;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbComprobantes;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTransacciones;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnCerrarImpuestos;
		private System.Windows.Forms.DateTimePicker dtPeriodo;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdDebitoRenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCierreImpuestos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidAsientoRenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidAsientoIVA;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidDebitoIVA;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmDetalleATS1()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc", -1, "cmbTipoRUC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario", -1, "cmbCreditoTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante", -1, "cmbComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoProveedor", -1, "cmbTipoProveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoDePago", -1, "cmbTipoPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFormaPago", -1, "cmbTipoFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPais", -1, "cmbPais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseNoImp");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalExcento");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento0");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base0");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalIva");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizacionRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FecRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVARetenido");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalRetenido");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BaseNoImp", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BaseNoImp", 16, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalExcento", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalExcento", 17, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Descuento0", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Descuento0", 18, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Base0", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Base0", 19, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalIva", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalIva", 20, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Descuento", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Descuento", 21, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Base", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Base", 22, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Iva", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Iva", 23, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Ice", 24, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Ice", 24, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 25, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 25, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVARetenido", 31, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVARetenido", 31, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Renta", 32, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Renta", 32, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retenido", 33, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retenido", 33, true);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalRetenido", 34, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalRetenido", 34, true);
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoRuc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCredTributario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idComprobante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoProveedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoDePago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPais");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SerieFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BaseNoImp");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtotalExcento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtotalIva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Iva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ice");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serie");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumRet");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AutorizacionRet");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FecRet");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVARetenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Renta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalRetenido");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc", -1, "cmbTipoRUC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario", -1, "cmbCreditoTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante", -1, "cmbComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseNoImp");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalExcento");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento0");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base0");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalIva");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizacionRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FecRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVARetenido");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalRetenido");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BaseNoImp", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BaseNoImp", 12, true);
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalExcento", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalExcento", 13, true);
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Descuento0", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Descuento0", 14, true);
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Base0", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Base0", 15, true);
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalIva", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalIva", 16, true);
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Descuento", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Descuento", 17, true);
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Base", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Base", 18, true);
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Iva", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Iva", 19, true);
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Ice", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Ice", 20, true);
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 21, true);
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVARetenido", 27, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVARetenido", 27, true);
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Renta", 28, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Renta", 28, true);
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings27 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retenido", 29, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retenido", 29, true);
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings28 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalRetenido", 30, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalRetenido", 30, true);
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoRuc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCredTributario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idComprobante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SerieFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BaseNoImp");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtotalExcento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtotalIva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Iva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ice");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serie");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumRet");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AutorizacionRet");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FecRet");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVARetenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Renta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalRetenido");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fiscal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base0");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseIVA");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetenidoIVA");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetenidoRenta");
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings29 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Base0", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Base0", 8, true);
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings30 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BaseIVA", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BaseIVA", 9, true);
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings31 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Iva", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Iva", 10, true);
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings32 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 11, true);
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings33 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetenidoIVA", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetenidoIVA", 12, true);
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings34 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetenidoRenta", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetenidoRenta", 13, true);
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fiscal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BaseIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Iva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RetenidoIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RetenidoRenta");
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.uGridCompras = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.uGridVentas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGeneraXML = new System.Windows.Forms.Button();
			this.cmbCreditoTributario = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbTipoRUC = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbComprobante = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbTipoProveedor = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbTipoPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbTipoFormaPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbPais = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btnTalonResumen = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.saveFileDialog2 = new System.Windows.Forms.SaveFileDialog();
			this.uGridTransacciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbComprobantes = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbTransacciones = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnCerrarImpuestos = new System.Windows.Forms.Button();
			this.dtPeriodo = new System.Windows.Forms.DateTimePicker();
			this.label4 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtIdDebitoRenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidCierreImpuestos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidAsientoRenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidAsientoIVA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidDebitoIVA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.uGridCompras)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCreditoTributario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRUC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPais)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTransacciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobantes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransacciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDebitoRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCierreImpuestos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsientoRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsientoIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDebitoIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridCompras
			// 
			this.uGridCompras.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.uGridCompras.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCompras.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridCompras.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 40;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 75;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Identificación";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 200;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Crédito Tributario";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 140;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Comprobante";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 100;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Proveedor";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 140;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Tipo De Pago";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn11.Width = 100;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Forma De Pago";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn12.Width = 120;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "Pais";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn13.Width = 100;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.Caption = "Serie";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 60;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 100;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 120;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance2;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 90;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance3;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 90;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance4;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 90;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance5;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 90;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance6;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 90;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance7;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 90;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance8;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 90;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance9;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 90;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance10;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Width = 90;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance11;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 90;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.Caption = "R-Serie";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Width = 60;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.Caption = "R-Numero";
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Width = 100;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.Caption = "R-Autorizacion";
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Width = 120;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.Caption = "R-Fecha";
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Width = 75;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance12;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.Caption = "R-IVA";
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Width = 90;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance13;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.Caption = "R-Renta";
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Width = 90;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance14;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.Caption = "R-Retenido";
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn34.Width = 90;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance15;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.Caption = "Total";
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn35.Width = 90;
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
																										 ultraGridColumn35});
			appearance16.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance16;
			ultraGridBand1.Header.Caption = "COMPRAS";
			ultraGridBand1.HeaderVisible = true;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance17;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance18;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance19;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance20;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance21;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance22;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance23;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance24;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance25;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance26;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance27;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance28;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance29;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance30;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
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
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13,
																																															summarySettings14});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridCompras.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance31.ForeColor = System.Drawing.Color.Black;
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCompras.DisplayLayout.Override.ActiveRowAppearance = appearance31;
			this.uGridCompras.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridCompras.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridCompras.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance32.BackColor = System.Drawing.Color.Transparent;
			this.uGridCompras.DisplayLayout.Override.CardAreaAppearance = appearance32;
			this.uGridCompras.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance33.FontData.BoldAsString = "True";
			appearance33.FontData.Name = "Arial";
			appearance33.FontData.SizeInPoints = 8F;
			appearance33.ForeColor = System.Drawing.Color.White;
			appearance33.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCompras.DisplayLayout.Override.HeaderAppearance = appearance33;
			appearance34.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance34.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCompras.DisplayLayout.Override.RowAlternateAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCompras.DisplayLayout.Override.RowSelectorAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCompras.DisplayLayout.Override.SelectedRowAppearance = appearance36;
			this.uGridCompras.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCompras.Location = new System.Drawing.Point(8, 368);
			this.uGridCompras.Name = "uGridCompras";
			this.uGridCompras.Size = new System.Drawing.Size(112, 32);
			this.uGridCompras.TabIndex = 19;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn4.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(System.Double);
			ultraDataColumn17.DefaultValue = 0;
			ultraDataColumn18.DataType = typeof(System.Double);
			ultraDataColumn18.DefaultValue = 0;
			ultraDataColumn19.DataType = typeof(System.Double);
			ultraDataColumn19.DefaultValue = 0;
			ultraDataColumn20.DataType = typeof(System.Double);
			ultraDataColumn20.DefaultValue = 0;
			ultraDataColumn21.DataType = typeof(System.Double);
			ultraDataColumn21.DefaultValue = 0;
			ultraDataColumn22.DataType = typeof(System.Double);
			ultraDataColumn22.DefaultValue = 0;
			ultraDataColumn23.DataType = typeof(System.Double);
			ultraDataColumn23.DefaultValue = 0;
			ultraDataColumn24.DataType = typeof(System.Double);
			ultraDataColumn24.DefaultValue = 0;
			ultraDataColumn25.DataType = typeof(System.Double);
			ultraDataColumn25.DefaultValue = 0;
			ultraDataColumn26.DataType = typeof(System.Double);
			ultraDataColumn26.DefaultValue = 0;
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(System.DateTime);
			ultraDataColumn31.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn32.DataType = typeof(System.Double);
			ultraDataColumn32.DefaultValue = 0;
			ultraDataColumn33.DataType = typeof(System.Double);
			ultraDataColumn33.DefaultValue = 0;
			ultraDataColumn34.DataType = typeof(System.Double);
			ultraDataColumn34.DefaultValue = 0;
			ultraDataColumn35.DataType = typeof(System.Double);
			ultraDataColumn35.DefaultValue = 0;
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
																																 ultraDataColumn35});
			// 
			// btnGenerar
			// 
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(424, 72);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.TabIndex = 216;
			this.btnGenerar.Text = "&Ver";
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// uGridVentas
			// 
			this.uGridVentas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.uGridVentas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVentas.DataSource = this.ultraDataSource2;
			appearance37.BackColor = System.Drawing.Color.White;
			this.uGridVentas.DisplayLayout.Appearance = appearance37;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn36.Header.VisiblePosition = 0;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn37.Header.VisiblePosition = 1;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn38.Header.VisiblePosition = 2;
			ultraGridColumn38.Width = 40;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn39.Header.VisiblePosition = 3;
			ultraGridColumn39.Width = 75;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn40.Header.Caption = "Identificacion";
			ultraGridColumn40.Header.VisiblePosition = 4;
			ultraGridColumn40.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn40.Width = 100;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn41.Header.VisiblePosition = 5;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn42.Header.VisiblePosition = 6;
			ultraGridColumn42.Width = 200;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn43.Header.Caption = "Crédito Tributario";
			ultraGridColumn43.Header.VisiblePosition = 7;
			ultraGridColumn43.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn43.Width = 140;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn44.Header.Caption = "Comprobante";
			ultraGridColumn44.Header.VisiblePosition = 8;
			ultraGridColumn44.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn44.Width = 100;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn45.Header.Caption = "Serie";
			ultraGridColumn45.Header.VisiblePosition = 9;
			ultraGridColumn45.Width = 60;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn46.Header.VisiblePosition = 10;
			ultraGridColumn46.Width = 90;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn47.Header.VisiblePosition = 11;
			ultraGridColumn47.Width = 150;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance38;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.Caption = "No Imponible";
			ultraGridColumn48.Header.VisiblePosition = 12;
			ultraGridColumn48.Width = 90;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance39;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.Caption = "Subtotal IVA 0";
			ultraGridColumn49.Header.VisiblePosition = 13;
			ultraGridColumn49.Width = 90;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance40;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.Caption = "Descuento IVA 0";
			ultraGridColumn50.Header.VisiblePosition = 14;
			ultraGridColumn50.Width = 90;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance41;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.Caption = "Base IVA 0";
			ultraGridColumn51.Header.VisiblePosition = 15;
			ultraGridColumn51.Width = 90;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance42;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.Caption = "Subtotal IVA ";
			ultraGridColumn52.Header.VisiblePosition = 16;
			ultraGridColumn52.Width = 90;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance43;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.Caption = "Descuento IVA";
			ultraGridColumn53.Header.VisiblePosition = 17;
			ultraGridColumn53.Width = 90;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance44;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.Caption = "Base IVA";
			ultraGridColumn54.Header.VisiblePosition = 18;
			ultraGridColumn54.Width = 90;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance45;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.Caption = "IVA";
			ultraGridColumn55.Header.VisiblePosition = 19;
			ultraGridColumn55.Width = 90;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance46;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.Caption = "ICE";
			ultraGridColumn56.Header.VisiblePosition = 20;
			ultraGridColumn56.Width = 90;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn57.CellAppearance = appearance47;
			ultraGridColumn57.Format = "#,##0.00";
			ultraGridColumn57.Header.VisiblePosition = 21;
			ultraGridColumn57.Width = 90;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn58.Header.VisiblePosition = 22;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn59.Header.Caption = "R-Serie";
			ultraGridColumn59.Header.VisiblePosition = 23;
			ultraGridColumn59.Width = 60;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn60.Header.Caption = "R-Numero";
			ultraGridColumn60.Header.VisiblePosition = 24;
			ultraGridColumn60.Width = 90;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn61.Header.Caption = "R-Autorizacion";
			ultraGridColumn61.Header.VisiblePosition = 25;
			ultraGridColumn61.Width = 150;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn62.Header.Caption = "R-Fecha";
			ultraGridColumn62.Header.VisiblePosition = 26;
			ultraGridColumn62.Width = 90;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance48;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.Caption = "R-IVA";
			ultraGridColumn63.Header.VisiblePosition = 27;
			ultraGridColumn63.Width = 90;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance49;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.Caption = "R-Renta";
			ultraGridColumn64.Header.VisiblePosition = 28;
			ultraGridColumn64.Width = 90;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance50;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.VisiblePosition = 29;
			ultraGridColumn65.Width = 90;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance51;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.Caption = "Total - Retención";
			ultraGridColumn66.Header.VisiblePosition = 30;
			ultraGridColumn66.Width = 90;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn66});
			appearance52.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand2.Header.Appearance = appearance52;
			ultraGridBand2.Header.Caption = "VENTAS";
			ultraGridBand2.HeaderVisible = true;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance53;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance54;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance55;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance56;
			summarySettings18.DisplayFormat = "{0: #,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance57;
			summarySettings19.DisplayFormat = "{0: #,##0.00}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance58;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance59;
			summarySettings21.DisplayFormat = "{0: #,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance60;
			summarySettings22.DisplayFormat = "{0: #,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance61;
			summarySettings23.DisplayFormat = "{0: #,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance62;
			summarySettings24.DisplayFormat = "{0: #,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance63;
			summarySettings25.DisplayFormat = "{0: #,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance64;
			summarySettings26.DisplayFormat = "{0: #,##0.00}";
			summarySettings26.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings27.Appearance = appearance65;
			summarySettings27.DisplayFormat = "{0: #,##0.00}";
			summarySettings27.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings28.Appearance = appearance66;
			summarySettings28.DisplayFormat = "{0: #,##0.00}";
			summarySettings28.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings15,
																																															summarySettings16,
																																															summarySettings17,
																																															summarySettings18,
																																															summarySettings19,
																																															summarySettings20,
																																															summarySettings21,
																																															summarySettings22,
																																															summarySettings23,
																																															summarySettings24,
																																															summarySettings25,
																																															summarySettings26,
																																															summarySettings27,
																																															summarySettings28});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridVentas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance67.ForeColor = System.Drawing.Color.Black;
			appearance67.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentas.DisplayLayout.Override.ActiveRowAppearance = appearance67;
			this.uGridVentas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridVentas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance68.BackColor = System.Drawing.Color.Transparent;
			this.uGridVentas.DisplayLayout.Override.CardAreaAppearance = appearance68;
			this.uGridVentas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance69.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance69.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance69.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance69.FontData.BoldAsString = "True";
			appearance69.FontData.Name = "Arial";
			appearance69.FontData.SizeInPoints = 8F;
			appearance69.ForeColor = System.Drawing.Color.White;
			appearance69.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVentas.DisplayLayout.Override.HeaderAppearance = appearance69;
			appearance70.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance70.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance70.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.RowAlternateAppearance = appearance70;
			appearance71.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance71.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance71.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.RowSelectorAppearance = appearance71;
			appearance72.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance72.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance72.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.SelectedRowAppearance = appearance72;
			this.uGridVentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVentas.Location = new System.Drawing.Point(8, 368);
			this.uGridVentas.Name = "uGridVentas";
			this.uGridVentas.Size = new System.Drawing.Size(112, 32);
			this.uGridVentas.TabIndex = 217;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn39.DataType = typeof(System.DateTime);
			ultraDataColumn39.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn40.DataType = typeof(int);
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(System.Double);
			ultraDataColumn48.DefaultValue = 0;
			ultraDataColumn49.DataType = typeof(System.Double);
			ultraDataColumn49.DefaultValue = 0;
			ultraDataColumn50.DataType = typeof(System.Double);
			ultraDataColumn50.DefaultValue = 0;
			ultraDataColumn51.DataType = typeof(System.Double);
			ultraDataColumn51.DefaultValue = 0;
			ultraDataColumn52.DataType = typeof(System.Double);
			ultraDataColumn52.DefaultValue = 0;
			ultraDataColumn53.DataType = typeof(System.Double);
			ultraDataColumn53.DefaultValue = 0;
			ultraDataColumn54.DataType = typeof(System.Double);
			ultraDataColumn54.DefaultValue = 0;
			ultraDataColumn55.DataType = typeof(System.Double);
			ultraDataColumn55.DefaultValue = 0;
			ultraDataColumn56.DataType = typeof(System.Double);
			ultraDataColumn56.DefaultValue = 0;
			ultraDataColumn57.DataType = typeof(System.Double);
			ultraDataColumn57.DefaultValue = 0;
			ultraDataColumn58.DataType = typeof(int);
			ultraDataColumn62.DataType = typeof(System.DateTime);
			ultraDataColumn62.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn63.DataType = typeof(System.Double);
			ultraDataColumn63.DefaultValue = 0;
			ultraDataColumn64.DataType = typeof(System.Double);
			ultraDataColumn64.DefaultValue = 0;
			ultraDataColumn65.DataType = typeof(System.Double);
			ultraDataColumn65.DefaultValue = 0;
			ultraDataColumn66.DataType = typeof(System.Double);
			ultraDataColumn66.DefaultValue = 0;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn58,
																																 ultraDataColumn59,
																																 ultraDataColumn60,
																																 ultraDataColumn61,
																																 ultraDataColumn62,
																																 ultraDataColumn63,
																																 ultraDataColumn64,
																																 ultraDataColumn65,
																																 ultraDataColumn66});
			// 
			// btnGeneraXML
			// 
			this.btnGeneraXML.CausesValidation = false;
			this.btnGeneraXML.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGeneraXML.Location = new System.Drawing.Point(864, 40);
			this.btnGeneraXML.Name = "btnGeneraXML";
			this.btnGeneraXML.Size = new System.Drawing.Size(120, 23);
			this.btnGeneraXML.TabIndex = 218;
			this.btnGeneraXML.Text = "&Generar XML";
			this.btnGeneraXML.Click += new System.EventHandler(this.btnGeneraXML_Click);
			// 
			// cmbCreditoTributario
			// 
			this.cmbCreditoTributario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbCreditoTributario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCreditoTributario.DisplayMember = "CredTributario";
			this.cmbCreditoTributario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCreditoTributario.Location = new System.Drawing.Point(120, 376);
			this.cmbCreditoTributario.Name = "cmbCreditoTributario";
			this.cmbCreditoTributario.Size = new System.Drawing.Size(96, 24);
			this.cmbCreditoTributario.TabIndex = 219;
			this.cmbCreditoTributario.ValueMember = "idCredTributario";
			this.cmbCreditoTributario.Visible = false;
			// 
			// cmbTipoRUC
			// 
			this.cmbTipoRUC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbTipoRUC.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRUC.DisplayMember = "TipoRuc";
			this.cmbTipoRUC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRUC.Location = new System.Drawing.Point(120, 376);
			this.cmbTipoRUC.Name = "cmbTipoRUC";
			this.cmbTipoRUC.Size = new System.Drawing.Size(96, 24);
			this.cmbTipoRUC.TabIndex = 220;
			this.cmbTipoRUC.ValueMember = "idTipoRuc";
			this.cmbTipoRUC.Visible = false;
			// 
			// cmbComprobante
			// 
			this.cmbComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbComprobante.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbComprobante.DisplayMember = "Comprobante";
			this.cmbComprobante.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbComprobante.Location = new System.Drawing.Point(120, 376);
			this.cmbComprobante.Name = "cmbComprobante";
			this.cmbComprobante.Size = new System.Drawing.Size(96, 24);
			this.cmbComprobante.TabIndex = 221;
			this.cmbComprobante.ValueMember = "idComprobante";
			this.cmbComprobante.Visible = false;
			// 
			// cmbTipoProveedor
			// 
			this.cmbTipoProveedor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbTipoProveedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoProveedor.DisplayMember = "Tipo";
			this.cmbTipoProveedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoProveedor.Location = new System.Drawing.Point(120, 376);
			this.cmbTipoProveedor.Name = "cmbTipoProveedor";
			this.cmbTipoProveedor.Size = new System.Drawing.Size(96, 24);
			this.cmbTipoProveedor.TabIndex = 222;
			this.cmbTipoProveedor.ValueMember = "idTipoProveedor";
			this.cmbTipoProveedor.Visible = false;
			// 
			// cmbTipoPago
			// 
			this.cmbTipoPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbTipoPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoPago.DisplayMember = "Descripcion";
			this.cmbTipoPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoPago.Location = new System.Drawing.Point(120, 376);
			this.cmbTipoPago.Name = "cmbTipoPago";
			this.cmbTipoPago.Size = new System.Drawing.Size(96, 24);
			this.cmbTipoPago.TabIndex = 223;
			this.cmbTipoPago.ValueMember = "idTipoDePago";
			this.cmbTipoPago.Visible = false;
			// 
			// cmbTipoFormaPago
			// 
			this.cmbTipoFormaPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbTipoFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFormaPago.DisplayMember = "FormaDePago";
			this.cmbTipoFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFormaPago.Location = new System.Drawing.Point(120, 376);
			this.cmbTipoFormaPago.Name = "cmbTipoFormaPago";
			this.cmbTipoFormaPago.Size = new System.Drawing.Size(96, 24);
			this.cmbTipoFormaPago.TabIndex = 224;
			this.cmbTipoFormaPago.ValueMember = "idTipoFormaPago";
			this.cmbTipoFormaPago.Visible = false;
			// 
			// cmbPais
			// 
			this.cmbPais.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbPais.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPais.DisplayMember = "Pais";
			this.cmbPais.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPais.Location = new System.Drawing.Point(120, 376);
			this.cmbPais.Name = "cmbPais";
			this.cmbPais.Size = new System.Drawing.Size(96, 24);
			this.cmbPais.TabIndex = 225;
			this.cmbPais.ValueMember = "idClientePais";
			this.cmbPais.Visible = false;
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
			// btnTalonResumen
			// 
			this.btnTalonResumen.CausesValidation = false;
			this.btnTalonResumen.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTalonResumen.Location = new System.Drawing.Point(864, 8);
			this.btnTalonResumen.Name = "btnTalonResumen";
			this.btnTalonResumen.Size = new System.Drawing.Size(120, 23);
			this.btnTalonResumen.TabIndex = 236;
			this.btnTalonResumen.Text = "&Ver Talón Resumen";
			this.btnTalonResumen.Click += new System.EventHandler(this.btnTalonResumen_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.CausesValidation = false;
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(864, 72);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(120, 23);
			this.btnExportar.TabIndex = 237;
			this.btnExportar.Text = "&Exportar Excel";
			this.btnExportar.Click += new System.EventHandler(this.btnExportarCompras_Click);
			// 
			// uGridTransacciones
			// 
			this.uGridTransacciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridTransacciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTransacciones.DataSource = this.ultraDataSource4;
			appearance73.BackColor = System.Drawing.Color.White;
			this.uGridTransacciones.DisplayLayout.Appearance = appearance73;
			this.uGridTransacciones.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "Serie";
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn67.Header.VisiblePosition = 0;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn67.Width = 96;
			ultraGridColumn68.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn68.Header.VisiblePosition = 1;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 112;
			ultraGridColumn69.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn69.Header.VisiblePosition = 2;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.Width = 110;
			ultraGridColumn70.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn70.Header.VisiblePosition = 3;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 110;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn71.Header.VisiblePosition = 6;
			ultraGridColumn71.Width = 90;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn72.Header.VisiblePosition = 7;
			ultraGridColumn72.Width = 235;
			ultraGridColumn73.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn73.Header.VisiblePosition = 5;
			ultraGridColumn73.Width = 65;
			ultraGridColumn74.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn74.Header.VisiblePosition = 4;
			ultraGridColumn74.Width = 85;
			ultraGridColumn75.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance74;
			ultraGridColumn75.Format = "#,##0.00";
			ultraGridColumn75.Header.Caption = "Base 0";
			ultraGridColumn75.Header.VisiblePosition = 8;
			ultraGridColumn75.Width = 80;
			ultraGridColumn76.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellAppearance = appearance75;
			ultraGridColumn76.Format = "#,##0.00";
			ultraGridColumn76.Header.Caption = "Base I.V.A.";
			ultraGridColumn76.Header.VisiblePosition = 9;
			ultraGridColumn76.Width = 80;
			ultraGridColumn77.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn77.CellAppearance = appearance76;
			ultraGridColumn77.Format = "#,##0.00";
			ultraGridColumn77.Header.Caption = "I.V.A.";
			ultraGridColumn77.Header.VisiblePosition = 10;
			ultraGridColumn77.Width = 80;
			ultraGridColumn78.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn78.CellAppearance = appearance77;
			ultraGridColumn78.Format = "#,##0.00";
			ultraGridColumn78.Header.VisiblePosition = 11;
			ultraGridColumn78.Width = 80;
			ultraGridColumn79.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn79.CellAppearance = appearance78;
			ultraGridColumn79.Format = "#,##0.00";
			ultraGridColumn79.Header.Caption = "Ret IVA";
			ultraGridColumn79.Header.VisiblePosition = 12;
			ultraGridColumn79.Width = 80;
			ultraGridColumn80.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance79;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.Caption = "Ret Renta";
			ultraGridColumn80.Header.VisiblePosition = 13;
			ultraGridColumn80.Width = 80;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn80});
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings29.Appearance = appearance80;
			summarySettings29.DisplayFormat = "{0: #,##0.00}";
			summarySettings29.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings30.Appearance = appearance81;
			summarySettings30.DisplayFormat = "{0: #,##0.00}";
			summarySettings30.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance82.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings31.Appearance = appearance82;
			summarySettings31.DisplayFormat = "{0: #,##0.00}";
			summarySettings31.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings32.Appearance = appearance83;
			summarySettings32.DisplayFormat = "{0: #,##0.00}";
			summarySettings32.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings33.Appearance = appearance84;
			summarySettings33.DisplayFormat = "{0: #,##0.00}";
			summarySettings33.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance85.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings34.Appearance = appearance85;
			summarySettings34.DisplayFormat = "{0: #,##0.00}";
			summarySettings34.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings29,
																																															summarySettings30,
																																															summarySettings31,
																																															summarySettings32,
																																															summarySettings33,
																																															summarySettings34});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridTransacciones.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance86.ForeColor = System.Drawing.Color.Black;
			appearance86.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTransacciones.DisplayLayout.Override.ActiveRowAppearance = appearance86;
			this.uGridTransacciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance87.BackColor = System.Drawing.Color.Transparent;
			this.uGridTransacciones.DisplayLayout.Override.CardAreaAppearance = appearance87;
			this.uGridTransacciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance88.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance88.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance88.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance88.FontData.BoldAsString = "True";
			appearance88.FontData.Name = "Arial";
			appearance88.FontData.SizeInPoints = 8.5F;
			appearance88.ForeColor = System.Drawing.Color.White;
			appearance88.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTransacciones.DisplayLayout.Override.HeaderAppearance = appearance88;
			this.uGridTransacciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance89.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance89.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance89.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransacciones.DisplayLayout.Override.RowSelectorAppearance = appearance89;
			appearance90.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance90.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance90.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransacciones.DisplayLayout.Override.SelectedRowAppearance = appearance90;
			this.uGridTransacciones.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridTransacciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTransacciones.Location = new System.Drawing.Point(8, 120);
			this.uGridTransacciones.Name = "uGridTransacciones";
			this.uGridTransacciones.Size = new System.Drawing.Size(976, 288);
			this.uGridTransacciones.TabIndex = 239;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn67.DataType = typeof(int);
			ultraDataColumn68.DataType = typeof(bool);
			ultraDataColumn68.DefaultValue = false;
			ultraDataColumn69.DataType = typeof(int);
			ultraDataColumn70.DataType = typeof(int);
			ultraDataColumn73.DataType = typeof(System.DateTime);
			ultraDataColumn73.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn75.DataType = typeof(System.Double);
			ultraDataColumn76.DataType = typeof(System.Double);
			ultraDataColumn77.DataType = typeof(System.Double);
			ultraDataColumn78.DataType = typeof(System.Double);
			ultraDataColumn79.DataType = typeof(System.Double);
			ultraDataColumn80.DataType = typeof(System.Double);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn80});
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(8, 72);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(71, 16);
			this.label9.TabIndex = 350;
			this.label9.Text = "Comprobante";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbComprobantes
			// 
			appearance91.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbComprobantes.Appearance = appearance91;
			this.cmbComprobantes.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbComprobantes.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn81.Header.VisiblePosition = 0;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 1;
			ultraGridColumn82.Width = 304;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn81,
																										 ultraGridColumn82});
			this.cmbComprobantes.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbComprobantes.DisplayMember = "Comprobante";
			this.cmbComprobantes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbComprobantes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbComprobantes.Location = new System.Drawing.Point(96, 72);
			this.cmbComprobantes.Name = "cmbComprobantes";
			this.cmbComprobantes.Size = new System.Drawing.Size(304, 21);
			this.cmbComprobantes.TabIndex = 349;
			this.cmbComprobantes.ValueMember = "idComprobante";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(64, 16);
			this.label3.TabIndex = 353;
			this.label3.Text = "Transaccion";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbTransacciones
			// 
			appearance92.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTransacciones.Appearance = appearance92;
			this.cmbTransacciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTransacciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTransacciones.DataSource = this.ultraDataSource3;
			ultraGridColumn83.Header.VisiblePosition = 0;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 1;
			ultraGridColumn84.Width = 304;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn83,
																										 ultraGridColumn84});
			this.cmbTransacciones.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbTransacciones.DisplayMember = "Nombre";
			this.cmbTransacciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTransacciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTransacciones.Location = new System.Drawing.Point(96, 40);
			this.cmbTransacciones.Name = "cmbTransacciones";
			this.cmbTransacciones.Size = new System.Drawing.Size(304, 21);
			this.cmbTransacciones.TabIndex = 352;
			this.cmbTransacciones.ValueMember = "idTipoFactura";
			// 
			// ultraDataSource3
			// 
			ultraDataColumn81.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn81,
																																 ultraDataColumn82});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(992, 8);
			this.groupBox1.TabIndex = 354;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(544, 72);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 649;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnCerrarImpuestos
			// 
			this.btnCerrarImpuestos.CausesValidation = false;
			this.btnCerrarImpuestos.Enabled = false;
			this.btnCerrarImpuestos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCerrarImpuestos.Location = new System.Drawing.Point(728, 8);
			this.btnCerrarImpuestos.Name = "btnCerrarImpuestos";
			this.btnCerrarImpuestos.Size = new System.Drawing.Size(120, 23);
			this.btnCerrarImpuestos.TabIndex = 650;
			this.btnCerrarImpuestos.Text = "&Cerrar Impuestos";
			this.btnCerrarImpuestos.Click += new System.EventHandler(this.btnCerrarImpuestos_Click);
			// 
			// dtPeriodo
			// 
			this.dtPeriodo.Checked = false;
			this.dtPeriodo.CustomFormat = "MM/yyyy";
			this.dtPeriodo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtPeriodo.Location = new System.Drawing.Point(96, 8);
			this.dtPeriodo.Name = "dtPeriodo";
			this.dtPeriodo.ShowUpDown = true;
			this.dtPeriodo.Size = new System.Drawing.Size(88, 20);
			this.dtPeriodo.TabIndex = 652;
			this.dtPeriodo.ValueChanged += new System.EventHandler(this.dtPeriodo_ValueChanged);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 11);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(42, 16);
			this.label4.TabIndex = 653;
			this.label4.Text = "Periodo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(216, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 654;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtIdDebitoRenta
			// 
			this.txtIdDebitoRenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdDebitoRenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdDebitoRenta.Enabled = false;
			this.txtIdDebitoRenta.Location = new System.Drawing.Point(632, 8);
			this.txtIdDebitoRenta.Name = "txtIdDebitoRenta";
			this.txtIdDebitoRenta.PromptChar = ' ';
			this.txtIdDebitoRenta.Size = new System.Drawing.Size(16, 22);
			this.txtIdDebitoRenta.TabIndex = 659;
			this.txtIdDebitoRenta.Visible = false;
			// 
			// txtidCierreImpuestos
			// 
			this.txtidCierreImpuestos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidCierreImpuestos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCierreImpuestos.Enabled = false;
			this.txtidCierreImpuestos.Location = new System.Drawing.Point(568, 8);
			this.txtidCierreImpuestos.Name = "txtidCierreImpuestos";
			this.txtidCierreImpuestos.PromptChar = ' ';
			this.txtidCierreImpuestos.Size = new System.Drawing.Size(16, 22);
			this.txtidCierreImpuestos.TabIndex = 658;
			this.txtidCierreImpuestos.Visible = false;
			// 
			// txtidAsientoRenta
			// 
			this.txtidAsientoRenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidAsientoRenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidAsientoRenta.Enabled = false;
			this.txtidAsientoRenta.Location = new System.Drawing.Point(600, 8);
			this.txtidAsientoRenta.Name = "txtidAsientoRenta";
			this.txtidAsientoRenta.PromptChar = ' ';
			this.txtidAsientoRenta.Size = new System.Drawing.Size(16, 22);
			this.txtidAsientoRenta.TabIndex = 657;
			this.txtidAsientoRenta.Visible = false;
			// 
			// txtidAsientoIVA
			// 
			this.txtidAsientoIVA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidAsientoIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidAsientoIVA.Enabled = false;
			this.txtidAsientoIVA.Location = new System.Drawing.Point(584, 8);
			this.txtidAsientoIVA.Name = "txtidAsientoIVA";
			this.txtidAsientoIVA.PromptChar = ' ';
			this.txtidAsientoIVA.Size = new System.Drawing.Size(16, 22);
			this.txtidAsientoIVA.TabIndex = 656;
			this.txtidAsientoIVA.Visible = false;
			// 
			// txtidDebitoIVA
			// 
			this.txtidDebitoIVA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidDebitoIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidDebitoIVA.Enabled = false;
			this.txtidDebitoIVA.Location = new System.Drawing.Point(616, 8);
			this.txtidDebitoIVA.Name = "txtidDebitoIVA";
			this.txtidDebitoIVA.PromptChar = ' ';
			this.txtidDebitoIVA.Size = new System.Drawing.Size(16, 22);
			this.txtidDebitoIVA.TabIndex = 655;
			this.txtidDebitoIVA.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(648, 8);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 660;
			this.txtEstado.Visible = false;
			// 
			// frmDetalleATS1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(992, 414);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdDebitoRenta);
			this.Controls.Add(this.txtidCierreImpuestos);
			this.Controls.Add(this.txtidAsientoRenta);
			this.Controls.Add(this.txtidAsientoIVA);
			this.Controls.Add(this.txtidDebitoIVA);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtPeriodo);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.btnCerrarImpuestos);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbTransacciones);
			this.Controls.Add(this.cmbComprobantes);
			this.Controls.Add(this.uGridTransacciones);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnTalonResumen);
			this.Controls.Add(this.cmbPais);
			this.Controls.Add(this.cmbTipoFormaPago);
			this.Controls.Add(this.cmbTipoPago);
			this.Controls.Add(this.cmbTipoProveedor);
			this.Controls.Add(this.cmbComprobante);
			this.Controls.Add(this.cmbTipoRUC);
			this.Controls.Add(this.cmbCreditoTributario);
			this.Controls.Add(this.btnGeneraXML);
			this.Controls.Add(this.uGridVentas);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.uGridCompras);
			this.Name = "frmDetalleATS1";
			this.Text = "FRMVE";
			this.Load += new System.EventHandler(this.frmDetalleATS_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridCompras)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCreditoTributario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRUC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPais)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTransacciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobantes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransacciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDebitoRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCierreImpuestos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsientoRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsientoIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDebitoIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		CultureInfo us = new CultureInfo("en-US");
		private Acceso miAcceso;

		private string NombreGenerado(string stArchivo)
		{
			DateTime Fecha = (DateTime)this.dtPeriodo.Value;
			int UDia = DateTime.DaysInMonth(Fecha.Year, Fecha.Month);
			int iMes = Fecha.Month;
			int iAño = Fecha.Year;			
			DateTime dtFecha = Convert.ToDateTime(UDia.ToString() + "/" + iMes.ToString() + "/" + iAño.ToString());
			
			stArchivo += dtFecha.ToString("MMyyyy") + ".XML";
			if (Directory.Exists(@"C:\XML_ANEXO_TRANSACCIONAL\")) stArchivo = @"C:\XML_ANEXO_TRANSACCIONAL\" + stArchivo;
			return stArchivo;
		}

		private void frmDetalleATS_Load(object sender, System.EventArgs e)
		{			
			miAcceso = new Acceso(cdsSeteoF, "040701");

			this.dtPeriodo.Value = MenuLatinium.dtFechaSistema;
			this.dtPeriodo.MaxDate = MenuLatinium.dtFechaSistema;

			this.cmbComprobantes.Value = System.DBNull.Value;

			this.cmbTransacciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoFactura, Nombre From CompraNumero Where idTipoFactura In (1, 4, 5, 6) Order By Nombre");
			this.cmbComprobantes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idComprobante, Comprobante From CompraComprobante Where idComprobante In (1, 2, 3) Order By Codigo");
			
			#region Listas
//			this.cmbTipoRUC.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga 4");
//			this.cmbCreditoTributario.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCredTributario, Codigo, CredTributario From CompraCredTribut Order By Codigo");
//			this.cmbComprobante.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idComprobante, Codigo, Comprobante From CompraComprobante Order By Codigo");
//			this.cmbTipoProveedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoProveedor, Codigo, Tipo From TipoDeProveedor Order By idTipoProveedor Asc");
//			this.cmbTipoPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoDePago, Codigo, Descripcion From TiposDePagos Order By Codigo");
//			this.cmbTipoFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoFormaPago, FormaDePago From TipoFormaDePago Where Activo = 1");			
//			this.cmbPais.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idClientePais, Pais From ClientePais Order by Pais");
			#endregion Listas
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{	
			if (this.cmbTransacciones.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Transaccion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else
			{
				DateTime dtFecha = (DateTime)this.dtPeriodo.Value;
				int UDia = DateTime.DaysInMonth(dtFecha.Year, dtFecha.Month);
				int iMes = dtFecha.Month;
				int iAño = dtFecha.Year;

				DateTime Desde = Convert.ToDateTime("01/" + iMes.ToString() + "/" + iAño.ToString());
				DateTime Hasta = Convert.ToDateTime(UDia.ToString() + "/" + iMes.ToString() + "/" + iAño.ToString());

				int idComprobante = 0;

				if (this.cmbComprobantes.ActiveRow != null) idComprobante = (int)this.cmbComprobantes.Value;

				string sSQL = string.Format("Exec ConsultaTransacciones '{0}', '{1}', {2}, {3}", Desde.ToString("yyyyMMdd"), Hasta.ToString("yyyyMMdd"), (int)this.cmbTransacciones.Value, idComprobante);
				this.uGridTransacciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				this.lblContador.Text = this.uGridTransacciones.Rows.Count + " REGISTROS";
			}
		}

		private void btnGeneraXML_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;

			if (DialogResult.Yes == MessageBox.Show("¿Desea Generar el Anexo Transaccional Simplificado?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
			{
				Funcion.TiempoInicia();

				#region Variables
				DateTime Fecha = (DateTime)this.dtPeriodo.Value;
				int UDia = DateTime.DaysInMonth(Fecha.Year, Fecha.Month);
				int iMes = Fecha.Month;
				int iAño = Fecha.Year;

				DateTime dtFecha = Convert.ToDateTime("01/" + iMes.ToString() + "/" + iAño.ToString());
				DateTime dtHasta = Convert.ToDateTime(UDia.ToString() + "/" + iMes.ToString() + "/" + iAño.ToString());

				string sTipoId = "";
				string sRUCEmpresa = "";
				string sEmpresa = "";
				char Pad = '0';
				int idCompra = 0;
				int idOrigen = 0;
				int idTipoFactura = 0;
				DateTime dtFechaRegistro = DateTime.Today;
				DateTime dtFechaEmision = DateTime.Today;
				decimal dBaseNGrvIva = 0.00m;
				decimal dBase0 = 0.00m;
				decimal dBase = 0.00m;
				decimal dICE = 0.00m;
				decimal dIVA = 0.00m;
				decimal dIVABienes10 = 0.00m;
				decimal dIVAServicios20 = 0.00m;
				decimal dIVABienes = 0.00m;
				decimal dIVAServicios = 0.00m;
				decimal dIVABienesServicios = 0.00m;

				decimal dVentaTotal = 0.00m;
				decimal dVentaEstablecimiento = 0.00m;
				decimal dVentaBaseNGIVA = 0.00m;
				decimal dVentaBase0 = 0.00m;
				decimal dVentaBase = 0.00m;
				decimal dVentaIVA = 0.00m;
				decimal dVentaRetIVA = 0.00m;
				decimal dVentaRetRenta = 0.00m;

				DateTime dtFechaRet = DateTime.Today;

				DataSet oDSCompras = new DataSet();
				cdsSeteoF.Schema.Connections[0].Open();
				C1.Data.SchemaObjects.Connection miConeccion = cdsSeteoF.Schema.Connections[0];
				miConeccion = cdsSeteoF.Schema.Connections[0];
				#endregion Variables

				this.saveFileDialog1.Filter = "XML (*.xml)|*.xml";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 			
					#region IVA
					string stArchivo = saveFileDialog1.FileName; // stArchivo = @"C:\XML_ANEXO_TRANSACCIONAL\" + stArchivo; NombreGenerado("AT");
					
					XmlTextWriter miXML = new XmlTextWriter(stArchivo, Encoding.GetEncoding("UTF-8"));					
					//miXML.Formatting = Formatting.Indented;					
					miXML.WriteStartDocument(true);
					miXML.WriteStartElement("iva");
			
					#region Encabezado
					SqlDataReader drEmpresa = Funcion.rEscalarSQL(cdsSeteoF, "Select TipoId, Ruc, Empresa From Seteo", true);
					drEmpresa.Read();
					if (drEmpresa.HasRows)
					{
						sTipoId = drEmpresa.GetString(0);
						sRUCEmpresa = drEmpresa.GetString(1);
						sEmpresa = drEmpresa.GetString(2);
					}
					drEmpresa.Close();

					miXML.WriteElementString("TipoIDInformante", sTipoId);
					miXML.WriteElementString("IdInformante", sRUCEmpresa);
					miXML.WriteElementString("razonSocial", sEmpresa);
					miXML.WriteElementString("Anio", dtFecha.ToString("yyyy"));
					miXML.WriteElementString("Mes", dtFecha.ToString("MM"));

					miXML.WriteElementString("numEstabRuc", Funcion.iEscalarSQL(cdsSeteoF, "Select COUNT(*) From Bodega Where Activo = 1 And Factura = 1 And Principal = 1").ToString().PadLeft(3, Pad)); //+ Iva - DescPorcIva  - DescPorcIva0 
					string sSQLVentaTotal = string.Format("Exec dbo.ATSVentasTotales '{0}', '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					dVentaTotal = Funcion.decEscalarSQL(cdsSeteoF, sSQLVentaTotal);

					miXML.WriteElementString("totalVentas", dVentaTotal.ToString("0.00", us));// Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select SUM((SubtotalIva - Descuento) + (SubtotalExcento - Descuento0)) From Compra Where idTipoFactura = 1 And Borrar = 0 And Bodega Not In (50) And Estado In (10, 11) And Fecha >= '{0}' And Fecha <= '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"))).ToString("0.00", us));
					miXML.WriteElementString("codigoOperativo", "IVA");
					#endregion encabezado

					#region Compras
					string sSQLCompras = string.Format("Exec ATSCompras '{0}', '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					SqlDataAdapter oDACompras = new SqlDataAdapter(sSQLCompras, (SqlConnection) miConeccion.DbConnection);
					oDACompras.Fill(oDSCompras);

					if (oDSCompras.Tables[0].Rows.Count > 1)
					{
						miXML.WriteStartElement("compras");

						foreach(DataRow oDRCompras in oDSCompras.Tables[0].Rows)
						{
							#region Inicia Variables
							idCompra = (int)oDRCompras["idCompra"];
							idOrigen = (int)oDRCompras["idOrigen"];
							idTipoFactura = (int)oDRCompras["idTipoFactura"];
							dBaseNGrvIva = 0.00m;
							dBase0 = 0.00m;
							dBase = 0.00m;							
							dICE = 0.00m;
							dIVA = 0.00m;
							dIVABienes10 = 0.00m;
							dIVAServicios20 = 0.00m;
							dIVABienes = 0.00m;
							dIVAServicios = 0.00m;
							dIVABienesServicios = 0.00m;
							#endregion Inicia Variables

							miXML.WriteStartElement("detalleCompras");
					
							#region Generales
							miXML.WriteElementString("codSustento", oDRCompras["Sustento"].ToString().Trim());
							miXML.WriteElementString("tpIdProv", oDRCompras["TipoRuc"].ToString().Trim());
							miXML.WriteElementString("idProv", oDRCompras["Ruc"].ToString().Trim());
							miXML.WriteElementString("tipoComprobante", oDRCompras["Comprobante"].ToString().Trim());
					
							miXML.WriteElementString("parteRel", "NO");

							if (oDRCompras["TipoRuc"].ToString().Trim() == "03")
							{
								miXML.WriteElementString("tipoProv", oDRCompras["TipoProveedor"].ToString().Trim());
							}

							dtFechaRegistro = (DateTime)oDRCompras["FechaRegistro"];
					
							miXML.WriteElementString("fechaRegistro", dtFechaRegistro.ToString("dd/MM/yyyy"));

							if (oDRCompras["TipoRuc"].ToString().Trim() != "03")
							{
								miXML.WriteElementString("establecimiento", oDRCompras["SerieFactura"].ToString().Trim().Substring(0, 3));
								miXML.WriteElementString("puntoEmision", oDRCompras["SerieFactura"].ToString().Trim().Substring(3, 3));
							}
							else
							{
								miXML.WriteElementString("establecimiento", "999");
								miXML.WriteElementString("puntoEmision", "999");
							}

							miXML.WriteElementString("secuencial", oDRCompras["Numero"].ToString().Trim());

							dtFechaEmision = (DateTime)oDRCompras["Fecha"];
					
							miXML.WriteElementString("fechaEmision", dtFechaEmision.ToString("dd/MM/yyyy"));
							
							if (oDRCompras["TipoRuc"].ToString().Trim() != "03") miXML.WriteElementString("autorizacion", oDRCompras["AutFactura"].ToString());
							else miXML.WriteElementString("autorizacion", "9999999999");
							#endregion Generales

							#region Base IVA
							dBaseNGrvIva = (decimal)oDRCompras["BaseNoImp"];
							dBase0 = (decimal)oDRCompras["Base0"];
							dBase = (decimal)oDRCompras["Base"];
							dICE = Convert.ToDecimal(oDRCompras["Ice"]);
							dIVA = (decimal)oDRCompras["Iva"];

							miXML.WriteElementString("baseNoGraIva", dBaseNGrvIva.ToString("0.00", us));
							miXML.WriteElementString("baseImponible", dBase0.ToString("0.00", us));
							miXML.WriteElementString("baseImpGrav", dBase.ToString("0.00", us));
							miXML.WriteElementString("baseImpExe", dBaseNGrvIva.ToString("0.00", us));
					
							miXML.WriteElementString("montoIce", dICE.ToString("0.00", us));
							miXML.WriteElementString("montoIva", dIVA.ToString("0.00", us));
							#endregion Base IVA

							#region Retenciones de IVA
							dIVABienes10 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ATSRetencionIVA {0}, 10", idCompra));
							dIVAServicios20 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ATSRetencionIVA {0}, 20", idCompra));
							dIVABienes = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ATSRetencionIVA {0}, 30", idCompra));
							dIVAServicios = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ATSRetencionIVA {0}, 70", idCompra));;
							dIVABienesServicios = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ATSRetencionIVA {0}, 100", idCompra));;
							decimal dBaseReembolsos = 0.00m;
							miXML.WriteElementString("valRetBien10", dIVABienes10.ToString("0.00", us));
							miXML.WriteElementString("valRetServ20", dIVAServicios20.ToString("0.00", us));
							miXML.WriteElementString("valorRetBienes", dIVABienes.ToString("0.00", us));
							miXML.WriteElementString("valorRetServicios", dIVAServicios.ToString("0.00", us));
							miXML.WriteElementString("valRetServ100", dIVABienesServicios.ToString("0.00", us));
							miXML.WriteElementString("totbasesImpReemb", dBaseReembolsos.ToString("0.00", us));
							#endregion Retenciones de IVA

							#region Pago Exterior o Local
							miXML.WriteStartElement("pagoExterior");
							miXML.WriteElementString("pagoLocExt", oDRCompras["TipoDePago"].ToString());

							if (oDRCompras["TipoDePago"].ToString() == "01")
							{
								miXML.WriteElementString("paisEfecPago", "NA");
								miXML.WriteElementString("aplicConvDobTrib", "NA");
								miXML.WriteElementString("pagExtSujRetNorLeg", "NA");
								miXML.WriteElementString("pagoRegFis", "NA");
							}
							else
							{
								miXML.WriteElementString("paisEfecPago", oDRCompras["CodigoPais"].ToString());
								string ACDT = "NO";
								miXML.WriteElementString("aplicConvDobTrib", ACDT);
								if (ACDT == "NO") miXML.WriteElementString("pagExtSujRetNorLeg", "SI");
							}				
							miXML.WriteEndElement();
							#endregion Pago Exterior o Local

							#region Forma de pago
							if ((decimal)oDRCompras["Total"] >= 1000.00m)
							{
								miXML.WriteStartElement("formasDePago");
								miXML.WriteElementString("formaPago", oDRCompras["TipoFormaPago"].ToString());
								miXML.WriteEndElement();
							}
							#endregion Forma de pago					

							#region Retencion de Renta
							miXML.WriteStartElement("air");
							
							string sSQLAIR = string.Format("Exec ATSRetencionRenta {0}", idCompra);
							SqlDataReader drAIR = Funcion.rEscalarSQL(cdsSeteoF, sSQLAIR, true);

							while(drAIR.Read())
							{
								miXML.WriteStartElement("detalleAir");
					
								miXML.WriteElementString("codRetAir", drAIR.GetString(0).Trim());
								miXML.WriteElementString("baseImpAir", drAIR.GetDecimal(1).ToString("0.00", us));
								miXML.WriteElementString("porcentajeAir", drAIR.GetDecimal(2).ToString("0.00", us));
								miXML.WriteElementString("valRetAir", drAIR.GetDecimal(3).ToString("0.00", us));
					
								miXML.WriteEndElement();
							}
							drAIR.Close();

							miXML.WriteEndElement();							
							#endregion Retencion de Renta

							#region Comprobante de Retencion
							if (Convert.ToDecimal(oDRCompras["Retenido"]) > 0.00m)
							{
								miXML.WriteElementString("estabRetencion1", oDRCompras["Serie"].ToString().Trim().Substring(0,3));
								miXML.WriteElementString("ptoEmiRetencion1", oDRCompras["Serie"].ToString().Trim().Substring(3,3));
								miXML.WriteElementString("secRetencion1", oDRCompras["NumRet"].ToString().Trim());
								miXML.WriteElementString("autRetencion1", oDRCompras["Autorizacion"].ToString().Trim());
								dtFechaRet = (DateTime)oDRCompras["FecRet"];
								miXML.WriteElementString("fechaEmiRet1", dtFechaRet.ToString("dd/MM/yyyy"));
							}
							#endregion Comprobante de Retencion

							#region Documento modificado
							if (idTipoFactura != 4)
							{
								if (idOrigen > 0)
								{
									string sSQLMod = string.Format("Select cc.Codigo, c.SerieFactura, Numero, AutFactura From Compra c Inner Join CompraComprobante cc On cc.idComprobante = c.idComprobante Where idTipoFactura = 4 And idCompra = {0}", idOrigen);
									SqlDataReader drModificado = Funcion.rEscalarSQL(cdsSeteoF, sSQLMod, true);
									drModificado.Read();
									if (drModificado.HasRows)
									{
										miXML.WriteElementString("docModificado", drModificado.GetString(0).ToString().Trim());// int.Parse(stDevCompTipo[k]).ToString());
										miXML.WriteElementString("estabModificado", drModificado.GetString(1).ToString().Trim().Substring(0,3));
										miXML.WriteElementString("ptoEmiModificado", drModificado.GetString(1).ToString().Trim().Substring(3,3));
										miXML.WriteElementString("secModificado", drModificado.GetString(2).ToString().Trim());
										miXML.WriteElementString("autModificado", drModificado.GetString(3).ToString().Trim());
									}
									drModificado.Close();
								}
								else
								{
									string sSQLMod = string.Format("Select cc.Codigo, c.SerieFactura, Numero, AutFactura From Compra c Inner Join CompraComprobante cc On cc.idComprobante = c.idComprobante Where idCompra = (Select idOrigen From AutorizacionDePagos Where idTipo = 4 And idAutorizacionPago = (Select idAutorizacionPago From AutPagoNC where idCompra = {0}))", idCompra);
									SqlDataReader drModificado = Funcion.rEscalarSQL(cdsSeteoF, sSQLMod, true);
									drModificado.Read();
									if (drModificado.HasRows)
									{
										miXML.WriteElementString("docModificado", drModificado.GetString(0).ToString().Trim());// int.Parse(stDevCompTipo[k]).ToString());
										miXML.WriteElementString("estabModificado", drModificado.GetString(1).ToString().Trim().Substring(0,3));
										miXML.WriteElementString("ptoEmiModificado", drModificado.GetString(1).ToString().Trim().Substring(3,3));
										miXML.WriteElementString("secModificado", drModificado.GetString(2).ToString().Trim());
										miXML.WriteElementString("autModificado", drModificado.GetString(3).ToString().Trim());
									}
									drModificado.Close();
								}
							}
							#endregion Documento modificado

							miXML.WriteEndElement();
						}			
						miXML.WriteEndElement();
					}
					#endregion Compras

					#region Ventas
					miXML.WriteStartElement("ventas");

					string sSQLVentas = string.Format("Exec ATSVentas '{0}', '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					SqlDataReader drVentas = Funcion.rEscalarSQL(cdsSeteoF, sSQLVentas, true);

					while (drVentas.Read())
					{
						dVentaBaseNGIVA = 0;
						dVentaBase0 = 0;
						dVentaBase = 0;
						dIVA = 0;
						dVentaRetIVA = 0;
						dVentaRetRenta = 0;
						dICE = 0.00m;

						miXML.WriteStartElement("detalleVentas");

						miXML.WriteElementString("tpIdCliente", drVentas.GetString(0).ToString().Trim());
						miXML.WriteElementString("idCliente", drVentas.GetString(1).ToString().Trim());
						miXML.WriteElementString("parteRelVtas", "NO");
						miXML.WriteElementString("tipoComprobante", drVentas.GetString(2).ToString().Trim());
						miXML.WriteElementString("numeroComprobantes", drVentas.GetValue(3).ToString().Trim());

						dVentaBaseNGIVA = drVentas.GetDecimal(4);
						dVentaBase0 = drVentas.GetDecimal(5);
						dVentaBase = drVentas.GetDecimal(6);
						dVentaIVA = drVentas.GetDecimal(7);
						dVentaRetIVA = drVentas.GetDecimal(8);
						dVentaRetRenta = drVentas.GetDecimal(9);
						
						miXML.WriteElementString("baseNoGraIva", dVentaBaseNGIVA.ToString("0.00", us));
						miXML.WriteElementString("baseImponible", dVentaBase0.ToString("0.00", us));
						miXML.WriteElementString("baseImpGrav", dVentaBase.ToString("0.00", us));
						miXML.WriteElementString("montoIva", dVentaIVA.ToString("0.00", us));
						miXML.WriteElementString("montoIce", dICE.ToString("0.00", us));
						miXML.WriteElementString("valorRetIva", dVentaRetIVA.ToString("0.00", us));
						miXML.WriteElementString("valorRetRenta", dVentaRetRenta.ToString("0.00", us));

						miXML.WriteEndElement();
					}
					drVentas.Close();

					miXML.WriteEndElement();
					#endregion Ventas

					#region Ventas Establecimiento
					int iContBodegas = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Bodega"));

					for (int l = 1; l <= iContBodegas; l++)
					{	
						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Activo Bit = 0 If Exists(Select Serie From Bodega Where Activo = 1 And Factura = 1 And Principal = 1 And Bodega = {0}) Set @Activo = 1 Select @Activo", l), true))
						{
							string sSerie = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Serie From Bodega Where Activo = 1 And Factura = 1 And Principal = 1 And Bodega = {0}", l), true);
							
							if (sSerie.Length > 0)
							{
								miXML.WriteStartElement("ventaEst");

								miXML.WriteElementString("codEstab", Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select SUBSTRING(Serie, 1, 3) From Bodega Where Bodega = {0} And Principal = 1 And Activo = 1 And Factura = 1", l)));

								string sSQLVentaEstab = string.Format("Exec dbo.ATSVentasEstablecimiento '{0}', '{1}', '{2}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), sSerie);
								dVentaEstablecimiento = Funcion.decEscalarSQL(cdsSeteoF, sSQLVentaEstab);

								miXML.WriteElementString("ventasEstab", dVentaEstablecimiento.ToString("0.00", us));//  Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(SUM((SubtotalIva - Descuento) + (SubtotalExcento - Descuento0)), 0) From Compra Where idTipoFactura = 1 And Borrar = 0 And Estado In (10, 11) And Fecha >= '{1}' And Fecha <= '{2}' And Bodega = {0}", l, dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"))).ToString("0.00", us));
					
								miXML.WriteEndElement();
							}
						}
					}
					miXML.WriteEndElement();
					#endregion Ventas Establecimiento

					#region Comprobantes Anulados
					#region Facturas - Notas de crédito	
					string sSQLFacAnuladas = string.Format("Exec ATSComprobantesAnulados '{0}', '{1}', 0", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					SqlDataReader drFacAnuladas = Funcion.rEscalarSQL(cdsSeteoF, sSQLFacAnuladas, true);

					miXML.WriteStartElement("anulados");

					while (drFacAnuladas.Read())
					{
						miXML.WriteStartElement("detalleAnulados");
						
						miXML.WriteElementString("tipoComprobante", drFacAnuladas.GetString(0).ToString().Trim());
						miXML.WriteElementString("establecimiento", drFacAnuladas.GetString(1).ToString().Trim().Substring(0, 3));
						miXML.WriteElementString("puntoEmision", drFacAnuladas.GetString(1).ToString().Trim().Substring(3, 3));
						miXML.WriteElementString("secuencialInicio", drFacAnuladas.GetString(2).ToString().Trim());
						miXML.WriteElementString("secuencialFin", drFacAnuladas.GetString(2).ToString().Trim());
						miXML.WriteElementString("autorizacion", drFacAnuladas.GetString(3).ToString().Trim());

						miXML.WriteEndElement();
					}
					drFacAnuladas.Close();
					#endregion Facturas - Notas de crédito
			
					#region Retenciones
					string sSQLRetAnuladas = string.Format("Exec ATSComprobantesAnulados '{0}', '{1}', 1", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					SqlDataReader drRetAnuladas = Funcion.rEscalarSQL(cdsSeteoF, sSQLRetAnuladas, true);

					while (drRetAnuladas.Read())
					{
						miXML.WriteStartElement("detalleAnulados");
				
						miXML.WriteElementString("tipoComprobante", drRetAnuladas.GetString(0).ToString().Trim());
						miXML.WriteElementString("establecimiento", drRetAnuladas.GetString(1).ToString().Trim().Substring(0, 3));
						miXML.WriteElementString("puntoEmision", drRetAnuladas.GetString(1).ToString().Trim().Substring(3, 3));
						miXML.WriteElementString("secuencialInicio", drRetAnuladas.GetString(2).ToString().Trim());
						miXML.WriteElementString("secuencialFin", drRetAnuladas.GetString(2).ToString().Trim());
						miXML.WriteElementString("autorizacion", drRetAnuladas.GetString(3).ToString().Trim());

						miXML.WriteEndElement();
					}
					drRetAnuladas.Close();
					#endregion Retenciones
					#endregion Comprobantes Anulados

					miXML.WriteEndElement();

					miXML.WriteEndElement();
					miXML.WriteEndDocument();

					miXML.Flush();
					miXML.Close();
					#endregion IVA
			
					MessageBox.Show("Anexo Transaccional Simplificado " + stArchivo + " Generado " + Funcion.Tiempo("En"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}

			Cursor = Cursors.Default;
		}

		private void btnTalonResumen_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime)this.dtPeriodo.Value;
			int UDia = DateTime.DaysInMonth(dtFecha.Year, dtFecha.Month);
			int iMes = dtFecha.Month;
			int iAño = dtFecha.Year;

			DateTime dtFechaDesde = Convert.ToDateTime("01/" + iMes.ToString() + "/" + iAño.ToString());
			DateTime dtFechaHasta = Convert.ToDateTime(UDia.ToString() + "/" + iMes.ToString() + "/" + iAño.ToString());

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@Desde";
			discreteVal.Value = dtFechaDesde;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField0 = new ParameterField ();
			ParameterDiscreteValue discreteVal0 = new ParameterDiscreteValue ();
			paramField0.ParameterFieldName = "@Hasta";
			discreteVal0.Value = dtFechaHasta;
			paramField0.CurrentValues.Add (discreteVal0);
			paramFields.Add (paramField0);

			Reporte miReporte = new Reporte("TalonResumen.rpt");
			miReporte.MdiParent = MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Titulo("Talón Resumen Anexo Transaccional Simplificado");
			miReporte.Show();
		}

		private void btnExportarCompras_Click(object sender, System.EventArgs e)
		{
			if (this.uGridTransacciones.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				 
				ultraGridExcelExporter1.Export(this.uGridTransacciones, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}

		private void btnCerrarImpuestos_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value == 0)
			{
				if (DialogResult.Yes == MessageBox.Show("El Cierre de Impuestos no puede ser revertido\n\n ¿Esta Seguro de Cerrar Impuestos?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				{
					DateTime dtFecha = (DateTime)this.dtPeriodo.Value;
					int UDia = DateTime.DaysInMonth(dtFecha.Year, dtFecha.Month);
					int iMes = dtFecha.Month;
					int iAño = dtFecha.Year;

					DateTime dtDesde = Convert.ToDateTime("01/" + iMes.ToString() + "/" + iAño.ToString());
					DateTime dtHasta = Convert.ToDateTime(UDia.ToString() + "/" + iMes.ToString() + "/" + iAño.ToString());

					string sSQLIVA = string.Format("Exec GeneraAsientoImpuestosIVA '{0}', '{1}', '{2}'", (int)this.txtidCierreImpuestos.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					this.txtidAsientoIVA.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLIVA);

					string sSQLRenta = string.Format("Exec GeneraAsientoImpuestosRenta '{0}', '{1}', '{2}'", (int)this.txtidCierreImpuestos.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					this.txtidAsientoRenta.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLRenta);

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update CierreDeImpuestos Set Estado = 1, FechaCierre = GETDATE(), Usuario = SUSER_SNAME() Where idCierreDeImpuestos = {0}", (int)this.txtidCierreImpuestos.Value));

					this.txtEstado.Value = 0;
					this.lblEstado.Text = "ABIERTO";
				}
			}
			else
			{
				if (DialogResult.Yes == MessageBox.Show("Al Abrir Impuestos se Anularan los Debitos Bancarios del Pago de Impuestos ¿Esta Seguro de Abrir Impuestos?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				{
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update CierreDeImpuestos Set Estado = 0 Where idCierreDeImpuestos = {0}", (int)this.txtidCierreImpuestos.Value));
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Asiento Set Borrar = 1 Where idAsiento = {0}", (int)this.txtidDebitoIVA.Value));
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Asiento Set Borrar = 1 Where idAsiento = {0}", (int)this.txtIdDebitoRenta.Value));

					this.txtEstado.Value = 1;
					this.lblEstado.Text = "CERRADO";
				}
			}
		}

		private void dtPeriodo_ValueChanged(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime)this.dtPeriodo.Value;
			int UDia = DateTime.DaysInMonth(dtFecha.Year, dtFecha.Month);
			int iMes = dtFecha.Month;
			int iAño = dtFecha.Year;
		
			DateTime dtHasta = Convert.ToDateTime(UDia.ToString() + "/" + iMes.ToString() + "/" + iAño.ToString());
			string sSQL = string.Format("Select idCierreDeImpuestos, idAsientoIVA, idAsientoRenta, idDebitoIVA, idDebitoRenta, Estado, Case Estado When 0 Then 'ABIERTO' When 1 Then 'CERRADO' End From CierreDeImpuestos Where CONVERT(Date, Periodo) = '{0}'", dtHasta.ToString("yyyyMMdd"));
			SqlDataReader drImpuestos = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drImpuestos.Read();
			if (drImpuestos.HasRows)
			{
				this.txtidCierreImpuestos.Value = drImpuestos.GetInt32(0);
				this.txtidAsientoIVA.Value = drImpuestos.GetInt32(1);
				this.txtidAsientoRenta.Value = drImpuestos.GetInt32(2);
				this.txtidDebitoIVA.Value = drImpuestos.GetInt32(3);
				this.txtIdDebitoRenta.Value = drImpuestos.GetInt32(4);
				this.txtEstado.Value = drImpuestos.GetInt32(5);
				this.lblEstado.Text = drImpuestos.GetString(6);
			}
			drImpuestos.Close();

			if ((int)this.txtEstado.Value == 0) 
			{
				this.btnCerrarImpuestos.Text = "Cerrar Impuestos";
				this.btnCerrarImpuestos.Enabled = true; 				
			}
			else if ((int)this.txtEstado.Value == 1)
			{
				this.btnCerrarImpuestos.Text = "Abrir Impuestos";
				if (miAcceso.BAbrirImpuestos) this.btnCerrarImpuestos.Enabled = true; else this.btnCerrarImpuestos.Enabled = false;
			}
		}
	}
}

