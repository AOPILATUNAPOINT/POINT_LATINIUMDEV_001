using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using C1.Data;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCapturaVouchers.
	/// </summary>
	public class frmCapturaVouchers : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private System.Windows.Forms.Label lblCuenta;
		private System.Windows.Forms.Label lblCtaCorriente;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label lblCaducidad;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVouchers;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDeposito;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBancos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTarjetaCredito;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPlanes;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPlazo;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.ComponentModel.IContainer components;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private DevExpress.XtraEditors.RadioGroup rdgTipoConsulta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdDeposito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private DevExpress.XtraEditors.SimpleButton btnBuscar;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSuma;
		private System.Windows.Forms.Label lblContadorSeleccionados;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnDepositar;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnReporte;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbLocal;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosBancarios;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;

		private C1DataRow drSeteoF;

		public frmCapturaVouchers()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Column 0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capturado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco", -1, "cmbBancos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta", -1, "cmbTarjetaCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlan", -1, "cmbPlanes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlazo", -1, "cmbPlazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcComision");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionIVA");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCobrar");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetencionIVA");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 13, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retenido", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retenido", 14, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comision", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comision", 16, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorCobrar", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorCobrar", 18, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ComisionIVA", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ComisionIVA", 17, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetencionIVA", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetencionIVA", 19, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Column 0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capturado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCobrar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RetencionIVA");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlan");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCapturaVouchers));
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblCuenta = new System.Windows.Forms.Label();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCtaCorriente = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtDeposito = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCaducidad = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridVouchers = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbBancos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTarjetaCredito = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbPlanes = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbPlazo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.rdgTipoConsulta = new DevExpress.XtraEditors.RadioGroup();
			this.txtIdDeposito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.label3 = new System.Windows.Forms.Label();
			this.txtSuma = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblContadorSeleccionados = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnDepositar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnReporte = new System.Windows.Forms.Button();
			this.label6 = new System.Windows.Forms.Label();
			this.cmbLocal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.txtGastosBancarios = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnExportar = new System.Windows.Forms.Button();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchers)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBancos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjetaCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoConsulta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSuma)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosBancarios)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// ultraDataSource4
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance1;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource5;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 224;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(296, 314);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(232, 21);
			this.cmbCuenta.TabIndex = 8;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbCuenta_KeyPress);
			this.cmbCuenta.ValueChanged += new System.EventHandler(this.cmbCuenta_ValueChanged);
			this.cmbCuenta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCuenta_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// lblCuenta
			// 
			this.lblCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblCuenta.AutoSize = true;
			this.lblCuenta.BackColor = System.Drawing.Color.Transparent;
			this.lblCuenta.Location = new System.Drawing.Point(232, 317);
			this.lblCuenta.Name = "lblCuenta";
			this.lblCuenta.Size = new System.Drawing.Size(41, 16);
			this.lblCuenta.TabIndex = 199;
			this.lblCuenta.Text = "Cuenta";
			this.lblCuenta.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtConcepto
			// 
			this.txtConcepto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtConcepto.Appearance = appearance2;
			this.txtConcepto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Location = new System.Drawing.Point(96, 345);
			this.txtConcepto.MaxLength = 200;
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(432, 21);
			this.txtConcepto.TabIndex = 9;
			this.txtConcepto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtConcepto_KeyPress);
			// 
			// lblCtaCorriente
			// 
			this.lblCtaCorriente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblCtaCorriente.AutoSize = true;
			this.lblCtaCorriente.BackColor = System.Drawing.Color.Transparent;
			this.lblCtaCorriente.Location = new System.Drawing.Point(16, 379);
			this.lblCtaCorriente.Name = "lblCtaCorriente";
			this.lblCtaCorriente.Size = new System.Drawing.Size(63, 16);
			this.lblCtaCorriente.TabIndex = 195;
			this.lblCtaCorriente.Text = "Depósito N.";
			this.lblCtaCorriente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(16, 348);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(53, 16);
			this.lblNumero.TabIndex = 193;
			this.lblNumero.Text = "Concepto";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDeposito
			// 
			this.txtDeposito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDeposito.Appearance = appearance3;
			this.txtDeposito.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDeposito.Location = new System.Drawing.Point(96, 376);
			this.txtDeposito.MaxLength = 15;
			this.txtDeposito.Name = "txtDeposito";
			this.txtDeposito.Size = new System.Drawing.Size(224, 21);
			this.txtDeposito.TabIndex = 10;
			this.txtDeposito.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDeposito_KeyPress);
			// 
			// lblCaducidad
			// 
			this.lblCaducidad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblCaducidad.AutoSize = true;
			this.lblCaducidad.Location = new System.Drawing.Point(16, 317);
			this.lblCaducidad.Name = "lblCaducidad";
			this.lblCaducidad.Size = new System.Drawing.Size(36, 16);
			this.lblCaducidad.TabIndex = 191;
			this.lblCaducidad.Text = "Fecha";
			this.lblCaducidad.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			this.dtFecha.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance4;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy HH:mm";
			this.dtFecha.Location = new System.Drawing.Point(96, 314);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 7;
			this.dtFecha.Value = new System.DateTime(2013, 8, 12, 0, 0, 0, 0);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// uGridVouchers
			// 
			this.uGridVouchers.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVouchers.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVouchers.DataSource = this.ultraDataSource1;
			appearance5.BackColor = System.Drawing.Color.White;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchers.DisplayLayout.Appearance = appearance5;
			ultraGridColumn3.AutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.None;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 67;
			ultraGridColumn4.Header.Caption = "";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 17;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 20;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 73;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.Caption = "Almacén";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 130;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 60;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.Caption = "Banco";
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 130;
			ultraGridColumn11.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.Caption = "Tarjeta";
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn11.Width = 130;
			ultraGridColumn12.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.Caption = "Plan";
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn12.Width = 110;
			ultraGridColumn13.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.Caption = "Plazo";
			ultraGridColumn13.Header.VisiblePosition = 10;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn13.Width = 100;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Width = 40;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Width = 65;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance6;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 13;
			ultraGridColumn16.Width = 65;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance7;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 65;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance8;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "% Com.";
			ultraGridColumn18.Header.VisiblePosition = 14;
			ultraGridColumn18.Width = 50;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance9;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Comisión";
			ultraGridColumn19.Header.VisiblePosition = 15;
			ultraGridColumn19.Width = 65;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance10;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "Comision IVA";
			ultraGridColumn20.Header.VisiblePosition = 18;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 65;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance11;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Cobro";
			ultraGridColumn21.Header.VisiblePosition = 19;
			ultraGridColumn21.Width = 65;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance12;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "Ret IVA";
			ultraGridColumn22.Header.VisiblePosition = 17;
			ultraGridColumn22.Width = 65;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn22});
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance13;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance14;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance15;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance16;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance17;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.Hidden = true;
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance18;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridVouchers.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchers.DisplayLayout.GroupByBox.Appearance = appearance19;
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchers.DisplayLayout.GroupByBox.BandLabelAppearance = appearance20;
			this.uGridVouchers.DisplayLayout.GroupByBox.Prompt = " ";
			appearance21.BackColor = System.Drawing.SystemColors.Control;
			appearance21.BackColor2 = System.Drawing.SystemColors.Control;
			appearance21.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance21.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance21.ForeColor = System.Drawing.Color.Black;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchers.DisplayLayout.GroupByBox.PromptAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.uGridVouchers.DisplayLayout.Override.CardAreaAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance23.FontData.BoldAsString = "True";
			appearance23.FontData.Name = "Arial";
			appearance23.FontData.SizeInPoints = 7.8F;
			appearance23.ForeColor = System.Drawing.Color.White;
			appearance23.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVouchers.DisplayLayout.Override.HeaderAppearance = appearance23;
			this.uGridVouchers.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchers.DisplayLayout.Override.RowSelectorAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchers.DisplayLayout.Override.SelectedRowAppearance = appearance25;
			this.uGridVouchers.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVouchers.Location = new System.Drawing.Point(8, 88);
			this.uGridVouchers.Name = "uGridVouchers";
			this.uGridVouchers.Size = new System.Drawing.Size(1272, 200);
			this.uGridVouchers.TabIndex = 6;
			this.uGridVouchers.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVouchers_CellChange);
			this.uGridVouchers.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridVouchers_BeforeCellUpdate);
			this.uGridVouchers.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVouchers_KeyDown);
			this.uGridVouchers.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVouchers_AfterCellUpdate);
			this.uGridVouchers.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVouchers_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn8.DataType = typeof(bool);
			ultraDataColumn8.DefaultValue = false;
			ultraDataColumn9.DataType = typeof(bool);
			ultraDataColumn9.DefaultValue = false;
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn20.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn21.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn22.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn23.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn24.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn25.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn26.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn26});
			// 
			// txtBuscar
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance26;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(496, 40);
			this.txtBuscar.MaxLength = 15;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(248, 21);
			this.txtBuscar.TabIndex = 4;
			this.txtBuscar.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscar_KeyPress);
			this.txtBuscar.ValueChanged += new System.EventHandler(this.txtBuscar_ValueChanged);
			this.txtBuscar.Enter += new System.EventHandler(this.txtBuscar_Enter);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(432, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 206;
			this.label1.Text = "Buscar";
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 296);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1292, 8);
			this.groupBox1.TabIndex = 207;
			this.groupBox1.TabStop = false;
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
			// cmbBancos
			// 
			this.cmbBancos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBancos.DataSource = this.ultraDataSource6;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbBancos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBancos.DisplayMember = "Nombre";
			this.cmbBancos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBancos.Location = new System.Drawing.Point(32, 176);
			this.cmbBancos.Name = "cmbBancos";
			this.cmbBancos.Size = new System.Drawing.Size(72, 72);
			this.cmbBancos.TabIndex = 208;
			this.cmbBancos.ValueMember = "idBanco";
			this.cmbBancos.Visible = false;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn27.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn27,
																																 ultraDataColumn28});
			// 
			// cmbTarjetaCredito
			// 
			this.cmbTarjetaCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTarjetaCredito.DataSource = this.ultraDataSource7;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 200;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbTarjetaCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTarjetaCredito.DisplayMember = "Nombre";
			this.cmbTarjetaCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTarjetaCredito.Location = new System.Drawing.Point(120, 176);
			this.cmbTarjetaCredito.Name = "cmbTarjetaCredito";
			this.cmbTarjetaCredito.Size = new System.Drawing.Size(72, 72);
			this.cmbTarjetaCredito.TabIndex = 209;
			this.cmbTarjetaCredito.ValueMember = "idTarjeta";
			this.cmbTarjetaCredito.Visible = false;
			// 
			// ultraDataSource7
			// 
			ultraDataColumn29.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn29,
																																 ultraDataColumn30});
			// 
			// cmbPlanes
			// 
			this.cmbPlanes.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlanes.DataSource = this.ultraDataSource8;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 180;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbPlanes.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbPlanes.DisplayMember = "Descripcion";
			this.cmbPlanes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlanes.Location = new System.Drawing.Point(208, 176);
			this.cmbPlanes.Name = "cmbPlanes";
			this.cmbPlanes.Size = new System.Drawing.Size(72, 72);
			this.cmbPlanes.TabIndex = 210;
			this.cmbPlanes.ValueMember = "idPlan";
			this.cmbPlanes.Visible = false;
			// 
			// ultraDataSource8
			// 
			ultraDataColumn31.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// cmbPlazo
			// 
			this.cmbPlazo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlazo.DataSource = this.ultraDataSource9;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 180;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30});
			this.cmbPlazo.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbPlazo.DisplayMember = "Descripcion";
			this.cmbPlazo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlazo.Location = new System.Drawing.Point(296, 176);
			this.cmbPlazo.Name = "cmbPlazo";
			this.cmbPlazo.Size = new System.Drawing.Size(72, 72);
			this.cmbPlazo.TabIndex = 211;
			this.cmbPlazo.ValueMember = "idPlazo";
			this.cmbPlazo.Visible = false;
			// 
			// ultraDataSource9
			// 
			ultraDataColumn33.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn33,
																																 ultraDataColumn34});
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(432, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(36, 16);
			this.label2.TabIndex = 213;
			this.label2.Text = "Banco";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn35.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn35,
																																 ultraDataColumn36});
			// 
			// rdgTipoConsulta
			// 
			this.rdgTipoConsulta.EditValue = 0;
			this.rdgTipoConsulta.Location = new System.Drawing.Point(8, 8);
			this.rdgTipoConsulta.Name = "rdgTipoConsulta";
			// 
			// rdgTipoConsulta.Properties
			// 
			this.rdgTipoConsulta.Properties.Appearance.BackColor = System.Drawing.SystemColors.Control;
			this.rdgTipoConsulta.Properties.Appearance.Options.UseBackColor = true;
			this.rdgTipoConsulta.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
			this.rdgTipoConsulta.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Pendiente"),
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Capturado"),
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Depositado")});
			this.rdgTipoConsulta.Size = new System.Drawing.Size(256, 24);
			this.rdgTipoConsulta.TabIndex = 0;
			this.rdgTipoConsulta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.rdgTipoConsulta_KeyDown);
			this.rdgTipoConsulta.SelectedIndexChanged += new System.EventHandler(this.rdgTipoConsulta_SelectedIndexChanged);
			// 
			// txtIdDeposito
			// 
			this.txtIdDeposito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdDeposito.Location = new System.Drawing.Point(464, 376);
			this.txtIdDeposito.Name = "txtIdDeposito";
			this.txtIdDeposito.PromptChar = ' ';
			this.txtIdDeposito.Size = new System.Drawing.Size(64, 21);
			this.txtIdDeposito.TabIndex = 216;
			this.txtIdDeposito.Visible = false;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Location = new System.Drawing.Point(392, 376);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(64, 21);
			this.txtIdAsiento.TabIndex = 218;
			this.txtIdAsiento.Visible = false;
			// 
			// btnBuscar
			// 
			this.btnBuscar.Location = new System.Drawing.Point(776, 40);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(64, 21);
			this.btnBuscar.TabIndex = 5;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(600, 348);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(29, 16);
			this.label3.TabIndex = 235;
			this.label3.Text = "Total";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSuma
			// 
			this.txtSuma.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSuma.Appearance = appearance27;
			this.txtSuma.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSuma.Enabled = false;
			this.txtSuma.FormatString = "#,##0.00";
			this.txtSuma.Location = new System.Drawing.Point(656, 345);
			this.txtSuma.Name = "txtSuma";
			this.txtSuma.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSuma.PromptChar = ' ';
			this.txtSuma.Size = new System.Drawing.Size(112, 21);
			this.txtSuma.TabIndex = 236;
			// 
			// lblContadorSeleccionados
			// 
			this.lblContadorSeleccionados.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblContadorSeleccionados.AutoSize = true;
			this.lblContadorSeleccionados.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContadorSeleccionados.Location = new System.Drawing.Point(792, 344);
			this.lblContadorSeleccionados.Name = "lblContadorSeleccionados";
			this.lblContadorSeleccionados.Size = new System.Drawing.Size(0, 16);
			this.lblContadorSeleccionados.TabIndex = 237;
			this.lblContadorSeleccionados.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.Location = new System.Drawing.Point(792, 312);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 16);
			this.lblContador.TabIndex = 238;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.ImageList = this.imageList1;
			this.btnAsiento.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnAsiento.Location = new System.Drawing.Point(1024, 376);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
			this.btnAsiento.TabIndex = 15;
			this.btnAsiento.Text = "&Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnImprimir.Location = new System.Drawing.Point(936, 376);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(80, 23);
			this.btnImprimir.TabIndex = 14;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.ImageIndex = 57;
			this.btnConsultar.ImageList = this.imageList1;
			this.btnConsultar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnConsultar.Location = new System.Drawing.Point(848, 376);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(80, 23);
			this.btnConsultar.TabIndex = 13;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnDepositar
			// 
			this.btnDepositar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnDepositar.CausesValidation = false;
			this.btnDepositar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDepositar.ImageIndex = 68;
			this.btnDepositar.ImageList = this.imageList1;
			this.btnDepositar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnDepositar.Location = new System.Drawing.Point(760, 376);
			this.btnDepositar.Name = "btnDepositar";
			this.btnDepositar.Size = new System.Drawing.Size(80, 23);
			this.btnDepositar.TabIndex = 12;
			this.btnDepositar.Text = "&Depositar";
			this.btnDepositar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDepositar.Click += new System.EventHandler(this.btnDepositar_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnNuevo.Location = new System.Drawing.Point(672, 376);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 23);
			this.btnNuevo.TabIndex = 11;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.ImageIndex = 62;
			this.button2.ImageList = this.imageList1;
			this.button2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.button2.Location = new System.Drawing.Point(1200, 376);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(80, 23);
			this.button2.TabIndex = 16;
			this.button2.Text = "&Salir";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(232, 43);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 366;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 43);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(37, 16);
			this.label5.TabIndex = 365;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance28;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 40);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 2;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
			this.dtDesde.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtDesde_BeforeDropDown);
			// 
			// dtHasta
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance29;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton3.Caption = "Today";
			dateButton3.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton3);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(288, 40);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 3;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 72);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1292, 8);
			this.groupBox2.TabIndex = 367;
			this.groupBox2.TabStop = false;
			// 
			// btnReporte
			// 
			this.btnReporte.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnReporte.CausesValidation = false;
			this.btnReporte.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnReporte.ImageIndex = 69;
			this.btnReporte.ImageList = this.imageList1;
			this.btnReporte.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnReporte.Location = new System.Drawing.Point(1112, 376);
			this.btnReporte.Name = "btnReporte";
			this.btnReporte.Size = new System.Drawing.Size(80, 23);
			this.btnReporte.TabIndex = 368;
			this.btnReporte.Text = "&Reporte";
			this.btnReporte.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnReporte.Click += new System.EventHandler(this.btnReporte_Click);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(776, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(31, 16);
			this.label6.TabIndex = 370;
			this.label6.Text = "Local";
			// 
			// cmbLocal
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbLocal.Appearance = appearance30;
			this.cmbLocal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLocal.DataSource = this.ultraDataSource10;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn33.Header.VisiblePosition = 2;
			ultraGridColumn33.Width = 224;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33});
			this.cmbLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbLocal.DisplayMember = "Nombre";
			this.cmbLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbLocal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLocal.Location = new System.Drawing.Point(840, 8);
			this.cmbLocal.Name = "cmbLocal";
			this.cmbLocal.Size = new System.Drawing.Size(248, 21);
			this.cmbLocal.TabIndex = 369;
			this.cmbLocal.ValueMember = "Bodega";
			// 
			// ultraDataSource10
			// 
			ultraDataColumn37.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn37,
																																	ultraDataColumn38,
																																	ultraDataColumn39});
			// 
			// cmbBanco
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance31;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataSource = this.ultraDataSource2;
			ultraGridColumn34.Header.VisiblePosition = 0;
			ultraGridColumn35.Header.VisiblePosition = 1;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn34,
																										 ultraGridColumn35});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(496, 8);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(248, 21);
			this.cmbBanco.TabIndex = 372;
			this.cmbBanco.ValueMember = "idBanco";
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(544, 317);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(94, 16);
			this.label7.TabIndex = 373;
			this.label7.Text = "Gastos Bancarios";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtGastosBancarios
			// 
			this.txtGastosBancarios.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosBancarios.Appearance = appearance32;
			this.txtGastosBancarios.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosBancarios.FormatString = "#,##0.00";
			this.txtGastosBancarios.Location = new System.Drawing.Point(656, 314);
			this.txtGastosBancarios.Name = "txtGastosBancarios";
			this.txtGastosBancarios.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosBancarios.PromptChar = ' ';
			this.txtGastosBancarios.Size = new System.Drawing.Size(112, 21);
			this.txtGastosBancarios.TabIndex = 374;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnExportar.Location = new System.Drawing.Point(1184, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(80, 23);
			this.btnExportar.TabIndex = 375;
			this.btnExportar.Text = "&Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// frmCapturaVouchers
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1288, 406);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtGastosBancarios);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblCuenta);
			this.Controls.Add(this.lblCtaCorriente);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.lblCaducidad);
			this.Controls.Add(this.cmbLocal);
			this.Controls.Add(this.btnReporte);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnDepositar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.lblContadorSeleccionados);
			this.Controls.Add(this.txtSuma);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtIdDeposito);
			this.Controls.Add(this.rdgTipoConsulta);
			this.Controls.Add(this.cmbPlazo);
			this.Controls.Add(this.cmbPlanes);
			this.Controls.Add(this.cmbTarjetaCredito);
			this.Controls.Add(this.cmbBancos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.uGridVouchers);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.txtConcepto);
			this.Controls.Add(this.txtDeposito);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.Name = "frmCapturaVouchers";
			this.Text = "Deposito de Vouchers";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCapturaVouchers_KeyDown);
			this.Load += new System.EventHandler(this.frmCapturaVouchers_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchers)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBancos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjetaCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoConsulta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSuma)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosBancarios)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		bool Modifica = false;
		
		private void Consulta()
		{
			int IdBanco = 0;

			if (this.cmbBanco.ActiveRow != null) IdBanco = (int)this.cmbBanco.Value;

			bool bFechas = false;

			DateTime dtDesde = DateTime.Today;
			DateTime dtHasta = DateTime.Today;

			if (this.dtDesde.Value != System.DBNull.Value || this.dtHasta.Value != System.DBNull.Value)
			{
				bFechas = true;

				dtDesde = (DateTime)this.dtDesde.Value;
				dtHasta = (DateTime)this.dtHasta.Value;
			}

			bool Estado = false;
			bool Capturado = false;

			if ((int)this.rdgTipoConsulta.EditValue == 1) Estado = true;
			if ((int)this.rdgTipoConsulta.EditValue == 2) Capturado = true;
	
			int IBodega = 0;
			if (this.cmbLocal.ActiveRow != null)IBodega = (int)this.cmbLocal.Value;

			this.uGridVouchers.DataSource = this.ultraDataSource1;

			string sSQL = string.Format("Exec ConsultaDepositosCapturaVoucher {0}, {1}, '{2}', {3}, {4}, '{5}', '{6}', {7}, {8}, {9}", 
				(int)this.txtIdDeposito.Value, IdBanco, this.txtBuscar.Text.ToString().Trim(), Estado, MenuLatinium.IdUsuario, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), bFechas, Capturado, IBodega);
			
			this.uGridVouchers.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			this.uGridVouchers.DisplayLayout.Bands[0].Expandable = true;

			this.lblContador.Text = this.uGridVouchers.Rows.Count + " Documentos encontrados";
			
			Suma();
		}

		private void ModificaGrilla(bool bModifica)
		{
			this.uGridVouchers.Enabled = true;

			if (bModifica) this.uGridVouchers.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True; else this.uGridVouchers.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
		}

		private void frmCapturaVouchers_Load(object sender, System.EventArgs e)
		{
			/*this.cmbTarjeta.Value = 0;
			this.cmbCaja.Value = 0; 			

			this.cmbCaja.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre From Bodega Where Activo = 1 Order By Bodega Asc");
			this.cmbTarjeta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTarjeta, Nombre From TarjetasDeCredito Order By Nombre Asc");*/

			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCuentasCierreCaja");
			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco Order By idBanco Asc");

			this.cmbBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco Order By idBanco Asc");
			this.cmbTarjetaCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTarjeta, Nombre From TarjetasDeCredito Order By Nombre Asc");
			this.cmbPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPlan, Descripcion From Planes Order By idPlan Asc");
			this.cmbPlazo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPlazo, Descripcion From Plazos Order By idPlazo Asc");	
			this.cmbLocal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 And Factura = 1 Order By Nombre");	
		
			this.dtFecha.Value = DateTime.Today;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
		}

		private void frmCapturaVouchers_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void txtBuscar_Enter(object sender, System.EventArgs e)
		{
		
		}

		
		private void txtBuscar_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnConsultar_Click(sender, e);
		}

		private void uGridVouchers_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
					
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
						
			#region Decimales
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decPrecio"]);
			
			e.Layout.Bands[0].Columns["Retenido"].Format = stFormato;			
			e.Layout.Bands[0].Columns["Retenido"].MaskInput = stInput;

			e.Layout.Bands[0].Columns["RetencionIVA"].Format = stFormato;			
			e.Layout.Bands[0].Columns["RetencionIVA"].MaskInput = stInput;
			
			e.Layout.Bands[0].Columns["Comision"].Format = stFormato;			
			e.Layout.Bands[0].Columns["Comision"].MaskInput = stInput;

			e.Layout.Bands[0].Columns["PorcComision"].Format = stFormato;			
			e.Layout.Bands[0].Columns["PorcComision"].MaskInput = stInput;

			e.Layout.Bands[0].Columns["ComisionIVA"].Format = stFormato;			
			e.Layout.Bands[0].Columns["ComisionIVA"].MaskInput = stInput;
			#endregion Decimales			
		}

		private void uGridVouchers_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (uGridVouchers.ActiveCell == null) return;
			try
			{
				if (uGridVouchers.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
						#region Page Down Abre lista desplegable
					case (int) Keys.PageDown:
						if(uGridVouchers.ActiveCell.DroppedDown == false)
							uGridVouchers.ActiveCell.DroppedDown = true;
						break;
						#endregion Page Down Abre lista desplegable
						#region Flechas
					case (int) Keys.Up:
						if (uGridVouchers.ActiveCell.DroppedDown) return;

						uGridVouchers.PerformAction(UltraGridAction.ExitEditMode, false, false);
						uGridVouchers.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						uGridVouchers.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						uGridVouchers.PerformAction(UltraGridAction.ExitEditMode, false, false);
						uGridVouchers.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						uGridVouchers.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (uGridVouchers.ActiveCell.DroppedDown) return;
			
						uGridVouchers.PerformAction(UltraGridAction.ExitEditMode, false, false);
						uGridVouchers.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						uGridVouchers.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
						#endregion Flechas
						#region Enter
					case (int) Keys.Enter:
						uGridVouchers.PerformAction(UltraGridAction.ExitEditMode, false, false);
						uGridVouchers.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						uGridVouchers.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(uGridVouchers.ActiveCell.DroppedDown == false)
							uGridVouchers.ActiveCell.DroppedDown = true;
						//						uGridVouchers.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Tecla");
			}
		}

		private void cmbCuenta_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void cmbCuenta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtBuscar_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbBanco_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void cmbBanco_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridVouchers_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Retenido")
			{
				e.Cell.Row.Cells["ValorCobrar"].Value = (decimal)e.Cell.Row.Cells["Valor"].Value - (decimal)e.Cell.Row.Cells["Retenido"].Value - (decimal)e.Cell.Row.Cells["RetencionIVA"].Value - (decimal)e.Cell.Row.Cells["Comision"].Value - (decimal)e.Cell.Row.Cells["ComisionIVA"].Value;
				Suma();
			}

			if (e.Cell.Column.ToString() == "Comision")
			{
				e.Cell.Row.Cells["ValorCobrar"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["Retenido"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["RetencionIVA"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["Comision"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["ComisionIVA"].Value);
				Suma();				
			}

			if (e.Cell.Column.ToString() == "RetencionIVA")
			{
				e.Cell.Row.Cells["ValorCobrar"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["Retenido"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["RetencionIVA"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["Comision"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["ComisionIVA"].Value);
				Suma();				
			}

			if (e.Cell.Column.ToString() == "idBanco")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set idBanco = {0} Where idDetFormaPago = {1}", (int)e.Cell.Row.Cells["idBanco"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}

			if (e.Cell.Column.ToString() == "idTarjeta")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set idTarjeta = {0} Where idDetFormaPago = {1}", (int)e.Cell.Row.Cells["idTarjeta"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}

			if (e.Cell.Column.ToString() == "idPlan")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set idPlan = {0} Where idDetFormaPago = {1}", (int)e.Cell.Row.Cells["idPlan"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}

			if (e.Cell.Column.ToString() == "idPlazo")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set idPlazo = {0} Where idDetFormaPago = {1}", (int)e.Cell.Row.Cells["idPlazo"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}

			if (e.Cell.Column.ToString() == "Lote")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set Lote = '{0}' Where idDetFormaPago = {1}", e.Cell.Row.Cells["Lote"].Value.ToString(), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}

			if (e.Cell.Column.ToString() == "Referencia")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set Referencia = '{0}' Where idDetFormaPago = {1}", e.Cell.Row.Cells["Referencia"].Value.ToString(), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}

			if (e.Cell.Column.ToString() == "PorcComision")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set PorcComision = {0} Where idDetFormaPago = {1}", Convert.ToDecimal(e.Cell.Row.Cells["PorcComision"].Value), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));

				e.Cell.Row.Cells["Comision"].Value = decimal.Round((Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["PorcComision"].Value))/100, 2);
				e.Cell.Row.Cells["ValorCobrar"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["Retenido"].Value)
					- Convert.ToDecimal(e.Cell.Row.Cells["RetencionIVA"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["Comision"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["ComisionIVA"].Value);
				Suma();
			}

			if (e.Cell.Column.ToString() == "ComisionIVA")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set ComisionIVA = {0} Where idDetFormaPago = {1}", 
					Convert.ToDecimal(e.Cell.Row.Cells["ComisionIVA"].Value), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
				
				e.Cell.Row.Cells["ValorCobrar"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["Retenido"].Value)
					- Convert.ToDecimal(e.Cell.Row.Cells["RetencionIVA"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["Comision"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["ComisionIVA"].Value);

				Suma();
			}
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbCuenta.Focus();
		}

		private void cmbBanco_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtBuscar.Focus();
		}

		private void cmbCuenta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtConcepto.Focus();				
		}

		private void txtConcepto_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDeposito.Focus();
		}

		private void txtDeposito_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13) this.btnDepositar.Focus();
		}

		private void txtDepositante_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnDepositar.Focus();
		}

		private void rdgTipoConsulta_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			Consulta();

			if ((int)this.rdgTipoConsulta.EditValue == 0) this.btnDepositar.Enabled = false;
			else if ((int)this.rdgTipoConsulta.EditValue == 1) this.btnDepositar.Enabled = true;
			else if ((int)this.rdgTipoConsulta.EditValue == 2) this.btnDepositar.Enabled = false;

			if ((int)this.rdgTipoConsulta.EditValue == 2) ModificaGrilla(false);
			else ModificaGrilla(true);
		}

		private void Suma()
		{
			this.txtSuma.Value = 0.00m;

			this.lblContadorSeleccionados.Text = "";

			int iContador = 0;
			decimal dSuma = 0.00m;

			string sSQL = string.Format("Exec ContadorSumatoriaVouchers {0}, {1}", (int)this.txtIdDeposito.Value, MenuLatinium.IdUsuario);
			SqlDataReader drSuma = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drSuma.Read();
			if (drSuma.HasRows)
			{
				iContador = (int)drSuma.GetValue(0);
				dSuma = (decimal)drSuma.GetValue(1);
			}
			drSuma.Close();

			this.lblContadorSeleccionados.Text = iContador.ToString() + " Documentos Capturados";

			this.txtSuma.Value = dSuma;
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void uGridVouchers_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idDetFormaPago"].Value != System.DBNull.Value)
			{
				if (e.Cell.Column.ToString() == "Capturado")
				{
					this.uGridVouchers.UpdateData();
				
					int bTemp = 0; /* ESTA COMO ENTERO NO RECONOCE COMO BIT AL CAMPO , Cobrado = 0 */
				
					if ((bool)e.Cell.Row.Cells["Capturado"].Value) bTemp = 1;
				
					if (Modifica)
					{
						if (bTemp == 0) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set Capturado = 0, Cobrado = 0 Where idDetFormaPago = {0}", (int)e.Cell.Row.Cells["idDetFormaPago"].Value), true);
						else if (bTemp == 1) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set Capturado = 1 Where idDetFormaPago = {0}", (int)e.Cell.Row.Cells["idDetFormaPago"].Value), true);
					}
					else
					{
						if (bTemp == 0)
						{
							string sSQLAct = string.Format("Update DetalleFormaPago Set Capturado = {1}, idUsuario = 0 Where idDetFormaPago = {0}", (int)e.Cell.Row.Cells["idDetFormaPago"].Value, bTemp);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
						}
						else if (bTemp == 1)
						{
							string sSQLAct = string.Format("Update DetalleFormaPago Set Capturado = {1}, idUsuario = {2} Where idDetFormaPago = {0}", (int)e.Cell.Row.Cells["idDetFormaPago"].Value, bTemp, MenuLatinium.IdUsuario);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
						}
					}

					Suma();				
				}
			}
		}

		private void uGridVouchers_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Capturado")
			{
				if ((bool)e.Cell.Row.Cells["Cobrado"].Value)
				{
					if	(DialogResult.No == MessageBox.Show("Este voucher ya ha sido cobrado, \n\n ¿Esta seguro de quitarlo de este deposito?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) e.Cancel = true;				
				}
			}
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			#region Controles
			this.txtIdDeposito.Value = 0;
			this.txtIdAsiento.Value = 0;
			this.dtFecha.Value = DateTime.Today;
			this.cmbCuenta.Value = 0;
			this.txtDeposito.Text = "";
			this.txtConcepto.Text = "";

			Modifica = false;

			this.btnImprimir.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.rdgTipoConsulta.EditValue = 0;

			Consulta();
			#endregion Controles
		}

		private void btnDepositar_Click(object sender, System.EventArgs e)
		{
			if (this.cmbCuenta.Text.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione la cuenta para hacer el depósito.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCuenta.Focus();			
			}
			else if (this.txtConcepto.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba el concepto para el depósito.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtConcepto.Focus();				
			}
			else if (this.txtDeposito.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba el número del depósito.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDeposito.Focus();				
			}
			else if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba la Fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();				
			}
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Bancos"))
			{
				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Depositos en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();			
			}
			else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AsientoDetalle ad Inner Join Asiento a On a.idAsiento = ad.idAsiento Where a.idTipoAsiento = 2 And a.Cheque = '{0}' And ad.idCuenta In ({1}) And a.idAsiento <> {2}", this.txtDeposito.Text.ToString(), (int)this.cmbCuenta.Value, (int)this.txtIdAsiento.Value)) > 0)
			{
				MessageBox.Show(string.Format("El Deposito '{0}' ya ha sido utilizado.", this.txtDeposito.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDeposito.Focus();			
			}
			else
			{
				int i = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVouchers.Rows.All)
				{
					if ((bool)dr.Cells["Capturado"].Value == true) i++;
				}

				if (i == 0)
				{
					MessageBox.Show("Seleccione el o los Vouchers que va a Depositar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
				else if (DialogResult.No == MessageBox.Show("ANTES DE GENERAR EL DEPOSITO. \n\n VERIQUE QUE LOS DATOS REGISTRADOS SON LOS CORRECTOS \n\n DEBE DE CUADRAR LA COMISION, LA RETENCION Y EL VALOR DEL DEPOSITO \n\n\n ¿Esta seguro de Generar el Deposito.?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
				{					
					#region Cálculo de valores
					decimal dTotal = 0;
					decimal dCobro = 0;
					decimal dRetencion = 0;
					decimal dRetencionIVA = 0;
					decimal dComision = 0;
					DateTime dtFechaDeposito = (DateTime)this.dtFecha.Value;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVouchers.Rows.All)
					{
						if ((bool)dr.Cells["Capturado"].Value == true)
						{
							dTotal = dTotal + Convert.ToDecimal(dr.Cells["Valor"].Value);
							dCobro = dCobro + Convert.ToDecimal(dr.Cells["ValorCobrar"].Value);
							dRetencion = dRetencion + Convert.ToDecimal(dr.Cells["Retenido"].Value);
							dRetencionIVA = dRetencionIVA + Convert.ToDecimal(dr.Cells["RetencionIVA"].Value);
							dComision = dComision + Convert.ToDecimal(dr.Cells["Comision"].Value) + Convert.ToDecimal(dr.Cells["ComisionIVA"].Value);
						}
					}
					#endregion Cálculo de valores

					#region Registra Deposito
					string sSQLDetDep = "";
					string sSQLDep = string.Format("Exec GrabaEncabezadoDeposito {0}, {1}, '{2}', '{3}', '{4}', {5}, {6}", 
						(int)this.txtIdDeposito.Value, (int)this.cmbCuenta.Value, this.txtDeposito.Text.ToString(), dtFechaDeposito.ToString("yyyyMMdd HH:mm"), this.txtConcepto.Text.ToString(), dTotal, 
						Convert.ToDecimal(this.txtGastosBancarios.Value));
					this.txtIdDeposito.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLDep, true);

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec EliminaDetalleDepositoTarjetas {0}", (int)this.txtIdDeposito.Value));

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVouchers.Rows.All)
					{
						if ((bool)dr.Cells["Capturado"].Value == true)
						{
							sSQLDetDep = string.Format("Exec GuardaDetalleDepositos {0}, {1}", (int)this.txtIdDeposito.Value, (int)dr.Cells["idDetFormaPago"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLDetDep, true);
						}
					}
					#endregion Registra Deposito

					#region Genera Asiento 
					DateTime Fecha = (DateTime)this.dtFecha.Value;
					dCobro = dCobro - Convert.ToDecimal(this.txtGastosBancarios.Value);
					string sSQL = string.Format("Exec AsientoTarjetasDeCredito {0}, {1}, '{2}', '{3}', '{4}', '{5}', {6}, {7}, {8}, {9}, {10}, {11}", 
						(int)this.txtIdAsiento.Value, (int)this.cmbCuenta.Value, this.cmbCuenta.Text.ToString(), this.txtConcepto.Text.ToString(), this.txtDeposito.Text.ToString(), Fecha.ToString("yyyyMMdd"), 
						dTotal, dCobro, dRetencion, dRetencionIVA, dComision, Convert.ToDecimal(this.txtGastosBancarios.Value));
					this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DepositosDeTarjetas Set idAsiento = {0} Where idDepositoTarjeta = {1}", (int)this.txtIdAsiento.Value, (int)this.txtIdDeposito.Value));
					#endregion Genera Asiento 

					#region Actualiza Detalle de la forma pago
					string sSQLActDFP = "";
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVouchers.Rows.All)
					{
						if ((bool)dr.Cells["Capturado"].Value == true)
						{
							int IdPlazo = 0;
							if (dr.Cells["idPlazo"].Value != System.DBNull.Value) IdPlazo = (int)dr.Cells["idPlazo"].Value;

							sSQLActDFP = string.Format("Exec ActualizaDetalleFormaPagoTarjetas {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}", 
								(int)dr.Cells["idDetFormaPago"].Value, (int)dr.Cells["idBanco"].Value, (int)dr.Cells["idTarjeta"].Value, (int)dr.Cells["idPlan"].Value, IdPlazo, 
								Convert.ToDecimal(dr.Cells["Retenido"].Value), Convert.ToDecimal(dr.Cells["RetencionIVA"].Value), Convert.ToDecimal(dr.Cells["Comision"].Value), Convert.ToDecimal(dr.Cells["ValorCobrar"].Value));
							Funcion.EjecutaSQL(cdsSeteoF, sSQLActDFP, true);
						}
					}
					#endregion Actualiza Detalle de la forma pago

					#region Imprime reporte
					ParameterFields paramFields = new ParameterFields ();
				
					ParameterField pfidAsiento = new ParameterField ();
					ParameterDiscreteValue ValidAsiento = new ParameterDiscreteValue ();
					pfidAsiento.ParameterFieldName = "@idAsiento";
					ValidAsiento.Value = (int)this.txtIdAsiento.Value;
					pfidAsiento.CurrentValues.Add (ValidAsiento);
					paramFields.Add (pfidAsiento);
				
					ParameterField pfidAsientoDet = new ParameterField ();
					ParameterDiscreteValue ValIdAsientoDet = new ParameterDiscreteValue ();
					pfidAsientoDet.ParameterFieldName = "@idAsientoDet";
					ValIdAsientoDet.Value = (int)this.txtIdAsiento.Value;
					pfidAsientoDet.CurrentValues.Add (ValIdAsientoDet);
					paramFields.Add (pfidAsientoDet);
							
					Reporte miRepor = new Reporte("DepositoCaja.rpt", "");
					miRepor.MdiParent = this.MdiParent;
					miRepor.Titulo("Depósito de caja");
					miRepor.crVista.ParameterFieldInfo = paramFields;
					miRepor.Show();	
					#endregion Imprime reporte

					#region Actualiza Grid 
					Consulta();
//					bool Estado = false;
//					int Banco = 0;
//					if (this.cmbBanco.Text.ToString().Length > 0) Banco = (int)this.cmbBanco.Value;
//					if ((int)this.rdgTipoConsulta.EditValue == 1) Estado = true;
//					string sSQLAct = string.Format("Exec ConsultaCapturaVoucher {0}, '{1}', {2}", Banco, this.txtBuscar.Text.ToString().Trim(), Estado);
//					this.uGridVouchers.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLAct);
					#endregion Actualiza Grid 

					MessageBox.Show("Depósito registrado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);			

					#region Controles
					this.btnNuevo_Click(sender, e);
					#endregion Controles
				}
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			using (frmConsultaDepositoVouchers miCon = new frmConsultaDepositoVouchers())
			{
				if (DialogResult.OK == miCon.ShowDialog())
				{	
					this.txtIdDeposito.Value = (int)miCon.uGridEncabezado.ActiveRow.Cells["idDepositoTarjeta"].Value;
					this.dtFecha.Value = (DateTime)miCon.uGridEncabezado.ActiveRow.Cells["Fecha"].Value;
					this.cmbCuenta.Value = (int)miCon.uGridEncabezado.ActiveRow.Cells["idCuenta"].Value;
					this.txtDeposito.Text = miCon.uGridEncabezado.ActiveRow.Cells["Numero"].Value.ToString();
					this.txtConcepto.Text = miCon.uGridEncabezado.ActiveRow.Cells["Concepto"].Value.ToString();
					this.txtIdAsiento.Value = (int)miCon.uGridEncabezado.ActiveRow.Cells["idAsiento"].Value;

					int IdBanco = 0;
					if (this.cmbBanco.ActiveRow != null) IdBanco = (int)this.cmbBanco.Value;

					bool bFechas = false;

					DateTime dtDesde = DateTime.Today;
					DateTime dtHasta = DateTime.Today;

					if (this.dtDesde.Value != System.DBNull.Value || this.dtHasta.Value != System.DBNull.Value)
					{
						bFechas = true;

						dtDesde = (DateTime)this.dtDesde.Value;
						dtHasta = (DateTime)this.dtHasta.Value;
					}

					bool Estado = false;
					bool Capturado = false;
					this.rdgTipoConsulta.EditValue = 1;

					if ((int)this.rdgTipoConsulta.EditValue == 1) Estado = true;

					string sSQL = string.Format("Exec ConsultaDepositosCapturaVoucher {0}, {1}, '{2}', {3}, {4}, '{5}', '{6}', {7}, {8}, 0", 
						(int)this.txtIdDeposito.Value, IdBanco, this.txtBuscar.Text.ToString().Trim(), Estado, MenuLatinium.IdUsuario, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), bFechas, Capturado);
					this.uGridVouchers.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
					this.uGridVouchers.DisplayLayout.Bands[0].Expandable = true;
					
					Modifica = true;

					this.btnImprimir.Enabled = true;
					this.btnAsiento.Enabled = true;
					this.btnAsiento.Enabled = true;
				}
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtIdAsiento.Value == 0)
			{
				MessageBox.Show("No Existe Deposito para Imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				#region Imprime reporte
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfidAsiento = new ParameterField ();
				ParameterDiscreteValue ValidAsiento = new ParameterDiscreteValue ();
				pfidAsiento.ParameterFieldName = "@idAsiento";
				ValidAsiento.Value = (int)this.txtIdAsiento.Value;
				pfidAsiento.CurrentValues.Add (ValidAsiento);
				paramFields.Add (pfidAsiento);
				
				ParameterField pfidAsientoDet = new ParameterField ();
				ParameterDiscreteValue ValIdAsientoDet = new ParameterDiscreteValue ();
				pfidAsientoDet.ParameterFieldName = "@idAsientoDet";
				ValIdAsientoDet.Value = (int)this.txtIdAsiento.Value;
				pfidAsientoDet.CurrentValues.Add (ValIdAsientoDet);
				paramFields.Add (pfidAsientoDet);
							
				Reporte miRepor = new Reporte("DepositoCaja.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Depósito de caja");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				miRepor.Show();	
				#endregion Imprime reporte
			}
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtIdAsiento.Value == 0)
			{
				MessageBox.Show("No Existe Asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				frmAsientos Asientos = new frmAsientos(((int)this.txtIdAsiento.Value));
				Asientos.MdiParent = MdiParent;
				Asientos.Show();
			}
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void rdgTipoConsulta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbBanco.Focus();
		}

		private void cmbBanco_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtDesde.Focus();
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void btnReporte_Click(object sender, System.EventArgs e)
		{
			frmReportes Reportes = new frmReportes();
			Reportes.MdiParent = this.MdiParent;						
			Reportes.Show();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.uGridVouchers.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				
					ultraGridExcelExporter1.Export(this.uGridVouchers, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void dtDesde_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}
	}
}
