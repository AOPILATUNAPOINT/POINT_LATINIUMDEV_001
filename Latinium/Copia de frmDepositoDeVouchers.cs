using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmDepositoDeVouchers.
	/// </summary>
	public class frmDepositoDeVouchers : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosBancarios;
		private System.Windows.Forms.Label lblCuenta;
		private System.Windows.Forms.Label lblCtaCorriente;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label lblCaducidad;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdDeposito;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDeposito;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl9;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdRetencion;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbLocal;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPlazo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVouchers;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBancos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPlanes;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTarjetaCredito;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label lblContadorSeleccionados;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVouchersConsulta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVouchersGeneral;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl tabsGeneral;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private System.Windows.Forms.Button btnExportar;
		private System.ComponentModel.IContainer components;

		public frmDepositoDeVouchers()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlan");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepositoTarjeta");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlan", -1, "cmbPlanes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlazo", -1, "cmbPlazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorUsoTarjeta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcComision");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetencionIVA");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCobrar");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 14, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retenido", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retenido", 19, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comision", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comision", 18, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorCobrar", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorCobrar", 21, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorUsoTarjeta", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorUsoTarjeta", 15, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetencionIVA", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetencionIVA", 20, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 16, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepositoTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorUsoTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RetencionIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCobrar");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepositoTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco", -1, "cmbBancos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta", -1, "cmbTarjetaCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlan", -1, "cmbPlanes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlazo", -1, "cmbPlazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorUsoTarjeta");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcComision");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetencionIVA");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCobrar");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 14, true);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retenido", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retenido", 19, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comision", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comision", 18, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorCobrar", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorCobrar", 21, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorUsoTarjeta", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorUsoTarjeta", 15, true);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetencionIVA", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetencionIVA", 20, true);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 16, true);
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepositoTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorUsoTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RetencionIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCobrar");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepositoTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDepositado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco", -1, "cmbBancos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta", -1, "cmbTarjetaCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlan", -1, "cmbPlanes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlazo", -1, "cmbPlazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorUsoTarjeta");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcComision");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetencionIVA");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCobrar");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 18, true);
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retenido", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retenido", 23, true);
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comision", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comision", 22, true);
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorCobrar", 25, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorCobrar", 25, true);
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorUsoTarjeta", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorUsoTarjeta", 19, true);
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetencionIVA", 24, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetencionIVA", 24, true);
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 20, true);
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepositoTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDepositado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorUsoTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RetencionIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCobrar");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDepositoDeVouchers));
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			this.ultraTabPageControl9 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbBancos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbPlanes = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTarjetaCredito = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbPlazo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridVouchers = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridVouchersConsulta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.uGridVouchersGeneral = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label7 = new System.Windows.Forms.Label();
			this.txtGastosBancarios = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblCuenta = new System.Windows.Forms.Label();
			this.lblCtaCorriente = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.lblCaducidad = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdDeposito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDeposito = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.tabsGeneral = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.txtIdRetencion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label6 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbLocal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.lblContadorSeleccionados = new System.Windows.Forms.Label();
			this.btnExportar = new System.Windows.Forms.Button();
			this.ultraTabPageControl9.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbBancos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjetaCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchers)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchersConsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchersGeneral)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosBancarios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tabsGeneral)).BeginInit();
			this.tabsGeneral.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl9
			// 
			this.ultraTabPageControl9.Controls.Add(this.cmbBancos);
			this.ultraTabPageControl9.Controls.Add(this.cmbPlanes);
			this.ultraTabPageControl9.Controls.Add(this.cmbTarjetaCredito);
			this.ultraTabPageControl9.Controls.Add(this.cmbPlazo);
			this.ultraTabPageControl9.Controls.Add(this.uGridVouchers);
			this.ultraTabPageControl9.Location = new System.Drawing.Point(1, 20);
			this.ultraTabPageControl9.Name = "ultraTabPageControl9";
			this.ultraTabPageControl9.Size = new System.Drawing.Size(1142, 227);
			// 
			// cmbBancos
			// 
			this.cmbBancos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBancos.DataSource = this.ultraDataSource6;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbBancos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBancos.DisplayMember = "Nombre";
			this.cmbBancos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBancos.Location = new System.Drawing.Point(32, 104);
			this.cmbBancos.Name = "cmbBancos";
			this.cmbBancos.Size = new System.Drawing.Size(72, 72);
			this.cmbBancos.TabIndex = 213;
			this.cmbBancos.ValueMember = "idBanco";
			this.cmbBancos.Visible = false;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbPlanes
			// 
			this.cmbPlanes.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlanes.DataSource = this.ultraDataSource8;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 180;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbPlanes.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbPlanes.DisplayMember = "Descripcion";
			this.cmbPlanes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlanes.Location = new System.Drawing.Point(208, 104);
			this.cmbPlanes.Name = "cmbPlanes";
			this.cmbPlanes.Size = new System.Drawing.Size(72, 72);
			this.cmbPlanes.TabIndex = 215;
			this.cmbPlanes.ValueMember = "idPlan";
			this.cmbPlanes.Visible = false;
			// 
			// ultraDataSource8
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// cmbTarjetaCredito
			// 
			this.cmbTarjetaCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTarjetaCredito.DataSource = this.ultraDataSource7;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbTarjetaCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTarjetaCredito.DisplayMember = "Nombre";
			this.cmbTarjetaCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTarjetaCredito.Location = new System.Drawing.Point(120, 104);
			this.cmbTarjetaCredito.Name = "cmbTarjetaCredito";
			this.cmbTarjetaCredito.Size = new System.Drawing.Size(72, 72);
			this.cmbTarjetaCredito.TabIndex = 214;
			this.cmbTarjetaCredito.ValueMember = "idTarjeta";
			this.cmbTarjetaCredito.Visible = false;
			// 
			// ultraDataSource7
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// cmbPlazo
			// 
			this.cmbPlazo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlazo.DataSource = this.ultraDataSource9;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 180;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbPlazo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbPlazo.DisplayMember = "Descripcion";
			this.cmbPlazo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlazo.Location = new System.Drawing.Point(296, 104);
			this.cmbPlazo.Name = "cmbPlazo";
			this.cmbPlazo.Size = new System.Drawing.Size(72, 72);
			this.cmbPlazo.TabIndex = 216;
			this.cmbPlazo.ValueMember = "idPlazo";
			this.cmbPlazo.Visible = false;
			// 
			// ultraDataSource9
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// uGridVouchers
			// 
			this.uGridVouchers.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVouchers.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVouchers.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchers.DisplayLayout.Appearance = appearance1;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 73;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance2;
			ultraGridColumn11.Header.VisiblePosition = 21;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.Caption = "";
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 20;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.Caption = "Doc";
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 35;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.Caption = "Almacén";
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Width = 140;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Width = 80;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.VisiblePosition = 6;
			ultraGridColumn16.Width = 70;
			ultraGridColumn17.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnCellActivate;
			ultraGridColumn17.Header.Caption = "Banco";
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn17.Width = 140;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnCellActivate;
			ultraGridColumn18.Header.Caption = "Tarjeta";
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn18.Width = 140;
			ultraGridColumn19.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnCellActivate;
			ultraGridColumn19.Header.Caption = "Plan";
			ultraGridColumn19.Header.VisiblePosition = 9;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn19.Width = 140;
			ultraGridColumn20.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnCellActivate;
			ultraGridColumn20.Header.Caption = "Plazo";
			ultraGridColumn20.Header.VisiblePosition = 10;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn20.Width = 140;
			ultraGridColumn21.Header.VisiblePosition = 11;
			ultraGridColumn21.Width = 40;
			ultraGridColumn22.Header.VisiblePosition = 12;
			ultraGridColumn22.Width = 65;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance3;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 13;
			ultraGridColumn23.Width = 75;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance4;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "Com Tarjeta";
			ultraGridColumn24.Header.VisiblePosition = 14;
			ultraGridColumn24.Width = 75;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance5;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 15;
			ultraGridColumn25.Width = 75;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance6;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.Caption = "% Com.";
			ultraGridColumn26.Header.VisiblePosition = 16;
			ultraGridColumn26.MaxValue = 100;
			ultraGridColumn26.MinValue = 0;
			ultraGridColumn26.Width = 50;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance7;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.Caption = "Comisión";
			ultraGridColumn27.Header.VisiblePosition = 17;
			ultraGridColumn27.Width = 75;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance8;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 18;
			ultraGridColumn28.Width = 75;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance9;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.Caption = "Ret IVA";
			ultraGridColumn29.Header.VisiblePosition = 19;
			ultraGridColumn29.Width = 75;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance10;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.Caption = "Cobro";
			ultraGridColumn30.Header.VisiblePosition = 20;
			ultraGridColumn30.Width = 75;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn30});
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance11;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance12;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance13;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance14;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance15;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.Hidden = true;
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance16;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance17;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7});
			ultraGridBand5.SummaryFooterCaption = "";
			this.uGridVouchers.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchers.DisplayLayout.GroupByBox.Appearance = appearance18;
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchers.DisplayLayout.GroupByBox.BandLabelAppearance = appearance19;
			this.uGridVouchers.DisplayLayout.GroupByBox.Prompt = " ";
			appearance20.BackColor = System.Drawing.SystemColors.Control;
			appearance20.BackColor2 = System.Drawing.SystemColors.Control;
			appearance20.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance20.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchers.DisplayLayout.GroupByBox.PromptAppearance = appearance20;
			this.uGridVouchers.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVouchers.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVouchers.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance21.BackColor = System.Drawing.Color.Transparent;
			this.uGridVouchers.DisplayLayout.Override.CardAreaAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance22.FontData.BoldAsString = "True";
			appearance22.FontData.Name = "Arial";
			appearance22.FontData.SizeInPoints = 8F;
			appearance22.ForeColor = System.Drawing.Color.White;
			appearance22.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVouchers.DisplayLayout.Override.HeaderAppearance = appearance22;
			this.uGridVouchers.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchers.DisplayLayout.Override.RowSelectorAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchers.DisplayLayout.Override.SelectedRowAppearance = appearance24;
			this.uGridVouchers.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVouchers.Location = new System.Drawing.Point(8, 8);
			this.uGridVouchers.Name = "uGridVouchers";
			this.uGridVouchers.Size = new System.Drawing.Size(1128, 216);
			this.uGridVouchers.TabIndex = 212;
			this.uGridVouchers.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVouchers_CellChange);
			this.uGridVouchers.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVouchers_KeyDown);
			this.uGridVouchers.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVouchers_AfterCellUpdate);
			this.uGridVouchers.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVouchers_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(bool);
			ultraDataColumn16.DataType = typeof(System.DateTime);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn30});
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.uGridVouchersConsulta);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(1142, 227);
			// 
			// uGridVouchersConsulta
			// 
			this.uGridVouchersConsulta.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVouchersConsulta.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVouchersConsulta.DataSource = this.ultraDataSource3;
			appearance25.BackColor = System.Drawing.Color.White;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersConsulta.DisplayLayout.Appearance = appearance25;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 73;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn33.Header.VisiblePosition = 21;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn34.Header.VisiblePosition = 2;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Header.Caption = "Doc";
			ultraGridColumn35.Header.VisiblePosition = 3;
			ultraGridColumn35.Width = 35;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn36.Header.Caption = "Almacén";
			ultraGridColumn36.Header.VisiblePosition = 4;
			ultraGridColumn36.Width = 140;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn37.Header.VisiblePosition = 5;
			ultraGridColumn37.Width = 80;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn38.Header.VisiblePosition = 6;
			ultraGridColumn38.Width = 70;
			ultraGridColumn39.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnCellActivate;
			ultraGridColumn39.Header.Caption = "Banco";
			ultraGridColumn39.Header.VisiblePosition = 7;
			ultraGridColumn39.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn39.Width = 140;
			ultraGridColumn40.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnCellActivate;
			ultraGridColumn40.Header.Caption = "Tarjeta";
			ultraGridColumn40.Header.VisiblePosition = 8;
			ultraGridColumn40.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn40.Width = 140;
			ultraGridColumn41.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnCellActivate;
			ultraGridColumn41.Header.Caption = "Plan";
			ultraGridColumn41.Header.VisiblePosition = 9;
			ultraGridColumn41.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn41.Width = 140;
			ultraGridColumn42.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnCellActivate;
			ultraGridColumn42.Header.Caption = "Plazo";
			ultraGridColumn42.Header.VisiblePosition = 10;
			ultraGridColumn42.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn42.Width = 140;
			ultraGridColumn43.Header.VisiblePosition = 11;
			ultraGridColumn43.Width = 40;
			ultraGridColumn44.Header.VisiblePosition = 12;
			ultraGridColumn44.Width = 65;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance26;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.VisiblePosition = 13;
			ultraGridColumn45.Width = 75;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance27;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.Caption = "Com Tarjeta";
			ultraGridColumn46.Header.VisiblePosition = 14;
			ultraGridColumn46.Width = 75;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance28;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.VisiblePosition = 15;
			ultraGridColumn47.Width = 75;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance29;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.Caption = "% Com.";
			ultraGridColumn48.Header.VisiblePosition = 16;
			ultraGridColumn48.MaxValue = 100;
			ultraGridColumn48.MinValue = 0;
			ultraGridColumn48.Width = 50;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance30;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.Caption = "Comisión";
			ultraGridColumn49.Header.VisiblePosition = 17;
			ultraGridColumn49.Width = 75;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance31;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.VisiblePosition = 18;
			ultraGridColumn50.Width = 75;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance32;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.Caption = "Ret IVA";
			ultraGridColumn51.Header.VisiblePosition = 19;
			ultraGridColumn51.Width = 75;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance33;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.Caption = "Cobro";
			ultraGridColumn52.Header.VisiblePosition = 20;
			ultraGridColumn52.Width = 75;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn52});
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance34;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance35;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance36;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance37;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance38;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.Hidden = true;
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance39;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance40;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand6.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13,
																																															summarySettings14});
			ultraGridBand6.SummaryFooterCaption = "";
			this.uGridVouchersConsulta.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance41.ForeColor = System.Drawing.Color.Black;
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersConsulta.DisplayLayout.GroupByBox.Appearance = appearance41;
			appearance42.ForeColor = System.Drawing.Color.Black;
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersConsulta.DisplayLayout.GroupByBox.BandLabelAppearance = appearance42;
			this.uGridVouchersConsulta.DisplayLayout.GroupByBox.Prompt = " ";
			appearance43.BackColor = System.Drawing.SystemColors.Control;
			appearance43.BackColor2 = System.Drawing.SystemColors.Control;
			appearance43.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance43.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance43.ForeColor = System.Drawing.Color.Black;
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersConsulta.DisplayLayout.GroupByBox.PromptAppearance = appearance43;
			this.uGridVouchersConsulta.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVouchersConsulta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVouchersConsulta.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance44.BackColor = System.Drawing.Color.Transparent;
			this.uGridVouchersConsulta.DisplayLayout.Override.CardAreaAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance45.FontData.BoldAsString = "True";
			appearance45.FontData.Name = "Arial";
			appearance45.FontData.SizeInPoints = 8F;
			appearance45.ForeColor = System.Drawing.Color.White;
			appearance45.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVouchersConsulta.DisplayLayout.Override.HeaderAppearance = appearance45;
			this.uGridVouchersConsulta.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance46.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance46.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchersConsulta.DisplayLayout.Override.RowSelectorAppearance = appearance46;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchersConsulta.DisplayLayout.Override.SelectedRowAppearance = appearance47;
			this.uGridVouchersConsulta.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVouchersConsulta.Location = new System.Drawing.Point(7, 5);
			this.uGridVouchersConsulta.Name = "uGridVouchersConsulta";
			this.uGridVouchersConsulta.Size = new System.Drawing.Size(1128, 216);
			this.uGridVouchersConsulta.TabIndex = 213;
			this.uGridVouchersConsulta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVouchersConsulta_KeyDown);
			this.uGridVouchersConsulta.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVouchersConsulta_AfterCellUpdate);
			this.uGridVouchersConsulta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVouchersConsulta_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn33.DefaultValue = 0;
			ultraDataColumn34.DataType = typeof(bool);
			ultraDataColumn38.DataType = typeof(System.DateTime);
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn40.DataType = typeof(int);
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn42.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn46.DataType = typeof(System.Decimal);
			ultraDataColumn47.DataType = typeof(System.Decimal);
			ultraDataColumn48.DataType = typeof(System.Decimal);
			ultraDataColumn49.DataType = typeof(System.Decimal);
			ultraDataColumn50.DataType = typeof(System.Decimal);
			ultraDataColumn51.DataType = typeof(System.Decimal);
			ultraDataColumn52.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn52});
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.optEstados);
			this.ultraTabPageControl1.Controls.Add(this.uGridVouchersGeneral);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1142, 227);
			// 
			// optEstados
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance48;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance49;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = -1;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 0;
			valueListItem2.DisplayText = "Pendientes";
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Capturados";
			valueListItem4.DataValue = 2;
			valueListItem4.DisplayText = "Depositados";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(8, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(336, 24);
			this.optEstados.TabIndex = 219;
			this.optEstados.Text = "Todos";
			// 
			// uGridVouchersGeneral
			// 
			this.uGridVouchersGeneral.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVouchersGeneral.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVouchersGeneral.DataSource = this.ultraDataSource4;
			appearance50.BackColor = System.Drawing.Color.White;
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersGeneral.DisplayLayout.Appearance = appearance50;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn53.Header.VisiblePosition = 2;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 73;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn54.Header.VisiblePosition = 0;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn55.Header.Caption = "Deposito";
			ultraGridColumn55.Header.VisiblePosition = 3;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn56.Header.Caption = "";
			ultraGridColumn56.Header.VisiblePosition = 1;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 20;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn57.Header.Caption = "Estado";
			ultraGridColumn57.Header.VisiblePosition = 4;
			ultraGridColumn57.Width = 120;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn58.Header.VisiblePosition = 5;
			ultraGridColumn58.Width = 50;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn59.Header.Caption = "Almacén";
			ultraGridColumn59.Header.VisiblePosition = 6;
			ultraGridColumn59.Width = 130;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn60.Header.VisiblePosition = 7;
			ultraGridColumn60.Width = 75;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn61.Header.VisiblePosition = 8;
			ultraGridColumn61.Width = 65;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn62.Header.Caption = "Fecha Deposito";
			ultraGridColumn62.Header.VisiblePosition = 9;
			ultraGridColumn62.Width = 90;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn63.Header.Caption = "Numero Deposito";
			ultraGridColumn63.Header.VisiblePosition = 10;
			ultraGridColumn63.Width = 100;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn64.Header.Caption = "Cuenta Bancaria";
			ultraGridColumn64.Header.VisiblePosition = 11;
			ultraGridColumn64.Width = 120;
			ultraGridColumn65.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnCellActivate;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn65.Header.Caption = "Banco";
			ultraGridColumn65.Header.VisiblePosition = 12;
			ultraGridColumn65.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn65.Width = 130;
			ultraGridColumn66.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnCellActivate;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn66.Header.Caption = "Tarjeta";
			ultraGridColumn66.Header.VisiblePosition = 13;
			ultraGridColumn66.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn66.Width = 130;
			ultraGridColumn67.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnCellActivate;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn67.Header.Caption = "Plan";
			ultraGridColumn67.Header.VisiblePosition = 14;
			ultraGridColumn67.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn67.Width = 110;
			ultraGridColumn68.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnCellActivate;
			ultraGridColumn68.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn68.Header.Caption = "Plazo";
			ultraGridColumn68.Header.VisiblePosition = 15;
			ultraGridColumn68.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn68.Width = 100;
			ultraGridColumn69.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn69.Header.VisiblePosition = 16;
			ultraGridColumn69.Width = 40;
			ultraGridColumn70.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn70.Header.VisiblePosition = 18;
			ultraGridColumn70.Width = 65;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance51;
			ultraGridColumn71.Format = "#,##0.00";
			ultraGridColumn71.Header.VisiblePosition = 20;
			ultraGridColumn71.Width = 65;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance52;
			ultraGridColumn72.Format = "#,##0.00";
			ultraGridColumn72.Header.Caption = "Com Tarjeta";
			ultraGridColumn72.Header.VisiblePosition = 22;
			ultraGridColumn72.Width = 65;
			ultraGridColumn73.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance53;
			ultraGridColumn73.Format = "#,##0.00";
			ultraGridColumn73.Header.VisiblePosition = 17;
			ultraGridColumn74.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance54;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.Caption = "% Com.";
			ultraGridColumn74.Header.VisiblePosition = 19;
			ultraGridColumn74.MaxValue = 100;
			ultraGridColumn74.MinValue = 0;
			ultraGridColumn74.Width = 50;
			ultraGridColumn75.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance55;
			ultraGridColumn75.Format = "#,##0.00";
			ultraGridColumn75.Header.Caption = "Comisión";
			ultraGridColumn75.Header.VisiblePosition = 21;
			ultraGridColumn75.Width = 65;
			ultraGridColumn76.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellAppearance = appearance56;
			ultraGridColumn76.Format = "#,##0.00";
			ultraGridColumn76.Header.VisiblePosition = 23;
			ultraGridColumn76.Width = 65;
			ultraGridColumn77.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn77.CellAppearance = appearance57;
			ultraGridColumn77.Format = "#,##0.00";
			ultraGridColumn77.Header.Caption = "Ret IVA";
			ultraGridColumn77.Header.VisiblePosition = 24;
			ultraGridColumn77.Width = 65;
			ultraGridColumn78.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn78.CellAppearance = appearance58;
			ultraGridColumn78.Format = "#,##0.00";
			ultraGridColumn78.Header.Caption = "Cobro";
			ultraGridColumn78.Header.VisiblePosition = 25;
			ultraGridColumn78.Width = 65;
			ultraGridBand7.Columns.AddRange(new object[] {
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
																										 ultraGridColumn78});
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance59;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance60;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance61;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance62;
			summarySettings18.DisplayFormat = "{0: #,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance63;
			summarySettings19.DisplayFormat = "{0: #,##0.00}";
			summarySettings19.Hidden = true;
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance64;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance65;
			summarySettings21.DisplayFormat = "{0: #,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand7.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings15,
																																															summarySettings16,
																																															summarySettings17,
																																															summarySettings18,
																																															summarySettings19,
																																															summarySettings20,
																																															summarySettings21});
			ultraGridBand7.SummaryFooterCaption = "";
			this.uGridVouchersGeneral.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance66.ForeColor = System.Drawing.Color.Black;
			appearance66.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersGeneral.DisplayLayout.GroupByBox.Appearance = appearance66;
			appearance67.ForeColor = System.Drawing.Color.Black;
			appearance67.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersGeneral.DisplayLayout.GroupByBox.BandLabelAppearance = appearance67;
			this.uGridVouchersGeneral.DisplayLayout.GroupByBox.Prompt = " ";
			appearance68.BackColor = System.Drawing.SystemColors.Control;
			appearance68.BackColor2 = System.Drawing.SystemColors.Control;
			appearance68.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance68.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance68.ForeColor = System.Drawing.Color.Black;
			appearance68.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersGeneral.DisplayLayout.GroupByBox.PromptAppearance = appearance68;
			appearance69.BackColor = System.Drawing.Color.Transparent;
			this.uGridVouchersGeneral.DisplayLayout.Override.CardAreaAppearance = appearance69;
			appearance70.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance70.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance70.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance70.FontData.BoldAsString = "True";
			appearance70.FontData.Name = "Arial";
			appearance70.FontData.SizeInPoints = 7.8F;
			appearance70.ForeColor = System.Drawing.Color.White;
			appearance70.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVouchersGeneral.DisplayLayout.Override.HeaderAppearance = appearance70;
			this.uGridVouchersGeneral.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance71.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance71.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance71.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchersGeneral.DisplayLayout.Override.RowSelectorAppearance = appearance71;
			appearance72.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance72.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance72.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchersGeneral.DisplayLayout.Override.SelectedRowAppearance = appearance72;
			this.uGridVouchersGeneral.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVouchersGeneral.Location = new System.Drawing.Point(7, 40);
			this.uGridVouchersGeneral.Name = "uGridVouchersGeneral";
			this.uGridVouchersGeneral.Size = new System.Drawing.Size(1128, 184);
			this.uGridVouchersGeneral.TabIndex = 213;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn53.DataType = typeof(int);
			ultraDataColumn54.DataType = typeof(int);
			ultraDataColumn55.DataType = typeof(int);
			ultraDataColumn56.DataType = typeof(bool);
			ultraDataColumn61.DataType = typeof(System.DateTime);
			ultraDataColumn62.DataType = typeof(System.DateTime);
			ultraDataColumn62.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn65.DataType = typeof(int);
			ultraDataColumn66.DataType = typeof(int);
			ultraDataColumn67.DataType = typeof(int);
			ultraDataColumn68.DataType = typeof(int);
			ultraDataColumn71.DataType = typeof(System.Decimal);
			ultraDataColumn72.DataType = typeof(System.Decimal);
			ultraDataColumn73.DataType = typeof(System.Decimal);
			ultraDataColumn74.DataType = typeof(System.Decimal);
			ultraDataColumn75.DataType = typeof(System.Decimal);
			ultraDataColumn76.DataType = typeof(System.Decimal);
			ultraDataColumn77.DataType = typeof(System.Decimal);
			ultraDataColumn78.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn78});
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label7.Location = new System.Drawing.Point(584, 10);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(94, 16);
			this.label7.TabIndex = 386;
			this.label7.Text = "Gastos Bancarios";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtGastosBancarios
			// 
			appearance73.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosBancarios.Appearance = appearance73;
			this.txtGastosBancarios.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosBancarios.Enabled = false;
			this.txtGastosBancarios.FormatString = "#,##0.00";
			this.txtGastosBancarios.Location = new System.Drawing.Point(696, 7);
			this.txtGastosBancarios.Name = "txtGastosBancarios";
			this.txtGastosBancarios.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosBancarios.PromptChar = ' ';
			this.txtGastosBancarios.Size = new System.Drawing.Size(88, 22);
			this.txtGastosBancarios.TabIndex = 387;
			this.txtGastosBancarios.Click += new System.EventHandler(this.txtGastosBancarios_Click);
			this.txtGastosBancarios.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGastosBancarios_KeyPress);
			this.txtGastosBancarios.Enter += new System.EventHandler(this.txtGastosBancarios_Enter);
			// 
			// lblCuenta
			// 
			this.lblCuenta.AutoSize = true;
			this.lblCuenta.BackColor = System.Drawing.Color.Transparent;
			this.lblCuenta.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblCuenta.Location = new System.Drawing.Point(240, 10);
			this.lblCuenta.Name = "lblCuenta";
			this.lblCuenta.Size = new System.Drawing.Size(39, 16);
			this.lblCuenta.TabIndex = 382;
			this.lblCuenta.Text = "Cuenta";
			this.lblCuenta.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCtaCorriente
			// 
			this.lblCtaCorriente.AutoSize = true;
			this.lblCtaCorriente.BackColor = System.Drawing.Color.Transparent;
			this.lblCtaCorriente.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblCtaCorriente.Location = new System.Drawing.Point(8, 32);
			this.lblCtaCorriente.Name = "lblCtaCorriente";
			this.lblCtaCorriente.Size = new System.Drawing.Size(44, 16);
			this.lblCtaCorriente.TabIndex = 381;
			this.lblCtaCorriente.Text = "Numero";
			this.lblCtaCorriente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblNumero.Location = new System.Drawing.Point(240, 32);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(51, 16);
			this.lblNumero.TabIndex = 380;
			this.lblNumero.Text = "Concepto";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCaducidad
			// 
			this.lblCaducidad.AutoSize = true;
			this.lblCaducidad.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblCaducidad.Location = new System.Drawing.Point(8, 10);
			this.lblCaducidad.Name = "lblCaducidad";
			this.lblCaducidad.Size = new System.Drawing.Size(34, 16);
			this.lblCaducidad.TabIndex = 379;
			this.lblCaducidad.Text = "Fecha";
			this.lblCaducidad.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance74.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance74;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(1064, 408);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(88, 22);
			this.txtTotal.TabIndex = 385;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Location = new System.Drawing.Point(1104, 456);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsiento.TabIndex = 384;
			this.txtIdAsiento.Visible = false;
			// 
			// txtIdDeposito
			// 
			this.txtIdDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdDeposito.Location = new System.Drawing.Point(1120, 456);
			this.txtIdDeposito.Name = "txtIdDeposito";
			this.txtIdDeposito.PromptChar = ' ';
			this.txtIdDeposito.Size = new System.Drawing.Size(16, 22);
			this.txtIdDeposito.TabIndex = 383;
			this.txtIdDeposito.Visible = false;
			// 
			// cmbCuenta
			// 
			appearance75.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance75;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource5;
			ultraGridColumn79.Header.VisiblePosition = 0;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 1;
			ultraGridColumn80.Width = 352;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn79,
																										 ultraGridColumn80});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.Enabled = false;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(312, 8);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(250, 21);
			this.cmbCuenta.TabIndex = 376;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbCuenta_KeyPress);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn79.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn79,
																																 ultraDataColumn80});
			// 
			// txtConcepto
			// 
			appearance76.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtConcepto.Appearance = appearance76;
			this.txtConcepto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Enabled = false;
			this.txtConcepto.Location = new System.Drawing.Point(312, 32);
			this.txtConcepto.MaxLength = 200;
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(250, 22);
			this.txtConcepto.TabIndex = 377;
			this.txtConcepto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtConcepto_KeyPress);
			// 
			// txtDeposito
			// 
			appearance77.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDeposito.Appearance = appearance77;
			this.txtDeposito.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDeposito.Enabled = false;
			this.txtDeposito.Location = new System.Drawing.Point(72, 32);
			this.txtDeposito.MaxLength = 15;
			this.txtDeposito.Name = "txtDeposito";
			this.txtDeposito.Size = new System.Drawing.Size(150, 22);
			this.txtDeposito.TabIndex = 378;
			this.txtDeposito.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDeposito_KeyPress);
			// 
			// dtFecha
			// 
			appearance78.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance78;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 375;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label1.Location = new System.Drawing.Point(1016, 411);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(29, 16);
			this.label1.TabIndex = 388;
			this.label1.Text = "Total";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 64);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1160, 8);
			this.groupBox1.TabIndex = 389;
			this.groupBox1.TabStop = false;
			// 
			// tabsGeneral
			// 
			this.tabsGeneral.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance79.ForeColorDisabled = System.Drawing.Color.Black;
			this.tabsGeneral.Appearance = appearance79;
			this.tabsGeneral.Controls.Add(this.ultraTabSharedControlsPage1);
			this.tabsGeneral.Controls.Add(this.ultraTabPageControl9);
			this.tabsGeneral.Controls.Add(this.ultraTabPageControl10);
			this.tabsGeneral.Controls.Add(this.ultraTabPageControl1);
			this.tabsGeneral.Location = new System.Drawing.Point(8, 144);
			this.tabsGeneral.Name = "tabsGeneral";
			this.tabsGeneral.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.tabsGeneral.Size = new System.Drawing.Size(1144, 248);
			this.tabsGeneral.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.tabsGeneral.TabIndex = 390;
			ultraTab1.TabPage = this.ultraTabPageControl9;
			ultraTab1.Text = "Capturar";
			ultraTab2.TabPage = this.ultraTabPageControl10;
			ultraTab2.Text = "Capturados";
			ultraTab3.TabPage = this.ultraTabPageControl1;
			ultraTab3.Text = "Consultar";
			this.tabsGeneral.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																													ultraTab1,
																																													ultraTab2,
																																													ultraTab3});
			this.tabsGeneral.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			this.tabsGeneral.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.utcTransaccion_SelectedTabChanged);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1142, 227);
			// 
			// txtIdRetencion
			// 
			this.txtIdRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdRetencion.Location = new System.Drawing.Point(1136, 456);
			this.txtIdRetencion.Name = "txtIdRetencion";
			this.txtIdRetencion.PromptChar = ' ';
			this.txtIdRetencion.Size = new System.Drawing.Size(16, 22);
			this.txtIdRetencion.TabIndex = 391;
			this.txtIdRetencion.Visible = false;
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.ImageList = this.imageList1;
			this.btnAsiento.Location = new System.Drawing.Point(417, 408);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
			this.btnAsiento.TabIndex = 399;
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
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 408);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 393;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(86, 408);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 401;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(170, 408);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 400;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(252, 408);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 395;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(335, 408);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 394;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(503, 408);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 402;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(585, 408);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 396;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.ImageList = this.imageList1;
			this.btnSalir.Location = new System.Drawing.Point(663, 408);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 397;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// cmbBanco
			// 
			appearance80.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance80;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataSource = this.ultraDataSource2;
			ultraGridColumn81.Header.VisiblePosition = 0;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 1;
			ultraGridColumn82.Width = 256;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn81,
																										 ultraGridColumn82});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(432, 80);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(200, 21);
			this.cmbBanco.TabIndex = 413;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBanco_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn81.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn81,
																																 ultraDataColumn82});
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label6.Location = new System.Drawing.Point(376, 107);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(30, 16);
			this.label6.TabIndex = 412;
			this.label6.Text = "Local";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label4.Location = new System.Drawing.Point(192, 82);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 16);
			this.label4.TabIndex = 410;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label5.Location = new System.Drawing.Point(8, 82);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(35, 16);
			this.label5.TabIndex = 409;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Location = new System.Drawing.Point(376, 82);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 408;
			this.label2.Text = "Banco";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label3.Location = new System.Drawing.Point(8, 107);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 16);
			this.label3.TabIndex = 407;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbLocal
			// 
			appearance81.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbLocal.Appearance = appearance81;
			this.cmbLocal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLocal.DataSource = this.ultraDataSource10;
			ultraGridColumn83.Header.VisiblePosition = 0;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 1;
			ultraGridColumn84.Width = 60;
			ultraGridColumn85.Header.VisiblePosition = 2;
			ultraGridColumn85.Width = 200;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn83,
																											ultraGridColumn84,
																											ultraGridColumn85});
			this.cmbLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbLocal.DisplayMember = "Nombre";
			this.cmbLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbLocal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLocal.Location = new System.Drawing.Point(432, 105);
			this.cmbLocal.Name = "cmbLocal";
			this.cmbLocal.Size = new System.Drawing.Size(200, 21);
			this.cmbLocal.TabIndex = 411;
			this.cmbLocal.ValueMember = "Bodega";
			this.cmbLocal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbLocal_KeyDown);
			// 
			// ultraDataSource10
			// 
			ultraDataColumn83.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn83,
																																	ultraDataColumn84,
																																	ultraDataColumn85});
			// 
			// dtHasta
			// 
			appearance82.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance82;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 80);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 405;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			// 
			// dtDesde
			// 
			appearance83.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance83;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton3.Caption = "Today";
			dateButton3.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton3);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 80);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 404;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
			// 
			// txtBuscar
			// 
			appearance84.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance84;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(64, 104);
			this.txtBuscar.MaxLength = 15;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(200, 22);
			this.txtBuscar.TabIndex = 406;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
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
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.ImageIndex = 2;
			this.btnBuscar.ImageList = this.imageList1;
			this.btnBuscar.Location = new System.Drawing.Point(280, 104);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(72, 23);
			this.btnBuscar.TabIndex = 414;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(816, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 423;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 10F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(760, 409);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 20);
			this.lblUsuario.TabIndex = 424;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Location = new System.Drawing.Point(1088, 456);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 425;
			this.txtEstado.Visible = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(656, 82);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 18);
			this.lblContador.TabIndex = 427;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblContadorSeleccionados
			// 
			this.lblContadorSeleccionados.AutoSize = true;
			this.lblContadorSeleccionados.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContadorSeleccionados.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContadorSeleccionados.Location = new System.Drawing.Point(656, 107);
			this.lblContadorSeleccionados.Name = "lblContadorSeleccionados";
			this.lblContadorSeleccionados.Size = new System.Drawing.Size(0, 18);
			this.lblContadorSeleccionados.TabIndex = 426;
			this.lblContadorSeleccionados.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1064, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 728;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// frmDepositoDeVouchers
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1160, 438);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.lblContadorSeleccionados);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblCuenta);
			this.Controls.Add(this.lblCtaCorriente);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.lblCaducidad);
			this.Controls.Add(this.cmbLocal);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtIdRetencion);
			this.Controls.Add(this.tabsGeneral);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtGastosBancarios);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtIdDeposito);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.txtConcepto);
			this.Controls.Add(this.txtDeposito);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.Name = "frmDepositoDeVouchers";
			this.Text = "Depositos Tarjetas de Crédito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDepositoDeVouchers_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmDepositoDeVouchers_Closing);
			this.Load += new System.EventHandler(this.frmDepositoDeVouchers_Load);
			this.ultraTabPageControl9.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbBancos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjetaCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchers)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchersConsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchersGeneral)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosBancarios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.tabsGeneral)).EndInit();
			this.tabsGeneral.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtIdRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;		
		int iContador = 0;
		int idBloqueaTransacciones = 0;
		bool bActualiza = false;
		
		private void frmDepositoDeVouchers_Load(object sender, System.EventArgs e)
		{
			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCuentasCierreCaja");
			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco Order By idBanco Asc");
			this.cmbLocal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 And Factura = 1 Order By Nombre");	

			this.cmbBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco Order By idBanco Asc");
			this.cmbTarjetaCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTarjeta, Nombre From TarjetasDeCredito Order By Nombre Asc");
			this.cmbPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPlan, Descripcion From Planes Order By idPlan Asc");
			this.cmbPlazo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPlazo, Descripcion From Plazos Order By idPlazo Asc");	
					
			this.dtFecha.Value = DateTime.Today;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;			
			this.dtDesde.Value = "01/01/2015";
			this.dtHasta.Value = DateTime.Today;

			this.dtDesde.CalendarInfo.MinDate = (DateTime)dtDesde.Value;
			this.dtHasta.CalendarInfo.MinDate = (DateTime)dtDesde.Value;
			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;			

			FuncionesProcedimientos.EstadoGrids(false, this.uGridVouchers);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridVouchersConsulta);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			this.dtFecha.Enabled = true;
			this.cmbCuenta.Enabled = true;
			this.txtDeposito.Enabled = true;
			this.txtConcepto.Enabled = true;
			this.txtGastosBancarios.Enabled = true;
						
			FuncionesProcedimientos.EstadoGrids(true, this.uGridVouchers);
			FuncionesProcedimientos.EstadoGrids(true, this.uGridVouchersConsulta);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAsiento.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;

			this.dtFecha.Value = DateTime.Today;			
						
			this.GuardaEncabezado();
			
			bActualiza = true;

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 8", (int)this.txtIdDeposito.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			bNuevo = true;
			bEdicion = true;

			this.dtFecha.Focus();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bNuevo)
			{
				//				if (iContador > 0)
				//				{
				if (DialogResult.Yes == MessageBox.Show(string.Format("¿Esta seguro de cancelar esta captura con {0} documentos seleccionados?", iContador), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec CVCancelaDepositoDeTarjetas {0}", (int)this.txtIdDeposito.Value));
				else return;
				//				}				
			}
					
			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 8", idBloqueaTransacciones, (int)this.txtIdDeposito.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
	
			bNuevo = false;
			bEdicion = false;

			this.txtIdDeposito.Value = 0;
			this.txtIdAsiento.Value = 0;
			this.txtIdRetencion.Value = 0;

			this.dtFecha.Value = System.DBNull.Value;
			this.cmbCuenta.Value = System.DBNull.Value;
			this.txtDeposito.Text = "";
			this.txtConcepto.Text = "";
			this.txtGastosBancarios.Value = 0.00m;
			this.txtTotal.Value = 0.00m;

			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			this.lblContador.Text = "";
			this.lblContadorSeleccionados.Text = "";

			this.txtBuscar.Text = "";
			this.cmbBanco.Value = System.DBNull.Value;
			this.cmbLocal.Value = System.DBNull.Value;

			this.dtFecha.Enabled = false;
			this.cmbCuenta.Enabled = false;
			this.txtDeposito.Enabled = false;
			this.txtConcepto.Enabled = false;
			this.txtGastosBancarios.Enabled = false;
			
			string sSQL = "Select dfp.idDetFormaPago, dfp.idCompra, 0 As idDepositoTarjeta, 0 As Estado, '' As Origen, '' As Almacen, '' As Numero, Null As Fecha, dfp.idBanco, dfp.idTarjeta, dfp.idPlan, dfp.idPlazo, dfp.Lote, dfp.Referencia, dfp.Valor, dfp.ValorUsoTarjeta, dfp.Total, dfp.PorcComision, dfp.Comision, dfp.Retenido, dfp.RetencionIVA, dfp.ValorCobrar From DetalleFormaPago dfp Where idDetFormaPago = 0";
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridVouchers);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridVouchersConsulta);

			FuncionesProcedimientos.EstadoGrids(false, this.uGridVouchers);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridVouchersConsulta);

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;			
			this.btnAsiento.Enabled = false;			
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;

			iContador = 0;
			idBloqueaTransacciones = 0;
			bActualiza = false;
		}

		private void cmbCuenta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDeposito.Focus();
		}

		private void txtDeposito_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13) this.txtConcepto.Focus();
		}

		private void txtConcepto_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtGastosBancarios.Focus();
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbBanco.Focus();
		}

		private void cmbBanco_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbLocal.Focus();
		}

		private void cmbLocal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			this.Cursor  = Cursors.WaitCursor;					
						
			if (bEdicion && bActualiza)
			{
				if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha del deposito", true, "Bancos", cdsSeteoF)) return;
				
				if (!Validacion.vbComboVacio(this.cmbCuenta, "Seleccione una cuenta")) return;

				if (!Validacion.vbTexto(this.txtDeposito, 5, 20, "Numero de deposito")) return;

				if (!this.bValidaNumero()) return;

				if (!Validacion.vbTexto(this.txtConcepto, 5, 100, "Concepto")) return;			

				this.GuardaEncabezado();

				bActualiza = false;
			}

			if (this.tabsGeneral.Tabs[0].Active && (int)this.txtIdDeposito.Value > 0)	this.ConsultaPendientes();
			if (this.tabsGeneral.Tabs[1].Active && (int)this.txtIdDeposito.Value > 0) this.ConsultaCapturados();
			if (this.tabsGeneral.Tabs[2].Active) this.ConsultaGeneral();	
			

			this.Cursor  = Cursors.Default;			
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnBuscar_Click(sender, e);
		}

		private void Consultar (int idDepositoTarjeta)
		{
			string sSQL = string.Format("Exec DTConsultaIndividual {0}", idDepositoTarjeta);
			SqlDataReader drDeposito = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drDeposito.Read();
			if (drDeposito.HasRows)
			{				
				this.txtIdDeposito.Value = drDeposito.GetInt32(0);
				this.txtIdAsiento.Value = drDeposito.GetInt32(1);
				this.txtIdRetencion.Value = drDeposito.GetInt32(2);
				this.dtFecha.Value = drDeposito.GetDateTime(3);
				this.cmbCuenta.Value = drDeposito.GetInt32(4);	
				this.txtDeposito.Value = drDeposito.GetString(5);	
				this.txtConcepto.Text = drDeposito.GetString(6);
				if (drDeposito.GetValue(7)!= System.DBNull.Value)this.txtGastosBancarios.Value = drDeposito.GetDecimal(7);
				this.txtTotal.Value = drDeposito.GetDecimal(8);
				this.txtEstado.Value = drDeposito.GetInt32(9);
				this.lblUsuario.Text = drDeposito.GetString(10);
			}
			drDeposito.Close();
			
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;

			if ((int)this.txtEstado.Value == 0) 
			{
				this.btnEditar.Enabled = true;
				this.btnImprimir.Enabled = true;
				this.btnAsiento.Enabled = true;			
				this.btnAnular.Enabled = true;
				
				this.lblEstado.Text = "DEPOSITADO";

				this.tabsGeneral.Tabs[1].Active = true;
				
				this.ConsultaPendientes();
				this.ConsultaCapturados();	
			}

			if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "ANULADO";

			this.btnCancelar.Enabled = true;				
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			using (frmConsultaDepositoVouchers Consulta = new frmConsultaDepositoVouchers())
				if (DialogResult.OK == Consulta.ShowDialog())					
					this.Consultar((int)Consulta.uGridEncabezado.ActiveRow.Cells["idDepositoTarjeta"].Value);									
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			if (Funcion.ValidaBloqueo((int)this.txtIdDeposito.Value, 8, cdsSeteoF)) return;

			int idCuenta = Funcion.iEscalarSQL(cdsSeteoF, "Select idCuenta From AsignaCuenta Where idAsignacion = 106");

			if (!FuncionesContabilidad.bValidaConciliado((int)this.txtIdAsiento.Value, idCuenta, cdsSeteoF)) return;
			
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha del deposito", true, "Bancos", cdsSeteoF)) return;

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 8", (int)this.txtIdDeposito.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			this.dtFecha.Enabled = true;
			this.cmbCuenta.Enabled = true;
			this.txtDeposito.Enabled = true;
			this.txtConcepto.Enabled = true;
			this.txtGastosBancarios.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;			
			this.btnAsiento.Enabled = false;			
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			
			FuncionesProcedimientos.EstadoGrids(true, this.uGridVouchers);			
			FuncionesProcedimientos.EstadoGrids(true, this.uGridVouchersConsulta);
			
			this.ConsultaPendientes();
			this.ConsultaCapturados();

			bNuevo = false;
			bEdicion = true;

			this.dtFecha.Focus();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtIdAsiento.Value == 0)
			{
				MessageBox.Show("No Existe Deposito para Imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
						
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

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtIdAsiento.Value == 0)
			{
				MessageBox.Show("No existe asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			frmAsientos Asientos = new frmAsientos(((int)this.txtIdAsiento.Value));
			Asientos.MdiParent = MdiParent;
			Asientos.Show();			
		}

		private void GuardaEncabezado ()
		{			
			int idCuenta = 0;
			if (this.cmbCuenta.ActiveRow != null) idCuenta = (int)this.cmbCuenta.Value;

			string sSQLDep = string.Format("Exec GrabaEncabezadoDeposito {0}, {1}, '{2}', '{3}', '{4}', {5}, {6}", 
				(int)this.txtIdDeposito.Value, idCuenta, this.txtDeposito.Text.ToString(), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
				this.txtConcepto.Text.ToString(), Convert.ToDecimal(this.txtTotal.Value), Convert.ToDecimal(this.txtGastosBancarios.Value));
			this.txtIdDeposito.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLDep, true);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			this.tabsGeneral.Tabs[1].Selected = true;

			#region Validacion
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha del deposito", true, "Bancos", cdsSeteoF)) return;
				
			if (!Validacion.vbComboVacio(this.cmbCuenta, "Seleccione una cuenta")) return;

			if (!Validacion.vbTexto(this.txtDeposito, 5, 20, "Numero de deposito")) return;

			if (!this.bValidaNumero()) return;

			if (!Validacion.vbTexto(this.txtConcepto, 5, 100, "Concepto")) return;
			
			if (this.uGridVouchersConsulta.Rows.Count == 0)
			{
				MessageBox.Show("Seleccione el o los Vouchers que va a Depositar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion

			if (DialogResult.Yes == MessageBox.Show("ANTES DE GENERAR EL DEPOSITO. \n\nVERIQUE QUE LOS DATOS REGISTRADOS SON LOS CORRECTOS \n\nDEBE DE CUADRAR LA COMISION, LAS RETENCIONES Y EL VALOR DEL DEPOSITO \n\n\n¿Esta seguro de Generar el Deposito?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{				
				#region Cálculo de valores
				decimal dTotal = 0.00m;
				decimal dCobro = 0.00m;
				decimal dRetencion = 0.00m;
				decimal dRetencionIVA = 0.00m;
				decimal dComision = 0.00m;
				DateTime dtFechaDeposito = (DateTime)this.dtFecha.Value;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVouchersConsulta.Rows.All)
				{
					dTotal = dTotal + Convert.ToDecimal(dr.Cells["Total"].Value);
					dCobro = dCobro + Convert.ToDecimal(dr.Cells["ValorCobrar"].Value);
					dRetencion = dRetencion + Convert.ToDecimal(dr.Cells["Retenido"].Value);
					dRetencionIVA = dRetencionIVA + Convert.ToDecimal(dr.Cells["RetencionIVA"].Value);
					dComision = dComision + Convert.ToDecimal(dr.Cells["Comision"].Value);					
				}
				#endregion Cálculo de valores
			
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

						string sSQL = string.Format("Exec GrabaEncabezadoDeposito {0}, {1}, '{2}', '{3}', '{4}', {5}, {6}", 
							(int)this.txtIdDeposito.Value, (int)this.cmbCuenta.Value, this.txtDeposito.Text.ToString(), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
							this.txtConcepto.Text.ToString(), Convert.ToDecimal(this.txtTotal.Value), Convert.ToDecimal(this.txtGastosBancarios.Value));
						oCmdActualiza.CommandText = sSQL;
						this.txtIdDeposito.Value = (int)oCmdActualiza.ExecuteScalar();
						
						#region Genera Asiento 						
						dCobro = dCobro - Convert.ToDecimal(this.txtGastosBancarios.Value);

						sSQL = string.Format("Exec AsientoTarjetasDeCredito {0}, {1}, '{2}', '{3}', '{4}', '{5}', {6}, {7}, {8}, {9}, {10}, {11}", 
							(int)this.txtIdAsiento.Value, (int)this.cmbCuenta.Value, this.cmbCuenta.Text.ToString(), this.txtConcepto.Text.ToString(), this.txtDeposito.Text.ToString(), 
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), dTotal, dCobro, dRetencion, dRetencionIVA, dComision, Convert.ToDecimal(this.txtGastosBancarios.Value));
						oCmdActualiza.CommandText = sSQL;
						this.txtIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar();

						oCmdActualiza.CommandText = string.Format("Update DepositosDeTarjetas Set idAsiento = {0} Where idDepositoTarjeta = {1}", (int)this.txtIdAsiento.Value, (int)this.txtIdDeposito.Value);
						oCmdActualiza.ExecuteNonQuery();
						#endregion Genera Asiento 

						#region Actualiza Detalle de la forma pago						
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVouchersConsulta.Rows.All)
						{
							if ((int)dr.Cells["Estado"].Value == 1)
							{
								oCmdActualiza.CommandText = string.Format("Update DetalleFormaPago Set Estado = 2 Where idDetFormaPago = {0}", (int)dr.Cells["idDetFormaPago"].Value);
								oCmdActualiza.ExecuteNonQuery();
								
								//						int IdPlazo = 0;
								//						if (dr.Cells["idPlazo"].Value != System.DBNull.Value) IdPlazo = (int)dr.Cells["idPlazo"].Value;
								//
								//						sSQLActDFP = string.Format("Exec ActualizaDetalleFormaPagoTarjetas {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}", 
								//							(int)dr.Cells["idDetFormaPago"].Value, (int)dr.Cells["idBanco"].Value, (int)dr.Cells["idTarjeta"].Value, (int)dr.Cells["idPlan"].Value, IdPlazo, 
								//							(decimal)dr.Cells["Retenido"].Value, (decimal)dr.Cells["RetencionIVA"].Value, (decimal)dr.Cells["Comision"].Value, (decimal)dr.Cells["ValorCobrar"].Value);
								//						Funcion.EjecutaSQL(cdsSeteoF, sSQLActDFP, true);
							}
						}
						#endregion Actualiza Detalle de la forma pago
						
						oCmdActualiza.CommandText = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 8", idBloqueaTransacciones, (int)this.txtIdDeposito.Value);
						oCmdActualiza.ExecuteNonQuery();

						oTransaction.Commit();

						MessageBox.Show("Depósito registrado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);			

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
							
						#region Controles
						this.dtFecha.Enabled = false;
						this.cmbCuenta.Enabled = false;
						this.txtDeposito.Enabled = false;
						this.txtConcepto.Enabled = false;		
						this.txtGastosBancarios.Enabled = false;

						FuncionesProcedimientos.EstadoGrids(false, this.uGridVouchers);
						FuncionesProcedimientos.EstadoGrids(false, this.uGridVouchersConsulta);

						this.btnNuevo.Enabled = true;
						this.btnConsultar.Enabled = true;
						this.btnEditar.Enabled = true;
						this.btnImprimir.Enabled = true;
						this.btnGuardar.Enabled = false;			
						this.btnAsiento.Enabled = false;					
						this.btnAnular.Enabled = false;				

						bEdicion = false;
						bNuevo = false;
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

		private void uGridVouchers_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
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
			stFormato = stMoneda + stCeros.Substring(0, 2);
			stInput = stMascara + stNumero.Substring(0, 2);
			
			e.Layout.Bands[0].Columns["Retenido"].Format = stFormato;			
			e.Layout.Bands[0].Columns["Retenido"].MaskInput = stInput;

			e.Layout.Bands[0].Columns["RetencionIVA"].Format = stFormato;			
			e.Layout.Bands[0].Columns["RetencionIVA"].MaskInput = stInput;
			
			e.Layout.Bands[0].Columns["Comision"].Format = stFormato;			
			e.Layout.Bands[0].Columns["Comision"].MaskInput = stInput;

			e.Layout.Bands[0].Columns["PorcComision"].Format = stFormato;			
			e.Layout.Bands[0].Columns["PorcComision"].MaskInput = stInput;			
			#endregion Decimales			
		}

		private void Suma()
		{
			this.txtTotal.Value = 0;

			this.lblContadorSeleccionados.Text = "";			
			decimal dSuma = 0;

			string sSQL = string.Format("Exec ContadorSumatoriaVouchersReg {0}, {1}", (int)this.txtIdDeposito.Value, MenuLatinium.IdUsuario);
			SqlDataReader drSuma = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drSuma.Read();
			if (drSuma.HasRows)
			{
				iContador = (int)drSuma.GetValue(0);
				dSuma = Convert.ToDecimal(drSuma.GetValue(1));
			}
			drSuma.Close();

			this.lblContadorSeleccionados.Text = iContador.ToString() + " Documentos Capturados";

			this.txtTotal.Value = dSuma;
		}

		private void uGridVouchers_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
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
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set PorcComision = {0} Where idDetFormaPago = {1}", (decimal)e.Cell.Row.Cells["PorcComision"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value));

				e.Cell.Row.Cells["Comision"].Value = Math.Round(((decimal)e.Cell.Row.Cells["Total"].Value * (decimal)e.Cell.Row.Cells["PorcComision"].Value)/100, 2);
				
				e.Cell.Row.Cells["ValorCobrar"].Value = (decimal)e.Cell.Row.Cells["Total"].Value - (decimal)e.Cell.Row.Cells["Retenido"].Value - 
					(decimal)e.Cell.Row.Cells["RetencionIVA"].Value - (decimal)e.Cell.Row.Cells["Comision"].Value;//  - (decimal)e.Cell.Row.Cells["ComisionIVA"].Value;
			
				Suma();

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set ValorCobrar = {0}, Comision = {1} Where idDetFormaPago = {2}", 
					Math.Round((decimal)e.Cell.Row.Cells["ValorCobrar"].Value, 2), Math.Round((decimal)e.Cell.Row.Cells["Comision"].Value, 2), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}

			if (e.Cell.Column.ToString() == "Comision")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set Comision = {0} Where idDetFormaPago = {1}", (decimal)e.Cell.Row.Cells["Comision"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value));

				e.Cell.Row.Cells["ValorCobrar"].Value = (decimal)e.Cell.Row.Cells["Total"].Value - (decimal)e.Cell.Row.Cells["Retenido"].Value - 
					(decimal)e.Cell.Row.Cells["RetencionIVA"].Value - (decimal)e.Cell.Row.Cells["Comision"].Value;// - (decimal)e.Cell.Row.Cells["ComisionIVA"].Value;
				
				Suma();

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set ValorCobrar = {0} Where idDetFormaPago = {1}", Math.Round((decimal)e.Cell.Row.Cells["ValorCobrar"].Value, 2), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}			

			if (e.Cell.Column.ToString() == "Retenido")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set Retenido = {0} Where idDetFormaPago = {1}", Math.Round((decimal)e.Cell.Row.Cells["Retenido"].Value, 2), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));

				e.Cell.Row.Cells["ValorCobrar"].Value = (decimal)e.Cell.Row.Cells["Total"].Value - (decimal)e.Cell.Row.Cells["Retenido"].Value - 
					(decimal)e.Cell.Row.Cells["RetencionIVA"].Value - (decimal)e.Cell.Row.Cells["Comision"].Value;//  - (decimal)e.Cell.Row.Cells["ComisionIVA"].Value;

				Suma();

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set ValorCobrar = {0} Where idDetFormaPago = {1}", Math.Round((decimal)e.Cell.Row.Cells["ValorCobrar"].Value, 2), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}

			if (e.Cell.Column.ToString() == "RetencionIVA")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set RetencionIVA = {0} Where idDetFormaPago = {1}", (decimal)e.Cell.Row.Cells["RetencionIVA"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value));

				e.Cell.Row.Cells["ValorCobrar"].Value = (decimal)e.Cell.Row.Cells["Total"].Value - (decimal)e.Cell.Row.Cells["Retenido"].Value - 
					(decimal)e.Cell.Row.Cells["RetencionIVA"].Value - (decimal)e.Cell.Row.Cells["Comision"].Value;// - (decimal)e.Cell.Row.Cells["ComisionIVA"].Value;
				
				Suma();

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set ValorCobrar = {0} Where idDetFormaPago = {1}", Math.Round((decimal)e.Cell.Row.Cells["ValorCobrar"].Value, 2), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}			
		}

		private void uGridVouchers_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{					
			if (e.Cell.Column.ToString() == "Estado")
			{
				this.uGridVouchers.UpdateData();
				
				string sSQL = string.Format("Exec CVValidaRegistroCapturado {0}, {1}", (int)e.Cell.Row.Cells["idDetFormaPago"].Value, (int)this.txtIdDeposito.Value);
				string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
				
				if (sMensaje.Length > 0)
				{
					e.Cell.Row.Cells["Estado"].Value = false;
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				sSQL = string.Format("Exec CVActualizaCapturado {0}, {1}, {2}", 
					(bool)e.Cell.Row.Cells["Estado"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value, (int)this.txtIdDeposito.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);									

				Suma();				
			}
		}

		private void uGridVouchers_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridVouchers);
		}
		
		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbCuenta.Focus();
		}

		private void txtGastosBancarios_Click(object sender, System.EventArgs e)
		{
			this.txtGastosBancarios.SelectAll();
		}

		private void txtGastosBancarios_Enter(object sender, System.EventArgs e)
		{
			this.txtGastosBancarios.SelectAll();
		}

		private void txtGastosBancarios_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtDesde.Focus();
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{						
			if ((int)this.txtEstado.Value == 1) 
			{
				MessageBox.Show("El Deposito "+ txtDeposito.Value  +" esta anulado","Point Technology",MessageBoxButtons.OK,MessageBoxIcon.Warning);
				return;
			}
						
			if(Funcion.ValidaBloqueo((int)this.txtIdDeposito.Value, 8, cdsSeteoF)) return;

			int idCuenta = Funcion.iEscalarSQL(cdsSeteoF, "Select idCuenta From AsignaCuenta Where idAsignacion = 106");

			if (!FuncionesContabilidad.bValidaConciliado((int)this.txtIdAsiento.Value, idCuenta, cdsSeteoF)) return;

			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha del deposito", true, "Bancos", cdsSeteoF)) return;
			
			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Anular este documento?", "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				try
				{
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AnulaAsientoDepositoTarjetas {0}, 1", (int)this.txtIdDeposito.Value));
			
					MessageBox.Show("Deposito Numeto  "+ this.txtDeposito.Text.ToString() +" anulado con exito","Point Technology",MessageBoxButtons.OK,MessageBoxIcon.Information);			
				}
				catch(Exception Exc)
				{
					MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}						
			}			
		}

		#region Consultas
		private void ConsultaPendientes()
		{
			bool bFechas = false;

			DateTime dtDesde = DateTime.Today;
			DateTime dtHasta = DateTime.Today;

			if (this.dtDesde.Value != System.DBNull.Value || this.dtHasta.Value != System.DBNull.Value)
			{
				bFechas = true;

				dtDesde = (DateTime)this.dtDesde.Value;
				dtHasta = (DateTime)this.dtHasta.Value;
			}

			int iBodega = 0;
			int idBanco = 0;

			if (this.cmbLocal.ActiveRow != null) iBodega = (int)this.cmbLocal.Value;			
			if (this.cmbBanco.ActiveRow != null) idBanco = (int)this.cmbBanco.Value;

			string sSQL = string.Format("Exec ListaVouchersPendientesDeDepositar '{0}', '{1}', {2}, {3}, '{4}', {5}", 
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), iBodega, idBanco, this.txtBuscar.Text.ToString(), (int)this.txtIdDeposito.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridVouchers);

			this.lblContador.Text = this.uGridVouchers.Rows.Count + " REGISTROS ENCONTRADOS";	
		}

		private void ConsultaCapturados()
		{			
			string sSQLc = string.Format("Exec ListaVouchersRegistrados {0}", (int)this.txtIdDeposito.Value);
			FuncionesProcedimientos.dsGeneral(sSQLc, this.uGridVouchersConsulta);			

			this.lblContador.Text = this.uGridVouchers.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void ConsultaGeneral()
		{
			bool bFechas = false;

			DateTime dtDesde = DateTime.Today;
			DateTime dtHasta = DateTime.Today;

			if (this.dtDesde.Value != System.DBNull.Value || this.dtHasta.Value != System.DBNull.Value)
			{
				bFechas = true;

				dtDesde = (DateTime)this.dtDesde.Value;
				dtHasta = (DateTime)this.dtHasta.Value;
			}

			int iBodega = 0;
			int idBanco = 0;

			if (this.cmbLocal.ActiveRow != null) iBodega = (int)this.cmbLocal.Value;			
			if (this.cmbBanco.ActiveRow != null) idBanco = (int)this.cmbBanco.Value;

			string sSQL = string.Format("Exec ListaVouchersDepositados '{0}', '{1}', {2}, {3}, '{4}', {5}", 
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), iBodega, idBanco, this.txtBuscar.Text.ToString(), (int)this.optEstados.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridVouchersGeneral);		

			this.lblContador.Text = this.uGridVouchersGeneral.Rows.Count + " REGISTROS ENCONTRADOS";
		}
		#endregion Consultas

		private bool bValidaNumero()
		{			
			if(Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaNumeroDocumento {0}, '{1}', 2, {2}", (int)this.cmbCuenta.Value, this.txtDeposito.Text.ToString(), (int)txtIdAsiento.Value), true))
			{
				MessageBox.Show("Numero de deposito " + this.txtDeposito.Value + " ya registrado","Point Technology",MessageBoxButtons.OK,MessageBoxIcon.Stop) ;
				return false;
			}		

			return true;
		}
		
		private void utcTransaccion_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{		
			if (this.tabsGeneral.Tabs[0].Active && (int)this.txtIdDeposito.Value > 0) this.ConsultaPendientes();
			if (this.tabsGeneral.Tabs[1].Active && (int)this.txtIdDeposito.Value > 0) this.ConsultaCapturados();
			if (this.tabsGeneral.Tabs[2].Active)	this.ConsultaGeneral();
		}

		private void frmDepositoDeVouchers_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras se encuentre realizando algun proceso","Point Technology" ,MessageBoxButtons.OK,MessageBoxIcon.Warning);
				e.Cancel = true;			
			}
		}
		
		private void uGridVouchersConsulta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridVouchersConsulta);
		}

		private void uGridVouchersConsulta_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
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
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set PorcComision = {0} Where idDetFormaPago = {1}", (decimal)e.Cell.Row.Cells["PorcComision"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value));

				e.Cell.Row.Cells["Comision"].Value = Math.Round(((decimal)e.Cell.Row.Cells["Total"].Value * (decimal)e.Cell.Row.Cells["PorcComision"].Value)/100, 2);
				
				e.Cell.Row.Cells["ValorCobrar"].Value = (decimal)e.Cell.Row.Cells["Total"].Value - (decimal)e.Cell.Row.Cells["Retenido"].Value - 
					(decimal)e.Cell.Row.Cells["RetencionIVA"].Value - (decimal)e.Cell.Row.Cells["Comision"].Value;//  - (decimal)e.Cell.Row.Cells["ComisionIVA"].Value;
			
				Suma();

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set ValorCobrar = {0}, Comision = {1} Where idDetFormaPago = {2}", 
					Math.Round((decimal)e.Cell.Row.Cells["ValorCobrar"].Value, 2), Math.Round((decimal)e.Cell.Row.Cells["Comision"].Value, 2), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}

			if (e.Cell.Column.ToString() == "Comision")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set Comision = {0} Where idDetFormaPago = {1}", (decimal)e.Cell.Row.Cells["Comision"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value));

				e.Cell.Row.Cells["ValorCobrar"].Value = (decimal)e.Cell.Row.Cells["Total"].Value - (decimal)e.Cell.Row.Cells["Retenido"].Value - 
					(decimal)e.Cell.Row.Cells["RetencionIVA"].Value - (decimal)e.Cell.Row.Cells["Comision"].Value;// - (decimal)e.Cell.Row.Cells["ComisionIVA"].Value;
				
				Suma();

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set ValorCobrar = {0} Where idDetFormaPago = {1}", Math.Round((decimal)e.Cell.Row.Cells["ValorCobrar"].Value, 2), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}			

			if (e.Cell.Column.ToString() == "Retenido")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set Retenido = {0} Where idDetFormaPago = {1}", Math.Round((decimal)e.Cell.Row.Cells["Retenido"].Value, 2), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));

				e.Cell.Row.Cells["ValorCobrar"].Value = (decimal)e.Cell.Row.Cells["Total"].Value - (decimal)e.Cell.Row.Cells["Retenido"].Value - 
					(decimal)e.Cell.Row.Cells["RetencionIVA"].Value - (decimal)e.Cell.Row.Cells["Comision"].Value;//  - (decimal)e.Cell.Row.Cells["ComisionIVA"].Value;

				Suma();

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set ValorCobrar = {0} Where idDetFormaPago = {1}", Math.Round((decimal)e.Cell.Row.Cells["ValorCobrar"].Value, 2), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}

			if (e.Cell.Column.ToString() == "RetencionIVA")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set RetencionIVA = {0} Where idDetFormaPago = {1}", (decimal)e.Cell.Row.Cells["RetencionIVA"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value));

				e.Cell.Row.Cells["ValorCobrar"].Value = (decimal)e.Cell.Row.Cells["Total"].Value - (decimal)e.Cell.Row.Cells["Retenido"].Value - 
					(decimal)e.Cell.Row.Cells["RetencionIVA"].Value - (decimal)e.Cell.Row.Cells["Comision"].Value;// - (decimal)e.Cell.Row.Cells["ComisionIVA"].Value;
				
				Suma();

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleFormaPago Set ValorCobrar = {0} Where idDetFormaPago = {1}", Math.Round((decimal)e.Cell.Row.Cells["ValorCobrar"].Value, 2), (int)e.Cell.Row.Cells["idDetFormaPago"].Value));
			}			
		}

		private void frmDepositoDeVouchers_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdDeposito.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);																	 
					if (e.KeyCode == Keys.F5)
					{
						//miAcceso = new Acceso(cdsSeteoF, "0703");

						this.Consultar((int)this.txtIdDeposito.Value);
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

		private void uGridVouchersConsulta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.tabsGeneral.Tabs[0].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridVouchers);
			if (this.tabsGeneral.Tabs[1].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridVouchersConsulta);
			if (this.tabsGeneral.Tabs[2].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridVouchersGeneral);
		}
	}
}

