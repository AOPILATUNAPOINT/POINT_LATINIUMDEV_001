using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmPrestamos.
	/// </summary>
	public class frmPrestamos : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdPrestamo;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		bool bNuevo = false;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTasaDeInteres;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoPrestamo;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFechaPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTasaNominal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCapital;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtInteres;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label lblFechaSistema;
		private System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaPrestamo;
		public Infragistics.Win.UltraWinEditors.UltraOptionSet optBancoProveedor;
		bool bEdicion = false;

		DateTime dtFechaPago = DateTime.Today;
		int iCuotas = 0;
		decimal dTasaNominal = 0.00m;
		decimal dMonto = 0.00m;
		private System.Windows.Forms.Button btnCalcular;
		private System.Windows.Forms.Button btnImprimir;
		decimal dInteres = 0.00m;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbLineaCredito;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsLineaCredito;
		private Acceso miAcceso;

		public frmPrestamos()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPrestamos));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetallePrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroCuota");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Inicia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vence");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuota");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pago");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnPago", 0);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorCuota", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorCuota", 5, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Capital", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Capital", 6, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Interes", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Interes", 7, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetallePrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Inicia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vence");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTiposDePrestamos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTiposDePrestamos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			this.txtIdPrestamo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label8 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label5 = new System.Windows.Forms.Label();
			this.txtTasaDeInteres = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.txtFechaPago = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label9 = new System.Windows.Forms.Label();
			this.txtTasaNominal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.txtCapital = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtInteres = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbTipoPrestamo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label13 = new System.Windows.Forms.Label();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.lblFechaSistema = new System.Windows.Forms.Label();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.optBancoProveedor = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label14 = new System.Windows.Forms.Label();
			this.cmbCuentaPrestamo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnCalcular = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.label15 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExcel = new System.Windows.Forms.Button();
			this.label16 = new System.Windows.Forms.Label();
			this.cmbLineaCredito = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsLineaCredito = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTasaDeInteres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFechaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTasaNominal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCapital)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optBancoProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLineaCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsLineaCredito)).BeginInit();
			this.SuspendLayout();
			// 
			// txtIdPrestamo
			// 
			this.txtIdPrestamo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdPrestamo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdPrestamo.Enabled = false;
			this.txtIdPrestamo.Location = new System.Drawing.Point(648, 586);
			this.txtIdPrestamo.Name = "txtIdPrestamo";
			this.txtIdPrestamo.PromptChar = ' ';
			this.txtIdPrestamo.Size = new System.Drawing.Size(16, 22);
			this.txtIdPrestamo.TabIndex = 291;
			this.txtIdPrestamo.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(480, 136);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 289;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 586);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 288;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(82, 586);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 287;
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
			this.btnEditar.Location = new System.Drawing.Point(162, 586);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 286;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(319, 586);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 285;
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
			this.btnAnular.Location = new System.Drawing.Point(397, 586);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 284;
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
			this.btnCancelar.Location = new System.Drawing.Point(475, 586);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 283;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(549, 586);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 282;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// cmbBanco
			// 
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataSource = this.ultraDataSource1;
			this.cmbBanco.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 175;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 285;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.Enabled = false;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(128, 40);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(304, 21);
			this.cmbBanco.TabIndex = 0;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBanco_KeyDown);
			this.cmbBanco.ValueChanged += new System.EventHandler(this.cmbBanco_ValueChanged);
			this.cmbBanco.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBanco_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			this.ultraDataSource1.Band.Key = "Personal";
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.groupBox1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.groupBox1.Location = new System.Drawing.Point(0, 570);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(704, 8);
			this.groupBox1.TabIndex = 280;
			this.groupBox1.TabStop = false;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(528, 40);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(44, 16);
			this.label8.TabIndex = 279;
			this.label8.Text = "Número";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(584, 40);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(96, 22);
			this.txtNumero.TabIndex = 1;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 40);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(34, 16);
			this.label7.TabIndex = 277;
			this.label7.Text = "Banco";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 168);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 275;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 264);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(35, 16);
			this.label3.TabIndex = 273;
			this.label3.Text = "Monto";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 296);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 16);
			this.label1.TabIndex = 268;
			this.label1.Text = "Concepto";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Header.Caption = "Cuota";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 36;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 68;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance3;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Valor Cuota";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 92;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance4;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 88;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance5;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Width = 89;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance6;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance7;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Width = 88;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 10;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Estado";
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Width = 113;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 13;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			appearance8.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance8.ImageVAlign = Infragistics.Win.VAlign.Middle;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			appearance8.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn17.CellAppearance = appearance8;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			appearance9.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance9.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn17.CellButtonAppearance = appearance9;
			ultraGridColumn17.Header.Caption = "Pago";
			ultraGridColumn17.Header.VisiblePosition = 14;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn17.Width = 85;
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
																										 ultraGridColumn17});
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance10;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance11;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance12;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand2.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 10F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 400);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(680, 160);
			this.ultraGrid1.TabIndex = 276;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn6.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn7.DefaultValue = new System.DateTime(((long)(0)));
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
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn13.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
			// txtFecha
			// 
			this.txtFecha.AllowMonthSelection = true;
			this.txtFecha.AutoSelectionUpdate = true;
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.DayDisplayStyle = Infragistics.Win.UltraWinSchedule.DayDisplayStyle.DayNumberAndImage;
			this.txtFecha.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.txtFecha.Enabled = false;
			this.txtFecha.Format = "dd/MM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(80, 168);
			this.txtFecha.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(112, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 3;
			this.txtFecha.Value = ((object)(resources.GetObject("txtFecha.Value")));
			this.txtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFecha_KeyDown);
			this.txtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.txtFecha_BeforeDropDown);
			// 
			// txtValor
			// 
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(80, 264);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(96, 22);
			this.txtValor.TabIndex = 9;
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			this.txtValor.Validated += new System.EventHandler(this.txtValor_Validated);
			this.txtValor.ValueChanged += new System.EventHandler(this.txtValor_ValueChanged);
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			// 
			// txtConcepto
			// 
			this.txtConcepto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Enabled = false;
			this.txtConcepto.Location = new System.Drawing.Point(80, 296);
			this.txtConcepto.Multiline = true;
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(600, 80);
			this.txtConcepto.TabIndex = 12;
			this.txtConcepto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtConcepto_KeyDown);
			this.txtConcepto.ValueChanged += new System.EventHandler(this.txtConcepto_ValueChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 232);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(38, 16);
			this.label2.TabIndex = 294;
			this.label2.Text = "Cuotas";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCuotas
			// 
			this.txtCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotas.Enabled = false;
			this.txtCuotas.FormatString = "";
			this.txtCuotas.Location = new System.Drawing.Point(80, 232);
			this.txtCuotas.Name = "txtCuotas";
			this.txtCuotas.PromptChar = ' ';
			this.txtCuotas.Size = new System.Drawing.Size(64, 22);
			this.txtCuotas.TabIndex = 5;
			this.txtCuotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuotas_KeyDown);
			this.txtCuotas.Click += new System.EventHandler(this.txtCuotas_Click);
			this.txtCuotas.Validated += new System.EventHandler(this.txtCuotas_Validated);
			this.txtCuotas.ValueChanged += new System.EventHandler(this.txtCuotas_ValueChanged);
			this.txtCuotas.Enter += new System.EventHandler(this.txtCuotas_Enter);
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(680, 586);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 295;
			this.txtEstado.Visible = false;
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
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(216, 232);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(84, 16);
			this.label5.TabIndex = 297;
			this.label5.Text = "Tasa De Interes";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTasaDeInteres
			// 
			this.txtTasaDeInteres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTasaDeInteres.Enabled = false;
			this.txtTasaDeInteres.FormatString = "#,##0.00";
			this.txtTasaDeInteres.Location = new System.Drawing.Point(320, 232);
			this.txtTasaDeInteres.MaxValue = 100;
			this.txtTasaDeInteres.MinValue = 0;
			this.txtTasaDeInteres.Name = "txtTasaDeInteres";
			this.txtTasaDeInteres.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTasaDeInteres.PromptChar = ' ';
			this.txtTasaDeInteres.Size = new System.Drawing.Size(96, 22);
			this.txtTasaDeInteres.TabIndex = 6;
			this.txtTasaDeInteres.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTasaDeInteres_KeyDown);
			this.txtTasaDeInteres.Click += new System.EventHandler(this.txtTasaDeInteres_Click);
			this.txtTasaDeInteres.Validated += new System.EventHandler(this.txtTasaDeInteres_Validated);
			this.txtTasaDeInteres.ValueChanged += new System.EventHandler(this.txtTasaDeInteres_ValueChanged);
			this.txtTasaDeInteres.Enter += new System.EventHandler(this.txtTasaDeInteres_Enter);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(216, 168);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(77, 16);
			this.label6.TabIndex = 299;
			this.label6.Text = "Fecha de Pago";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFechaPago
			// 
			this.txtFechaPago.AllowMonthSelection = true;
			this.txtFechaPago.AutoSelectionUpdate = true;
			this.txtFechaPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.txtFechaPago.DateButtons.Add(dateButton2);
			this.txtFechaPago.DayDisplayStyle = Infragistics.Win.UltraWinSchedule.DayDisplayStyle.DayNumberAndImage;
			this.txtFechaPago.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.txtFechaPago.Enabled = false;
			this.txtFechaPago.Format = "dd/MM/yyyy";
			this.txtFechaPago.Location = new System.Drawing.Point(320, 168);
			this.txtFechaPago.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.txtFechaPago.Name = "txtFechaPago";
			this.txtFechaPago.NonAutoSizeHeight = 23;
			this.txtFechaPago.Size = new System.Drawing.Size(112, 21);
			this.txtFechaPago.SpinButtonsVisible = true;
			this.txtFechaPago.TabIndex = 4;
			this.txtFechaPago.Value = ((object)(resources.GetObject("txtFechaPago.Value")));
			this.txtFechaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFechaPago_KeyDown);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(496, 232);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(72, 16);
			this.label9.TabIndex = 301;
			this.label9.Text = "Tasa Nominal";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTasaNominal
			// 
			this.txtTasaNominal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTasaNominal.Enabled = false;
			this.txtTasaNominal.FormatString = "#,##0.00";
			this.txtTasaNominal.Location = new System.Drawing.Point(584, 232);
			this.txtTasaNominal.MaxValue = 100;
			this.txtTasaNominal.MinValue = 0;
			this.txtTasaNominal.Name = "txtTasaNominal";
			this.txtTasaNominal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTasaNominal.PromptChar = ' ';
			this.txtTasaNominal.Size = new System.Drawing.Size(96, 22);
			this.txtTasaNominal.TabIndex = 7;
			this.txtTasaNominal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTasaNominal_KeyDown);
			this.txtTasaNominal.Click += new System.EventHandler(this.txtTasaNominal_Click);
			this.txtTasaNominal.Validated += new System.EventHandler(this.txtTasaNominal_Validated);
			this.txtTasaNominal.Enter += new System.EventHandler(this.txtTasaNominal_Enter);
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource3;
			this.cmbCuenta.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "Personal";
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 110;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Width = 75;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 210;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.Enabled = false;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(128, 72);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(304, 21);
			this.cmbCuenta.TabIndex = 2;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCuenta_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(8, 72);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(86, 16);
			this.label10.TabIndex = 303;
			this.label10.Text = "Cuenta Bancaria";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(440, 264);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(38, 16);
			this.label11.TabIndex = 305;
			this.label11.Text = "Capital";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label11.Visible = false;
			// 
			// txtCapital
			// 
			this.txtCapital.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCapital.Enabled = false;
			this.txtCapital.FormatString = "#,##0.00";
			this.txtCapital.Location = new System.Drawing.Point(584, 264);
			this.txtCapital.Name = "txtCapital";
			this.txtCapital.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCapital.PromptChar = ' ';
			this.txtCapital.Size = new System.Drawing.Size(96, 22);
			this.txtCapital.TabIndex = 10;
			this.txtCapital.Visible = false;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(216, 264);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(40, 16);
			this.label12.TabIndex = 307;
			this.label12.Text = "Interes";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtInteres
			// 
			this.txtInteres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInteres.Enabled = false;
			this.txtInteres.FormatString = "#,##0.00";
			this.txtInteres.Location = new System.Drawing.Point(320, 264);
			this.txtInteres.Name = "txtInteres";
			this.txtInteres.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtInteres.PromptChar = ' ';
			this.txtInteres.Size = new System.Drawing.Size(96, 22);
			this.txtInteres.TabIndex = 11;
			this.txtInteres.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtInteres_KeyDown);
			this.txtInteres.Click += new System.EventHandler(this.txtInteres_Click);
			this.txtInteres.ValueChanged += new System.EventHandler(this.txtInteres_ValueChanged);
			this.txtInteres.Enter += new System.EventHandler(this.txtInteres_Enter);
			// 
			// cmbTipoPrestamo
			// 
			this.cmbTipoPrestamo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoPrestamo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoPrestamo.DataSource = this.ultraDataSource4;
			this.cmbTipoPrestamo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "Personal";
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 157;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 221;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbTipoPrestamo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipoPrestamo.DisplayMember = "Tipo";
			this.cmbTipoPrestamo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoPrestamo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoPrestamo.Enabled = false;
			this.cmbTipoPrestamo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoPrestamo.Location = new System.Drawing.Point(80, 200);
			this.cmbTipoPrestamo.Name = "cmbTipoPrestamo";
			this.cmbTipoPrestamo.Size = new System.Drawing.Size(240, 21);
			this.cmbTipoPrestamo.TabIndex = 8;
			this.cmbTipoPrestamo.ValueMember = "idTiposDePrestamos";
			this.cmbTipoPrestamo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoPrestamo_KeyDown);
			this.cmbTipoPrestamo.ValueChanged += new System.EventHandler(this.cmbTipoPrestamo_ValueChanged);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn20.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(8, 200);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(26, 16);
			this.label13.TabIndex = 309;
			this.label13.Text = "Tipo";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(664, 586);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsiento.TabIndex = 310;
			this.txtIdAsiento.Visible = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 384);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(696, 8);
			this.groupBox2.TabIndex = 311;
			this.groupBox2.TabStop = false;
			// 
			// lblFechaSistema
			// 
			this.lblFechaSistema.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblFechaSistema.AutoSize = true;
			this.lblFechaSistema.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFechaSistema.ForeColor = System.Drawing.Color.Firebrick;
			this.lblFechaSistema.Location = new System.Drawing.Point(344, 8);
			this.lblFechaSistema.Name = "lblFechaSistema";
			this.lblFechaSistema.Size = new System.Drawing.Size(0, 19);
			this.lblFechaSistema.TabIndex = 313;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(496, 8);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 19);
			this.lblUsuario.TabIndex = 314;
			// 
			// optBancoProveedor
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.optBancoProveedor.Appearance = appearance17;
			this.optBancoProveedor.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optBancoProveedor.Enabled = false;
			this.optBancoProveedor.ItemAppearance = appearance18;
			this.optBancoProveedor.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Banco";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Terceros";
			this.optBancoProveedor.Items.Add(valueListItem1);
			this.optBancoProveedor.Items.Add(valueListItem2);
			this.optBancoProveedor.ItemSpacingVertical = 10;
			this.optBancoProveedor.Location = new System.Drawing.Point(8, 8);
			this.optBancoProveedor.Name = "optBancoProveedor";
			this.optBancoProveedor.Size = new System.Drawing.Size(160, 26);
			this.optBancoProveedor.TabIndex = 315;
			this.optBancoProveedor.ValueChanged += new System.EventHandler(this.optBancoProveedor_ValueChanged);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(8, 104);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(90, 16);
			this.label14.TabIndex = 317;
			this.label14.Text = "Cuenta Prestamo";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCuentaPrestamo
			// 
			this.cmbCuentaPrestamo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaPrestamo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaPrestamo.DataSource = this.ultraDataSource3;
			this.cmbCuentaPrestamo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand5.AddButtonCaption = "Personal";
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 110;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 106;
			ultraGridColumn25.Header.VisiblePosition = 2;
			ultraGridColumn25.Width = 179;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbCuentaPrestamo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCuentaPrestamo.DisplayMember = "Descripcion";
			this.cmbCuentaPrestamo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaPrestamo.Enabled = false;
			this.cmbCuentaPrestamo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaPrestamo.Location = new System.Drawing.Point(128, 104);
			this.cmbCuentaPrestamo.Name = "cmbCuentaPrestamo";
			this.cmbCuentaPrestamo.Size = new System.Drawing.Size(304, 21);
			this.cmbCuentaPrestamo.TabIndex = 316;
			this.cmbCuentaPrestamo.ValueMember = "idCuenta";
			this.cmbCuentaPrestamo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCuentaPrestamo_KeyDown);
			// 
			// btnCalcular
			// 
			this.btnCalcular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCalcular.CausesValidation = false;
			this.btnCalcular.Enabled = false;
			this.btnCalcular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCalcular.Location = new System.Drawing.Point(472, 168);
			this.btnCalcular.Name = "btnCalcular";
			this.btnCalcular.Size = new System.Drawing.Size(88, 23);
			this.btnCalcular.TabIndex = 318;
			this.btnCalcular.Text = "&Calcular";
			this.btnCalcular.Click += new System.EventHandler(this.btnCalcular_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(240, 586);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 374;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(344, 200);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(80, 16);
			this.label15.TabIndex = 376;
			this.label15.Text = "Forma de Pago";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource5;
			this.cmbFormaPago.DisplayLayout.AutoFitColumns = true;
			ultraGridBand6.AddButtonCaption = "Personal";
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 101;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 221;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(440, 200);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(240, 21);
			this.cmbFormaPago.TabIndex = 375;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbFormaPago_KeyDown);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn22.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22,
																																 ultraDataColumn23});
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(600, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(83, 24);
			this.btnExcel.TabIndex = 691;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(8, 136);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(84, 16);
			this.label16.TabIndex = 693;
			this.label16.Text = "Línea de credito";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbLineaCredito
			// 
			this.cmbLineaCredito.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbLineaCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLineaCredito.DataSource = this.udsLineaCredito;
			this.cmbLineaCredito.DisplayLayout.AutoFitColumns = true;
			ultraGridBand7.AddButtonCaption = "Personal";
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 135;
			ultraGridColumn29.Header.Caption = "Número";
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Width = 285;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			this.cmbLineaCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbLineaCredito.DisplayMember = "Numero";
			this.cmbLineaCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbLineaCredito.Enabled = false;
			this.cmbLineaCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLineaCredito.Location = new System.Drawing.Point(128, 136);
			this.cmbLineaCredito.Name = "cmbLineaCredito";
			this.cmbLineaCredito.Size = new System.Drawing.Size(304, 21);
			this.cmbLineaCredito.TabIndex = 692;
			this.cmbLineaCredito.ValueMember = "idLineaCredito";
			// 
			// udsLineaCredito
			// 
			ultraDataColumn24.DataType = typeof(int);
			this.udsLineaCredito.Band.Columns.AddRange(new object[] {
																																ultraDataColumn24,
																																ultraDataColumn25});
			// 
			// frmPrestamos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(698, 616);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbLineaCredito);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnCalcular);
			this.Controls.Add(this.cmbCuentaPrestamo);
			this.Controls.Add(this.optBancoProveedor);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.lblFechaSistema);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.cmbTipoPrestamo);
			this.Controls.Add(this.txtInteres);
			this.Controls.Add(this.txtCapital);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.txtTasaNominal);
			this.Controls.Add(this.txtFechaPago);
			this.Controls.Add(this.txtTasaDeInteres);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtCuotas);
			this.Controls.Add(this.txtIdPrestamo);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.txtConcepto);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.Name = "frmPrestamos";
			this.Text = "Prestamos Por Pagar";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmPrestamos_Closing);
			this.Load += new System.EventHandler(this.frmPrestamos_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtIdPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTasaDeInteres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFechaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTasaNominal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCapital)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optBancoProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLineaCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsLineaCredito)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			this.txtIdPrestamo.Value = 0;
			this.txtEstado.Value = 0;
			this.txtIdAsiento.Value = 0;

			this.optBancoProveedor.Value = 0;
			this.cmbBanco.Value = System.DBNull.Value;			
			this.txtNumero.Text = "";
			this.cmbCuenta.Value = System.DBNull.Value;
			this.cmbCuentaPrestamo.Value = System.DBNull.Value;
			this.txtFecha.Value = System.DBNull.Value;
			this.txtFechaPago.Value = System.DBNull.Value;
			this.txtCuotas.Value = 0;
			this.txtTasaDeInteres.Value = 0;
			this.txtTasaNominal.Value = 0;
			this.cmbTipoPrestamo.Value = System.DBNull.Value;
			this.cmbFormaPago.Value = System.DBNull.Value;
			this.txtValor.Value = 0;
			this.txtCapital.Value = 0;
			this.txtInteres.Value = 0;
			this.txtConcepto.Text = "";
			this.lblEstado.Text = "";
			this.lblFechaSistema.Text = "";
			this.lblUsuario.Text = "";

			this.optBancoProveedor.Enabled = false;
			this.cmbBanco.Enabled = false;
			this.cmbCuenta.Enabled = false;
			this.cmbCuentaPrestamo.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtFecha.Enabled = false;
			this.txtFechaPago.Enabled = false;
			this.txtCuotas.Enabled = false;
			this.txtTasaDeInteres.Enabled = false;
			this.txtTasaNominal.Enabled = false;
			this.cmbTipoPrestamo.Enabled = false;
			this.cmbFormaPago.Enabled = false;
			this.txtValor.Enabled = false;
			this.txtCapital.Enabled = false;
			this.txtInteres.Enabled = false;
			this.txtConcepto.Enabled = false;

			#region Linea credito
			//Línea de crédito
			ActivarLineaCredito(false, true);
			#endregion Linea credito


			ModificaGrilla(false);

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec PrestamosConsultaDetalle -1");
	
			this.btnCalcular.Enabled = false;
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;

			dtFechaPago = DateTime.Today;
			iCuotas = 0;
			dTasaNominal = 0.00m;
			dMonto = 0.00m;
			dInteres = 0.00m;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.optBancoProveedor.Enabled = true;
			this.cmbBanco.Enabled = true;
			this.txtNumero.Enabled = true;
			this.cmbCuenta.Enabled = true;
			this.cmbCuentaPrestamo.Enabled = true;
			this.txtFecha.Enabled = true;
			this.txtFechaPago.Enabled = true;
			this.txtCuotas.Enabled = true;
			this.txtTasaDeInteres.Enabled = true;
			this.txtTasaNominal.Enabled = true;
			this.cmbTipoPrestamo.Enabled = true;
			this.cmbFormaPago.Enabled = true;
			this.txtValor.Enabled = true;
			this.txtCapital.Enabled = true;
			//this.txtInteres.Enabled = true;
			this.txtConcepto.Enabled = true;

			this.ModificaGrilla(true);

			this.btnCalcular.Enabled = true;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;

			bNuevo = true;
			bEdicion = true;

			this.txtFecha.Value = DateTime.Today;
			this.txtFechaPago.Value = DateTime.Today;
			
			this.optBancoProveedor.Value = 1;
			this.cmbBanco.Focus();
			
			#region Linea credito
			//Línea de crédito
			ActivarLineaCredito(true, true);
			#endregion Linea
		}

		private void ActivarLineaCredito(bool Activa, bool Setear)
		{
			this.cmbLineaCredito.Enabled = Activa;
			if(Setear == true)
			{
				this.cmbLineaCredito.Value = System.DBNull.Value;
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private int iCuotasPagadas()
		{
			int i = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if ((int)dr.Cells["idAsiento"].Value > 0) i++;
			}

			return i;
		}

		private int iUltimaCuotasPagada()
		{
			int i = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				if ((int)dr.Cells["Estado"].Value > 0) i++;

			i++;

			return i;
		}

		private void frmPrestamos_Load(object sender, System.EventArgs e)
		{
			string stTipo = "040806";
			string stPermiso = "468";
							
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			
			if (!Funcion.Permiso(stPermiso,cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Prestamos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				this.ModificaGrilla(false);
						
				this.cmbTipoPrestamo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTiposDePrestamos, Tipo From TiposDePrestamos");

				//this.cmbCuentaPrestamo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCuenta, Codigo, Descripcion From Cuenta Where Tipo = 2 And Grupo = 0 Order By Codigo");

				this.cmbCuentaPrestamo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ConsultaCuentaPrestamos  "));
				this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (2, 16)");
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			}
			//Cargar línea de créditos
			LineaCreditoCmb();

		}

		public void LineaCreditoCmb()
		{
			string squery = "";
			try
			{
				squery = String.Format("Exec LineaCreditoCmb");
				this.cmbLineaCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, squery);
			}
			catch(Exception ex)
			{
				Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
				MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void txtCuotas_Click(object sender, System.EventArgs e)
		{
			this.txtCuotas.SelectAll();
		}

		private void txtCuotas_Enter(object sender, System.EventArgs e)
		{
			this.txtCuotas.SelectAll();
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtValor_Enter(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtCuotas_Validated(object sender, System.EventArgs e)
		{
			if (this.txtCuotas.Value == System.DBNull.Value) this.txtCuotas.Value = 0;			
		}

		private void cmbBanco_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNumero.Focus();
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCuenta.Focus();
		}

		private void txtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtFechaPago.Focus();
		}

		private void txtCuotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtTasaDeInteres.Focus();
		}

		private void txtValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.cmbTipoPrestamo.ActiveRow != null)
				{
					if (Convert.ToInt32(this.cmbTipoPrestamo.Value) == 1) this.txtConcepto.Focus();
					else if (Convert.ToInt32(this.cmbTipoPrestamo.Value) == 1) this.txtInteres.Focus();
				}
				else this.cmbTipoPrestamo.Focus();
			}
		}

		private void txtConcepto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdPrestamo.Value);

			if (this.iCuotasPagadas() >= 0)//==0
			{
				this.optBancoProveedor.Enabled = true;
				this.cmbBanco.Enabled = true;
				this.cmbCuenta.Enabled = true;
				this.cmbCuentaPrestamo.Enabled = true;
				this.txtFecha.Enabled = true;
				this.txtFechaPago.Enabled = true;
				this.cmbTipoPrestamo.Enabled = true;
				this.cmbFormaPago.Enabled = true;
				this.txtCuotas.Enabled = true;
				this.txtTasaDeInteres.Enabled = false;
				this.txtTasaNominal.Enabled = false;
				this.txtValor.Enabled = true;
				if ((int)this.cmbTipoPrestamo.Value == 2) this.txtInteres.Enabled = true;
				this.txtConcepto.Enabled = true;
				
				#region Linea credito
				//Línea de crédito
				ActivarLineaCredito(true, false);
				#endregion Linea

				bNuevo = false;
				bEdicion = true;

				this.ModificaGrilla(true);

				this.btnCalcular.Enabled = true;
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = true;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = true;

				dtFechaPago = Convert.ToDateTime(this.txtFechaPago.Value);
				iCuotas = Convert.ToInt32(this.txtCuotas.Value);
				dTasaNominal = Convert.ToDecimal(this.txtTasaNominal.Value);
				dMonto = Convert.ToDecimal(this.txtValor.Value);
				dInteres = Convert.ToDecimal(this.txtInteres.Value);
			}
			else
			{
				MessageBox.Show("No puede Editar tiene Cuotas ya pagadas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				return;
			}
		}

		private void txtTasaDeInteres_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtTasaNominal.Focus();
		}

		private void txtTasaDeInteres_Click(object sender, System.EventArgs e)
		{
			this.txtTasaDeInteres.SelectAll();
		}

		private void txtTasaDeInteres_Enter(object sender, System.EventArgs e)
		{
			this.txtTasaDeInteres.SelectAll();
		}

		private void txtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCuotas_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtTasaDeInteres_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtValor_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtConcepto_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdPrestamo.Value);

			if (this.iCuotasPagadas() == 0)
			{
				if (DialogResult.No == MessageBox.Show("¿Esta seguro de anular este registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;

				string sSQL = string.Format("Update Prestamos Set Estado = 2 Where idPrestamo = {0} Update Asiento Set Borrar = 1 Where idAsiento = {1}", (int)this.txtIdPrestamo.Value, (int)this.txtIdAsiento.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				this.lblEstado.Text = "ANULADO";
				this.txtEstado.Value = 2;

				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnAnular.Enabled = false;
			}
			else
			{
				MessageBox.Show("No puede Anular tiene Cuotas ya pagadas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				return;
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			#region Campos Vacios

			#region Linea credito
			if (this.cmbLineaCredito.ActiveRow == null)
			{
					this.cmbLineaCredito.Value = 0;
//				MessageBox.Show("Seleccione una línea de crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.cmbLineaCredito.Focus();
//				return;
			}
			#endregion Linea credito

			if (this.txtNumero.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el número del prestamo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				return;
			}

			if (this.cmbBanco.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Banco", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBanco.Focus();
				return;
			}

			if (this.cmbCuenta.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCuenta.Focus();
				return;
			}

			if (this.cmbCuentaPrestamo.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la Cuenta Contable del Prestamo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCuentaPrestamo.Focus();
				return;
			}

			if (this.txtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtFecha.Focus();
				return;
			}

			if (this.txtFechaPago.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtFechaPago.Focus();
				return;
			}

			if (this.cmbTipoPrestamo.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el tipo de prestamo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipoPrestamo.Focus();
				return;
			}

			if (this.cmbFormaPago.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la Forma de pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbFormaPago.Focus();
				return;
			}

			if (Convert.ToInt32(this.txtCuotas.Value) == 0)
			{
				MessageBox.Show("Ingrese el numero de cuotas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCuotas.Focus();
				return;
			}	

			if (Convert.ToDecimal(this.txtTasaDeInteres.Value) == 0)
			{
				MessageBox.Show("Ingrese la Tasa de Interes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTasaDeInteres.Focus();
				return;
			}

			if (Convert.ToDecimal(this.txtTasaNominal.Value) == 0)
			{
				MessageBox.Show("Ingrese la Tasa Nominal", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTasaNominal.Focus();
				return;
			}

//			if (Convert.ToDecimal(this.txtCapital.Value) == 0)
//			{
//				MessageBox.Show("Ingrese el Monto del Credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.txtCuotas.Focus();
//				return;
//			}

			if (this.txtConcepto.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el concepto del prestamo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtConcepto.Focus();
				return;
			}
			#endregion Campos Vacios

			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No ha calculado la Tabla de Amortización", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.btnCalcular_Click(sender, e);				
			}

			if ((int)this.txtCuotas.Value != this.ultraGrid1.Rows.Count)
			{
				MessageBox.Show("El numero de cuotas fue modificado vuelva a Calcular la Tabla de Amortización", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.btnCalcular_Click(sender, e);				
			}			



			#endregion Validacion
			
			#region Guarda Registro
			/*
			 @idPrestamos Int, @ProveedorBanco, @idBanco Int, @Numero Varchar(20), @idCuenta Int, @idCuentaPrestamo, @Fecha Datetime, @FechaPago Datetime, @NumeroCuotas Int, @TasaInteres Decimal(18, 2), @TasaNominal Decimal(18, 2), 
			 @idTipoPrestamo Int, @Valor Decimal(18, 2), @Capital Decimal(18, 2), @Interes Decimal(18, 2),  @Observaciones Varchar(600), @Estado Int
			
			if (bNuevo) this.txtNumero.Value = Funcion.sEscalarSQL(cdsSeteoF, "Select dbo.NumeracionPrestamos()"); 
			*/
			DateTime dtFecha = (DateTime)this.txtFecha.Value;
			DateTime dtFechaPagos = (DateTime)this.txtFechaPago.Value;
			string squery = "";
			try
			{
				  squery = string.Format("Exec PrestamosGuardaEncabezado {0}, {1}, {2}, '{3}', {4}, {5}, '{6}', '{7}', {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, '{16}', {17}, {18}", 
					(int)this.txtIdPrestamo.Value, (int)this.optBancoProveedor.Value, (int)this.cmbBanco.Value, this.txtNumero.Text.ToString(), (int)this.cmbCuenta.Value, (int)this.cmbCuentaPrestamo.Value, 
					dtFecha.ToString("yyyyMMdd"), dtFechaPagos.ToString("yyyyMMdd"), 
					(int)this.txtCuotas.Value, Convert.ToDecimal(this.txtTasaDeInteres.Value), Convert.ToDecimal(this.txtTasaNominal.Value), 
					(int)this.cmbTipoPrestamo.Value, (int)this.cmbFormaPago.Value, Convert.ToDecimal(this.txtValor.Value), Convert.ToDecimal(this.txtCapital.Value), Convert.ToDecimal(this.txtInteres.Value), 
					this.txtConcepto.Text.ToString(), (int)this.txtEstado.Value, (int)this.cmbLineaCredito.Value);
				this.txtIdPrestamo.Value = Funcion.iEscalarSQL(cdsSeteoF, squery);
			}
			catch(Exception ex)
			{
				Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
				MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{	
				DateTime dtInicia = DateTime.Today;
				if (dr.Cells["Inicia"].Value != System.DBNull.Value) dtInicia = (DateTime)dr.Cells["Inicia"].Value;
				DateTime dtFin = (DateTime)dr.Cells["Vence"].Value;
				DateTime dtFechaPago = DateTime.Today;
				if (dr.Cells["FechaPago"].Value != System.DBNull.Value) dtFechaPago = (DateTime)dr.Cells["FechaPago"].Value;

				string sSQLDetalle = string.Format("Exec PrestamosGuardaDetalle {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, '{10}', {11}",
					(int)dr.Cells["idDetallePrestamo"].Value, (int)this.txtIdPrestamo.Value, (int)dr.Cells["NumeroCuota"].Value, dtInicia.ToString("yyyyMMdd"), dtFin.ToString("yyyyMMdd"), Convert.ToDecimal(dr.Cells["ValorCuota"].Value), 
					Convert.ToDecimal(dr.Cells["Capital"].Value), Convert.ToDecimal(dr.Cells["Interes"].Value), Convert.ToDecimal(dr.Cells["Pago"].Value), Convert.ToDecimal(dr.Cells["Saldo"].Value), dtFechaPago.ToString("yyyyMMdd"), (int)dr.Cells["Estado"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle, true);
			}
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PrestamosConsultaDetalle {0}", (int)this.txtIdPrestamo.Value));
			#endregion Guarda Registro

			#region Asiento
			if (bNuevo)
			{
				string sSQLAsiento = string.Format("Exec PrestamosCreaAsiento {0}", Convert.ToInt32(this.txtIdPrestamo.Value));
				this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento);
			}
			#endregion Asiento

			#region Autorizacion de Pago
			if ((int)this.cmbFormaPago.Value == 2)
			{	
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					if (Convert.ToInt32(dr.Cells["Estado"].Value) == 0)
					{
						DateTime dtFechaAut = Convert.ToDateTime(dr.Cells["Vence"].Value);
						string sNumPago = dr.Cells["NumeroCuota"].Value + " DE " + this.txtCuotas.Value.ToString();
						string sNotas = string.Format("PAGO PRESTAMO N. '{1}' CUOTA N. {0}", (int)dr.Cells["NumeroCuota"].Value, this.txtNumero.Text.ToString());

						string sSQLAutMod = string.Format("Exec CreaAutorizacionesDePago 17, {0}, '{1}', '{2}', {3}, {4}, '{5}', '{6}', {7}", 
							Convert.ToInt32(dr.Cells["idDetallePrestamo"].Value), dtFechaAut.ToString("yyyyMMdd"), dtFechaAut.ToString("yyyyMMdd"), Convert.ToDecimal(dr.Cells["ValorCuota"].Value),
							2, sNumPago, sNotas, Convert.ToInt32(this.cmbCuenta.Value));
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAutMod, true);
					}
				}
			}			
			#endregion Autorizacion de Pago

			#region Controles
			this.optBancoProveedor.Enabled = false;
			this.cmbBanco.Enabled = false;
			this.txtNumero.Enabled = false;
			this.cmbCuenta.Enabled = false;
			this.cmbCuentaPrestamo.Enabled = false;
			this.txtFecha.Enabled = false;
			this.txtFechaPago.Enabled = false;
			this.cmbTipoPrestamo.Enabled = false;
			this.cmbFormaPago.Enabled = false;
			this.txtCuotas.Enabled = false;
			this.txtTasaDeInteres.Enabled = false;
			this.txtTasaNominal.Enabled = false;
			this.txtValor.Enabled = false;
			this.txtCapital.Enabled = false;
			this.txtInteres.Enabled = false;
			this.txtConcepto.Enabled = false;

			#region Linea credito
			//Línea de crédito
			ActivarLineaCredito(false, false);
			#endregion Linea credito

			ModificaGrilla(false);

			this.btnCalcular.Enabled = false;
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
			this.btnGuardar.Enabled = false;
			if (miAcceso.Anular) this.btnAnular.Enabled = true;
			this.btnCancelar.Enabled = true;

			bNuevo = false;
			bEdicion = false;
			#endregion Controles
		}

		private void Consulta (int idPrestamo)
		{
			try
			{
				#region Consulta y Asignacion de Valores
				string sSQL = string.Format("Exec PrestamosConsultaIndividual {0}", idPrestamo);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdPrestamo.Value = dr.GetInt32(0);
					this.optBancoProveedor.Value = dr.GetInt32(1);
					this.txtIdAsiento.Value = dr.GetInt32(2);
					this.cmbBanco.Value = dr.GetInt32(3);
					this.txtNumero.Text = dr.GetString(4);
					this.cmbCuenta.Value = dr.GetInt32(5);
					this.cmbCuentaPrestamo.Value = dr.GetInt32(6);
					this.txtFecha.Value = dr.GetDateTime(7);
					this.txtFechaPago.Value = dr.GetDateTime(8);
					this.txtCuotas.Value = dr.GetInt32(9);
					this.txtTasaDeInteres.Value = dr.GetDecimal(10);
					this.txtTasaNominal.Value = dr.GetDecimal(11);
					this.cmbTipoPrestamo.Value = dr.GetInt32(12);
					this.txtValor.Value = dr.GetDecimal(13);
					this.txtCapital.Value = dr.GetDecimal(14);
					this.txtInteres.Value = dr.GetDecimal(15);
					this.txtConcepto.Text = dr.GetString(16);
					this.txtEstado.Value = dr.GetInt32(17);
					this.lblEstado.Text = dr.GetString(18);
					this.cmbFormaPago.Value = dr.GetInt32(19);
					this.lblFechaSistema.Text = dr.GetValue(20).ToString();
					this.lblUsuario.Text = dr.GetString(21);
					this.cmbLineaCredito.Value = dr.GetInt32(22);
				}
				dr.Close();
				#endregion Consulta y Asignacion de Valores

				this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PrestamosListaProveedores {0}", Convert.ToInt32(this.optBancoProveedor.Value)));

				this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select c.idCuenta, c.Descripcion From Cuenta c Inner Join CuentaCorriente cc On c.idCuenta = cc.idCuenta And cc.idBanco = {0}", (int)this.cmbBanco.Value)); //And cc.Tipo = 0

				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PrestamosConsultaDetalle {0}", (int)this.txtIdPrestamo.Value));

				if ((int)this.txtEstado.Value == 0)
				{
					if (miAcceso.Editar ) this.btnEditar.Enabled = true;
					//&& iCuotasPagadas() == 0
					if (miAcceso.Anular && iCuotasPagadas() == 0) this.btnAnular.Enabled = true;
				}

				if ((int)this.txtEstado.Value < 2) if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;

				this.btnCancelar.Enabled = true;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}	
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscarPrestamos Busqueda = new frmBuscarPrestamos())
			{
				if (Busqueda.ShowDialog() == DialogResult.OK)
				{												
					this.Consulta((int)Busqueda.ultraGrid1.ActiveRow.Cells["idPrestamo"].Value);
				}
			}
		}

		private void txtTasaDeInteres_Validated(object sender, System.EventArgs e)
		{
			if (this.txtTasaDeInteres.Value == System.DBNull.Value) this.txtTasaDeInteres.Value = 0;			
		}

		private void TablaDeAmortizacion ()
		{
			if (bEdicion)
			{
				if (this.cmbTipoPrestamo.ActiveRow != null)	
				{								
					if (Convert.ToDecimal(this.txtTasaNominal.Value) > 0)
					{
						if (Convert.ToDecimal(this.txtValor.Value) > 0)
						{
							DateTime dtFecha = (DateTime)this.txtFechaPago.Value;
			
							string sSQL = string.Format("Exec PrestamosTablaDeAmortizacion {0}, {1}, {2}, {3}, '{4}', {5}",
								Convert.ToDecimal(this.txtValor.Value), Convert.ToDecimal(this.txtInteres.Value), (int)this.txtCuotas.Value, 
								Convert.ToDecimal(this.txtTasaNominal.Value), dtFecha.ToString("yyyyMMdd"), (int)this.cmbTipoPrestamo.Value);
							this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);					
						}
					}
				}
			}
		}

		private void txtValor_Validated(object sender, System.EventArgs e)
		{
			if (this.txtValor.Value == System.DBNull.Value) this.txtValor.Value = 0.00m;
		}

		private void cmbBanco_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtTasaNominal_Validated(object sender, System.EventArgs e)
		{
			if (this.txtTasaNominal.Value == System.DBNull.Value) this.txtTasaNominal.Value = 0;			
		}

		private void cmbBanco_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				int idBanco = 0;
				if (this.cmbBanco.ActiveRow != null) idBanco = Convert.ToInt32(this.cmbBanco.Value);
				if ((int)this.optBancoProveedor.Value == 1)	this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select c.idCuenta, c.Descripcion From Cuenta c Inner Join CuentaCorriente cc On c.idCuenta = cc.idCuenta And cc.idBanco = {0}", idBanco));
				if ((int)this.optBancoProveedor.Value == 2)	this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select c.idCuenta, c.Descripcion From Cuenta c Inner Join CuentaCorriente cc On c.idCuenta = cc.idCuenta");
			}
		}

		private void cmbCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCuentaPrestamo.Focus();
		}

		private void txtFechaPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTipoPrestamo.Focus(); 
		}

		private void txtTasaNominal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtValor.Focus();
		}

		private void cmbTipoPrestamo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbFormaPago.Focus();
		}

		private void txtInteres_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtConcepto.Focus();
		}

		private void txtTasaNominal_Enter(object sender, System.EventArgs e)
		{
			this.txtTasaNominal.SelectAll();
		}

		private void txtTasaNominal_Click(object sender, System.EventArgs e)
		{
			this.txtTasaNominal.SelectAll();
		}

		private void txtInteres_Click(object sender, System.EventArgs e)
		{
			this.txtInteres.SelectAll();
		}

		private void txtInteres_Enter(object sender, System.EventArgs e)
		{
			this.txtInteres.SelectAll();
		}

		private void cmbTipoPrestamo_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbTipoPrestamo.ActiveRow != null)
				{
					if ((int)this.cmbTipoPrestamo.Value == 1)
					{
						this.txtInteres.Value = 0;
						this.txtInteres.Enabled = false;
					}
					else if ((int)this.cmbTipoPrestamo.Value == 2)
					{
						this.txtInteres.Enabled = true;
					}
				}
			}
		}

		private void txtInteres_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void cmbCuentaPrestamo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtFecha.Focus();
		}

		private void optBancoProveedor_ValueChanged(object sender, System.EventArgs e)
		{
			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PrestamosListaProveedores {0}", Convert.ToInt32(this.optBancoProveedor.Value)));
		}

		private void btnCalcular_Click(object sender, System.EventArgs e)
		{
			TablaDeAmortizacion();
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnPago")
			{	
				if (!bEdicion)
				{
					if ((int)this.cmbFormaPago.Value == 2)
					{
						MessageBox.Show("Solo para prestamos con Debito Bancario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					if ((int)e.Cell.Row.Cells["Estado"].Value == 1)
					{
						MessageBox.Show("Ya esta autorizado el Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					if ((int)e.Cell.Row.Cells["Estado"].Value == 2)
					{
						MessageBox.Show("Ya esta registrado el Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					int iCuota = Convert.ToInt32(e.Cell.Row.Cells["NumeroCuota"].Value);

					if (iCuota > this.iUltimaCuotasPagada())
					{
						MessageBox.Show(string.Format("La cuota n. {0} esta pendiente de pago", this.iUltimaCuotasPagada()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					using (frmNumeroNotaDeDebito NNDD = new frmNumeroNotaDeDebito(Convert.ToDateTime(Convert.ToDateTime(e.Cell.Row.Cells["Vence"].Value))))
					{
						if (DialogResult.OK == NNDD.ShowDialog())
						{
							DateTime dtFecha = (DateTime)NNDD.txtFecha.Value;

							string sSQL = string.Format("Exec PrestamosCreaAsientoDePago {0}, '{1}', '{2}', 1", (int)e.Cell.Row.Cells["idDetallePrestamo"].Value, NNDD.txtDebitoInteres.Text.ToString(), dtFecha.ToString("yyyyMMdd"));
							int idAsiento = Funcion.iEscalarSQL(cdsSeteoF, sSQL, true);

							sSQL = string.Format("Exec PrestamosCreaAsientoDePago {0}, '{1}', '{2}', 0", (int)e.Cell.Row.Cells["idDetallePrestamo"].Value, NNDD.txtDebitoCapital.Text.ToString(), dtFecha.ToString("yyyyMMdd"));
							int idAsiento1 = Funcion.iEscalarSQL(cdsSeteoF, sSQL, true);

							Reporte ReporteInteres = new Reporte("Asientos.rpt", "{Asiento.idAsiento} = " + idAsiento.ToString());
							ReporteInteres.MdiParent = this.MdiParent;									
							ReporteInteres.Show();	

							Reporte ReporteCapital = new Reporte("Asientos.rpt", "{Asiento.idAsiento} = " + idAsiento1.ToString());
							ReporteCapital.MdiParent = this.MdiParent;							
							ReporteCapital.Show();	

							this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PrestamosConsultaDetalle {0}", (int)this.txtIdPrestamo.Value));
						}
					}
				}
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{			
			FuncionesProcedimientos.FormatoGrid(e, "Interes", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Capital", 2);
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ParameterFields paramFields = new ParameterFields ();

			ParameterField Compra = new ParameterField ();
			ParameterDiscreteValue IdCompra = new ParameterDiscreteValue ();
			Compra.ParameterFieldName = "@idPrestamo";
			IdCompra.Value = int.Parse(this.txtIdPrestamo.Value.ToString());
			Compra.CurrentValues.Add (IdCompra);
			paramFields.Add (Compra);
			
			Reporte miRep = new Reporte("Prestamos.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Prestamos");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
						
			Cursor = Cursors.Default;
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Capital")
			{
				e.Cell.Row.Cells["ValorCuota"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Capital"].Value) + Convert.ToDecimal(e.Cell.Row.Cells["Interes"].Value);
			}

			if (e.Cell.Column.ToString() == "Interes")
			{
				e.Cell.Row.Cells["ValorCuota"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Capital"].Value) + Convert.ToDecimal(e.Cell.Row.Cells["Interes"].Value);
			}
		}

		private void cmbFormaPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCuotas.Focus();
		}

		private void frmPrestamos_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un Prestamo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}
	}
}

