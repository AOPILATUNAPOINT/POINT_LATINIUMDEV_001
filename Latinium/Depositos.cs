using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Depositos.
	/// </summary>
	public class Depositos : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private C1.Data.C1DataSet cdsPagos;
		private C1.Data.C1DataSet cdsSeteoF;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCliente;
		private Infragistics.Win.Misc.UltraButton btPago;
		private C1.Data.C1DataSet cdsFormaPago;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdPagos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaCorr;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBanco;
		private C1.Data.C1DataView cdvPersonal;
		private C1.Data.C1DataSet cdsPersonal;
		private Infragistics.Win.Misc.UltraButton btEgreso;
		private C1.Data.C1DataSet cdsAsiento;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorr1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkIncluyenombre;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDesglosa;
		private C1.Data.C1DataSet cdsAsignaCuenta;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		private C1.Data.C1DataView cdvCliente;
		private Infragistics.Win.Misc.UltraButton btReporte;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Depositos()
		{
			InitializeComponent();
		}

		int IdTipo = 1;
		public Depositos(int idTipo)
		{
			InitializeComponent();
			IdTipo = idTipo;
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Pago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco", -1, "cmbBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobrador", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente", -1, "cmbCuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pago");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pendiente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PendLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PendLote1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AsiEgr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDeposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Pago", 10, true, "Pago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Pago", 10, true);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraFormaPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraFormaPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.grdPagos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsPagos = new C1.Data.C1DataSet();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCliente = new C1.Data.C1DataView();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.chkCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btPago = new Infragistics.Win.Misc.UltraButton();
			this.cdsFormaPago = new C1.Data.C1DataSet();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.cmbCuentaCorr = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.btEgreso = new Infragistics.Win.Misc.UltraButton();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbCuentaCorr1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbFormaPago1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkIncluyenombre = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDesglosa = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsAsignaCuenta = new C1.Data.C1DataSet();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.btReporte = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsignaCuenta)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbDesde
			// 
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(80, 16);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 24;
			this.cmbDesde.Size = new System.Drawing.Size(112, 21);
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 0;
			this.cmbDesde.Value = new System.DateTime(2007, 2, 27, 0, 0, 0, 0);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(24, 16);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 14);
			this.ultraLabel1.TabIndex = 1;
			this.ultraLabel1.Text = "Desde";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(24, 48);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(34, 14);
			this.ultraLabel2.TabIndex = 3;
			this.ultraLabel2.Text = "Hasta";
			// 
			// cmbHasta
			// 
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(80, 48);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 24;
			this.cmbHasta.Size = new System.Drawing.Size(112, 21);
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 2;
			this.cmbHasta.Value = new System.DateTime(2007, 2, 27, 0, 0, 0, 0);
			// 
			// grdPagos
			// 
			this.grdPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdPagos.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdPagos.DataMember = "Pago";
			this.grdPagos.DataSource = this.cdsPagos;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdPagos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 15;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 23;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 14;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 36;
			ultraGridColumn4.Header.VisiblePosition = 13;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.Caption = "Forma Pago";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn5.Width = 80;
			ultraGridColumn6.Header.Caption = "Banco";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.Header.Caption = "Cobrador";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 141;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.Caption = "CuentaCorriente";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn8.Width = 99;
			ultraGridColumn9.Header.VisiblePosition = 24;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Format = "dd/MMM/yyyy";
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Width = 79;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance2;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Width = 71;
			ultraGridColumn12.Header.VisiblePosition = 16;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 10;
			ultraGridColumn13.Width = 58;
			ultraGridColumn14.Header.VisiblePosition = 19;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.Width = 73;
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 9;
			ultraGridColumn18.Width = 71;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 20;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn21.Header.VisiblePosition = 4;
			ultraGridColumn21.Width = 36;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 12;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 11;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 22;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn34.Header.VisiblePosition = 33;
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
																										 ultraGridColumn34});
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.grdPagos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdPagos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdPagos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdPagos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.grdPagos.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdPagos.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.grdPagos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdPagos.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdPagos.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.grdPagos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdPagos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdPagos.Location = new System.Drawing.Point(8, 136);
			this.grdPagos.Name = "grdPagos";
			this.grdPagos.Size = new System.Drawing.Size(592, 208);
			this.grdPagos.TabIndex = 4;
			// 
			// cdsPagos
			// 
			this.cdsPagos.BindingContextCtrl = this;
			this.cdsPagos.DataLibrary = "LibFacturacion";
			this.cdsPagos.DataLibraryUrl = "";
			this.cdsPagos.DataSetDef = "dsPago";
			this.cdsPagos.FillOnRequest = false;
			this.cdsPagos.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPagos.Name = "cdsPagos";
			this.cdsPagos.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPagos.SchemaDef = null;
			this.cdsPagos.BeforeFill += new C1.Data.FillEventHandler(this.cdsPagos_BeforeFill);
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
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			this.cdsSeguridad.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeguridad_BeforeFill);
			// 
			// btFiltro
			// 
			this.btFiltro.Location = new System.Drawing.Point(224, 104);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(56, 24);
			this.btFiltro.TabIndex = 5;
			this.btFiltro.Text = "Filtro";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// optTipo
			// 
			this.optTipo.CheckedIndex = 1;
			this.optTipo.ItemAppearance = appearance7;
			this.optTipo.ItemOrigin = new System.Drawing.Point(6, 2);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendientes";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Confirmados";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.Items.Add(valueListItem3);
			this.optTipo.ItemSpacingHorizontal = 2;
			this.optTipo.ItemSpacingVertical = 2;
			this.optTipo.Location = new System.Drawing.Point(208, 16);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(104, 56);
			this.optTipo.TabIndex = 6;
			this.optTipo.Text = "Pendientes";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.FillOnRequest = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataSource = this.cdvCliente;
			ultraGridColumn35.Header.VisiblePosition = 1;
			ultraGridColumn36.Header.VisiblePosition = 3;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 0;
			ultraGridColumn37.Width = 201;
			ultraGridColumn38.Header.VisiblePosition = 4;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 5;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 2;
			ultraGridColumn40.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.Enabled = false;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(128, 80);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(184, 21);
			this.cmbCliente.TabIndex = 7;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// chkCliente
			// 
			this.chkCliente.Checked = true;
			this.chkCliente.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCliente.Location = new System.Drawing.Point(24, 80);
			this.chkCliente.Name = "chkCliente";
			this.chkCliente.Size = new System.Drawing.Size(96, 20);
			this.chkCliente.TabIndex = 9;
			this.chkCliente.Text = "Todo Cliente";
			this.chkCliente.CheckedChanged += new System.EventHandler(this.chkCliente_CheckedChanged);
			// 
			// btPago
			// 
			this.btPago.Location = new System.Drawing.Point(416, 104);
			this.btPago.Name = "btPago";
			this.btPago.Size = new System.Drawing.Size(56, 23);
			this.btPago.TabIndex = 10;
			this.btPago.Text = "Pago";
			this.btPago.Click += new System.EventHandler(this.btPago_Click);
			// 
			// cdsFormaPago
			// 
			this.cdsFormaPago.BindingContextCtrl = this;
			this.cdsFormaPago.DataLibrary = "LibFacturacion";
			this.cdsFormaPago.DataLibraryUrl = "";
			this.cdsFormaPago.DataSetDef = "dsCompraTabla";
			this.cdsFormaPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsFormaPago.Name = "cdsFormaPago";
			this.cdsFormaPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFormaPago.SchemaDef = null;
			this.cdsFormaPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsFormaPago_BeforeFill);
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataMember = "CompraFormaPago";
			this.cmbFormaPago.DataSource = this.cdsFormaPago;
			ultraGridColumn41.Header.VisiblePosition = 3;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 4;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Width = 187;
			ultraGridColumn44.Header.VisiblePosition = 2;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 1;
			ultraGridColumn45.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(304, 224);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(208, 72);
			this.cmbFormaPago.TabIndex = 11;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Visible = false;
			// 
			// cdsCuentaCorr
			// 
			this.cdsCuentaCorr.BindingContextCtrl = this;
			this.cdsCuentaCorr.DataLibrary = "LibContabilidad";
			this.cdsCuentaCorr.DataLibraryUrl = "";
			this.cdsCuentaCorr.DataSetDef = "dsCuentaCorriente";
			this.cdsCuentaCorr.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentaCorr.Name = "cdsCuentaCorr";
			this.cdsCuentaCorr.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaCorr.SchemaDef = null;
			this.cdsCuentaCorr.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuentaCorr_BeforeFill);
			// 
			// cmbCuentaCorr
			// 
			this.cmbCuentaCorr.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorr.DataMember = "CuentaCorriente";
			this.cmbCuentaCorr.DataSource = this.cdsCuentaCorr;
			ultraGridColumn46.Header.VisiblePosition = 6;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 9;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 4;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 5;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 0;
			ultraGridColumn50.Width = 155;
			ultraGridColumn51.Header.VisiblePosition = 7;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 1;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 8;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 2;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 3;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 10;
			ultraGridColumn57.Header.VisiblePosition = 11;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn57});
			this.cmbCuentaCorr.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCuentaCorr.DisplayMember = "NumCuenta";
			this.cmbCuentaCorr.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorr.Location = new System.Drawing.Point(176, 208);
			this.cmbCuentaCorr.Name = "cmbCuentaCorr";
			this.cmbCuentaCorr.Size = new System.Drawing.Size(176, 72);
			this.cmbCuentaCorr.TabIndex = 12;
			this.cmbCuentaCorr.ValueMember = "idCuentaCorriente";
			this.cmbCuentaCorr.Visible = false;
			// 
			// cmbBanco
			// 
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataMember = "Banco";
			this.cmbBanco.DataSource = this.cdsCuentaCorr;
			ultraGridColumn58.Header.VisiblePosition = 3;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 0;
			ultraGridColumn59.Width = 186;
			ultraGridColumn60.Header.VisiblePosition = 1;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 4;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 2;
			ultraGridColumn62.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(16, 224);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(208, 80);
			this.cmbBanco.TabIndex = 13;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.Visible = false;
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.RowFilter = "Cobrador = 1";
			this.cdvPersonal.Sort = "Nombre";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// btEgreso
			// 
			this.btEgreso.Location = new System.Drawing.Point(536, 104);
			this.btEgreso.Name = "btEgreso";
			this.btEgreso.Size = new System.Drawing.Size(56, 24);
			this.btEgreso.TabIndex = 14;
			this.btEgreso.Text = "Ingreso";
			this.btEgreso.Click += new System.EventHandler(this.btEgreso_Click);
			// 
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsAsiento";
			this.cdsAsiento.FillOnRequest = false;
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(328, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(69, 16);
			this.label4.TabIndex = 20;
			this.label4.Text = "Cuenta Corr.";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(328, 48);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(82, 16);
			this.label3.TabIndex = 18;
			this.label3.Text = "Forma de Pago";
			// 
			// cmbCuentaCorr1
			// 
			this.cmbCuentaCorr1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorr1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorr1.DataMember = "CuentaCorriente";
			this.cmbCuentaCorr1.DataSource = this.cdsCuentaCorr;
			this.cmbCuentaCorr1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand6.AddButtonCaption = "CompraComprobante";
			ultraGridColumn63.Header.VisiblePosition = 7;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 9;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 155;
			ultraGridColumn65.Header.VisiblePosition = 6;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 5;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 0;
			ultraGridColumn67.Width = 41;
			ultraGridColumn68.Header.VisiblePosition = 1;
			ultraGridColumn68.Width = 34;
			ultraGridColumn69.Header.VisiblePosition = 2;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 8;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 3;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 4;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 10;
			ultraGridColumn73.Width = 26;
			ultraGridColumn74.Header.VisiblePosition = 11;
			ultraGridColumn74.Width = 57;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn74});
			this.cmbCuentaCorr1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCuentaCorr1.DisplayMember = "NumCuenta";
			this.cmbCuentaCorr1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCorr1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCorr1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorr1.Location = new System.Drawing.Point(416, 16);
			this.cmbCuentaCorr1.Name = "cmbCuentaCorr1";
			this.cmbCuentaCorr1.Size = new System.Drawing.Size(176, 21);
			this.cmbCuentaCorr1.TabIndex = 19;
			this.cmbCuentaCorr1.ValueMember = "idCuentaCorriente";
			// 
			// cmbFormaPago1
			// 
			this.cmbFormaPago1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago1.DataMember = "CompraFormaPago";
			this.cmbFormaPago1.DataSource = this.cdsFormaPago;
			this.cmbFormaPago1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand7.AddButtonCaption = "CompraComprobante";
			ultraGridColumn75.Header.VisiblePosition = 3;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 4;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 0;
			ultraGridColumn77.Width = 158;
			ultraGridColumn78.Header.VisiblePosition = 2;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 1;
			ultraGridColumn79.Hidden = true;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79});
			this.cmbFormaPago1.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbFormaPago1.DisplayMember = "FormaPago";
			this.cmbFormaPago1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago1.Location = new System.Drawing.Point(416, 48);
			this.cmbFormaPago1.Name = "cmbFormaPago1";
			this.cmbFormaPago1.Size = new System.Drawing.Size(176, 21);
			this.cmbFormaPago1.TabIndex = 17;
			this.cmbFormaPago1.ValueMember = "idFormaPago";
			// 
			// chkIncluyenombre
			// 
			this.chkIncluyenombre.Location = new System.Drawing.Point(456, 72);
			this.chkIncluyenombre.Name = "chkIncluyenombre";
			this.chkIncluyenombre.Size = new System.Drawing.Size(104, 20);
			this.chkIncluyenombre.TabIndex = 28;
			this.chkIncluyenombre.Text = "Incluir Nombre";
			// 
			// chkDesglosa
			// 
			this.chkDesglosa.Location = new System.Drawing.Point(328, 72);
			this.chkDesglosa.Name = "chkDesglosa";
			this.chkDesglosa.Size = new System.Drawing.Size(104, 20);
			this.chkDesglosa.TabIndex = 27;
			this.chkDesglosa.Text = "Desglosar Pago";
			// 
			// cdsAsignaCuenta
			// 
			this.cdsAsignaCuenta.BindingContextCtrl = this;
			this.cdsAsignaCuenta.DataLibrary = "LibContabilidad";
			this.cdsAsignaCuenta.DataLibraryUrl = "";
			this.cdsAsignaCuenta.DataSetDef = "dsAsignaCuenta";
			this.cdsAsignaCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsignaCuenta.Name = "cdsAsignaCuenta";
			this.cdsAsignaCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsignaCuenta.SchemaDef = null;
			// 
			// btGrabar
			// 
			this.btGrabar.Location = new System.Drawing.Point(24, 104);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(64, 23);
			this.btGrabar.TabIndex = 29;
			this.btGrabar.Text = "Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// btReporte
			// 
			this.btReporte.Location = new System.Drawing.Point(328, 104);
			this.btReporte.Name = "btReporte";
			this.btReporte.Size = new System.Drawing.Size(56, 24);
			this.btReporte.TabIndex = 30;
			this.btReporte.Text = "Reporte";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// Depositos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(624, 358);
			this.Controls.Add(this.btReporte);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.chkIncluyenombre);
			this.Controls.Add(this.chkDesglosa);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbCuentaCorr1);
			this.Controls.Add(this.cmbFormaPago1);
			this.Controls.Add(this.btEgreso);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.cmbCuentaCorr);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.btPago);
			this.Controls.Add(this.chkCliente);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.grdPagos);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbDesde);
			this.Name = "Depositos";
			this.Text = "Depositos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Depositos_Closing);
			this.Load += new System.EventHandler(this.Depositos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsignaCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsPagos_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPagos.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);				
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private string Filtro()
		{
			DateTime dtFechaIni = (DateTime) cmbDesde.Value;
			DateTime dtFechaFin = (DateTime) cmbHasta.Value;
			string stFiltro = "Pago.Fecha >= '" + dtFechaIni.ToString("yyyyMMdd")
				+ "' And Pago.Fecha < '" + dtFechaFin.AddDays(1).ToString("yyyyMMdd") + "'";
			if ((int) optTipo.Value == 1)
			{
				stFiltro += " And Pago.Pendiente = 1";
			}
			if ((int) optTipo.Value == 2)
			{
				stFiltro += " And Pago.Pendiente = 0";
			}
			if (IdTipo!= 1)	stFiltro += " And Pago.PendLote = 1 ";
			else stFiltro += " And Pago.PendLote1 = 1 ";
			if (chkCliente.Checked) FiltraTabla(0);
			else FiltraTabla((int) cmbCliente.Value);
			return stFiltro;
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			BorraErrores();
			C1.Data.FilterCondition fcPadre = new 
					C1.Data.FilterCondition(cdsPagos, "Pago", Filtro());
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			cdsPagos.Fill(fcTotal, false);
//			MessageBox.Show("Filtro Generado", "Informacion");
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			string stbProveedor = "";
			switch(IdTipo)
			{
				case 1:
				case 3:
				case 5:
				case 13:
				case 15:
					stbProveedor = "False";
					break;
				case 2:
				case 4:
				case 6:
				case 14:
				case 16:
					stbProveedor = "True";
					break;
			}
			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsCliente"];

			if (stbProveedor == "True") stbProveedor = " <> 0";
			else stbProveedor = " = 0";

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"],
				"[Ambos] <> 0 OR [Proveedor] " + stbProveedor));
		}

		private void chkCliente_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCliente.Enabled = !chkCliente.Checked;
			if (!chkCliente.Checked) cmbCliente.Select();
		}

		private void btPago_Click(object sender, System.EventArgs e)
		{
			BorraErrores();
			if (grdPagos.ActiveRow == null)
			{
				errorProvider.SetError(grdPagos, "Escoja un Pago para Editar");
				return;
			}
			int IdPago = (int) grdPagos.ActiveRow.Cells["idPago"].Value;
			using (Pago miPago = new Pago(IdPago, true, 0))
			{
				if (IdTipo == 1 || IdTipo == 12) miPago.bCobro = false;
				miPago.ShowDialog();
			}
		}

		private void cdsFormaPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsFormaPago.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCuentaCorr_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCuentaCorr.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void BorraErrores()
		{
			errorProvider.SetError(optTipo, "");
			errorProvider.SetError(cmbFormaPago1, "");
			errorProvider.SetError(cmbCuentaCorr1, "");
			errorProvider.SetError(grdPagos, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if ((int) optTipo.Value != 1)
			{
				errorProvider.SetError(optTipo, 
					"En esta pantalla solo se generan asientos de pagos pendientes");
				bOk =false;
			}
			if (cmbFormaPago1.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbFormaPago1, "Ingrese Forma de Pago");
				bOk = false;
			}
			if(cmbCuentaCorr1.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbCuentaCorr1, "Ingrese Cuenta corriente");
				bOk = false;
			}
			if(grdPagos.Rows.Count == 0)
			{
				errorProvider.SetError(grdPagos, "No existen Pagos para realizar el Asiento");
				bOk = false;
			}
			return bOk;
		}

		private void GrupoCxC(int IdAsiento, double dTotal)
		{
			int IdCuenta = 0;
			double dValor = 0;
			C1.Data.C1DataRow drAsigna;
			C1.Data.C1DataRow drCC = this.cdsCuentaCorr.TableViews[
				"CuentaCorriente"].Rows.Find((int) cmbCuentaCorr1.Value);
			if ((int) cmbCuentaCorr1.Value != 0 || drCC["idCuenta"] != null || (int) drCC["idCuenta"] == 0)
				IdCuenta = (int) drCC["idCuenta"];
			else
			{
				drAsigna = cdsAsignaCuenta.TableViews["AsignaCuenta"].Rows[0];
				IdCuenta = (int) drAsigna["idCuenta"]; // Caja o Bancos en Ventas
			}

			C1.Data.C1DataRow drDetAsi;
			int iCuenta1 = 0;
			if (IdTipo == 1) 
			{
				drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
				drDetAsi["Debe"] = dTotal;
				drDetAsi["idCuenta"] = IdCuenta;
				drDetAsi["idAsiento"] = IdAsiento;
			}
			else iCuenta1 = IdCuenta;

			#region Detalle en asiento por cada linea
			string strSelect = "";
			if (chkDesglosa.Checked) 
			{
				string stNombre = "";
				if (chkIncluyenombre.Checked)
					stNombre += ", MIN(ISNULL(Cliente.Nombre, '')), MIN(IsNull(Compra.Numero, ''))";
				strSelect = "SELECT SUM(Pago.Pago), " +
					"MIN(IsNull(ClienteGrupo.idCuenta, 0)) " + stNombre + 
					" FROM Pago " +
					"Inner Join Compra ON Pago.idCompra = Compra.idCompra " +
					"Inner Join Cliente ON Compra.idCliente = Cliente.idCliente " +
					"Full Join ClienteGrupo ON ClienteGrupo.idGrupoCliente = " +
					"Cliente.idGrupoCliente Where " + Filtro() +
					" Group by Pago.idPago ";
			}
			else
			{
				strSelect = "SELECT SUM(Pago.Pago), " +
					"MIN(ClienteGrupo.idCuenta) FROM Pago " +
					"Inner Join Compra ON Pago.idCompra = Compra.idCompra " +
					"Inner Join Cliente ON Compra.idCliente = Cliente.idCliente " +
					"Full Join ClienteGrupo ON ClienteGrupo.idGrupoCliente = " +
					"Cliente.idGrupoCliente Where " + Filtro() +
					" Group by idCuenta Order by idCuenta";
			}
			SqlDataReader miReader = 
				Funcion.rEscalarSQL(cdsCliente, strSelect, true);
			try
			{
				while (miReader.Read())
				{
					dValor = miReader.GetDouble(0);
					IdCuenta = 0;
					if (miReader.GetValue(1) != System.DBNull.Value)
						IdCuenta =  miReader.GetInt32(1);
					if (IdCuenta == 0) 
					{
						int iTipo = 3;
						if (IdTipo != 1) iTipo = 12;
						drAsigna = cdsAsignaCuenta.TableViews["AsignaCuenta"].Rows[iTipo];
						IdCuenta = (int) drAsigna["idCuenta"]; // Cuentas por cobrar
					}

					drDetAsi = this.cdsAsiento.TableViews["AsientoDetalle"].AddNew();
					if (IdTipo == 1) drDetAsi["Haber"] = dValor;
					else drDetAsi["Debe"] = dValor;
					drDetAsi["idCuenta"] = IdCuenta;
					drDetAsi["idAsiento"] = IdAsiento;
					if (chkDesglosa.Checked)
					{
						if (chkIncluyenombre.Checked)
						{
							string stDetalle = miReader.GetString(2).Trim()
								+ " " + miReader.GetString(3).Trim();
							drDetAsi["Detalle"] = stDetalle;
						}
					}
				}
			}
			catch(System.Exception ex) 
			{
				MessageBox.Show(ex.Message, "Detalle de asiento");
			}
			miReader.Close();
		
			#endregion // Fin de Detalle de asiento por pagos no detallados


			if (IdTipo != 1) 
			{
				drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
				drDetAsi["Haber"] = dTotal;
				drDetAsi["idCuenta"] = iCuenta1;
				drDetAsi["idAsiento"] = IdAsiento;
			}
			this.cdsAsiento.Update();
		}

		private string stCodigo = "";
		private string stCheque = "";
		private void NumeroPago()
		{
			int iTipo = 0;
			stCodigo = "";
			stCheque = "";
			if (IdTipo != 1) iTipo = 1;
			string stNumero = Funcion.NumeraFactura(cdsCliente, 21 + iTipo);
			if (stNumero.Trim().Length > 0) stCodigo = stNumero;

			if (iTipo != 1) return;

			int iNumCheque = Contabilidad.NumeroCheque(
				cdsCuentaCorr, (int) cmbCuentaCorr1.Value, true);
			if (iNumCheque != -1)	stCheque = iNumCheque.ToString();
		}

		private void btEgreso_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			btFiltro_Click(this, e);
			// Crea asiento
			C1.Data.C1DataRow drAsiento = cdsAsiento.TableViews["Asiento"].AddNew();

			drAsiento["Fecha"] = DateTime.Today;
			if (IdTipo == 1)	drAsiento["idTipoAsiento"] = 2; // Ingreso
			else drAsiento["idTipoAsiento"] = 3; // Egreso
			drAsiento["Descripcion"] = this.cmbCliente.Text;

//			drAsiento["Numero"] = this.txtCodigo.Value;

			drAsiento["CodAsiento"] = 
				Contabilidad.NumeroAsiento(cdsAsiento, (int) drAsiento["idTipoAsiento"]);
//			string stFactura = "";
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
//				stFactura += dr.Cells["Concepto"].Value + " ";
//			drAsiento["Notas"] = stFactura;
			drAsiento["Descripcion"] = "Pago por Lotes Pendientes";

			// Actualiza el codigo del asiento en cada uno de los pagos

			double dTotal = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdPagos.Rows.All)
				dTotal += (double) dr.Cells["Pago"].Value;
			dTotal = Math.Round(dTotal, 2);
			drAsiento["Total"] = dTotal;
			drAsiento["TotalH"] = dTotal;
			NumeroPago();
			if (IdTipo!= 1)	drAsiento["Cheque"] = stCheque;
			drAsiento["Numero"] = stCodigo;
			cdsAsiento.Update();

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdPagos.Rows.All)
			{
				dr.Cells["NumeroIngreso"].Value = drAsiento["CodAsiento"];
				dr.Cells["Numero"].Value = stCodigo;
				dr.Cells["idAsiento"].Value = drAsiento["idAsiento"];
				dr.Cells["Pendiente"].Value = false;
				dr.Cells["idCuentaCorriente"].Value = (int) cmbCuentaCorr1.Value;
				if (IdTipo!= 1)	dr.Cells["Cheque"].Value = stCheque;
			}
			GrupoCxC((int) drAsiento["idAsiento"], dTotal);

			cdsPagos.Update();
			// Creacion de Detalle de Asiento
			MessageBox.Show("Asiento Generado", "Informacion");
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			BorraErrores();
			try
			{
				cdsPagos.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "No se puede grabar Pagos");
			}
			MessageBox.Show("Pagos Grabados", "Informacion");
		}

		private void FiltraTabla(int IdCliente)
		{
			string stPendLote = "PendLote1";
			if (IdTipo != 1) stPendLote = "PendLote";
			string stSelect = "Update Pago Set PendLote = 0 Where PendLote = 1";
			string stFiltroCliente = "";
			if (IdCliente > 0) 
				stFiltroCliente = " And Compra.IdCliente = " + IdCliente.ToString();
			Funcion.EjecutaSQL(cdsPagos, stSelect, true);
			stSelect = "Update Pago Set " + stPendLote + " = 1 From Pago Inner Join "
				+ "Compra On Compra.idCompra = Pago.IdCompra Where "
				+ " Compra.idTipoFactura = " + IdTipo.ToString()
				+ " And Pago.Pendiente = 1" + stFiltroCliente;
			Funcion.EjecutaSQL(cdsPagos, stSelect, true);
		}

		private void Depositos_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Depositos'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
			string stFiltro="";
			if (IdTipo != 1)
			{
				Text = "Pagos Pendientes";
				btEgreso.Text = "Egreso";
				stFiltro = " Proveedor <> 0 Or Ambos <> 0";
			}
			else
			{
				stFiltro = "Proveedor = 0 Or Ambos <> 0";
			}

			C1.Data.FilterCondition fcCliente = new 
				C1.Data.FilterCondition(cdsCliente, "Cliente", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcCliente);
			cdsCliente.Fill(fcTotal, false);
		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			string stFiltro = "{Pago.Pendiente} = true ";
			DateTime dtFechaDesde = (DateTime) cmbDesde.Value;
			DateTime dtFechaHasta = (DateTime) cmbHasta.Value;
			stFiltro += " And {Pago.Fecha} >= #" + dtFechaDesde.ToString("MM/dd/yyyy")
				+ "# And {Pago.Fecha} < #" + dtFechaHasta.AddDays(1).ToString("MM/dd/yyyy")
				+ "# And {Compra.idTipoFactura} = " + IdTipo.ToString();
			Reporte miRep = new Reporte("AbonoPendiente", stFiltro);
//			miRep.MdiParent = this.MdiParent;
			miRep.Show();
		}

		private void Depositos_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Depositos'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
		}

	}
}
