using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Tarjeta.
	/// </summary>
	public class Tarjeta : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsTarjeta;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo2;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private Infragistics.Win.Misc.UltraButton btCalculo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsientoV;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsientoD;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumVoucher;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVoucher;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton btAsientoD;
		private Infragistics.Win.Misc.UltraButton btMostrarAV;
		private Infragistics.Win.Misc.UltraButton btMostrarAD;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumDeposito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDeposito;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorriente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorrD;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumFactura;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalFactura;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btNumero;
		private Infragistics.Win.Misc.UltraButton btFecha;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaBusca;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscaNumero;
		private Infragistics.Win.Misc.UltraButton btPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPago;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRetencion;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaVoucher;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtComisionTar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtExceso;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetenido;
		private C1.Data.C1DataSet cdsCompraTabla;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private System.ComponentModel.IContainer components;

		public Tarjeta()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDep");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Voucher");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDep");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Voucher");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraFormaPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			this.cdsTarjeta = new C1.Data.C1DataSet();
			this.cmbFechaVoucher = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtNumVoucher = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtExceso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtComisionTar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtRetenido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumDeposito = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.ultraCalendarCombo2 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtIdTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAsientoV = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAsientoD = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtTotalFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btCalculo = new Infragistics.Win.Misc.UltraButton();
			this.txtVoucher = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.btAsientoD = new Infragistics.Win.Misc.UltraButton();
			this.btMostrarAV = new Infragistics.Win.Misc.UltraButton();
			this.btMostrarAD = new Infragistics.Win.Misc.UltraButton();
			this.txtDeposito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.cmbCuentaCorriente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCuentaCorrD = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsTarjeta = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtBuscaNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btNumero = new Infragistics.Win.Misc.UltraButton();
			this.btFecha = new Infragistics.Win.Misc.UltraButton();
			this.cmbFechaBusca = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btPago = new Infragistics.Win.Misc.UltraButton();
			this.txtRetencion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cdsCompraTabla = new C1.Data.C1DataSet();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.cdsTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaVoucher)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumVoucher)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtExceso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComisionTar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsientoV)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsientoD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVoucher)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorrD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscaNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsTarjeta
			// 
			this.cdsTarjeta.BindingContextCtrl = this;
			this.cdsTarjeta.DataLibrary = "LibFacturacion";
			this.cdsTarjeta.DataLibraryUrl = "";
			this.cdsTarjeta.DataSetDef = "dsTarjeta";
			this.cdsTarjeta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTarjeta.Name = "cdsTarjeta";
			this.cdsTarjeta.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTarjeta.SchemaDef = null;
			this.cdsTarjeta.BeforeFill += new C1.Data.FillEventHandler(this.cdsTarjeta_BeforeFill);
			// 
			// cmbFechaVoucher
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaVoucher.Appearance = appearance1;
			this.cmbFechaVoucher.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbFechaVoucher.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.Fecha"));
			dateButton1.Caption = "Today";
			this.cmbFechaVoucher.DateButtons.Add(dateButton1);
			this.cmbFechaVoucher.Format = "dd/MMM/yyyy";
			this.cmbFechaVoucher.Location = new System.Drawing.Point(192, 32);
			this.cmbFechaVoucher.Name = "cmbFechaVoucher";
			this.cmbFechaVoucher.NonAutoSizeHeight = 23;
			this.cmbFechaVoucher.Size = new System.Drawing.Size(120, 21);
			this.cmbFechaVoucher.SpinButtonsVisible = true;
			this.cmbFechaVoucher.TabIndex = 0;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(32, 32);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(81, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Fecha Voucher";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(32, 64);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(64, 16);
			this.label2.TabIndex = 2;
			this.label2.Text = "No Voucher";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(32, 96);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(58, 16);
			this.label3.TabIndex = 3;
			this.label3.Text = "Porcentaje";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(32, 128);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(145, 16);
			this.label4.TabIndex = 4;
			this.label4.Text = "Comisión Tarjeta de Crédito";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(32, 160);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(60, 16);
			this.label5.TabIndex = 5;
			this.label5.Text = "No Factura";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(32, 192);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(94, 16);
			this.label6.TabIndex = 6;
			this.label6.Text = "Exceso por Pagar";
			// 
			// txtNumVoucher
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumVoucher.Appearance = appearance2;
			this.txtNumVoucher.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.NumVoucher"));
			this.txtNumVoucher.Location = new System.Drawing.Point(104, 64);
			this.txtNumVoucher.Name = "txtNumVoucher";
			this.txtNumVoucher.Size = new System.Drawing.Size(80, 21);
			this.txtNumVoucher.TabIndex = 7;
			// 
			// ultraNumericEditor1
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor1.Appearance = appearance3;
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.Porcentaje"));
			this.ultraNumericEditor1.Location = new System.Drawing.Point(104, 96);
			this.ultraNumericEditor1.MaskInput = "nn,nn";
			this.ultraNumericEditor1.MaxValue = 100;
			this.ultraNumericEditor1.MinValue = 0;
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(80, 21);
			this.ultraNumericEditor1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor1.TabIndex = 8;
			// 
			// txtPago
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPago.Appearance = appearance4;
			this.txtPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.Agencia"));
			this.txtPago.FormatString = "#,##0.00";
			this.txtPago.Location = new System.Drawing.Point(208, 96);
			this.txtPago.Name = "txtPago";
			this.txtPago.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPago.PromptChar = ' ';
			this.txtPago.ReadOnly = true;
			this.txtPago.Size = new System.Drawing.Size(104, 21);
			this.txtPago.TabIndex = 9;
			// 
			// txtNumFactura
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumFactura.Appearance = appearance5;
			this.txtNumFactura.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.NumFactura"));
			this.txtNumFactura.Location = new System.Drawing.Point(104, 160);
			this.txtNumFactura.Name = "txtNumFactura";
			this.txtNumFactura.Size = new System.Drawing.Size(80, 21);
			this.txtNumFactura.TabIndex = 10;
			// 
			// txtExceso
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtExceso.Appearance = appearance6;
			this.txtExceso.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.Exceso"));
			this.txtExceso.FormatString = "#,##0.00";
			this.txtExceso.Location = new System.Drawing.Point(208, 192);
			this.txtExceso.Name = "txtExceso";
			this.txtExceso.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtExceso.PromptChar = ' ';
			this.txtExceso.ReadOnly = true;
			this.txtExceso.Size = new System.Drawing.Size(104, 21);
			this.txtExceso.TabIndex = 11;
			// 
			// txtComisionTar
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComisionTar.Appearance = appearance7;
			this.txtComisionTar.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.Comision"));
			this.txtComisionTar.FormatString = "#,##0.00";
			this.txtComisionTar.Location = new System.Drawing.Point(208, 128);
			this.txtComisionTar.Name = "txtComisionTar";
			this.txtComisionTar.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtComisionTar.PromptChar = ' ';
			this.txtComisionTar.ReadOnly = true;
			this.txtComisionTar.Size = new System.Drawing.Size(104, 21);
			this.txtComisionTar.TabIndex = 12;
			// 
			// txtIdCompra
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdCompra.Appearance = appearance8;
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.idCompra"));
			this.txtIdCompra.FormatString = "";
			this.txtIdCompra.Location = new System.Drawing.Point(96, 0);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(72, 21);
			this.txtIdCompra.TabIndex = 13;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(32, 224);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(66, 16);
			this.label7.TabIndex = 14;
			this.label7.Text = "No Depósito";
			// 
			// txtRetenido
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetenido.Appearance = appearance9;
			this.txtRetenido.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.Retenido"));
			this.txtRetenido.FormatString = "#,##0.00";
			this.txtRetenido.Location = new System.Drawing.Point(208, 256);
			this.txtRetenido.Name = "txtRetenido";
			this.txtRetenido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetenido.PromptChar = ' ';
			this.txtRetenido.ReadOnly = true;
			this.txtRetenido.Size = new System.Drawing.Size(104, 21);
			this.txtRetenido.TabIndex = 15;
			// 
			// txtNumDeposito
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumDeposito.Appearance = appearance10;
			this.txtNumDeposito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.NumDeposito"));
			this.txtNumDeposito.Location = new System.Drawing.Point(104, 224);
			this.txtNumDeposito.Name = "txtNumDeposito";
			this.txtNumDeposito.Size = new System.Drawing.Size(80, 21);
			this.txtNumDeposito.TabIndex = 16;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(32, 256);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(50, 16);
			this.label8.TabIndex = 17;
			this.label8.Text = "Retenido";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(32, 288);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(83, 16);
			this.label9.TabIndex = 19;
			this.label9.Text = "Fecha Depósito";
			// 
			// ultraCalendarCombo2
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCalendarCombo2.Appearance = appearance11;
			this.ultraCalendarCombo2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCalendarCombo2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.FechaDep"));
			dateButton2.Caption = "Today";
			this.ultraCalendarCombo2.DateButtons.Add(dateButton2);
			this.ultraCalendarCombo2.Format = "dd/MMM/yyyy";
			this.ultraCalendarCombo2.Location = new System.Drawing.Point(192, 288);
			this.ultraCalendarCombo2.Name = "ultraCalendarCombo2";
			this.ultraCalendarCombo2.NonAutoSizeHeight = 23;
			this.ultraCalendarCombo2.Size = new System.Drawing.Size(120, 21);
			this.ultraCalendarCombo2.SpinButtonsVisible = true;
			this.ultraCalendarCombo2.TabIndex = 18;
			// 
			// txtIdTarjeta
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdTarjeta.Appearance = appearance12;
			this.txtIdTarjeta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.idTarjeta"));
			this.txtIdTarjeta.FormatString = "";
			this.txtIdTarjeta.Location = new System.Drawing.Point(256, 0);
			this.txtIdTarjeta.Name = "txtIdTarjeta";
			this.txtIdTarjeta.PromptChar = ' ';
			this.txtIdTarjeta.Size = new System.Drawing.Size(100, 21);
			this.txtIdTarjeta.TabIndex = 20;
			// 
			// txtIdAsientoV
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdAsientoV.Appearance = appearance13;
			this.txtIdAsientoV.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.idAsientoV"));
			this.txtIdAsientoV.FormatString = "";
			this.txtIdAsientoV.Location = new System.Drawing.Point(24, 0);
			this.txtIdAsientoV.Name = "txtIdAsientoV";
			this.txtIdAsientoV.PromptChar = ' ';
			this.txtIdAsientoV.Size = new System.Drawing.Size(64, 21);
			this.txtIdAsientoV.TabIndex = 21;
			// 
			// txtIdAsientoD
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdAsientoD.Appearance = appearance14;
			this.txtIdAsientoD.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.idAsientoD"));
			this.txtIdAsientoD.FormatString = "";
			this.txtIdAsientoD.Location = new System.Drawing.Point(176, 0);
			this.txtIdAsientoD.Name = "txtIdAsientoD";
			this.txtIdAsientoD.PromptChar = ' ';
			this.txtIdAsientoD.Size = new System.Drawing.Size(72, 21);
			this.txtIdAsientoD.TabIndex = 22;
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.BarraMovimiento = true;
			this.barraDatoSQL1.DataMember = "Tarjeta";
			this.barraDatoSQL1.DataNombreId = "idTarjeta";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsTarjeta;
			this.barraDatoSQL1.DataTabla = "Tarjeta";
			this.barraDatoSQL1.DataTablaHija = "";
			this.barraDatoSQL1.Location = new System.Drawing.Point(40, 328);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(304, 24);
			this.barraDatoSQL1.TabIndex = 23;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = false;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDatoSQL1_Buscar);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDatoSQL1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDatoSQL1_Imprimir);
			this.barraDatoSQL1.Restaurar += new System.EventHandler(this.barraDatoSQL1_Restaurar);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDatoSQL1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDatoSQL1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDatoSQL1.Impresora += new System.EventHandler(this.barraDatoSQL1_Impresora);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDatoSQL1_Grabar);
			// 
			// txtTotalFactura
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalFactura.Appearance = appearance15;
			this.txtTotalFactura.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.TotalFactura"));
			this.txtTotalFactura.FormatString = "#,##0.00";
			this.txtTotalFactura.Location = new System.Drawing.Point(208, 160);
			this.txtTotalFactura.Name = "txtTotalFactura";
			this.txtTotalFactura.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalFactura.PromptChar = ' ';
			this.txtTotalFactura.Size = new System.Drawing.Size(104, 21);
			this.txtTotalFactura.TabIndex = 24;
			// 
			// btCalculo
			// 
			this.btCalculo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCalculo.Location = new System.Drawing.Point(336, 128);
			this.btCalculo.Name = "btCalculo";
			this.btCalculo.Size = new System.Drawing.Size(88, 23);
			this.btCalculo.TabIndex = 25;
			this.btCalculo.Text = "&Cálculo";
			this.btCalculo.Click += new System.EventHandler(this.btCalculo_Click);
			// 
			// txtVoucher
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVoucher.Appearance = appearance16;
			this.txtVoucher.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.Voucher"));
			this.txtVoucher.FormatString = "#,##0.00";
			this.txtVoucher.Location = new System.Drawing.Point(208, 64);
			this.txtVoucher.Name = "txtVoucher";
			this.txtVoucher.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVoucher.PromptChar = ' ';
			this.txtVoucher.Size = new System.Drawing.Size(104, 21);
			this.txtVoucher.TabIndex = 26;
			// 
			// ultraButton1
			// 
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton1.Location = new System.Drawing.Point(336, 96);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(88, 23);
			this.ultraButton1.TabIndex = 27;
			this.ultraButton1.Text = "Asi. &Voucher";
			this.ultraButton1.Click += new System.EventHandler(this.btAsientoV_Click);
			// 
			// btAsientoD
			// 
			this.btAsientoD.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAsientoD.Location = new System.Drawing.Point(336, 224);
			this.btAsientoD.Name = "btAsientoD";
			this.btAsientoD.Size = new System.Drawing.Size(88, 23);
			this.btAsientoD.TabIndex = 28;
			this.btAsientoD.Text = "Asi. &Depósito";
			this.btAsientoD.Click += new System.EventHandler(this.btAsientoD_Click);
			// 
			// btMostrarAV
			// 
			this.btMostrarAV.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btMostrarAV.Location = new System.Drawing.Point(432, 96);
			this.btMostrarAV.Name = "btMostrarAV";
			this.btMostrarAV.Size = new System.Drawing.Size(24, 23);
			this.btMostrarAV.TabIndex = 29;
			this.btMostrarAV.Text = "...";
			this.btMostrarAV.Click += new System.EventHandler(this.btMostrarAV_Click);
			// 
			// btMostrarAD
			// 
			this.btMostrarAD.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btMostrarAD.Location = new System.Drawing.Point(432, 224);
			this.btMostrarAD.Name = "btMostrarAD";
			this.btMostrarAD.Size = new System.Drawing.Size(24, 23);
			this.btMostrarAD.TabIndex = 30;
			this.btMostrarAD.Text = "...";
			this.btMostrarAD.Click += new System.EventHandler(this.btMostrarAD_Click);
			// 
			// txtDeposito
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDeposito.Appearance = appearance17;
			this.txtDeposito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.Deposito"));
			this.txtDeposito.FormatString = "#,##0.00";
			this.txtDeposito.Location = new System.Drawing.Point(208, 224);
			this.txtDeposito.Name = "txtDeposito";
			this.txtDeposito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDeposito.PromptChar = ' ';
			this.txtDeposito.Size = new System.Drawing.Size(104, 21);
			this.txtDeposito.TabIndex = 31;
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
			// 
			// cmbCuentaCorriente
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaCorriente.Appearance = appearance18;
			this.cmbCuentaCorriente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorriente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorriente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.idCuentaCorriente"));
			this.cmbCuentaCorriente.DataMember = "CuentaCorriente";
			this.cmbCuentaCorriente.DataSource = this.cdsCuentaCorr;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
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
																										 ultraGridColumn12});
			this.cmbCuentaCorriente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuentaCorriente.DisplayMember = "NumCuenta";
			this.cmbCuentaCorriente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCorriente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCorriente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorriente.Location = new System.Drawing.Point(336, 64);
			this.cmbCuentaCorriente.Name = "cmbCuentaCorriente";
			this.cmbCuentaCorriente.Size = new System.Drawing.Size(136, 21);
			this.cmbCuentaCorriente.TabIndex = 32;
			this.cmbCuentaCorriente.ValueMember = "idCuentaCorriente";
			// 
			// cmbCuentaCorrD
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaCorrD.Appearance = appearance19;
			this.cmbCuentaCorrD.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorrD.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorrD.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.idCuentaCorrD"));
			this.cmbCuentaCorrD.DataMember = "CuentaCorriente";
			this.cmbCuentaCorrD.DataSource = this.cdsCuentaCorr;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn18.Header.VisiblePosition = 5;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 6;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 7;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 8;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 9;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 10;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 11;
			ultraGridColumn24.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn24});
			this.cmbCuentaCorrD.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuentaCorrD.DisplayMember = "NumCuenta";
			this.cmbCuentaCorrD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCorrD.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCorrD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorrD.Location = new System.Drawing.Point(336, 256);
			this.cmbCuentaCorrD.Name = "cmbCuentaCorrD";
			this.cmbCuentaCorrD.Size = new System.Drawing.Size(136, 21);
			this.cmbCuentaCorrD.TabIndex = 33;
			this.cmbCuentaCorrD.ValueMember = "idCuentaCorriente";
			// 
			// btFactura
			// 
			this.btFactura.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFactura.Location = new System.Drawing.Point(336, 160);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(88, 23);
			this.btFactura.TabIndex = 34;
			this.btFactura.Text = "&Factura";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Band 0";
			this.ultraGrid1.DataSource = this.udsTarjeta;
			appearance20.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance20;
			ultraGridBand3.AddButtonCaption = "Tarjeta";
			ultraGridColumn25.Header.VisiblePosition = 5;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Format = "dd/MMM/yyyy";
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Width = 77;
			ultraGridColumn27.Format = "dd/MMM/yyyy";
			ultraGridColumn27.Header.Caption = "F. Dep";
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 76;
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.Width = 83;
			ultraGridColumn29.Header.VisiblePosition = 4;
			ultraGridColumn29.Width = 74;
			ultraGridColumn30.Header.VisiblePosition = 3;
			ultraGridColumn30.Width = 76;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance21.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance21;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance22.FontData.BoldAsString = "True";
			appearance22.FontData.Name = "Arial";
			appearance22.FontData.SizeInPoints = 10F;
			appearance22.ForeColor = System.Drawing.Color.White;
			appearance22.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance22;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance24;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(512, 64);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(288, 296);
			this.ultraGrid1.TabIndex = 35;
			this.ultraGrid1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.ultraGrid1_MouseUp);
			// 
			// udsTarjeta
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			this.udsTarjeta.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn1,
																													 ultraDataColumn2,
																													 ultraDataColumn3,
																													 ultraDataColumn4,
																													 ultraDataColumn5,
																													 ultraDataColumn6});
			// 
			// txtBuscaNumero
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscaNumero.Appearance = appearance25;
			this.txtBuscaNumero.Location = new System.Drawing.Point(624, 8);
			this.txtBuscaNumero.Name = "txtBuscaNumero";
			this.txtBuscaNumero.Size = new System.Drawing.Size(80, 21);
			this.txtBuscaNumero.TabIndex = 36;
			// 
			// btNumero
			// 
			this.btNumero.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btNumero.Location = new System.Drawing.Point(528, 8);
			this.btNumero.Name = "btNumero";
			this.btNumero.Size = new System.Drawing.Size(64, 24);
			this.btNumero.TabIndex = 37;
			this.btNumero.Text = "Número";
			this.toolTip1.SetToolTip(this.btNumero, "Bucar por Numero / Editar Bloqueado");
			this.btNumero.Click += new System.EventHandler(this.btNumero_Click);
			this.btNumero.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btNumero_MouseDown);
			// 
			// btFecha
			// 
			this.btFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFecha.Location = new System.Drawing.Point(528, 32);
			this.btFecha.Name = "btFecha";
			this.btFecha.Size = new System.Drawing.Size(64, 24);
			this.btFecha.TabIndex = 38;
			this.btFecha.Text = "Fecha";
			this.btFecha.Click += new System.EventHandler(this.btFecha_Click);
			// 
			// cmbFechaBusca
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaBusca.Appearance = appearance26;
			this.cmbFechaBusca.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.cmbFechaBusca.DateButtons.Add(dateButton3);
			this.cmbFechaBusca.Format = "dd/MMM/yyyy";
			this.cmbFechaBusca.Location = new System.Drawing.Point(608, 32);
			this.cmbFechaBusca.Name = "cmbFechaBusca";
			this.cmbFechaBusca.NonAutoSizeHeight = 23;
			this.cmbFechaBusca.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaBusca.SpinButtonsVisible = true;
			this.cmbFechaBusca.TabIndex = 39;
			// 
			// btPago
			// 
			this.btPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPago.Location = new System.Drawing.Point(336, 192);
			this.btPago.Name = "btPago";
			this.btPago.Size = new System.Drawing.Size(88, 23);
			this.btPago.TabIndex = 40;
			this.btPago.Text = "&Pago";
			this.btPago.Click += new System.EventHandler(this.btPago_Click);
			// 
			// txtRetencion
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetencion.Appearance = appearance27;
			this.txtRetencion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.Retencion"));
			this.txtRetencion.Location = new System.Drawing.Point(104, 256);
			this.txtRetencion.Name = "txtRetencion";
			this.txtRetencion.Size = new System.Drawing.Size(80, 21);
			this.txtRetencion.TabIndex = 41;
			this.toolTip1.SetToolTip(this.txtRetencion, "Numero de Retencion de Tarjeta de crediti");
			// 
			// cdsCompraTabla
			// 
			this.cdsCompraTabla.BindingContextCtrl = this;
			this.cdsCompraTabla.DataLibrary = "LibFacturacion";
			this.cdsCompraTabla.DataLibraryUrl = "";
			this.cdsCompraTabla.DataSetDef = "dsCompraTabla";
			this.cdsCompraTabla.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsCompraTabla.Name = "cdsCompraTabla";
			this.cdsCompraTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraTabla.SchemaDef = null;
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTarjeta, "Tarjeta.idFormaPago"));
			this.cmbFormaPago.DataMember = "CompraFormaPago";
			this.cmbFormaPago.DataSource = this.cdsCompraTabla;
			this.cmbFormaPago.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 22;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 15;
			ultraGridColumn33.Header.VisiblePosition = 2;
			ultraGridColumn33.Width = 117;
			ultraGridColumn34.Header.VisiblePosition = 3;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 16;
			ultraGridColumn35.Header.VisiblePosition = 4;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 17;
			ultraGridColumn36.Header.VisiblePosition = 5;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 21;
			ultraGridColumn37.Header.VisiblePosition = 6;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 55;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(336, 32);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(136, 21);
			this.cmbFormaPago.TabIndex = 42;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			// 
			// Tarjeta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(808, 374);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.txtRetencion);
			this.Controls.Add(this.btPago);
			this.Controls.Add(this.cmbFechaBusca);
			this.Controls.Add(this.btFecha);
			this.Controls.Add(this.btNumero);
			this.Controls.Add(this.txtBuscaNumero);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.cmbCuentaCorrD);
			this.Controls.Add(this.cmbCuentaCorriente);
			this.Controls.Add(this.txtDeposito);
			this.Controls.Add(this.btMostrarAD);
			this.Controls.Add(this.btMostrarAV);
			this.Controls.Add(this.btAsientoD);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.txtVoucher);
			this.Controls.Add(this.btCalculo);
			this.Controls.Add(this.txtTotalFactura);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.txtIdAsientoD);
			this.Controls.Add(this.txtIdAsientoV);
			this.Controls.Add(this.txtIdTarjeta);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.ultraCalendarCombo2);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.txtNumDeposito);
			this.Controls.Add(this.txtRetenido);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtComisionTar);
			this.Controls.Add(this.txtExceso);
			this.Controls.Add(this.txtNumFactura);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.txtNumVoucher);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbFechaVoucher);
			this.MaximizeBox = false;
			this.Name = "Tarjeta";
			this.Text = "Tarjeta";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Tarjeta_Closing);
			this.Load += new System.EventHandler(this.Tarjeta_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaVoucher)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumVoucher)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtExceso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComisionTar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsientoV)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsientoD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVoucher)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorrD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscaNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void Tarjeta_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Tarjeta'";
			Funcion.EjecutaSQL(cdsCuentaCorr, stAudita, true);
			txtIdAsientoD.Width = 0;
			txtIdAsientoV.Width = 0;
			txtIdCompra.Width = 0;
			txtIdTarjeta.Width = 0;
			barraDatoSQL1.ProximoId(4);
			cmbFechaBusca.Value = DateTime.Today;
		}

		private void cdsTarjeta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTarjeta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Borrar(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Borrar éste Registro?"
				+ "\nNo puede ser revertido este proceso", "Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Warning)) return;
      string stSelect = "Exec TarjetaBorrar " + txtIdTarjeta.Value.ToString();
			Funcion.EjecutaSQL(cdsCuentaCorr, stSelect, true);
			barraDatoSQL1.ProximoId(2);
		}

		private void barraDatoSQL1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Crear(object sender, System.EventArgs e)
		{
			barraDatoSQL1.CrearRegistro();
			barraDatoSQL1.GrabaRegistro();
			barraDatoSQL1.EditarRegistro();
		}

		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			barraDatoSQL1.DeshacerRegistro();
			txtBuscaNumero.Enabled = true;
			cmbFechaBusca.Enabled = true;
		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			barraDatoSQL1.EditarRegistro();
		}

		private void barraDatoSQL1_Grabar(object sender, System.EventArgs e)
		{
			barraDatoSQL1.GrabaRegistro();
			txtBuscaNumero.Enabled = true;
			cmbFechaBusca.Enabled = true;
		}

		private void barraDatoSQL1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Imprimir(object sender, System.EventArgs e)
		{
			TarjetaImprime miTarjeta = new TarjetaImprime();
			miTarjeta.MdiParent = this.MdiParent;
			miTarjeta.Show();
		}

		private void barraDatoSQL1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private bool Verifica()
		{
			if (barraDatoSQL1.bEditar) 
			{
				MessageBox.Show("Debe grabar el registro para generar el calculo", "Información");
				return false;
			}
			return true;
		}

		private void btCalculo_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string stSelect = "Exec TarjetaCalculo " + txtIdTarjeta.Value.ToString();
			Funcion.EjecutaSQL(cdsTarjeta, stSelect, true);
			barraDatoSQL1.ProximoId(5);
		}

		private void btAsientoV_Click(object sender, System.EventArgs e)
		{
			if (txtIdAsientoV.Value != DBNull.Value && (int) txtIdAsientoV.Value > 0)
			{
				if (DialogResult.No == MessageBox.Show("Asiento de Voucher ya generado. "
					+ "¿Desea sobreescribirlo?",	"Confirmación", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			}
			string stSelect = "Exec TarjetaAsientoV " + txtIdTarjeta.Value.ToString();
			Funcion.EjecutaSQL(cdsTarjeta, stSelect, true);
			barraDatoSQL1.ProximoId(5);
			MessageBox.Show("Asiento de Voucher Generado", "Información");
		}

		private void btMostrarAV_Click(object sender, System.EventArgs e)
		{
			if (txtIdAsientoV.Value == DBNull.Value || (int) txtIdAsientoV.Value == 0)
			{
				MessageBox.Show("Asiento de Voucher No generado. "
					,	"Información", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			Asiento miAsiento = new Asiento((int) txtIdAsientoV.Value);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
		}

		private void btAsientoD_Click(object sender, System.EventArgs e)
		{
			if (txtIdAsientoD.Value != DBNull.Value && (int) txtIdAsientoD.Value > 0)
			{
				if (DialogResult.No == MessageBox.Show("Asiento de Deposito ya generado. "
					+ "¿Desea sobreescribirlo?",	"Confirmación", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			}
			string stSelect = "Exec TarjetaAsientoD " + txtIdTarjeta.Value.ToString();
			Funcion.EjecutaSQL(cdsTarjeta, stSelect, true);
			barraDatoSQL1.ProximoId(5);
			MessageBox.Show("Asiento de Deposito Generado", "Información");
		}

		private void btMostrarAD_Click(object sender, System.EventArgs e)
		{
			if (txtIdAsientoD.Value == DBNull.Value || (int) txtIdAsientoD.Value == 0)
			{
				MessageBox.Show("Asiento de Deposito No generado. "
					,	"Información", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			Asiento miAsiento = new Asiento((int) txtIdAsientoD.Value);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
		}

		private void BorraErrores()
		{
			errorProvider.SetError(txtNumFactura, "");
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			BorraErrores();
			if (txtNumFactura.Value == null || txtNumFactura.Value.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtNumFactura, "Ingrese Factura");
				return;
			}
			string stSelect = "Exec TarjetaFactura " + txtIdTarjeta.Value.ToString();
			Funcion.EjecutaSQL(cdsTarjeta, stSelect, true);
			barraDatoSQL1.ProximoId(5);
			MessageBox.Show("Valores de Factura Generados", "Información");
		}

		private void LlenaGrilla(SqlDataReader miReader)
		{
			udsTarjeta.Rows.Clear();
			int iLinea = 0;
			while (miReader.Read())
			{
				udsTarjeta.Rows.Add();
				udsTarjeta.Rows[iLinea]["Fecha"] = miReader.GetDateTime(0);
				udsTarjeta.Rows[iLinea]["FechaDep"] = miReader.GetDateTime(1);
				udsTarjeta.Rows[iLinea]["Voucher"] = miReader.GetString(2);
				udsTarjeta.Rows[iLinea]["Factura"] = miReader.GetString(3);
				udsTarjeta.Rows[iLinea]["Deposito"] = miReader.GetString(4);
				udsTarjeta.Rows[iLinea]["idTarjeta"] = miReader.GetInt32(5);
				iLinea ++;
			}
			miReader.Close();
		}

		private void btNumero_Click(object sender, System.EventArgs e)
		{
			string stNumero = txtBuscaNumero.Text.Trim();
			string stBusca = "'NumVoucher Like ''%" + stNumero
				+ "%'' Or NumFactura Like ''%" + stNumero + "%'' "
				+ "Or NumDeposito Like ''%" + stNumero + "%'' '";
			string stSelect = "Exec TarjetaBuscar " + stBusca;
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCuentaCorr, stSelect, true);
			LlenaGrilla(miReader);
		}

		private void ultraGrid1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (ultraGrid1.ActiveRow == null) return;
			barraDatoSQL1.IdRegistro = (int) ultraGrid1.ActiveRow.Cells["idTarjeta"].Value;
			barraDatoSQL1.ProximoId(5);
		}

		private void btFecha_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime) cmbFechaBusca.Value;
			string stFecha = "''" + dtFecha.ToString("yyyyMMdd") + "''";
			string stFecha2 = "''" + dtFecha.AddDays(1).ToString("yyyyMMdd") + "''";
			string stBusca = "'(Fecha >= " + stFecha + " And Fecha < " + stFecha2
				+ ") Or (FechaDep >= " + stFecha + " And Fecha < " + stFecha2 + ") '";
			string stSelect = "Exec TarjetaBuscar " + stBusca;
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCuentaCorr, stSelect, true);
			LlenaGrilla(miReader);
		}

		private void Tarjeta_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Tarjeta'";
			Funcion.EjecutaSQL(cdsCuentaCorr, stAudita, true);
		}
		CultureInfo us = new CultureInfo("en-US");
		private void btPago_Click(object sender, System.EventArgs e)
		{
			double dPago = (double) txtPago.Value;
			string stRetencion = txtRetencion.Text.Trim();
//			string stExec = "Exec TarjetaPago '" + txtNumFactura.Text.Trim() + "', "
//				+ dPago.ToString("0.00", us) + ", '" + stRetencion + "', "
//				+ txtIdAsientoV.Value.ToString();
			string stExec = "Exec TarjetaPago " + txtIdTarjeta.Text.Trim();
			int iResultado = Funcion.iEscalarSQL(cdsCuentaCorr, stExec);
			if (iResultado == 1)
				MessageBox.Show("Pago Generado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btNumero_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				txtPago.ReadOnly = !txtPago.ReadOnly;
				txtComisionTar.ReadOnly = !txtComisionTar.ReadOnly;
				txtExceso.ReadOnly = !txtExceso.ReadOnly;
				txtRetenido.ReadOnly = !txtRetenido.ReadOnly;
			}
		}
	}
}
