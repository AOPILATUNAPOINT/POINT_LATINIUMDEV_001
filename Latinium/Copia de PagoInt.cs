using System;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Pago1.
	/// </summary>
	public class PagoInt : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private C1.Data.C1DataSet cdsPagos;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		
		private bool bNuevo = true;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private C1.Data.C1DataSet cdsCompraTabla;
		private C1.Data.C1DataView cdvCompraTabla;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private C1.Data.C1DataView cdvPago;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorr;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNota;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraLabel ultraLabel8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCheque;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel9;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAsiento;
		private Infragistics.Win.Misc.UltraLabel ultraLabel10;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReferencia;
		private Infragistics.Win.Misc.UltraLabel ultraLabel12;
		private Infragistics.Win.Misc.UltraLabel ultraLabel14;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroCosto;
		private Infragistics.Win.Misc.UltraButton btInteres;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsTipoPrestamo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoPago;
		private Infragistics.Win.Misc.UltraButton btTotal;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataView cdvProyecto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdPago;
		private Infragistics.Win.Misc.UltraButton btCapital;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel13;
		private Infragistics.Win.Misc.UltraLabel ultraLabel11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAmortizacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMora;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProntoPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtInteres;
		private int IdCompra = 0;

		double nSaldo = 0;
		public PagoInt(int idBusqueda, bool bNuevos, double dSaldo)
		{
			InitializeComponent();
			IdCompra = idBusqueda;
			bNuevo = bNuevos;
			nSaldo = dSaldo;
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraFormaPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hoja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProcesoOk");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPrestamo");
			this.cdsPagos = new C1.Data.C1DataSet();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCompraTabla = new C1.Data.C1DataView();
			this.cdsCompraTabla = new C1.Data.C1DataSet();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.cdvPago = new C1.Data.C1DataView();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbCuentaCorr = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNota = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraLabel8 = new Infragistics.Win.Misc.UltraLabel();
			this.txtCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel9 = new Infragistics.Win.Misc.UltraLabel();
			this.txtAsiento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel10 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel12 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel14 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbCentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btInteres = new Infragistics.Win.Misc.UltraButton();
			this.btTotal = new Infragistics.Win.Misc.UltraButton();
			this.cmbTipoPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsTipoPrestamo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btCapital = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel13 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel11 = new Infragistics.Win.Misc.UltraLabel();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtAmortizacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMora = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProntoPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtInteres = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipoPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAmortizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMora)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProntoPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsPagos
			// 
			this.cdsPagos.BindingContextCtrl = this;
			this.cdsPagos.DataLibrary = "LibFacturacion";
			this.cdsPagos.DataLibraryUrl = "";
			this.cdsPagos.DataSetDef = "dsPagos";
			this.cdsPagos.FillOnRequest = false;
			this.cdsPagos.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsPagos.Name = "cdsPagos";
			this.cdsPagos.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPagos.SchemaDef = null;
			this.cdsPagos.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsPagos_PositionChanged);
			this.cdsPagos.BeforeFill += new C1.Data.FillEventHandler(this.cdsPagos_BeforeFill);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.Location = new System.Drawing.Point(16, 272);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(48, 16);
			this.ultraLabel1.TabIndex = 1;
			this.ultraLabel1.Text = "Notas";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.Location = new System.Drawing.Point(16, 40);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(56, 16);
			this.ultraLabel3.TabIndex = 7;
			this.ultraLabel3.Text = "Fecha:";
			// 
			// txtFecha
			// 
			this.txtFecha.AllowMonthSelection = true;
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.Fecha"));
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Format = "dd/MMM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(96, 40);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(112, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 1;
			this.txtFecha.Value = new System.DateTime(2006, 3, 22, 0, 0, 0, 0);
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.idFormaPago"));
			this.cmbFormaPago.DataSource = this.cdvCompraTabla;
			this.cmbFormaPago.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.Caption = "Forma de Pago";
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 261;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 62;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 73;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(96, 160);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(280, 21);
			this.cmbFormaPago.TabIndex = 4;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			// 
			// cdvCompraTabla
			// 
			this.cdvCompraTabla.BindingContextCtrl = this;
			this.cdvCompraTabla.DataSet = this.cdsCompraTabla;
			this.cdvCompraTabla.TableName = "";
			this.cdvCompraTabla.TableViewName = "CompraFormaPago";
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
			this.cdsCompraTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompraTabla_BeforeFill);
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.Location = new System.Drawing.Point(16, 160);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(72, 16);
			this.ultraLabel4.TabIndex = 12;
			this.ultraLabel4.Text = "Forma Pago:";
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.Location = new System.Drawing.Point(16, 240);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(56, 16);
			this.ultraLabel5.TabIndex = 13;
			this.ultraLabel5.Text = "Concepto";
			// 
			// cdvPago
			// 
			this.cdvPago.BindingContextCtrl = this;
			this.cdvPago.DataSet = this.cdsPagos;
			this.cdvPago.Sort = "idPago";
			this.cdvPago.TableName = "";
			this.cdvPago.TableViewName = "Pago";
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.Location = new System.Drawing.Point(16, 136);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(69, 14);
			this.ultraLabel6.TabIndex = 16;
			this.ultraLabel6.Text = "Cuenta Corr:";
			// 
			// cmbCuentaCorr
			// 
			this.cmbCuentaCorr.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorr.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorr.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.idCuentaCorriente"));
			this.cmbCuentaCorr.DataMember = "CuentaCorriente";
			this.cmbCuentaCorr.DataSource = this.cdsCuentaCorr;
			this.cmbCuentaCorr.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "CompraFormaPago";
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 54;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.Caption = "Nº Cuenta";
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Width = 261;
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 10;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 46;
			ultraGridColumn19.Header.VisiblePosition = 11;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 64;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn19});
			this.cmbCuentaCorr.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuentaCorr.DisplayMember = "NumCuenta";
			this.cmbCuentaCorr.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCorr.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCorr.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorr.Location = new System.Drawing.Point(96, 136);
			this.cmbCuentaCorr.Name = "cmbCuentaCorr";
			this.cmbCuentaCorr.Size = new System.Drawing.Size(280, 21);
			this.cmbCuentaCorr.TabIndex = 3;
			this.cmbCuentaCorr.ValueMember = "idCuentaCorriente";
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
			// txtConcepto
			// 
			this.txtConcepto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.Concepto"));
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Location = new System.Drawing.Point(96, 240);
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(128, 21);
			this.txtConcepto.TabIndex = 6;
			// 
			// txtNota
			// 
			this.txtNota.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.Notas"));
			this.txtNota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNota.Location = new System.Drawing.Point(96, 272);
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(128, 21);
			this.txtNota.TabIndex = 7;
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
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
			// ultraLabel7
			// 
			this.ultraLabel7.Location = new System.Drawing.Point(16, 184);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(48, 16);
			this.ultraLabel7.TabIndex = 22;
			this.ultraLabel7.Text = "Banco";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// ultraLabel8
			// 
			this.ultraLabel8.Location = new System.Drawing.Point(240, 240);
			this.ultraLabel8.Name = "ultraLabel8";
			this.ultraLabel8.Size = new System.Drawing.Size(48, 16);
			this.ultraLabel8.TabIndex = 24;
			this.ultraLabel8.Text = "Cheque";
			// 
			// txtCheque
			// 
			this.txtCheque.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.Cheque"));
			this.txtCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCheque.Location = new System.Drawing.Point(296, 240);
			this.txtCheque.Name = "txtCheque";
			this.txtCheque.Size = new System.Drawing.Size(80, 21);
			this.txtCheque.TabIndex = 25;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.idCompra"));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.FormatString = "#,##0.00";
			this.txtIdCompra.Location = new System.Drawing.Point(360, 328);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(24, 21);
			this.txtIdCompra.TabIndex = 26;
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Pago";
			this.barraDato1.DataNombreId = "idPago";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsPagos;
			this.barraDato1.DataTabla = "Pago";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(24, 336);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 27;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtNumero
			// 
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.Numero"));
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(296, 272);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(80, 21);
			this.txtNumero.TabIndex = 29;
			// 
			// ultraLabel9
			// 
			this.ultraLabel9.Location = new System.Drawing.Point(240, 272);
			this.ultraLabel9.Name = "ultraLabel9";
			this.ultraLabel9.Size = new System.Drawing.Size(48, 16);
			this.ultraLabel9.TabIndex = 28;
			this.ultraLabel9.Text = "Número";
			// 
			// txtAsiento
			// 
			this.txtAsiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.NumeroIngreso"));
			this.txtAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAsiento.Location = new System.Drawing.Point(296, 304);
			this.txtAsiento.Name = "txtAsiento";
			this.txtAsiento.Size = new System.Drawing.Size(80, 21);
			this.txtAsiento.TabIndex = 31;
			// 
			// ultraLabel10
			// 
			this.ultraLabel10.Location = new System.Drawing.Point(240, 304);
			this.ultraLabel10.Name = "ultraLabel10";
			this.ultraLabel10.Size = new System.Drawing.Size(48, 16);
			this.ultraLabel10.TabIndex = 30;
			this.ultraLabel10.Text = "Asiento";
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
			// cmbBanco
			// 
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.idBanco"));
			this.cmbBanco.DataMember = "Banco";
			this.cmbBanco.DataSource = this.cdsCuentaCorr;
			this.cmbBanco.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "CompraFormaPago";
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Width = 261;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 81;
			ultraGridColumn23.Header.VisiblePosition = 3;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 83;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 5;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 128;
			ultraGridColumn26.Header.VisiblePosition = 6;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 130;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(96, 184);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(280, 21);
			this.cmbBanco.TabIndex = 34;
			this.cmbBanco.ValueMember = "idBanco";
			// 
			// txtReferencia
			// 
			this.txtReferencia.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.Referencia"));
			this.txtReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtReferencia.Location = new System.Drawing.Point(96, 304);
			this.txtReferencia.Name = "txtReferencia";
			this.txtReferencia.Size = new System.Drawing.Size(128, 21);
			this.txtReferencia.TabIndex = 37;
			// 
			// ultraLabel12
			// 
			this.ultraLabel12.Location = new System.Drawing.Point(16, 304);
			this.ultraLabel12.Name = "ultraLabel12";
			this.ultraLabel12.Size = new System.Drawing.Size(72, 16);
			this.ultraLabel12.TabIndex = 36;
			this.ultraLabel12.Text = "Referencia";
			// 
			// ultraLabel14
			// 
			this.ultraLabel14.Location = new System.Drawing.Point(16, 208);
			this.ultraLabel14.Name = "ultraLabel14";
			this.ultraLabel14.Size = new System.Drawing.Size(56, 16);
			this.ultraLabel14.TabIndex = 42;
			this.ultraLabel14.Text = "C. Costo";
			// 
			// cmbCentroCosto
			// 
			this.cmbCentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroCosto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.idProyecto"));
			this.cmbCentroCosto.DisplayLayout.AutoFitColumns = true;
			this.cmbCentroCosto.DisplayMember = "Nombre";
			this.cmbCentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentroCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroCosto.Location = new System.Drawing.Point(96, 208);
			this.cmbCentroCosto.Name = "cmbCentroCosto";
			this.cmbCentroCosto.Size = new System.Drawing.Size(280, 21);
			this.cmbCentroCosto.TabIndex = 41;
			this.cmbCentroCosto.ValueMember = "idProyecto";
			// 
			// btInteres
			// 
			this.btInteres.Location = new System.Drawing.Point(16, 72);
			this.btInteres.Name = "btInteres";
			this.btInteres.Size = new System.Drawing.Size(64, 23);
			this.btInteres.TabIndex = 47;
			this.btInteres.Text = "Interes";
			this.btInteres.Click += new System.EventHandler(this.btInteres_Click);
			// 
			// btTotal
			// 
			this.btTotal.Location = new System.Drawing.Point(16, 104);
			this.btTotal.Name = "btTotal";
			this.btTotal.Size = new System.Drawing.Size(64, 23);
			this.btTotal.TabIndex = 48;
			this.btTotal.Text = "Total";
			this.btTotal.Click += new System.EventHandler(this.btTotal_Click);
			// 
			// cmbTipoPago
			// 
			this.cmbTipoPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoPago.DataMember = "Band 0";
			this.cmbTipoPago.DataSource = this.udsTipoPrestamo;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbTipoPago.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipoPago.DisplayMember = "TipoPrestamo";
			this.cmbTipoPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoPago.Location = new System.Drawing.Point(224, 104);
			this.cmbTipoPago.Name = "cmbTipoPago";
			this.cmbTipoPago.Size = new System.Drawing.Size(152, 21);
			this.cmbTipoPago.TabIndex = 49;
			this.cmbTipoPago.ValueMember = "idTipoPrestamo";
			// 
			// udsTipoPrestamo
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsTipoPrestamo.Band.Columns.AddRange(new object[] {
																																ultraDataColumn1,
																																ultraDataColumn2});
			this.udsTipoPrestamo.Rows.AddRange(new object[] {
																												new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														"idTipoPrestamo",
																																																														1,
																																																														"TipoPrestamo",
																																																														"Pago de Cuota"}, new object[0]),
																												new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														"idTipoPrestamo",
																																																														2,
																																																														"TipoPrestamo",
																																																														"Pago Total"}, new object[0]),
																												new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														"idTipoPrestamo",
																																																														3,
																																																														"TipoPrestamo",
																																																														"Pago Parcial"}, new object[0])});
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible =1";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// ultraNumericEditor1
			// 
			this.ultraNumericEditor1.Location = new System.Drawing.Point(336, 128);
			this.ultraNumericEditor1.MaskInput = "nnn";
			this.ultraNumericEditor1.MinValue = 1;
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(64, 21);
			this.ultraNumericEditor1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor1.TabIndex = 50;
			this.ultraNumericEditor1.Value = 1;
			this.ultraNumericEditor1.Visible = false;
			// 
			// txtIdPago
			// 
			this.txtIdPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.idPago"));
			this.txtIdPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdPago.FormatString = "#,##0.00";
			this.txtIdPago.Location = new System.Drawing.Point(336, 336);
			this.txtIdPago.Name = "txtIdPago";
			this.txtIdPago.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIdPago.PromptChar = ' ';
			this.txtIdPago.Size = new System.Drawing.Size(24, 21);
			this.txtIdPago.TabIndex = 52;
			// 
			// btCapital
			// 
			this.btCapital.Location = new System.Drawing.Point(16, 8);
			this.btCapital.Name = "btCapital";
			this.btCapital.Size = new System.Drawing.Size(64, 23);
			this.btCapital.TabIndex = 53;
			this.btCapital.Text = "Capital";
			this.btCapital.Click += new System.EventHandler(this.btCapital_Click);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(224, 40);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(33, 14);
			this.ultraLabel2.TabIndex = 55;
			this.ultraLabel2.Text = "Mora:";
			// 
			// ultraLabel13
			// 
			this.ultraLabel13.AutoSize = true;
			this.ultraLabel13.Location = new System.Drawing.Point(224, 72);
			this.ultraLabel13.Name = "ultraLabel13";
			this.ultraLabel13.Size = new System.Drawing.Size(51, 14);
			this.ultraLabel13.TabIndex = 57;
			this.ultraLabel13.Text = "Pronto P.";
			// 
			// ultraLabel11
			// 
			this.ultraLabel11.AutoSize = true;
			this.ultraLabel11.Location = new System.Drawing.Point(224, 8);
			this.ultraLabel11.Name = "ultraLabel11";
			this.ultraLabel11.Size = new System.Drawing.Size(46, 14);
			this.ultraLabel11.TabIndex = 59;
			this.ultraLabel11.Text = "Amortiz.";
			// 
			// txtPago
			// 
			this.txtPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.Pago"));
			this.txtPago.FormatString = "#,##0.00";
			this.txtPago.Location = new System.Drawing.Point(96, 8);
			this.txtPago.Name = "txtPago";
			this.txtPago.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPago.PromptChar = ' ';
			this.txtPago.Size = new System.Drawing.Size(112, 21);
			this.txtPago.TabIndex = 60;
			// 
			// txtAmortizacion
			// 
			this.txtAmortizacion.FormatString = "#,##0.00";
			this.txtAmortizacion.Location = new System.Drawing.Point(280, 8);
			this.txtAmortizacion.Name = "txtAmortizacion";
			this.txtAmortizacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAmortizacion.PromptChar = ' ';
			this.txtAmortizacion.Size = new System.Drawing.Size(96, 21);
			this.txtAmortizacion.TabIndex = 61;
			// 
			// txtMora
			// 
			this.txtMora.FormatString = "#,##0.00";
			this.txtMora.Location = new System.Drawing.Point(280, 40);
			this.txtMora.Name = "txtMora";
			this.txtMora.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMora.PromptChar = ' ';
			this.txtMora.Size = new System.Drawing.Size(96, 21);
			this.txtMora.TabIndex = 62;
			// 
			// txtProntoPago
			// 
			this.txtProntoPago.FormatString = "#,##0.00";
			this.txtProntoPago.Location = new System.Drawing.Point(280, 72);
			this.txtProntoPago.Name = "txtProntoPago";
			this.txtProntoPago.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProntoPago.PromptChar = ' ';
			this.txtProntoPago.Size = new System.Drawing.Size(96, 21);
			this.txtProntoPago.TabIndex = 63;
			// 
			// txtTotal
			// 
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(96, 104);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(112, 21);
			this.txtTotal.TabIndex = 64;
			// 
			// txtInteres
			// 
			this.txtInteres.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagos, "Pago.Interes"));
			this.txtInteres.FormatString = "#,##0.00";
			this.txtInteres.Location = new System.Drawing.Point(96, 72);
			this.txtInteres.Name = "txtInteres";
			this.txtInteres.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtInteres.PromptChar = ' ';
			this.txtInteres.Size = new System.Drawing.Size(112, 21);
			this.txtInteres.TabIndex = 65;
			// 
			// PagoInt
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(400, 373);
			this.Controls.Add(this.txtInteres);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtProntoPago);
			this.Controls.Add(this.txtMora);
			this.Controls.Add(this.txtAmortizacion);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.ultraLabel11);
			this.Controls.Add(this.ultraLabel13);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.btCapital);
			this.Controls.Add(this.txtIdPago);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.cmbTipoPago);
			this.Controls.Add(this.btTotal);
			this.Controls.Add(this.btInteres);
			this.Controls.Add(this.ultraLabel14);
			this.Controls.Add(this.cmbCentroCosto);
			this.Controls.Add(this.txtReferencia);
			this.Controls.Add(this.ultraLabel12);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.txtAsiento);
			this.Controls.Add(this.ultraLabel10);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.ultraLabel9);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtCheque);
			this.Controls.Add(this.ultraLabel8);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.txtNota);
			this.Controls.Add(this.txtConcepto);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.cmbCuentaCorr);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.ultraLabel1);
			this.MaximizeBox = false;
			this.Name = "PagoInt";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Pago con interés";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Pago_Closing);
			this.Load += new System.EventHandler(this.Pago1_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipoPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAmortizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMora)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProntoPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsPagos_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPagos.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		C1.Data.C1DataRow drPago;

		private void cdsPagos_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
//			drPago = e.Row;		
		}

		private void ValoresIniciales()
		{
			drPago["idCompra"] = IdCompra;
			drPago["Fecha"] = DateTime.Today;
			drPago["Vencimiento"] = DateTime.Today;
			drPago["Pago"] = nSaldo;

			txtTotal.Value = nSaldo;
			Interes();
			drPago["Pago"] = nSaldo;
			txtTotal.Value = nSaldo + (double) drPago["Interes"]; 
			this.txtTotal.Select();
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea eliminar éste registro?", 
        "Confirmación", MessageBoxButtons.YesNo,
				MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
				== DialogResult.No) return;
			string stExec = "Exec PagoBorrar " + txtIdPago.Value.ToString();
			Funcion.EjecutaSQL(cdsPagos, stExec, true);
			MessageBox.Show("Registro Borrado", "Información");
			Close();
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			MessageBox.Show("Pagos los crea desde la pantalla anterior", "Información");
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			BorrarErrores();
			this.barraDato1.DeshacerRegistro();
		}

		private double nPagoAntes=0;
		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			MessageBox.Show("No se puede editar con interes", "Información");
//			try
//			{
//				drPago = cdsPagos.TableViews["Pago"].Rows[0];
//			}
//			catch
//			{
//				barraDato1.ProximoId(1);
//				drPago = cdsPagos.TableViews["Pago"].Rows[0];
//			}
//			nPagoAntes = double.Parse(this.txtPago.Value.ToString());
//			this.barraDato1.EditarRegistro();
//			chkAsiento.Enabled = false;
		}

		private bool Verifica()
		{
			bool bOk = true;
			this.BorrarErrores();
			if (double.Parse(this.txtPago.Value.ToString()) == 0)
			{
				this.errorProvider.SetError(this.txtPago, "Ingrese Pago");
				this.txtPago.Select();
				bOk = false;
			}

//			if (!chkPendiente.Checked)
			{
				if (this.cmbCuentaCorr.Value == System.DBNull.Value)
				{
					this.errorProvider.SetError(this.cmbCuentaCorr, "Ingrese Cuenta Corriente");
					this.cmbCuentaCorr.Select();
					bOk = false;
				}
				if (this.cmbFormaPago.Value == System.DBNull.Value)
				{
					this.errorProvider.SetError(this.cmbFormaPago, "Ingrese Forma de Pago");
					this.cmbFormaPago.Select();
					bOk = false;
				}
			}
			return bOk;
		}

		private void BorrarErrores()
		{
			this.errorProvider.SetError(this.txtPago, "");
			this.errorProvider.SetError(this.cmbCuentaCorr, "");
			this.errorProvider.SetError(this.cmbFormaPago, "");	
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			double dPago = double.Parse(this.txtPago.Value.ToString());
			if (drSeteoF["PagoNegativo"] == System.DBNull.Value ||
				(bool) drSeteoF["PagoNegativo"] == false)
			{
				//if (Math.Round(dPago, 2) > Math.Round(nSaldo, 2) + Math.Round(nPagoAntes, 2))
				if (dPago - nSaldo - nPagoAntes >= 0.01)
					{
					MessageBox.Show("Pago sobrepasa saldo " + nSaldo.ToString("#,##0.00"));
					this.txtPago.Focus();
					return;
				}
			}

			NumeroPago();
			if (!barraDato1.bNuevo)
			{
				string stSelect = "Select GetDate()";
				DateTime dtFecha = Funcion.dtEscalarSQL(cdsPagos, stSelect);
				drPago["FechaEditado"] = dtFecha;
			}
			this.barraDato1.GrabaRegistro();
			double dPagoA = double.Parse(txtIdPago.Value.ToString());
			string stExec = "Exec AmortizacionPago " + txtIdCompra.Value.ToString()
				+ ", " + dPagoA.ToString("0.00", us);
			Funcion.EjecutaSQL(cdsCompraTabla, stExec, true);
		}

		private void NumeroPago()
		{
			bool bOk = false;
			int iIngEgr = int.Parse(txtIdCompra.Value.ToString());
			string stSelect = "Select idTipoFactura From Compra Where idCompra = "
				+ iIngEgr.ToString();
			iIngEgr = Funcion.iEscalarSQL(cdsCompraTabla, stSelect, false);

			if (!barraDato1.bNuevo && txtCheque.Text.Trim().Length == 0 && iIngEgr == 4) // viene de compra
				bOk = true;
			if (barraDato1.bNuevo && iIngEgr == 4) // viene de compra
				bOk = true;

			if (txtNumero.Text.Trim().Length == 0 )
			{
				int iTipo = 0;
				if (iIngEgr != 1) iTipo = 1;
				string stNumero = Funcion.NumeraFactura(this.cdsCompraTabla, 21 + iTipo);
				if (stNumero.Trim().Length > 0) drPago["Numero"] = stNumero;
			}
			if (!bOk) return;

			int iNumCheque = Contabilidad.NumeroCheque(
				cdsCuentaCorr, (int) cmbCuentaCorr.Value, true);
			if (iNumCheque != -1)	drPago["Cheque"] = iNumCheque.ToString();
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
//			string stFitlro = "";
//			Reporte miRep = new Reporte("", stFiltro);
//			miRep.MdiParent = this.MdiParent;
//			miRep.Show();
		}

		private void BuscarPago(int idPago)
		{
			if (idPago > 0) 
			{
				barraDato1.IdRegistro = idPago;
				this.barraDato1.ProximoId(5);
			}
		}

		public bool bCobro = true;
		C1.Data.C1DataRow drSeteoF;
		private void Pago1_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Pago'";
			cmbTipoPago.Value = 2;
			Funcion.EjecutaSQL(cdsCompraTabla, stAudita, true);
			Cursor = Cursors.Default;
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			if (!bCobro) this.Text = "Cobro";
			txtIdCompra.Width = 0;
			txtIdPago.Width = 0;
			this.barraDato1.HabilitaControles(false);
			txtPago.Enabled = false;
			if (bNuevo) 
			{
				this.barraDato1.CrearRegistro();
				this.ValoresIniciales();
			}
			else
			{
				BuscarPago(IdCompra);
			}
		}

		private void cdsCompraTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompraTabla.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);			
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drPago = cdsPagos.TableViews["Pago"].Rows[0];
			}
			catch{};
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void Pago_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Pago'";
			Funcion.EjecutaSQL(cdsCompraTabla, stAudita, true);
		}

		CultureInfo us = new CultureInfo("en-US");
		private void Interes()
		{
			DateTime dtFecha = (DateTime) txtFecha.Value;
			int iTipo = (int) cmbTipoPago.Value;
			double dTotal = 0;
			double dCapitalP = 0;
			if (txtTotal.Text.Trim().Length > 0)
        dTotal = double.Parse(txtTotal.Text);
			if (txtPago.Text.Trim().Length > 0)
				dCapitalP = double.Parse(txtPago.Text);
			string stSelect = "Exec AmortizacionInteres "
				+ txtIdCompra.Value.ToString() + ", '" 
				+ dtFecha.ToString("yyyyMMdd") + "', "
				+ dTotal.ToString("0.00", us);

			double dInteres = 0;//Funcion.dEscalarSQL(cdsCompraTabla, stSelect, true);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsCompraTabla, stSelect, true);
			while (dr.Read())
			{
				double dMora = dr.GetDouble(0);
				double dProntoP = dr.GetDouble(1);
				double dAmortizacion = dr.GetDouble(2);
				double dCapital = dr.GetDouble(3);

				txtMora.Value = dMora;
				txtAmortizacion.Value = dAmortizacion;
				txtProntoPago.Value = dProntoP;
//				if (iTipo == 3)
//				{
//					dInteres = dMora + dAmortizacion - dProntoP;
//					drPago["Pago"] = dCapital;
//				}
//				else
				{
					dInteres = dMora + dAmortizacion - dProntoP;
					drPago["Pago"] = dCapital;
					if (dInteres + dCapital < dTotal)
						txtTotal.Value = dInteres + dCapital;
				}
			}
			dr.Close();
			drPago["Interes"] = dInteres.ToString("0.00", us);
		}
		private void btInteres_Click(object sender, System.EventArgs e)
		{
			Interes();
		}

		private void btTotal_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime) txtFecha.Value;
			int iTipo = (int) cmbTipoPago.Value;
			if (iTipo == 1 || iTipo == 2)
			{
				string stSelect = "Exec AmortizacionCapital "
					+ txtIdCompra.Value.ToString() + ", '" 
					+ dtFecha.ToString("yyyyMMdd") + "', "
					+ cmbTipoPago.Value.ToString();
				double dCapital = Funcion.dEscalarSQL(cdsCompraTabla, stSelect, true);
				if (iTipo == 2) txtTotal.Value = "9999999";
				else 
				{
					txtTotal.Value = dCapital.ToString("0.00", us);
					Interes();
					dCapital += double.Parse(txtInteres.Value.ToString());
					txtTotal.Value = dCapital.ToString("0.00", us);
				}
			}

			Interes();
		}

		private void btCapital_Click(object sender, System.EventArgs e)
		{
			double dTotal = 0;
			if (txtTotal.Text.Trim().Length > 0) 
				dTotal =	double.Parse(txtTotal.Value.ToString());
			double dInteres = 0;
			if (txtInteres.Text.Trim().Length > 0) 
				dInteres =	double.Parse(txtInteres.Value.ToString());
			drPago["Pago"] = dTotal - dInteres;
		}
	}
}
