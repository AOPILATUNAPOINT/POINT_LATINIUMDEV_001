using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for DetalleFormaPago.
	/// </summary>
	public class DetalleFormaPago : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsDetFormaPago;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label11;
		private C1.Data.C1DataSet cdsBanco;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTarjeta;
		private C1.Data.C1DataSet cdsCompraTabla;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label1;
		
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuentaCheque;
		private System.Windows.Forms.ErrorProvider errorProvider1;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.ComboBox cmbPlan;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumVoucher;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutorizacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnSaldo;
		private C1.Data.C1DataSet cdsFormaPago;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private bool bNuevo = true;
		
		private int IdFP = 0;
		private double Valor = 0;
		private double DSaldo = 0;
		private string sDocNC;
		private int iCompra = 0;
		private double dSaldoFac = 0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private DevExpress.XtraEditors.RadioGroup rdgCorrDiferido;
		private DevExpress.XtraEditors.RadioGroup rdgInteres;
		private Infragistics.Win.Misc.UltraButton btTarjetaDeCredito;
		public System.Windows.Forms.Button btGrabar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnValor;
		
		public DetalleFormaPago(int idFP, double valor, double dSaldo, string sDoc, int iIdCompra)
		{
			InitializeComponent();
			IdFP = idFP;
			Valor = valor;
			DSaldo = dSaldo;
			sDocNC = sDoc;
			iCompra = iIdCompra;
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("TarjetasDeCredito", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hoja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProcesoOk");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraFormaPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsCompraTabla = new C1.Data.C1DataSet();
			this.cmbTarjeta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label15 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsDetFormaPago = new C1.Data.C1DataSet();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsBanco = new C1.Data.C1DataSet();
			this.txtCuentaCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.errorProvider1 = new System.Windows.Forms.ErrorProvider();
			this.label16 = new System.Windows.Forms.Label();
			this.spnSaldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.cmbPlan = new System.Windows.Forms.ComboBox();
			this.label7 = new System.Windows.Forms.Label();
			this.txtNumVoucher = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtAutorizacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsFormaPago = new C1.Data.C1DataSet();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.spnValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.rdgCorrDiferido = new DevExpress.XtraEditors.RadioGroup();
			this.rdgInteres = new DevExpress.XtraEditors.RadioGroup();
			this.btTarjetaDeCredito = new Infragistics.Win.Misc.UltraButton();
			this.btGrabar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDetFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentaCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnSaldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumVoucher)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgCorrDiferido.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgInteres.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(20, 112);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(89, 16);
			this.label6.TabIndex = 32;
			this.label6.Text = "Fecha Caducidad";
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MMM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(124, 112);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 31;
			this.dtFecha.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			// 
			// cdsCompraTabla
			// 
			this.cdsCompraTabla.BindingContextCtrl = this;
			this.cdsCompraTabla.DataLibrary = "LibFacturacion";
			this.cdsCompraTabla.DataLibraryUrl = "";
			this.cdsCompraTabla.DataSetDef = "dsCompraTabla";
			this.cdsCompraTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompraTabla.Name = "cdsCompraTabla";
			this.cdsCompraTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraTabla.SchemaDef = null;
			this.cdsCompraTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompraTabla_BeforeFill);
			// 
			// cmbTarjeta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTarjeta.Appearance = appearance2;
			this.cmbTarjeta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTarjeta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTarjeta.DataMember = "TarjetasDeCredito";
			this.cmbTarjeta.DataSource = this.cdsCompraTabla;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 149;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTarjeta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTarjeta.DisplayMember = "Nombre";
			this.cmbTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTarjeta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTarjeta.Enabled = false;
			this.cmbTarjeta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTarjeta.Location = new System.Drawing.Point(372, 16);
			this.cmbTarjeta.Name = "cmbTarjeta";
			this.cmbTarjeta.Size = new System.Drawing.Size(149, 21);
			this.cmbTarjeta.TabIndex = 92;
			this.cmbTarjeta.ValueMember = "idTarjeta";
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(268, 16);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(93, 16);
			this.label15.TabIndex = 91;
			this.label15.Text = "Tarjeta de crédito";
			// 
			// txtNumero
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance3;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(108, 80);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(152, 22);
			this.txtNumero.TabIndex = 114;
			// 
			// cdsDetFormaPago
			// 
			this.cdsDetFormaPago.BindingContextCtrl = this;
			this.cdsDetFormaPago.DataLibrary = "LibFacturacion";
			this.cdsDetFormaPago.DataLibraryUrl = "";
			this.cdsDetFormaPago.DataSetDef = "dsCompraTabla";
			this.cdsDetFormaPago.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsDetFormaPago.Name = "cdsDetFormaPago";
			this.cdsDetFormaPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsDetFormaPago.SchemaDef = null;
			this.cdsDetFormaPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsDetFormaPago_BeforeFill);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(20, 80);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(44, 16);
			this.label11.TabIndex = 113;
			this.label11.Text = "Número";
			// 
			// cmbBanco
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance4;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataMember = "Banco";
			this.cmbBanco.DataSource = this.cdsBanco;
			ultraGridBand2.AddButtonCaption = "Compra";
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 150;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Enabled = false;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(108, 48);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(149, 21);
			this.cmbBanco.TabIndex = 119;
			this.cmbBanco.ValueMember = "idBanco";
			// 
			// cdsBanco
			// 
			this.cdsBanco.BindingContextCtrl = this;
			this.cdsBanco.DataLibrary = "LibContabilidad";
			this.cdsBanco.DataLibraryUrl = "";
			this.cdsBanco.DataSetDef = "dsCuentaCorriente";
			this.cdsBanco.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsBanco.Name = "cdsBanco";
			this.cdsBanco.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsBanco.SchemaDef = null;
			this.cdsBanco.BeforeFill += new C1.Data.FillEventHandler(this.cdsBanco_BeforeFill);
			// 
			// txtCuentaCheque
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuentaCheque.Appearance = appearance5;
			this.txtCuentaCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuentaCheque.Enabled = false;
			this.txtCuentaCheque.Location = new System.Drawing.Point(372, 48);
			this.txtCuentaCheque.Name = "txtCuentaCheque";
			this.txtCuentaCheque.Size = new System.Drawing.Size(152, 22);
			this.txtCuentaCheque.TabIndex = 122;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(268, 48);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(89, 16);
			this.label9.TabIndex = 128;
			this.label9.Text = "Cuenta Corriente";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(20, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 125;
			this.label1.Text = "Banco";
			// 
			// errorProvider1
			// 
			this.errorProvider1.ContainerControl = this;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(212, 248);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(31, 16);
			this.label16.TabIndex = 132;
			this.label16.Text = "Saldo";
			// 
			// spnSaldo
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnSaldo.Appearance = appearance6;
			this.spnSaldo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnSaldo.Enabled = false;
			this.spnSaldo.FormatString = "#,##0.00";
			this.spnSaldo.Location = new System.Drawing.Point(252, 248);
			this.spnSaldo.MinValue = 0;
			this.spnSaldo.Name = "spnSaldo";
			this.spnSaldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnSaldo.PromptChar = ' ';
			this.spnSaldo.ReadOnly = true;
			this.spnSaldo.Size = new System.Drawing.Size(120, 22);
			this.spnSaldo.TabIndex = 131;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(20, 16);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(80, 16);
			this.label10.TabIndex = 122;
			this.label10.Text = "Forma de pago";
			// 
			// cmbPlan
			// 
			this.cmbPlan.Enabled = false;
			this.cmbPlan.Location = new System.Drawing.Point(196, 176);
			this.cmbPlan.Name = "cmbPlan";
			this.cmbPlan.Size = new System.Drawing.Size(121, 21);
			this.cmbPlan.TabIndex = 137;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(156, 176);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(25, 16);
			this.label7.TabIndex = 136;
			this.label7.Text = "Plan";
			// 
			// txtNumVoucher
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumVoucher.Appearance = appearance7;
			this.txtNumVoucher.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumVoucher.Enabled = false;
			this.txtNumVoucher.Location = new System.Drawing.Point(108, 144);
			this.txtNumVoucher.Name = "txtNumVoucher";
			this.txtNumVoucher.Size = new System.Drawing.Size(152, 22);
			this.txtNumVoucher.TabIndex = 130;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(268, 144);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(65, 16);
			this.label4.TabIndex = 131;
			this.label4.Text = "Autorización";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(20, 144);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(72, 16);
			this.label3.TabIndex = 129;
			this.label3.Text = "Num Voucher";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(268, 80);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(30, 16);
			this.label5.TabIndex = 133;
			this.label5.Text = "Valor";
			// 
			// txtAutorizacion
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutorizacion.Appearance = appearance8;
			this.txtAutorizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutorizacion.Enabled = false;
			this.txtAutorizacion.Location = new System.Drawing.Point(372, 144);
			this.txtAutorizacion.Name = "txtAutorizacion";
			this.txtAutorizacion.Size = new System.Drawing.Size(152, 22);
			this.txtAutorizacion.TabIndex = 132;
			// 
			// cdsFormaPago
			// 
			this.cdsFormaPago.BindingContextCtrl = this;
			this.cdsFormaPago.DataLibrary = "LibFacturacion";
			this.cdsFormaPago.DataLibraryUrl = "";
			this.cdsFormaPago.DataSetDef = "dsCompraTabla";
			this.cdsFormaPago.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsFormaPago.Name = "cdsFormaPago";
			this.cdsFormaPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFormaPago.SchemaDef = null;
			this.cdsFormaPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsFormaPago_BeforeFill);
			// 
			// cmbFormaPago
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance9;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataMember = "CompraFormaPago";
			this.cmbFormaPago.DataSource = this.cdsFormaPago;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Width = 149;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 6;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(108, 16);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(149, 21);
			this.cmbFormaPago.TabIndex = 141;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Validating += new System.ComponentModel.CancelEventHandler(this.cmbFormaPago_Validating);
			this.cmbFormaPago.ValueChanged += new System.EventHandler(this.cmbFormaPago_ValueChanged);
			this.cmbFormaPago.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbFormaPago_InitializeLayout);
			// 
			// spnValor
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnValor.Appearance = appearance10;
			this.spnValor.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnValor.Enabled = false;
			this.spnValor.FormatString = "#,##0.00";
			this.spnValor.Location = new System.Drawing.Point(372, 80);
			this.spnValor.MaskInput = "nnn";
			this.spnValor.MinValue = 0;
			this.spnValor.Name = "spnValor";
			this.spnValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnValor.PromptChar = ' ';
			this.spnValor.Size = new System.Drawing.Size(152, 22);
			this.spnValor.TabIndex = 144;
			this.spnValor.ValueChanged += new System.EventHandler(this.spnValor_ValueChanged);
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.FormatString = "";
			this.txtIdCompra.Location = new System.Drawing.Point(460, 256);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(24, 22);
			this.txtIdCompra.TabIndex = 146;
			// 
			// rdgCorrDiferido
			// 
			this.rdgCorrDiferido.Enabled = false;
			this.rdgCorrDiferido.Location = new System.Drawing.Point(20, 168);
			this.rdgCorrDiferido.Name = "rdgCorrDiferido";
			// 
			// rdgCorrDiferido.Properties
			// 
			this.rdgCorrDiferido.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Corriente"),
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Diferido")});
			this.rdgCorrDiferido.Size = new System.Drawing.Size(88, 72);
			this.rdgCorrDiferido.TabIndex = 147;
			this.rdgCorrDiferido.SelectedIndexChanged += new System.EventHandler(this.rdgCorrDiferido_SelectedIndexChanged);
			// 
			// rdgInteres
			// 
			this.rdgInteres.Enabled = false;
			this.rdgInteres.Location = new System.Drawing.Point(412, 168);
			this.rdgInteres.Name = "rdgInteres";
			// 
			// rdgInteres.Properties
			// 
			this.rdgInteres.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			 new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Con Intereses"),
																																																			 new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Sin Intereses")});
			this.rdgInteres.Size = new System.Drawing.Size(96, 72);
			this.rdgInteres.TabIndex = 148;
			// 
			// btTarjetaDeCredito
			// 
			this.btTarjetaDeCredito.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btTarjetaDeCredito.Enabled = false;
			this.btTarjetaDeCredito.Location = new System.Drawing.Point(532, 16);
			this.btTarjetaDeCredito.Name = "btTarjetaDeCredito";
			this.btTarjetaDeCredito.Size = new System.Drawing.Size(24, 24);
			this.btTarjetaDeCredito.TabIndex = 149;
			this.btTarjetaDeCredito.Text = "...";
			this.btTarjetaDeCredito.Click += new System.EventHandler(this.btTarjetaDeCredito_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btGrabar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btGrabar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btGrabar.Location = new System.Drawing.Point(20, 248);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(176, 23);
			this.btGrabar.TabIndex = 145;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// DetalleFormaPago
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(576, 286);
			this.Controls.Add(this.btTarjetaDeCredito);
			this.Controls.Add(this.rdgInteres);
			this.Controls.Add(this.rdgCorrDiferido);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.spnValor);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.cmbPlan);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtNumVoucher);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtAutorizacion);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.cmbTarjeta);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.txtCuentaCheque);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.spnSaldo);
			this.Controls.Add(this.label16);
			this.Name = "DetalleFormaPago";
			this.Text = "Detalle forma de pago";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.DetalleFormaPago_Closing);
			this.Load += new System.EventHandler(this.DetalleFormaPago_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDetFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentaCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnSaldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumVoucher)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgCorrDiferido.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgInteres.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		public void DetalleFormaPago_Load(object sender, System.EventArgs e)
		{
			this.txtIdCompra.Width = 0;
			this.txtIdCompra.Value = int.Parse(iCompra.ToString());			

			this.spnValor.Value = Valor;
			dSaldoFac = Valor;

			if (IdFP == 1)
			{
				this.cmbFormaPago.Value = 1;
			}

			if (IdFP == 2)
			{
				cmbFormaPago.Value = 2;

				this.cmbBanco.Enabled = true;
				this.txtCuentaCheque.Enabled = true;
				this.txtNumero.Enabled = true;
				this.spnValor.Enabled = true;
				this.cmbTarjeta.Enabled = false;
				this.dtFecha.Enabled = false;
				this.txtNumVoucher.Enabled = false;
				this.txtAutorizacion.Enabled = false;
				this.rdgCorrDiferido.Enabled = false;
				this.cmbPlan.Enabled = false;
				this.rdgInteres.Enabled = false;
				this.cmbBanco.Select();
			}

			if (IdFP == 3)
			{
				cmbFormaPago.Value = 3;		
				this.cmbTarjeta.Enabled = true;
				this.cmbBanco.Enabled = true;
				this.txtNumero.Enabled = true;
				this.spnValor.Enabled = true;
				this.dtFecha.Enabled = true;
				this.txtNumVoucher.Enabled = true;
				this.txtAutorizacion.Enabled = true;
				this.rdgCorrDiferido.Enabled = true;
				this.cmbPlan.Enabled = false;
				this.rdgInteres.Enabled = false;

				this.cmbPlan.Items.Add("3 Meses");
				this.cmbPlan.Items.Add("6 Meses");
				this.cmbPlan.Items.Add("9 Meses");
				this.cmbPlan.Items.Add("12 Meses");
				this.cmbPlan.Items.Add("18 Meses");
				this.cmbPlan.Items.Add("24 Meses");
				this.cmbTarjeta.Select();
			}

			if (IdFP == 4)
			{
				this.cmbFormaPago.Value = 4;
				this.txtNumero.Text = sDocNC;
				this.spnSaldo.Value = DSaldo;
			}

			if (IdFP == 5)
			{
				this.cmbFormaPago.Value = 1;
				this.cmbFormaPago.Enabled = true;
				this.spnValor.Enabled = true;
				this.cmbFormaPago.Select();
			}
		}		

		private void cdsCompraTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompraTabla.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirNumero);
		}

		private bool Valida()
		{
			bool bOk = true;
			if ((int) this.cmbFormaPago.Value == 2)
			{				
				if (this.cmbBanco.Text == "")
				{
					this.errorProvider1.SetError(this.cmbBanco, "Seleccione el banco");
					bOk = false;
				}
				if (this.txtCuentaCheque.Text == "")
				{
					this.errorProvider1.SetError(this.txtCuentaCheque, "Ingrese el número de cuenta");
					bOk = false;
				}
				if (this.txtNumero.Text == "")
				{
					this.errorProvider1.SetError(this.txtNumero, "Ingrese el número de cheque");
					bOk = false;
				}
			}

			if ((int) this.cmbFormaPago.Value == 3)
			{
				if (this.cmbTarjeta.Text == "")
				{
					this.errorProvider1.SetError(this.cmbTarjeta, "Seleccione la tarjeta de crédito");
					bOk = false;
				}
				if (this.cmbBanco.Text == "")
				{
					this.errorProvider1.SetError(this.cmbBanco, "Seleccione el banco");
					bOk = false;
				}
				DateTime Fecha = (DateTime) this.dtFecha.Value;
				if (Fecha.Date.CompareTo(DateTime.Now)==1)
				{					
					this.errorProvider1.SetError(this.dtFecha, "La Tarjeta esta caducada");
					bOk = false;
				}
				if (this.txtNumero.Text == "")
				{
					this.errorProvider1.SetError(this.txtNumero, "Ingrese el número de tarjeta");
					bOk = false;
				}
				if (this.txtNumVoucher.Text == "")
				{
					this.errorProvider1.SetError(this.txtNumVoucher, "Ingrese el número de voucher");
					bOk = false;
				}
				if (this.txtAutorizacion.Text == "")
				{
					this.errorProvider1.SetError(this.txtAutorizacion, "Ingrese la autorización del voucher");
					bOk = false;
				}
			}
			return bOk;
		}

		private void rdgCorrienteDiferido_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void cdsDetFormaPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsDetFormaPago.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirNumero);
		}

		private void cdsFormaPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsFormaPago.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirNumero);
		}

		private void cdsBanco_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsBanco.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cmbFormaPago_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
		}

		public void btGrabar_Click(object sender, System.EventArgs e)
		{
			if (!Valida()) return;

			int idTarjeta = 0;
			int idBanco = 0;
			int Corriente = 0;
			int Interes = 0;
			string Plan = "";

			if ((int) this.cmbFormaPago.Value == 2)
			{
				idTarjeta = 0;
				Corriente = 0;
				Interes = 0;
				idBanco = (int) this.cmbBanco.Value;
			}
			else if ((int) this.cmbFormaPago.Value == 3)
			{
				idTarjeta = (int) this.cmbFormaPago.Value;
				idBanco = (int) this.cmbBanco.Value;
				Plan = this.cmbPlan.Text.ToString();
				Corriente = (int) this.rdgCorrDiferido.EditValue;
				Interes = (int) this.rdgInteres.EditValue;
			}
			
			DateTime dtFec = (DateTime) this.dtFecha.Value;
			int idCompra = (int) this.txtIdCompra.Value;
			int idFormaPago = (int) this.cmbFormaPago.Value;			
			
			double Valor = (double) this.spnValor.Value;

			string stSelect = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, '{4}', '{5}', {6}, '{7}', '{8}', '{9}', {10}, '{11}', {12}",
				idCompra, idFormaPago, idTarjeta, idBanco, this.txtCuentaCheque.Text.ToString(),
				this.txtNumero.Text, Valor,  dtFec.ToString("yyyyMMdd"), this.txtNumVoucher.Text.ToString(), 
				this.txtAutorizacion.Text.ToString(), Corriente, Plan, Interes);
Funcion.EjecutaSQL(cdsFormaPago, stSelect);
			dSaldoFac = dSaldoFac - (double) this.spnValor.Value;
			if ((int) IdFP != 5)
			{
				if ((double) dSaldoFac > 0.0)
				{
					if (DialogResult.Yes == MessageBox.Show(string.Format("Hay una diferencia de {0}. ¿Grabar este registro para grabar la diferencia con otro(a) '{1}'?.", dSaldoFac, this.cmbFormaPago.Text), "Información",
						MessageBoxButtons.YesNo, MessageBoxIcon.Question))
					{
						Funcion.EjecutaSQL(cdsFormaPago, stSelect);
						this.spnValor.Value = dSaldoFac;
					}
					else
					{
						Funcion.EjecutaSQL(cdsFormaPago, stSelect);
						Close();
					}
				}
				else
				{
					Funcion.EjecutaSQL(cdsFormaPago, stSelect);
					Close();
				}
			}
			else
			{
				if ((double) dSaldoFac > 0.0)
				{
					Funcion.EjecutaSQL(cdsFormaPago, stSelect);
					this.spnValor.Value = dSaldoFac;
				}
				else
				{
					Funcion.EjecutaSQL(cdsFormaPago, stSelect);
					Close();
				}
			}
		}
		
		private void rdgCorrDiferido_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if ((int) this.rdgCorrDiferido.EditValue == 1)
			{
				this.cmbPlan.Enabled = false;
				this.cmbPlan.Text = "";
				this.rdgInteres.Enabled = false;
				this.rdgInteres.EditValue = 2;
			}
			else if ((int) this.rdgCorrDiferido.EditValue == 2)
			{
				this.cmbPlan.Enabled = true;
				this.rdgInteres.Enabled = true;
				this.rdgInteres.EditValue = 1;
			}		
		}

		private void DetalleFormaPago_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
		//	e.Cancel = true;
		}

		private void cmbFormaPago_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbFormaPago_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int) this.cmbFormaPago.Value == 2)
			{
				cmbFormaPago.Value = 2;
				this.cmbBanco.Enabled = true;
				this.txtCuentaCheque.Enabled = true;
				this.txtNumero.Enabled = true;
				this.spnValor.Enabled = true;
				this.cmbTarjeta.Enabled = false;
				this.btTarjetaDeCredito.Enabled = false;
				this.dtFecha.Enabled = false;
				this.txtNumVoucher.Enabled = false;
				this.txtAutorizacion.Enabled = false;
				this.rdgCorrDiferido.Enabled = false;
				this.cmbPlan.Enabled = false;
				this.rdgInteres.Enabled = false;
				this.cmbBanco.Select();
			}

			if ((int) this.cmbFormaPago.Value == 3)
			{
				cmbFormaPago.Value = 3;		
				this.cmbTarjeta.Enabled = true;
				this.btTarjetaDeCredito.Enabled = true;
				this.cmbBanco.Enabled = true;
				this.txtNumero.Enabled = true;
				this.spnValor.Enabled = true;
				this.dtFecha.Enabled = true;
				this.txtNumVoucher.Enabled = true;
				this.txtAutorizacion.Enabled = true;
				this.rdgCorrDiferido.Enabled = true;
				this.cmbPlan.Enabled = false;
				this.rdgInteres.Enabled = false;

				this.cmbPlan.Items.Add("3 Meses");
				this.cmbPlan.Items.Add("6 Meses");
				this.cmbPlan.Items.Add("9 Meses");
				this.cmbPlan.Items.Add("12 Meses");
				this.cmbPlan.Items.Add("18 Meses");
				this.cmbPlan.Items.Add("24 Meses");
				this.cmbTarjeta.Select();
			}

			if ((int) this.cmbFormaPago.Value == 4)
			{
				
			}

			if ((int) this.cmbFormaPago.Value == 5)
			{
				this.cmbFormaPago.ReadOnly = false;
				this.cmbFormaPago.Select();
			}
		}

		private void btTarjetaDeCredito_Click(object sender, System.EventArgs e)
		{
			using(TarjetasDeCredito miTarjeta = new TarjetasDeCredito())
			{
				miTarjeta.ShowDialog();
			}
		}

		private void spnValor_ValueChanged(object sender, System.EventArgs e)
		{
		
		}		
	}
}

