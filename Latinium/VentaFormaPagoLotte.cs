using System;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Shared;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Diagnostics;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Summary description for VentaFormaPagoLotte.
	/// </summary>
	public class VentaFormaPagoLotte : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnSubTotal;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnSaldo;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor spnTotal;
		private System.Windows.Forms.Label label5;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		bool bEliminaAlValidar = false;
		int IdCompra = 0;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1DataRow drSeteoF;
		decimal  DValor = 0.00m;
		int IFormaPag = 0;
		decimal DEfectivo = 0.00m;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnEfectivo;
		private System.Windows.Forms.Label lblEfectivo;
		decimal DSaldo = 0.00m;
		bool BAnticipo = false;
		private C1.Data.C1DataSet cdsSeteoF;
		decimal DAnticipo = 0.00m;
		bool BNuevo = false;
		string SNumNC = "";
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnRetencion;
		decimal DRetenido = 0.00m;
		int IOrigen = 0;
		int IEstado = 0;
		private System.Windows.Forms.GroupBox groupBox1;
		bool BPagoCredito = false;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btGrabar;
		private System.Windows.Forms.Button btCancelar;
		private System.Windows.Forms.Label lblFecha;
		int IdCliente = 0;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Acceso miAcceso;
		
		public VentaFormaPagoLotte()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		#region Variables Factura
		int IdCompraFact = 0;
		int IdFormaDePago = 0;
		decimal DValorFact = 0.00m;
		string SNumero = "";
		bool BNotaDeCreditoFact = false;
		int IdNotaDeCredito = 0;
		decimal DValorNotaDeCredito = 0.00m;
		bool BRetencionFact = false;
		decimal DValorRetencion = 0.00m;
		bool BAnticipoFact = false;
		decimal DValorAnticipo = 0.00m;
		bool BCredito = false;
		decimal DValorCrediPoint = 0.00m;
		decimal DEntrada = 0.00m;
		int IdOrigen = 0;
		int IdBanco = 0;
		int IdTarjeta = 0;
		int IdPlan = 0;
		int IdPlazo = 0;
		int iBodega = 0;
		Infragistics.Win.UltraWinGrid.UltraGrid uGridVenta;
		#endregion Variables Factura

		public VentaFormaPagoLotte(int idFormaPago, decimal  dValor, int iOrigen, bool bPagoCredito)
		{
			InitializeComponent();
			
			IdFormaDePago = idFormaPago;
			DValorFact = dValor;
			IdOrigen = iOrigen;
			BPagoCredito = bPagoCredito;
		}

		public VentaFormaPagoLotte(int idCompra, int idFormaPago, decimal  dValor, int iOrigen)
		{
			InitializeComponent();

			IdCompraFact = idCompra;
			IdFormaDePago = idFormaPago;
			DValorFact = dValor;
			IdOrigen = iOrigen;
		}

		public VentaFormaPagoLotte(int idCompra, int idFormaPago, decimal dValor, string sNumero, bool bNotaDeCredito, int idNotaDeCredito, decimal  dValorNotaDeCredito, bool bRetencion, decimal  dValorRetencion, bool bAnticipo, decimal  dValorAnticipo, bool bCrediPoint, decimal  dValorCrediPoint, decimal  dEntrada, int idOrigen, int idCliente, int idBanco, int idTarjeta, int idPlan, int idPlazo, int IBodega, Infragistics.Win.UltraWinGrid.UltraGrid UGridVenta)
		{
			InitializeComponent();
			
			IdCompraFact = idCompra;
			IdFormaDePago = idFormaPago;
			DValorFact = dValor;
			SNumero = sNumero;
			BNotaDeCreditoFact = bNotaDeCredito;
			IdNotaDeCredito = idNotaDeCredito;
			DValorNotaDeCredito = dValorNotaDeCredito;
			BRetencionFact = bRetencion;
			DValorRetencion = dValorRetencion;
			BAnticipoFact = bAnticipo;
			DValorAnticipo = dValorAnticipo;
			BCredito = bCrediPoint;
			DValorCrediPoint = dValorCrediPoint;
			DEntrada = dEntrada;
			IdOrigen = idOrigen;
			IdCliente = idCliente;
			IdBanco = idBanco;
			IdTarjeta = idTarjeta;
			IdPlan = idPlan;
			IdPlazo = idPlazo;
			iBodega = IBodega;
			uGridVenta = UGridVenta;
		}

		public VentaFormaPagoLotte(int idCompra, int iFormaPag, decimal  dValor, decimal  dEfectivo, decimal  dSaldo, bool bAnticipo, decimal  dAnticipo, bool bNuevo, string sNumNC, decimal dRetenido, int iOrigen, int iEstado, bool bPagoCredito, int idCliente)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
			IdCompra = idCompra;
			IFormaPag = iFormaPag;
			DValor = dValor;
			DEfectivo = dEfectivo;
			DSaldo = dSaldo;
			BAnticipo = bAnticipo;
			DAnticipo = dAnticipo;
			BNuevo = bNuevo;
			SNumNC = sNumNC;
			DRetenido = dRetenido;
			IdOrigen = iOrigen;
			IEstado = iEstado;
			BPagoCredito = bPagoCredito;
			IdCliente = idCliente;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCaducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumVoucher");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlan");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorUsoTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDocumento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("I", 0);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuentaCorriente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCaducidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumVoucher");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorUsoTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDocumento");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.spnSubTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.spnSaldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.spnTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.spnEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEfectivo = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label3 = new System.Windows.Forms.Label();
			this.spnRetencion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblFecha = new System.Windows.Forms.Label();
			this.btGrabar = new System.Windows.Forms.Button();
			this.btCancelar = new System.Windows.Forms.Button();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnSubTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnSaldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnRetencion)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource1;
			ultraGridBand1.AddButtonCaption = "CompraFormaPago";
			ultraGridBand1.ColHeadersVisible = false;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 186;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(24, 64);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(176, 86);
			this.cmbFormaPago.TabIndex = 159;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// spnSubTotal
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnSubTotal.Appearance = appearance1;
			this.spnSubTotal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnSubTotal.Enabled = false;
			this.spnSubTotal.FormatString = "#,##0.00";
			this.spnSubTotal.Location = new System.Drawing.Point(136, 272);
			this.spnSubTotal.MaskInput = "nnn";
			this.spnSubTotal.Name = "spnSubTotal";
			this.spnSubTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnSubTotal.PromptChar = ' ';
			this.spnSubTotal.Size = new System.Drawing.Size(162, 22);
			this.spnSubTotal.TabIndex = 166;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label2.Location = new System.Drawing.Point(8, 272);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(58, 17);
			this.label2.TabIndex = 165;
			this.label2.Text = "SUBTOTAL";
			// 
			// spnSaldo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnSaldo.Appearance = appearance2;
			this.spnSaldo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnSaldo.Enabled = false;
			this.spnSaldo.FormatString = "#,##0.00";
			this.spnSaldo.Location = new System.Drawing.Point(136, 320);
			this.spnSaldo.MaskInput = "nnn";
			this.spnSaldo.MinValue = -2147483647;
			this.spnSaldo.Name = "spnSaldo";
			this.spnSaldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnSaldo.PromptChar = ' ';
			this.spnSaldo.Size = new System.Drawing.Size(162, 22);
			this.spnSaldo.TabIndex = 164;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label1.Location = new System.Drawing.Point(8, 320);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 17);
			this.label1.TabIndex = 163;
			this.label1.Text = "SALDO";
			// 
			// spnTotal
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnTotal.Appearance = appearance3;
			this.spnTotal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnTotal.Enabled = false;
			this.spnTotal.FormatString = "#,##0.00";
			this.spnTotal.Location = new System.Drawing.Point(136, 344);
			this.spnTotal.MaskInput = "nnn";
			this.spnTotal.MinValue = -2147483647;
			this.spnTotal.Name = "spnTotal";
			this.spnTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnTotal.PromptChar = ' ';
			this.spnTotal.Size = new System.Drawing.Size(162, 22);
			this.spnTotal.TabIndex = 161;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label5.Location = new System.Drawing.Point(8, 344);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(38, 17);
			this.label5.TabIndex = 160;
			this.label5.Text = "TOTAL";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance4.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance4;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 8;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 8;
			ultraGridColumn5.Header.Caption = "Forma de Pago";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn5.Width = 174;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 8;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 8;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 8;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 8;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance5;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.PromptChar = ' ';
			ultraGridColumn10.Width = 73;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 17;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 8;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 8;
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 8;
			ultraGridColumn15.Header.VisiblePosition = 17;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 8;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 8;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 8;
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 8;
			ultraGridColumn19.Header.VisiblePosition = 14;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 8;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 18;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 147;
			ultraGridColumn21.Header.VisiblePosition = 19;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 67;
			ultraGridColumn22.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn22.Header.Caption = "...";
			ultraGridColumn22.Header.VisiblePosition = 8;
			ultraGridColumn22.NullText = "...";
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn22.Width = 52;
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
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(1, 0);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(320, 240);
			this.ultraGrid1.TabIndex = 1;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.ultraGrid1_BeforeRowInsert);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.Validating += new System.ComponentModel.CancelEventHandler(this.ultraGrid1_Validating);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.DateTime);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn20.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn21.DataType = typeof(int);
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
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// spnEfectivo
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnEfectivo.Appearance = appearance11;
			this.spnEfectivo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnEfectivo.Enabled = false;
			this.spnEfectivo.FormatString = "#,##0.00";
			this.spnEfectivo.Location = new System.Drawing.Point(136, 272);
			this.spnEfectivo.MaskInput = "nnn";
			this.spnEfectivo.Name = "spnEfectivo";
			this.spnEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnEfectivo.PromptChar = ' ';
			this.spnEfectivo.Size = new System.Drawing.Size(162, 22);
			this.spnEfectivo.TabIndex = 169;
			this.spnEfectivo.Visible = false;
			// 
			// lblEfectivo
			// 
			this.lblEfectivo.AutoSize = true;
			this.lblEfectivo.BackColor = System.Drawing.Color.Transparent;
			this.lblEfectivo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblEfectivo.Location = new System.Drawing.Point(8, 272);
			this.lblEfectivo.Name = "lblEfectivo";
			this.lblEfectivo.Size = new System.Drawing.Size(55, 17);
			this.lblEfectivo.TabIndex = 168;
			this.lblEfectivo.Text = "EFECTIVO";
			this.lblEfectivo.Visible = false;
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
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label3.Location = new System.Drawing.Point(8, 296);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(80, 17);
			this.label3.TabIndex = 170;
			this.label3.Text = "COM. TARJETA";
			// 
			// spnRetencion
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnRetencion.Appearance = appearance12;
			this.spnRetencion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnRetencion.Enabled = false;
			this.spnRetencion.FormatString = "#,##0.00";
			this.spnRetencion.Location = new System.Drawing.Point(136, 296);
			this.spnRetencion.MaskInput = "nnn";
			this.spnRetencion.Name = "spnRetencion";
			this.spnRetencion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnRetencion.PromptChar = ' ';
			this.spnRetencion.Size = new System.Drawing.Size(162, 22);
			this.spnRetencion.TabIndex = 171;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.lblFecha);
			this.groupBox1.Location = new System.Drawing.Point(0, 248);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(328, 128);
			this.groupBox1.TabIndex = 172;
			this.groupBox1.TabStop = false;
			// 
			// lblFecha
			// 
			this.lblFecha.AutoSize = true;
			this.lblFecha.Location = new System.Drawing.Point(8, 24);
			this.lblFecha.Name = "lblFecha";
			this.lblFecha.Size = new System.Drawing.Size(38, 16);
			this.lblFecha.TabIndex = 169;
			this.lblFecha.Text = "FECHA";
			this.lblFecha.Visible = false;
			// 
			// btGrabar
			// 
			this.btGrabar.Location = new System.Drawing.Point(68, 384);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(91, 23);
			this.btGrabar.TabIndex = 0;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(164, 384);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(91, 23);
			this.btCancelar.TabIndex = 173;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// dtFecha
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance13;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(136, 272);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 170;
			this.dtFecha.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			this.dtFecha.Visible = false;
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			// 
			// VentaFormaPagoLotte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(322, 408);
			this.ControlBox = false;
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.spnRetencion);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblEfectivo);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.spnEfectivo);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.spnSubTotal);
			this.Controls.Add(this.spnSaldo);
			this.Controls.Add(this.spnTotal);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.groupBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "VentaFormaPagoLotte";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Detalle de Cobro";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.VentaFormaPagoLotte_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.VentaFormaPagoLotte_Closing);
			this.Load += new System.EventHandler(this.VentaFormaPagoLotte_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnSubTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnSaldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnRetencion)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void UnloadMe()
		{
			this.DialogResult = DialogResult.OK;
			this.Close();
		}

		bool bLoad = false;
		public void VentaFormaPagoLotte_Load(object sender, System.EventArgs e)
		{
			string stTipo = "0703";
										
			miAcceso = new Acceso(cdsSeteoF, stTipo);

			bLoad = true;
			bool bDeUna = false;
			if (IdFormaDePago == 32 && DValorFact >= 500.00m) bDeUna = true;

			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec FormasDePagoIngresoDeCaja {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}", 
				IdFormaDePago, IdOrigen, BAnticipoFact, BRetencionFact, BNotaDeCreditoFact, BCredito, false, iBodega, false, bDeUna));

			if (!BPagoCredito)
			{
				this.spnTotal.Value = Math.Round(DValorFact, 2);
							
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From DetalleFormaPago Where idCompra = {0} And Origen = {1}", IdCompraFact, IdOrigen)) == 0)
				{	
					#region Anticipos - Retenciones - Notas de Crédito - Bono De Una
					if (IdOrigen == 1)
					{					
						if (BAnticipoFact)
						{
							decimal dVAnticipo = 0.00m;
							if (DValorAnticipo > DValorFact) dVAnticipo = DValorFact;
							else if (DValorAnticipo <= DValorFact) dVAnticipo = DValorAnticipo;
							
							this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
							this.ultraGrid1.ActiveRow.Cells["idDetFormaPago"].Value = 0;
							this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value = 8;
							this.ultraGrid1.ActiveRow.Cells["Valor"].Value = dVAnticipo;
							this.ultraGrid1.ActiveRow.Cells["Origen"].Value = IdOrigen;
							if (BCredito) DEntrada = DEntrada - DValorAnticipo;
						}
						if (BRetencionFact)
						{
							this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
							this.ultraGrid1.ActiveRow.Cells["idDetFormaPago"].Value = 0;
							this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value = 10;
							this.ultraGrid1.ActiveRow.Cells["Valor"].Value = DValorRetencion;
							this.ultraGrid1.ActiveRow.Cells["Origen"].Value = IdOrigen;
						}
						if (BNotaDeCreditoFact)
						{
							this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
							this.ultraGrid1.ActiveRow.Cells["idDetFormaPago"].Value = 0;
							this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value = 4;
							this.ultraGrid1.ActiveRow.Cells["Valor"].Value = DValorNotaDeCredito;
							this.ultraGrid1.ActiveRow.Cells["Origen"].Value = IdOrigen;
							this.ultraGrid1.ActiveRow.Cells["idDocumento"].Value = IdNotaDeCredito;
							IdFormaDePago = 1;
						}
						if (BCredito)
						{
							this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
							this.ultraGrid1.ActiveRow.Cells["idDetFormaPago"].Value = 0;
							this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value = IdFormaDePago;
							this.ultraGrid1.ActiveRow.Cells["Valor"].Value = DValorCrediPoint;
							this.ultraGrid1.ActiveRow.Cells["Origen"].Value = IdOrigen;
						}
						if (bDeUna)
						{	
							this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
							this.ultraGrid1.ActiveRow.Cells["idDetFormaPago"].Value = 0;
							this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value = 33;
							this.ultraGrid1.ActiveRow.Cells["Valor"].Value = 50.00m;
							this.ultraGrid1.ActiveRow.Cells["Origen"].Value = IdOrigen;							
						}
					}
					#endregion Anticipos - Retenciones - Notas de Crédito

					#region Carga fila con Forma de pago
					if (!BCredito)
					{
						decimal dValorFactura = Math.Round(DValorFact - DValorAnticipo - DValorRetencion - DValorNotaDeCredito, 2);
						if (dValorFactura > 0)
						{
							this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
							this.ultraGrid1.ActiveRow.Cells["idDetFormaPago"].Value = 0;
							this.ultraGrid1.ActiveRow.Cells["idBanco"].Value = IdBanco;
							this.ultraGrid1.ActiveRow.Cells["idTarjeta"].Value = IdTarjeta;						
							this.ultraGrid1.ActiveRow.Cells["idPlan"].Value = IdPlan;
							this.ultraGrid1.ActiveRow.Cells["idPlazo"].Value = IdPlazo;
							this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value = IdFormaDePago;
							this.ultraGrid1.ActiveRow.Cells["Valor"].Value = Math.Round(DValorFact - DValorAnticipo - DValorRetencion - DValorNotaDeCredito, 2);
							this.ultraGrid1.ActiveRow.Cells["Origen"].Value = IdOrigen;
						}
						//						this.ultraGrid1.ActiveRow.Cells["FechaCaducidad"].Value = DateTime.Today;
						this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
						this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells[0];
					}
					else if (BCredito)
					{
						if (DEntrada > 0)
						{
							this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
							this.ultraGrid1.ActiveRow.Cells["idDetFormaPago"].Value = 0;
							this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value = 1;
							this.ultraGrid1.ActiveRow.Cells["Valor"].Value = DEntrada;
							this.ultraGrid1.ActiveRow.Cells["Origen"].Value = IdOrigen;
							//				this.ultraGrid1.ActiveRow.Cells["idBanco"].Value = 0;
							//				this.ultraGrid1.ActiveRow.Cells["idTarjeta"].Value = 0;
							//				this.ultraGrid1.ActiveRow.Cells["idCuenta"].Value = 0;
							//				this.ultraGrid1.ActiveRow.Cells["idPlan"].Value = 0;
							//				this.ultraGrid1.ActiveRow.Cells["idPlazo"].Value = 0;
							//				this.ultraGrid1.ActiveRow.Cells["FechaCaducidad"].Value = DateTime.Today;
							this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
							this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells[0];
						}
					}
					#endregion Carga fila con Forma de pago

					#region Cheque - Tarjetas de crédito y débito
					//				if (IdFormaDePago == 2 || IdFormaDePago == 3 || IdFormaDePago == 7 || IdFormaDePago == 11)
					//				{
					//					#region Variables
					//					int IdBanco = 0;
					//					int IdTarjeta = 0;
					//					int IdCuenta = 0;
					//					string SNumero = "";
					//					DateTime DTFecha = DateTime.Now;
					//					string SCuentaCorriente = "";
					//					string SVoucher = "";
					//					string SAutorizacion = "";
					//					string SLote = "";
					//					string SReferencia = "";
					//					int IdPlan = 0;
					//					int IdPlazo = 0;
					//					#endregion Variables
					//
					//					using(ChequeTarjeta miCheque = new ChequeTarjeta(IFormaPag, IdBanco, IdTarjeta, IdCuenta, SNumero, DTFecha, SCuentaCorriente, SVoucher, SAutorizacion, SLote, SReferencia, IdPlan, IdPlazo))
					//					{
					//						if (DialogResult.OK == miCheque.ShowDialog())
					//						{
					//							if (IdFormaDePago == 3 || IdFormaDePago == 11) /* TARJETAS DE DEBITO Y CREDITO */
					//							{
					//								this.ultraGrid1.ActiveRow.Cells["idTarjeta"].Value = (int) miCheque.cmbTarjeta.Value;
					//								this.ultraGrid1.ActiveRow.Cells["idBanco"].Value = (int) miCheque.cmbBanco.Value;
					//								this.ultraGrid1.ActiveRow.Cells["Numero"].Value = miCheque.txtNumero.Text.ToString();
					//								this.ultraGrid1.ActiveRow.Cells["FechaCaducidad"].Value = (DateTime) miCheque.dtFecha.Value;
					//								this.ultraGrid1.ActiveRow.Cells["NumVoucher"].Value = miCheque.txtNumVoucher.Text.ToString();
					//								this.ultraGrid1.ActiveRow.Cells["Autorizacion"].Value = miCheque.txtAutorizacion.Text.ToString(); 
					//								this.ultraGrid1.ActiveRow.Cells["idPlan"].Value = (int) miCheque.cmbPlanes.Value;
					//								if (miCheque.cmbPlazo.Text.ToString().Length > 0)	this.ultraGrid1.ActiveRow.Cells["idPlazo"].Value = (int) miCheque.cmbPlazo.Value;
					//								this.ultraGrid1.ActiveRow.Cells["Lote"].Value = miCheque.txtLote.Text.ToString();
					//								this.ultraGrid1.ActiveRow.Cells["Referencia"].Value = miCheque.txtReferencia.Text.ToString();
					//							}
					//							else if (IdFormaDePago == 2) /* CHEQUES */
					//							{								
					//								this.ultraGrid1.ActiveRow.Cells["idBanco"].Value = (int) miCheque.cmbBanco.Value;
					//								this.ultraGrid1.ActiveRow.Cells["CuentaCorriente"].Value = miCheque.txtCuentaCheque.Text.ToString();
					//								this.ultraGrid1.ActiveRow.Cells["Numero"].Value = miCheque.txtNumero.Text.ToString();
					//							}
					//							else if (IdFormaDePago == 7) /* TRANSFERENCIAS */
					//							{								
					//								this.ultraGrid1.ActiveRow.Cells["idCuenta"].Value = (int) miCheque.cmbCuenta.Value;
					//								this.ultraGrid1.ActiveRow.Cells["Numero"].Value = miCheque.txtNumero.Text.ToString();
					//							}
					//						}
					//					}
					//				}
					#endregion Cheque - Tarjetas de crédito y débito					
				}
				else
				{
					string sSQL = string.Format("Exec ConsultaDetalleFormaPago {0}, {1}", IdCompraFact, IdOrigen);
					this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

					#region Anticipos - Retenciones - Notas de Crédito
					if (IdOrigen == 1)
					{					
						if (BAnticipoFact)
						{
							if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleFormaPago Where idCompra = {0} And Origen = {1} And idFormaPago = 8", IdCompraFact, IdOrigen)) == 0)
							{
								decimal dVAnticipo = 0.00m;
								if (DValorAnticipo > DValorFact) dVAnticipo = DValorFact;
								else if (DValorAnticipo <= DValorFact) dVAnticipo = DValorAnticipo;

								this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
								this.ultraGrid1.ActiveRow.Cells["idDetFormaPago"].Value = 0;
								this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value = 8;
								this.ultraGrid1.ActiveRow.Cells["Valor"].Value = DValorAnticipo;
								this.ultraGrid1.ActiveRow.Cells["Origen"].Value = IdOrigen;
								this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
							}
						}
						if (BRetencionFact)
						{
							if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleFormaPago Where idCompra = {0} And Origen = {1} And idFormaPago = 10", IdCompraFact, IdOrigen)) == 0)
							{
								this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
								this.ultraGrid1.ActiveRow.Cells["idDetFormaPago"].Value = 0;
								this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value = 10;
								this.ultraGrid1.ActiveRow.Cells["Valor"].Value = DValorRetencion;
								this.ultraGrid1.ActiveRow.Cells["Origen"].Value = IdOrigen;
								this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
							}
						}
						if (BNotaDeCreditoFact)
						{
							if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleFormaPago Where idCompra = {0} And Origen = {1} And idFormaPago = 4", IdCompraFact, IdOrigen)) == 0)
							{
								this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
								this.ultraGrid1.ActiveRow.Cells["idDetFormaPago"].Value = 0;
								this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value = 4;
								this.ultraGrid1.ActiveRow.Cells["Valor"].Value = DValorNotaDeCredito;
								this.ultraGrid1.ActiveRow.Cells["Origen"].Value = IdOrigen;
							}
						}
					}
					#endregion Anticipos - Retenciones - Notas de Crédito
				}
			}
			else
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveRow.Cells["idDetFormaPago"].Value = 0;
				this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value = 1;
				this.ultraGrid1.ActiveRow.Cells["Valor"].Value = DValorFact;
				this.ultraGrid1.ActiveRow.Cells["Origen"].Value = IdOrigen;
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells[0];

				this.spnSaldo.Value = DValorFact;
				
				this.label2.Visible = false;
				this.spnSubTotal.Visible = false;

				this.lblFecha.Visible = true;
				this.dtFecha.Visible = true;
				
				this.dtFecha.Value = DateTime.Today;
				this.dtFecha.CalendarInfo.MaxDate = (DateTime)this.dtFecha.Value;

				this.label1.Text = "SALDO FACTURA";
				this.label5.Text = "PAGO";
			}
			bLoad = false;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			stFormato = stMoneda + stCeros.Substring(0, 2);
			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["Valor"].Format = stFormato; 
			e.Layout.Bands[0].Columns["Valor"].MaskInput = stInput;
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			//e.Row.Cells["idCompra"].Value = IdCompra;
			e.Row.Cells["Origen"].Value = IOrigen;
		}
		
		int idActaRecibido = 0;
		decimal dValorCupon = 0.00m;

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
//			if (e.Cell.Column.ToString() == "idFormaPago" && !BPagoCredito) if ((decimal ) this.spnSaldo.Value > 0) e.Cell.Row.Cells["Valor"].Value = (decimal ) this.spnSaldo.Value;
//			else e.Cell.Row.Cells["Valor"].Value = 0;
			
			#region Editor de información de Tarjetas de crédito - Transferencias - Cheques - Notas de crédito
			if (e.Cell.Column.ToString() == "idFormaPago")
			{
				int iFormaPago = (int) e.Cell.Row.Cells["idFormaPago"].Value;

				#region Notas De Credito
				if (iFormaPago == 4 && IdOrigen == 4)
				{
					SNumNC = e.Cell.Row.Cells["Numero"].Value.ToString();
					DValorNC = (decimal )e.Cell.Row.Cells["Valor"].Value;
					using(frmCruzarNotaDeCredito miNC = new frmCruzarNotaDeCredito(IdCliente, 0, SNumNC, DValorNC, DSaldo, DValorFact))
					{
						if (DialogResult.OK == miNC.ShowDialog())
						{
							e.Cell.Row.Cells["Numero"].Value = miNC.txtNumero.Text;
							e.Cell.Row.Cells["Valor"].Value = Convert.ToDecimal(miNC.txtValor.Value);
							DSaldo = Convert.ToDecimal(miNC.txtSaldo.Value);
							e.Cell.Row.Cells["idDocumento"].Value = (int)miNC.IdCompra;
						}
					}
				}
				#endregion Notas De Credito

				#region Cheque Tarjetas Transferencias Depositos
				if (iFormaPago == 2 || iFormaPago == 3 || iFormaPago == 7 || iFormaPago == 11 || iFormaPago == 13)// && e.Cell.Row.Cells["idFormaPago"].Value != DBNull.Value) //(decimal ) e.Cell.Row.Cells["Valor"].Value > 0 && 
				{
					#region Variables
					int IdBanco = 0;
					int IdTarjeta = 0;
					int IdCuenta = 0;
					string SNumero = "";
					DateTime DTFecha = DateTime.Now;
					string SCuentaCorriente = "";
					string SVoucher = "";
					string SAutorizacion = "";
					string SLote = "";
					string SReferencia = "";
					int IdPlan = 0;
					int IdPlazo = 0;
					#endregion Variables

					#region Asignación de valores
					if (iFormaPago == 3)
					{
						if (e.Cell.Row.Cells["idBanco"].Value != DBNull.Value)
						{
							IdBanco = (int)e.Cell.Row.Cells["idBanco"].Value;
							IdTarjeta = (int)e.Cell.Row.Cells["idTarjeta"].Value;
							SNumero = e.Cell.Row.Cells["Numero"].Value.ToString();
							if (e.Cell.Row.Cells["FechaCaducidad"].Value != System.DBNull.Value) DTFecha = (DateTime)e.Cell.Row.Cells["FechaCaducidad"].Value;
							SVoucher = e.Cell.Row.Cells["NumVoucher"].Value.ToString();
							SAutorizacion = e.Cell.Row.Cells["Autorizacion"].Value.ToString();
							SLote = e.Cell.Row.Cells["Lote"].Value.ToString();
							SReferencia = e.Cell.Row.Cells["Referencia"].Value.ToString();
							IdPlan = (int)e.Cell.Row.Cells["idPlan"].Value;
							IdPlazo = (int)e.Cell.Row.Cells["idPlazo"].Value; 
						}
					}
					else if (iFormaPago == 2)
					{
						if (e.Cell.Row.Cells["idBanco"].Value != DBNull.Value) 
						{
							IdBanco = (int)e.Cell.Row.Cells["idBanco"].Value;
							SCuentaCorriente = e.Cell.Row.Cells["CuentaCorriente"].Value.ToString();
							SNumero = e.Cell.Row.Cells["Numero"].Value.ToString();
						}
					}
					else if (iFormaPago == 7)
					{
						if (e.Cell.Row.Cells["idCuenta"].Value != DBNull.Value) 
						{
							IdCuenta = (int)e.Cell.Row.Cells["idCuenta"].Value;
							SNumero = (string)e.Cell.Row.Cells["Numero"].Value;
						}
					}
					#endregion Asignación de valores

					using(ChequeTarjeta miCheque = new ChequeTarjeta(iFormaPago, IdBanco, IdTarjeta, IdCuenta, SNumero, DTFecha, SCuentaCorriente, SVoucher, SAutorizacion, SLote, SReferencia, IdPlan, IdPlazo, iBodega, DateTime.Today))
					{
						if (DialogResult.OK == miCheque.ShowDialog())
						{
							Suma();

							#region Forma de pago Tarjeta de crédito
							if (iFormaPago == 3 || iFormaPago == 11)
							{
								e.Cell.Row.Cells["idTarjeta"].Value = (int) miCheque.cmbTarjeta.Value;
								e.Cell.Row.Cells["idBanco"].Value = (int) miCheque.cmbBanco.Value;
								e.Cell.Row.Cells["Numero"].Value = (string) miCheque.txtNumero.Text;
								if (iFormaPago == 3) e.Cell.Row.Cells["FechaCaducidad"].Value = (DateTime) miCheque.dtFecha.Value;
								e.Cell.Row.Cells["NumVoucher"].Value = (string) miCheque.txtNumVoucher.Text;
								e.Cell.Row.Cells["Autorizacion"].Value = (string) miCheque.txtAutorizacion.Text; 
								e.Cell.Row.Cells["idPlan"].Value = (int) miCheque.cmbPlanes.Value;
								if (miCheque.cmbPlazo.Text.ToString().Length > 0) e.Cell.Row.Cells["idPlazo"].Value = (int) miCheque.cmbPlazo.Value;
								e.Cell.Row.Cells["Lote"].Value = (string) miCheque.txtLote.Text;
								e.Cell.Row.Cells["Referencia"].Value = (string) miCheque.txtReferencia.Text;
							}
							#endregion Forma de pago Tarjeta de crédito

							#region Forma de pago Cheque
							else if (iFormaPago == 2)
							{
								e.Cell.Row.Cells["idBanco"].Value = (int) miCheque.cmbBanco.Value;
								e.Cell.Row.Cells["CuentaCorriente"].Value = (string) miCheque.txtCuentaCheque.Text;
								e.Cell.Row.Cells["Numero"].Value = (string) miCheque.txtNumero.Text;
							}
							#endregion Forma de pago Cheque

							#region Transferencia - Deposito
							else if (iFormaPago == 7 || iFormaPago == 13)
							{
								e.Cell.Row.Cells["idCuenta"].Value = (int) miCheque.cmbCuenta.Value;
								e.Cell.Row.Cells["Numero"].Value = (string) miCheque.txtNumero.Text;
							}
							#endregion Transferencia - Deposito
						}
					}
				}
				#endregion Cheque Tarjetas Transferencias Depositos

				#region Cupon
				if (iFormaPago == 14)
				{
					using(frmNumeracionCupones Cupon = new frmNumeracionCupones(uGridVenta))
					{
						if (DialogResult.OK == Cupon.ShowDialog())
						{
							idActaRecibido = Cupon.idActaRecibido;
							dValorCupon = Cupon.dValor;

							e.Cell.Row.Cells["Numero"].Value = (string)Cupon.txtNumero.Text;
							e.Cell.Row.Cells["Valor"].Value = dValorCupon;
							Suma();
							return;
						}
					}
				}
				#endregion Cupon

				Suma();

				e.Cell.Row.Cells["Valor"].Value = Convert.ToDecimal(this.spnSaldo.Value);
				return;
			}
			#endregion Editor de información de Tarjetas de crédito - Transferencias - Cheques - Notas de crédito

			#region Suma valores 
			if (e.Cell.Column.ToString() == "Valor")
			{
				#region Valida Cupon 
				if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 14)
				{
					if (Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value) != dValorCupon)
					{
						MessageBox.Show(string.Format("El Valor del cupón no puede ser distinto de {0}", dValorCupon), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cell.Row.Cells["Valor"].Value = dValorCupon;
					}
				}
				#endregion Valida Cupon

				Suma();				
			}			
			#endregion Suma valores 
		}

		private void Suma()
		{
			if (!BPagoCredito)
			{
				this.spnRetencion.Value = 0.00m;
				this.spnSubTotal.Value = 0.00m;
				this.spnSaldo.Value = 0.00m;
				this.spnTotal.Value = Math.Round(DValorFact, 2);

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (dr.Cells["Valor"].Value != DBNull.Value)
					{
						if ((IdOrigen == 2 || IdOrigen == 3 || IdOrigen == 4) || IdFormaDePago == 9)
						{
							dr.Cells["ValorUsoTarjeta"].Value = 0;
							if ((int)dr.Cells["idFormaPago"].Value == 3 || (int)dr.Cells["idFormaPago"].Value == 11)
							{
								//decimal  dIvaNuevo = Math.Round((decimal ) dr.Cells["Valor"].Value / 0.92, 2);
								dr.Cells["ValorUsoTarjeta"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Valor"].Value) * 5.00m / 100.00m, 2);
								this.spnRetencion.Value = Math.Round(Convert.ToDecimal(this.spnRetencion.Value) + Convert.ToDecimal(dr.Cells["ValorUsoTarjeta"].Value), 2);
							}
						}
						this.spnSubTotal.Value = Math.Round(Convert.ToDecimal(this.spnSubTotal.Value) + Convert.ToDecimal(dr.Cells["Valor"].Value), 2);
					}				
				}

				this.spnSubTotal.Value = Math.Round(Convert.ToDecimal(this.spnSubTotal.Value) + DAnticipo + DRetenido, 2);
				this.spnSaldo.Value = Math.Round(Convert.ToDecimal(this.spnTotal.Value) - Convert.ToDecimal(this.spnSubTotal.Value), 2); //  + (decimal )this.spnRetencion.Value)
				this.spnTotal.Value = Convert.ToDecimal(this.spnTotal.Value) + Convert.ToDecimal(this.spnRetencion.Value);
			}
			else
			{
				this.spnTotal.Value = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{					
					if (dr.Cells["Valor"].Value != DBNull.Value)
					{						
						this.spnTotal.Value = Math.Round(Convert.ToDecimal(this.spnTotal.Value) + Convert.ToDecimal(dr.Cells["Valor"].Value), 2);
					}				
				}
			}
		}

		private void VentaFormaPagoLotte_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			switch(e.KeyValue)
			{
					#region Forma de pago
				case (int) Keys.F3:
					if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "IDFORMAPAGO" || ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "VALOR")
					{
						if (ultraGrid1.ActiveRow.Cells["idFormaPago"].Value != DBNull.Value)
						{
							int iFormaPago = (int) ultraGrid1.ActiveRow.Cells["idFormaPago"].Value;
														
							if (iFormaPago == 2 || iFormaPago == 3 || iFormaPago == 7 || iFormaPago == 11)
							{
								#region Variables
								int IdBanco = 0;
								int IdTarjeta = 0;
								int IdCuenta = 0;
								string SNumero = "";
								DateTime DTFecha = DateTime.Now;
								string SCuentaCorriente = "";
								string SVoucher = "";
								string SAutorizacion = "";
								string SLote = "";
								string SReferencia = "";
								int IdPlan = 0;
								int IdPlazo = 0;
								#endregion Variables

								#region Asignación de valores
								if (iFormaPago == 3)
								{
									if (ultraGrid1.ActiveRow.Cells["idBanco"].Value != DBNull.Value) 
									{
										IdBanco = (int)ultraGrid1.ActiveRow.Cells["idBanco"].Value;
										IdTarjeta = (int)ultraGrid1.ActiveRow.Cells["idTarjeta"].Value;
										SNumero = (string)ultraGrid1.ActiveRow.Cells["Numero"].Value;
										DTFecha = (DateTime)ultraGrid1.ActiveRow.Cells["FechaCaducidad"].Value;
										SVoucher = (string) ultraGrid1.ActiveRow.Cells["NumVoucher"].Value;
										SAutorizacion = (string)ultraGrid1.ActiveRow.Cells["Autorizacion"].Value;
										SLote = (string)ultraGrid1.ActiveRow.Cells["Lote"].Value;
										SReferencia = (string)ultraGrid1.ActiveRow.Cells["Referencia"].Value;
										IdPlan = (int)ultraGrid1.ActiveRow.Cells["idPlan"].Value;
										IdPlazo = (int)ultraGrid1.ActiveRow.Cells["idPlazo"].Value; 
									}
								}
								else if (iFormaPago == 2)
								{
									if (ultraGrid1.ActiveRow.Cells["idBanco"].Value != DBNull.Value) 
									{
										IdBanco = (int)ultraGrid1.ActiveRow.Cells["idBanco"].Value;
										SCuentaCorriente = (string)ultraGrid1.ActiveRow.Cells["CuentaCorriente"].Value;
										SNumero = (string)ultraGrid1.ActiveRow.Cells["Numero"].Value;
									}
								}
								else if (iFormaPago == 7)
								{
									if (ultraGrid1.ActiveRow.Cells["idCuenta"].Value != DBNull.Value) 
									{
										IdCuenta = (int)ultraGrid1.ActiveRow.Cells["idCuenta"].Value;
										SNumero = (string)ultraGrid1.ActiveRow.Cells["Numero"].Value;
									}
								}
								#endregion Asignación de valores

								using(ChequeTarjeta miCheque = new ChequeTarjeta(iFormaPago, IdBanco, IdTarjeta, IdCuenta, SNumero, DTFecha, SCuentaCorriente, SVoucher, SAutorizacion, SLote, SReferencia, IdPlan, IdPlazo, iBodega, DateTime.Today))
								{
									if (DialogResult.OK == miCheque.ShowDialog())
									{
										if (iFormaPago == 3 || iFormaPago == 11)
										{
											ultraGrid1.ActiveRow.Cells["idTarjeta"].Value = (int) miCheque.cmbTarjeta.Value;
											ultraGrid1.ActiveRow.Cells["idBanco"].Value = (int) miCheque.cmbBanco.Value;
											ultraGrid1.ActiveRow.Cells["Numero"].Value = (string) miCheque.txtNumero.Text;
											ultraGrid1.ActiveRow.Cells["FechaCaducidad"].Value = (DateTime) miCheque.dtFecha.Value;
											ultraGrid1.ActiveRow.Cells["NumVoucher"].Value = (string) miCheque.txtNumVoucher.Text;
											ultraGrid1.ActiveRow.Cells["Autorizacion"].Value = (string) miCheque.txtAutorizacion.Text; 
											ultraGrid1.ActiveRow.Cells["idPlan"].Value = (int) miCheque.cmbPlanes.Value;
											if (miCheque.cmbPlazo.Text.ToString().Length > 0) ultraGrid1.ActiveRow.Cells["idPlazo"].Value = (int) miCheque.cmbPlazo.Value;
											ultraGrid1.ActiveRow.Cells["Lote"].Value = (string) miCheque.txtLote.Text;
											ultraGrid1.ActiveRow.Cells["Referencia"].Value = (string) miCheque.txtReferencia.Text;
										}
										else if (iFormaPago == 2)
										{
											ultraGrid1.ActiveRow.Cells["idBanco"].Value = (int) miCheque.cmbBanco.Value;
											ultraGrid1.ActiveRow.Cells["CuentaCorriente"].Value = (string) miCheque.txtCuentaCheque.Text;
											ultraGrid1.ActiveRow.Cells["Numero"].Value = (string) miCheque.txtNumero.Text;
										}
										else if (iFormaPago == 7)
										{
											ultraGrid1.ActiveRow.Cells["idCuenta"].Value  = (int) miCheque.cmbCuenta.Value;
											ultraGrid1.ActiveRow.Cells["Numero"].Value = (string) miCheque.txtNumero.Text;
										}
									}
								}
							}
						}
					}
					break;
					#endregion Forma de pago

					#region Flechas
				case (int) Keys.Up:
					if (ultraGrid1.ActiveCell.DroppedDown) return;

					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Left:
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (ultraGrid1.ActiveCell.DroppedDown) return;
			
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
					#endregion Flechas

					#region Enter
					
				case (int) Keys.Enter:
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
					if (!bEliminaAlValidar) return;

					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
					//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
					break;
					#endregion Enter
			}
		}

		private void ultraGrid1_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
//			if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "IDFORMAPAGO")
//			{
//				int iCont = 0;
//
//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
//				{
//					if ((int)dr.Cells["idFormaPago"].Value == 1) iCont++;
//				}
//
//				if (iCont > 1)
//				{
//					MessageBox.Show("Ya registro la forma de pago Efectivo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
//					e.Cancel = true;
//				}
//			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!bEliminaAlValidar)
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if ((int)e.Rows[i].Cells["idFormaPago"].Value == 6 || (int)e.Rows[i].Cells["idFormaPago"].Value == 9)// || (int)e.Rows[i].Cells["idFormaPago"].Value == 10)
					{
						e.Cancel = true;
					}
					else if ((int)e.Rows[i].Cells["idDetFormaPago"].Value > 0)
					{	
						if (miAcceso.BEliminaFormaDePago)
						{
							if (DialogResult.Yes ==	MessageBox.Show("Si elimina esta fila este proceso no puede ser revertido, \n\n\n ¿Esta seguro de borrar la linea seleccionada?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
							{
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From DetalleFormaPago Where idDetFormaPago = {0}", (int)e.Rows[i].Cells["idDetFormaPago"].Value));
								e.DisplayPromptMsg = false;
							}
							else e.Cancel = true;
						}
						else
						{
							MessageBox.Show("No tiene Acceso a Eliminar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cancel = true;
						}						
					}
					else if (DialogResult.Yes ==	MessageBox.Show("¿Desea Borrar Las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))  e.DisplayPromptMsg = false;
					else e.Cancel = true;
				}
			}
			else
			{
				e.DisplayPromptMsg = false;
			}
		}

		private void ultraGrid1_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idFormaPago"].Value == System.DBNull.Value || dr.Cells["Valor"].Value == System.DBNull.Value || Convert.ToDecimal(dr.Cells["Valor"].Value) == 0.00m)
				{
					bEliminaAlValidar = true;
					dr.Delete();						
				}
			}

			if (bLoad) return;

			//if (miAcceso.BSobrantesFaltantesCaja) return;

			if (!BPagoCredito)
			{
				if (Math.Round((decimal )spnSaldo.Value, 2) <= 0.00m)			
				{
					MessageBox.Show("No puede ingresar mas formas de pago el saldo es 0.00", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}

				if (Math.Round((decimal )spnSaldo.Value, 2) < 0.00m)			
				{
					MessageBox.Show("No puede ingresar mas formas de pago el saldo es menor a 0.00 Corrija los valores para continuar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}
			else
			{
				if (Math.Round(Convert.ToDecimal(spnTotal.Value), 2) > Math.Round(Convert.ToDecimal(spnSaldo.Value), 2))
				{
					MessageBox.Show(string.Format("No puede ingresar mas formas de pago excede el saldo de la factura ({0}).", Convert.ToDecimal(spnSaldo.Value)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
				if (Math.Round(Convert.ToDecimal(spnTotal.Value), 2) == Math.Round(Convert.ToDecimal(spnSaldo.Value), 2))
				{
					MessageBox.Show(string.Format("No puede ingresar mas formas de pago. El valor del pago es igual al saldo de la factura ({0}).", Convert.ToDecimal(spnSaldo.Value)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}
		}

		private void VentaFormaPagoLotte_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) 
			{
				if (DialogResult.Yes ==	MessageBox.Show("No ha grabado el cobro. \n\n ¿Esta seguro de salir?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					this.DialogResult = DialogResult.No;
					Close();
				}
			}

			if (e.KeyCode == Keys.F12) this.btGrabar_Click(sender, e); 
		}

		decimal  DValorNC = 0;
		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null) return;
			if (e.Cell.Column.ToString() == "I")
			{
				if (e.Cell.Row.Cells["idFormaPago"].Value == DBNull.Value) return;
				
				int iFormaPago = (int)e.Cell.Row.Cells["idFormaPago"].Value;

				#region Notas de Credito
				if (iFormaPago == 4)
				{
					SNumNC = e.Cell.Row.Cells["Numero"].Value.ToString();
					DValorNC = (decimal )e.Cell.Row.Cells["Valor"].Value;
					using(frmCruzarNotaDeCredito miNC = new frmCruzarNotaDeCredito(IdCliente, 0, SNumNC, DValorNC, DSaldo, DValorFact))
					{
						if (DialogResult.OK == miNC.ShowDialog())
						{
							e.Cell.Row.Cells["idDocumento"].Value = miNC.IdCompra;
							e.Cell.Row.Cells["Numero"].Value = miNC.txtNumero.Text;
							e.Cell.Row.Cells["Valor"].Value = (decimal ) miNC.txtValor.Value;

							DSaldo = (decimal )miNC.txtSaldo.Value;
						}
					}
				}
				#endregion Notas de Credito

				if (iFormaPago == 2 || iFormaPago == 3 || iFormaPago == 7 || iFormaPago == 11 || iFormaPago == 13)
				{
					#region Variables
					int IdBanco = 0;
					int IdTarjeta = 0;
					int IdCuenta = 0;
					string SNumero = "";
					DateTime DTFecha = DateTime.Now;
					string SCuentaCorriente = "";
					string SVoucher = "";
					string SAutorizacion = "";
					string SLote = "";
					string SReferencia = "";
					int IdPlan = 0;
					int IdPlazo = 0;
					#endregion Variables

					#region Asignación de valores
					if (iFormaPago == 3 || iFormaPago == 11)
					{
						if (ultraGrid1.ActiveRow.Cells["idBanco"].Value != DBNull.Value) 
						{
							IdBanco = (int)ultraGrid1.ActiveRow.Cells["idBanco"].Value;
							IdTarjeta = (int)ultraGrid1.ActiveRow.Cells["idTarjeta"].Value;
							SNumero = (string)ultraGrid1.ActiveRow.Cells["Numero"].Value;
							if (ultraGrid1.ActiveRow.Cells["FechaCaducidad"].Value != System.DBNull.Value) DTFecha = (DateTime)ultraGrid1.ActiveRow.Cells["FechaCaducidad"].Value;
							SVoucher = ultraGrid1.ActiveRow.Cells["NumVoucher"].Value.ToString();
							SAutorizacion = ultraGrid1.ActiveRow.Cells["Autorizacion"].Value.ToString();
							SLote = ultraGrid1.ActiveRow.Cells["Lote"].Value.ToString();
							SReferencia = ultraGrid1.ActiveRow.Cells["Referencia"].Value.ToString();
							IdPlan = (int)ultraGrid1.ActiveRow.Cells["idPlan"].Value;
							IdPlazo = (int)ultraGrid1.ActiveRow.Cells["idPlazo"].Value; 
						}
					}
					else if (iFormaPago == 2)
					{
						if (ultraGrid1.ActiveRow.Cells["idBanco"].Value != DBNull.Value)
						{
							IdBanco = (int)ultraGrid1.ActiveRow.Cells["idBanco"].Value;
							SCuentaCorriente = (string)ultraGrid1.ActiveRow.Cells["CuentaCorriente"].Value;
							SNumero = (string)ultraGrid1.ActiveRow.Cells["Numero"].Value;
						}
					}
					else if (iFormaPago == 7)
					{
						if (ultraGrid1.ActiveRow.Cells["idCuenta"].Value != DBNull.Value)
						{
							IdCuenta = (int)ultraGrid1.ActiveRow.Cells["idCuenta"].Value;
							SNumero = (string)ultraGrid1.ActiveRow.Cells["Numero"].Value;
						}
					}
					#endregion Asignación de valores

					//using(ChequeTarjeta miCheque = new ChequeTarjeta(iFormaPago, IdBanco, IdTarjeta, IdCuenta, SNumero, DTFecha, SCuentaCorriente, SVoucher, SAutorizacion, SLote, SReferencia, IdPlan, IdPlazo))
					using(ChequeTCLot miCheque = new ChequeTCLot(iFormaPago, IdBanco, IdTarjeta, IdCuenta, SNumero, DTFecha, SCuentaCorriente, SVoucher, SAutorizacion, SLote, SReferencia, IdPlan, IdPlazo))
					{
						if (DialogResult.OK == miCheque.ShowDialog())
						{
							if (iFormaPago == 3 || iFormaPago == 11)
							{
								ultraGrid1.ActiveRow.Cells["idTarjeta"].Value = (int) miCheque.cmbTarjeta.Value;
								ultraGrid1.ActiveRow.Cells["idBanco"].Value = (int) miCheque.cmbBanco.Value;
								ultraGrid1.ActiveRow.Cells["Numero"].Value = (string) miCheque.txtNumero.Text;
								if (miCheque.dtFecha.Value != System.DBNull.Value) this.ultraGrid1.ActiveRow.Cells["FechaCaducidad"].Value = (DateTime) miCheque.dtFecha.Value;
								ultraGrid1.ActiveRow.Cells["NumVoucher"].Value = (string) miCheque.txtNumVoucher.Text;
								ultraGrid1.ActiveRow.Cells["Autorizacion"].Value = (string) miCheque.txtAutorizacion.Text; 
								ultraGrid1.ActiveRow.Cells["idPlan"].Value = (int) miCheque.cmbPlanes.Value;
								ultraGrid1.ActiveRow.Cells["idPlazo"].Value = (int) miCheque.cmbPlazo.Value;
								ultraGrid1.ActiveRow.Cells["Lote"].Value = (string) miCheque.txtLote.Text;
								ultraGrid1.ActiveRow.Cells["Referencia"].Value = (string) miCheque.txtReferencia.Text;
							}
							else if (iFormaPago == 2)
							{
								ultraGrid1.ActiveRow.Cells["idBanco"].Value = (int) miCheque.cmbBanco.Value;
								ultraGrid1.ActiveRow.Cells["CuentaCorriente"].Value = (string) miCheque.txtCuentaCheque.Text;
								ultraGrid1.ActiveRow.Cells["Numero"].Value = (string) miCheque.txtNumero.Text;
							}
							else if (iFormaPago == 7 || iFormaPago == 13)
							{
								ultraGrid1.ActiveRow.Cells["idCuenta"].Value  = (int) miCheque.cmbCuenta.Value;
								ultraGrid1.ActiveRow.Cells["Numero"].Value = (string) miCheque.txtNumero.Text;
							}
						}
					}
				}				

				#region Prenda
				if (iFormaPago == 16)
				{
					#region Variables
					int IdBanco = 0;
					int IdTarjeta = 0;
					int IdCuenta = 0;
					string SNumero = "";
					DateTime DTFecha = DateTime.Now;
					string SCuentaCorriente = "";
					string SVoucher = "";
					string SAutorizacion = "";
					string SLote = "";
					string SReferencia = "";
					int IdPlan = 0;
					int IdPlazo = 0;
					#endregion Variables

					#region Asignación de valores
					if (iFormaPago == 3 || iFormaPago == 11)
					{
						if (ultraGrid1.ActiveRow.Cells["idBanco"].Value != DBNull.Value) 
						{
							IdBanco = (int)ultraGrid1.ActiveRow.Cells["idBanco"].Value;
							IdTarjeta = (int)ultraGrid1.ActiveRow.Cells["idTarjeta"].Value;
							SNumero = (string)ultraGrid1.ActiveRow.Cells["Numero"].Value;
							if (ultraGrid1.ActiveRow.Cells["FechaCaducidad"].Value != System.DBNull.Value) DTFecha = (DateTime)ultraGrid1.ActiveRow.Cells["FechaCaducidad"].Value;
							SVoucher = ultraGrid1.ActiveRow.Cells["NumVoucher"].Value.ToString();
							SAutorizacion = ultraGrid1.ActiveRow.Cells["Autorizacion"].Value.ToString();
							SLote = ultraGrid1.ActiveRow.Cells["Lote"].Value.ToString();
							SReferencia = ultraGrid1.ActiveRow.Cells["Referencia"].Value.ToString();
							IdPlan = (int)ultraGrid1.ActiveRow.Cells["idPlan"].Value;
							IdPlazo = (int)ultraGrid1.ActiveRow.Cells["idPlazo"].Value; 
						}
					}
					else if (iFormaPago == 2)
					{
						if (ultraGrid1.ActiveRow.Cells["idBanco"].Value != DBNull.Value)
						{
							IdBanco = (int)ultraGrid1.ActiveRow.Cells["idBanco"].Value;
							SCuentaCorriente = (string)ultraGrid1.ActiveRow.Cells["CuentaCorriente"].Value;
							SNumero = (string)ultraGrid1.ActiveRow.Cells["Numero"].Value;
						}
					}
					else if (iFormaPago == 7)
					{
						if (ultraGrid1.ActiveRow.Cells["idCuenta"].Value != DBNull.Value)
						{
							IdCuenta = (int)ultraGrid1.ActiveRow.Cells["idCuenta"].Value;
							SNumero = (string)ultraGrid1.ActiveRow.Cells["Numero"].Value;
						}
					}
					#endregion Asignación de valores

					using(ChequeTCLot miCheque = new ChequeTCLot(iFormaPago, IdBanco, IdTarjeta, IdCuenta, SNumero, DTFecha, SCuentaCorriente, SVoucher, SAutorizacion, SLote, SReferencia, IdPlan, IdPlazo))
					{
						if (DialogResult.OK == miCheque.ShowDialog())
						{
							if (iFormaPago == 3 || iFormaPago == 11)
							{
								ultraGrid1.ActiveRow.Cells["idTarjeta"].Value = (int) miCheque.cmbTarjeta.Value;
								ultraGrid1.ActiveRow.Cells["idBanco"].Value = (int) miCheque.cmbBanco.Value;
								ultraGrid1.ActiveRow.Cells["Numero"].Value = (string) miCheque.txtNumero.Text;
								if (miCheque.dtFecha.Value != System.DBNull.Value) this.ultraGrid1.ActiveRow.Cells["FechaCaducidad"].Value = (DateTime) miCheque.dtFecha.Value;
								ultraGrid1.ActiveRow.Cells["NumVoucher"].Value = (string) miCheque.txtNumVoucher.Text;
								ultraGrid1.ActiveRow.Cells["Autorizacion"].Value = (string) miCheque.txtAutorizacion.Text; 
								ultraGrid1.ActiveRow.Cells["idPlan"].Value = (int) miCheque.cmbPlanes.Value;
								ultraGrid1.ActiveRow.Cells["idPlazo"].Value = (int) miCheque.cmbPlazo.Value;
								ultraGrid1.ActiveRow.Cells["Lote"].Value = (string) miCheque.txtLote.Text;
								ultraGrid1.ActiveRow.Cells["Referencia"].Value = (string) miCheque.txtReferencia.Text;
							}
							else if (iFormaPago == 2)
							{
								ultraGrid1.ActiveRow.Cells["idBanco"].Value = (int) miCheque.cmbBanco.Value;
								ultraGrid1.ActiveRow.Cells["CuentaCorriente"].Value = (string) miCheque.txtCuentaCheque.Text;
								ultraGrid1.ActiveRow.Cells["Numero"].Value = (string) miCheque.txtNumero.Text;
							}
							else if (iFormaPago == 7 || iFormaPago == 13)
							{
								ultraGrid1.ActiveRow.Cells["idCuenta"].Value  = (int) miCheque.cmbCuenta.Value;
								ultraGrid1.ActiveRow.Cells["Numero"].Value = (string) miCheque.txtNumero.Text;
							}
							else if (iFormaPago == 16)
							{
								//Ingresa datos de la prenda
							}
						}
					}
				}
				#endregion Prenda

				#region Cupon
				if (iFormaPago == 14)
				{
					using(frmNumeracionCupones Cupon = new frmNumeracionCupones(uGridVenta))
					{
						if (DialogResult.OK == Cupon.ShowDialog())
						{
							ultraGrid1.ActiveRow.Cells["Numero"].Value = (string)Cupon.txtNumero.Text;
							ultraGrid1.ActiveRow.Cells["Valor"].Value = Cupon.dValor;
						}
					}
				}
				#endregion Cupon
			}
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (bLoad) return;
			
			if (e.Cell.Column.ToString() == "Valor")
			{
				if (e.Cell.Row.Cells["idFormaPago"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 9 || (int)e.Cell.Row.Cells["idFormaPago"].Value == 12 
						|| (int)e.Cell.Row.Cells["idFormaPago"].Value == 10 || (int)e.Cell.Row.Cells["idFormaPago"].Value == 6)// || (int)e.Cell.Row.Cells["idFormaPago"].Value == 14)
					{
						e.Cancel = true;
					}
				}
			}

			if (e.Cell.Column.ToString() == "idFormaPago")
			{
				if (e.Cell.Row.Cells["idFormaPago"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 9 || (int)e.Cell.Row.Cells["idFormaPago"].Value == 12 || (int)e.Cell.Row.Cells["idFormaPago"].Value == 6 
						|| (int)e.Cell.Row.Cells["idFormaPago"].Value == 10 || (int)e.Cell.Row.Cells["idFormaPago"].Value == 11)
					{
						e.Cancel = true;
					}
				}
			}
		}

		private void cdsDetalleFormaPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			try
			{
				#region Elimina filas vacias
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (dr.Cells["idFormaPago"].Value == System.DBNull.Value || dr.Cells["Valor"].Value == System.DBNull.Value)// || (decimal)dr.Cells["Valor"].Value == 0.00m)
					{
						bEliminaAlValidar = true;
						dr.Delete();						
					}
				}
				#endregion Elimina filas vacias
				Suma();
				#region Validaciones
				if (BPagoCredito)
				{
					#region Valida Cierre de Periodo
					if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Cobros"))
					{
						DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
						MessageBox.Show(string.Format("No puede ingresar Cobros en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.dtFecha.Focus();
						return;
					}
					#endregion Valida Cierre de Periodo
				}

				#region Valida Cupones
				int iNumeroCupones = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drCupon in ultraGrid1.Rows.All)
				{
					if (drCupon.Cells["idFormaPago"].Value != System.DBNull.Value)
					{
						if ((int)drCupon.Cells["idFormaPago"].Value == 14) iNumeroCupones++;						
					}
				}

				if (iNumeroCupones > 1)
				{
					MessageBox.Show(string.Format("Esta ingresando {0} cupones, En una Factura solo puede utilizar un Cupon", iNumeroCupones), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#endregion Valida Cupones

				#region Valida Informacion de las forma de pago
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (dr.Cells["idFormaPago"].Value == System.DBNull.Value && (Convert.ToDecimal(dr.Cells["Valor"].Value) == 0.00m || dr.Cells["Valor"].Value == System.DBNull.Value)) return;
					int iFP = (int) dr.Cells["idFormaPago"].Value;
					string FormaPago = "";

					if (iFP == 1) FormaPago = "l Efectivo  ";
					if (iFP == 2) FormaPago = "l Cheque ";
					if (iFP == 3) FormaPago = " la Tarjeta de crédito ";
					if (iFP == 4) FormaPago = " la Nota de crédito ";
					if (iFP == 7) FormaPago = " la Transferencia";
					if (iFP == 11) FormaPago = " la Tarjeta de debito ";
					if (iFP == 14) FormaPago = "l Cupón";

					if (iFP == 0 || dr.Cells["idFormaPago"].Value == DBNull.Value)
					{
						MessageBox.Show("Seleccione la forma de pago.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						return;
					}

					if (Convert.ToDecimal(dr.Cells["Valor"].Value) == 0.00m)
					{
						MessageBox.Show("Ingrese la cantidad de" + FormaPago + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						return;
					}
				
					if (iFP == 3 || iFP == 2 || iFP == 11)
					{
						if (dr.Cells["idBanco"].Value == DBNull.Value)
						{
							MessageBox.Show("Ingrese los datos de" + FormaPago + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.ultraGrid1.ActiveRow = dr;
							return;
						}
					}

					if (iFP == 7)
					{
						if (dr.Cells["idCuenta"].Value == DBNull.Value)
						{
							MessageBox.Show("Ingrese los datos de" + FormaPago + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.ultraGrid1.ActiveRow = dr;
							return;
						}
					}

					if (iFP == 14)
					{
						if (dr.Cells["Numero"].Value.ToString().Trim().Length == 0)
						{
							MessageBox.Show("Ingrese los datos de" + FormaPago + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.ultraGrid1.ActiveRow = dr;
							return;
						}
					}
				}
				#endregion Valida Informacion de las forma de pago

				#region Sobrantes Faltantes
				if (!miAcceso.BSobrantesFaltantesCaja)
				{
					if (!BPagoCredito)
					{
						decimal  dSaldo = Convert.ToDecimal(this.spnSaldo.Value);
						if (Math.Round(Convert.ToDecimal(spnSaldo.Value), 2) > 0.00m)
						{
							MessageBox.Show("El saldo es mayor a 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return;
						}

						if (Math.Round(Convert.ToDecimal(spnSaldo.Value), 2) < 0.00m) 
						{
							MessageBox.Show("El saldo es menor a 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return;
						}
					}
					else
					{
						if (Math.Round(Convert.ToDecimal(spnTotal.Value), 2) > Math.Round(Convert.ToDecimal(spnSaldo.Value), 2))			
						{
							MessageBox.Show(string.Format("El pago es mayor al saldo de la factura ({0}).", Convert.ToDecimal(spnSaldo.Value)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
					}
				}
				#endregion Sobrantes Faltantes
				#endregion Validaciones

				#region Mensaje de ingreso de valores a caja
				int idFormaPago = 0;
				decimal  dValor = 0.00m;
				decimal  dTotal = 0.00m;
				string Mensaje = "Detalle de ingreso a caja: \n\n";

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					idFormaPago = (int)dr.Cells["idFormaPago"].Value;
					dValor = Convert.ToDecimal(dr.Cells["Valor"].Value);

					if (idFormaPago == 1) Mensaje = Mensaje + "\t Efectivo           \t" + dValor.ToString("#,##0.00") + "\n";
					if (idFormaPago == 2) Mensaje = Mensaje + "\t Cheque             \t" + dValor.ToString("#,##0.00") + "\n";
					if (idFormaPago == 3) Mensaje = Mensaje + "\t Tarjeta de crédito \t" + dValor.ToString("#,##0.00") + "\n";
					if (idFormaPago == 4) Mensaje = Mensaje + "\t Nota de crédito    \t" + dValor.ToString("#,##0.00") + "\n";
					if (idFormaPago == 7) Mensaje = Mensaje + "\t Transferencia      \t" + dValor.ToString("#,##0.00") + "\n";
					if (idFormaPago == 8) Mensaje = Mensaje + "\t Anticipo		       \t" + dValor.ToString("#,##0.00") + "\n";
					if (idFormaPago == 10 && !BPagoCredito) Mensaje = Mensaje + "\t Retención      \t" + dValor.ToString("#,##0.00") + "\n";
					if (idFormaPago == 11) Mensaje = Mensaje + "\t Tarjeta de débito \t" + dValor.ToString("#,##0.00") + "\n";
				}
				decimal dComisionTarjeta = Convert.ToDecimal(this.spnRetencion.Value);
				if (Convert.ToDecimal(this.spnRetencion.Value) > 0) Mensaje = Mensaje + "\t Comisión Tarjeta  \t" +  dComisionTarjeta.ToString("#,##0.00") + "\n";

				dTotal = Convert.ToDecimal(this.spnTotal.Value);
				#endregion Mensaje de ingreso de valores a caja

				#region Mensaje de confirmacion 
				if	(DialogResult.Yes == MessageBox.Show(Mensaje + "\n\n\t TOTAL              \t" + dTotal.ToString("#,##0.00") + "\n\n ¿Esta seguro de grabar?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActaRecibido Set Estado = 2 Where idActaRecibido = {0}", idActaRecibido));
					this.DialogResult = DialogResult.OK;
					this.Close();
				}
				else
					return;
				#endregion Mensaje de confirmacion 
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			if (IdCompraFact > 0)
			{
				if (this.ultraGrid1.Rows.Count == 0)
				{
					MessageBox.Show("Ha Eliminado Formas De Pago no puede Salir sin Grabar el Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
      if (DialogResult.Yes ==	MessageBox.Show("No ha grabado el cobro. \n\n ¿Esta seguro de salir?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
			{
				this.DialogResult = DialogResult.No;
				Close();
			}
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			Suma();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btGrabar.Focus();
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{			
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Cobros"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Cobros en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				e.Cancel = true;					
			}
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value) this.dtFecha.Value = DateTime.Today;
		}
	}
}

