using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de TarjetaImprime.
	/// </summary>
	public class TarjetaImprime : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorrD;
		private System.Windows.Forms.CheckBox chkDeposito;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaVoucherDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaDepDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaDepHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaVoucherHasta;
		private System.Windows.Forms.CheckBox chkVoucher;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorriente;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private System.Windows.Forms.CheckBox chkCuentaDep;
		private System.Windows.Forms.CheckBox chkCuentaVoucher;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReporte;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private System.Windows.Forms.Button btCuentas;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public TarjetaImprime()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
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
			this.cmbFechaVoucherDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbFechaDepDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbCuentaCorrD = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.chkDeposito = new System.Windows.Forms.CheckBox();
			this.cmbFechaDepHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbFechaVoucherHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.chkVoucher = new System.Windows.Forms.CheckBox();
			this.cmbCuentaCorriente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.chkCuentaDep = new System.Windows.Forms.CheckBox();
			this.chkCuentaVoucher = new System.Windows.Forms.CheckBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtReporte = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.btCuentas = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaVoucherDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaDepDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorrD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaDepHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaVoucherHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtReporte)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbFechaVoucherDesde
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaVoucherDesde.Appearance = appearance1;
			this.cmbFechaVoucherDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFechaVoucherDesde.DateButtons.Add(dateButton1);
			this.cmbFechaVoucherDesde.Format = "dd/MMM/yyyy";
			this.cmbFechaVoucherDesde.Location = new System.Drawing.Point(152, 24);
			this.cmbFechaVoucherDesde.Name = "cmbFechaVoucherDesde";
			this.cmbFechaVoucherDesde.NonAutoSizeHeight = 23;
			this.cmbFechaVoucherDesde.Size = new System.Drawing.Size(120, 21);
			this.cmbFechaVoucherDesde.SpinButtonsVisible = true;
			this.cmbFechaVoucherDesde.TabIndex = 2;
			// 
			// cmbFechaDepDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaDepDesde.Appearance = appearance2;
			this.cmbFechaDepDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbFechaDepDesde.DateButtons.Add(dateButton2);
			this.cmbFechaDepDesde.Format = "dd/MMM/yyyy";
			this.cmbFechaDepDesde.Location = new System.Drawing.Point(152, 152);
			this.cmbFechaDepDesde.Name = "cmbFechaDepDesde";
			this.cmbFechaDepDesde.NonAutoSizeHeight = 23;
			this.cmbFechaDepDesde.Size = new System.Drawing.Size(120, 21);
			this.cmbFechaDepDesde.SpinButtonsVisible = true;
			this.cmbFechaDepDesde.TabIndex = 20;
			// 
			// cmbCuentaCorrD
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaCorrD.Appearance = appearance3;
			this.cmbCuentaCorrD.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorrD.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorrD.DataMember = "CuentaCorriente";
			this.cmbCuentaCorrD.DataSource = this.cdsCuentaCorr;
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
			this.cmbCuentaCorrD.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuentaCorrD.DisplayMember = "NumCuenta";
			this.cmbCuentaCorrD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCorrD.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCorrD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorrD.Location = new System.Drawing.Point(128, 80);
			this.cmbCuentaCorrD.Name = "cmbCuentaCorrD";
			this.cmbCuentaCorrD.Size = new System.Drawing.Size(120, 21);
			this.cmbCuentaCorrD.TabIndex = 34;
			this.cmbCuentaCorrD.ValueMember = "idCuentaCorriente";
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
			// chkDeposito
			// 
			this.chkDeposito.Checked = true;
			this.chkDeposito.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDeposito.Location = new System.Drawing.Point(32, 168);
			this.chkDeposito.Name = "chkDeposito";
			this.chkDeposito.TabIndex = 35;
			this.chkDeposito.Text = "Toda Fecha";
			// 
			// cmbFechaDepHasta
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaDepHasta.Appearance = appearance4;
			this.cmbFechaDepHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.cmbFechaDepHasta.DateButtons.Add(dateButton3);
			this.cmbFechaDepHasta.Format = "dd/MMM/yyyy";
			this.cmbFechaDepHasta.Location = new System.Drawing.Point(128, 48);
			this.cmbFechaDepHasta.Name = "cmbFechaDepHasta";
			this.cmbFechaDepHasta.NonAutoSizeHeight = 23;
			this.cmbFechaDepHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbFechaDepHasta.SpinButtonsVisible = true;
			this.cmbFechaDepHasta.TabIndex = 36;
			// 
			// cmbFechaVoucherHasta
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaVoucherHasta.Appearance = appearance5;
			this.cmbFechaVoucherHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.cmbFechaVoucherHasta.DateButtons.Add(dateButton4);
			this.cmbFechaVoucherHasta.Format = "dd/MMM/yyyy";
			this.cmbFechaVoucherHasta.Location = new System.Drawing.Point(152, 56);
			this.cmbFechaVoucherHasta.Name = "cmbFechaVoucherHasta";
			this.cmbFechaVoucherHasta.NonAutoSizeHeight = 23;
			this.cmbFechaVoucherHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbFechaVoucherHasta.SpinButtonsVisible = true;
			this.cmbFechaVoucherHasta.TabIndex = 37;
			// 
			// chkVoucher
			// 
			this.chkVoucher.Checked = true;
			this.chkVoucher.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkVoucher.Location = new System.Drawing.Point(32, 32);
			this.chkVoucher.Name = "chkVoucher";
			this.chkVoucher.Size = new System.Drawing.Size(96, 24);
			this.chkVoucher.TabIndex = 38;
			this.chkVoucher.Text = "Toda Fecha";
			// 
			// cmbCuentaCorriente
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaCorriente.Appearance = appearance6;
			this.cmbCuentaCorriente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorriente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorriente.DataMember = "CuentaCorriente";
			this.cmbCuentaCorriente.DataSource = this.cdsCuentaCorr;
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
			this.cmbCuentaCorriente.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuentaCorriente.DisplayMember = "NumCuenta";
			this.cmbCuentaCorriente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCorriente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCorriente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorriente.Location = new System.Drawing.Point(152, 88);
			this.cmbCuentaCorriente.Name = "cmbCuentaCorriente";
			this.cmbCuentaCorriente.Size = new System.Drawing.Size(120, 21);
			this.cmbCuentaCorriente.TabIndex = 39;
			this.cmbCuentaCorriente.ValueMember = "idCuentaCorriente";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(40, 304);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(88, 24);
			this.btAceptar.TabIndex = 40;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(192, 304);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(88, 24);
			this.btCancelar.TabIndex = 41;
			this.btCancelar.Text = "&Cancelar";
			// 
			// chkCuentaDep
			// 
			this.chkCuentaDep.Checked = true;
			this.chkCuentaDep.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCuentaDep.Location = new System.Drawing.Point(32, 216);
			this.chkCuentaDep.Name = "chkCuentaDep";
			this.chkCuentaDep.TabIndex = 44;
			this.chkCuentaDep.Text = "Toda Cuenta";
			// 
			// chkCuentaVoucher
			// 
			this.chkCuentaVoucher.Checked = true;
			this.chkCuentaVoucher.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCuentaVoucher.Location = new System.Drawing.Point(32, 80);
			this.chkCuentaVoucher.Name = "chkCuentaVoucher";
			this.chkCuentaVoucher.Size = new System.Drawing.Size(96, 24);
			this.chkCuentaVoucher.TabIndex = 45;
			this.chkCuentaVoucher.Text = "Toda Cuenta";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cmbCuentaCorrD);
			this.groupBox1.Controls.Add(this.cmbFechaDepHasta);
			this.groupBox1.Location = new System.Drawing.Point(24, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(264, 120);
			this.groupBox1.TabIndex = 46;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Deposito";
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(16, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(272, 120);
			this.groupBox2.TabIndex = 47;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Voucher";
			// 
			// txtReporte
			// 
			this.txtReporte.Location = new System.Drawing.Point(152, 264);
			this.txtReporte.Name = "txtReporte";
			this.txtReporte.Size = new System.Drawing.Size(120, 21);
			this.txtReporte.TabIndex = 48;
			this.txtReporte.Text = "Tarjeta.Rpt";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.Location = new System.Drawing.Point(40, 272);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(48, 16);
			this.ultraLabel1.TabIndex = 49;
			this.ultraLabel1.Text = "Reporte";
			// 
			// btCuentas
			// 
			this.btCuentas.Location = new System.Drawing.Point(272, 264);
			this.btCuentas.Name = "btCuentas";
			this.btCuentas.Size = new System.Drawing.Size(24, 24);
			this.btCuentas.TabIndex = 50;
			this.btCuentas.Text = "C";
			this.btCuentas.Click += new System.EventHandler(this.btCuentas_Click);
			// 
			// TarjetaImprime
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(312, 350);
			this.Controls.Add(this.btCuentas);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtReporte);
			this.Controls.Add(this.chkCuentaVoucher);
			this.Controls.Add(this.chkCuentaDep);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbCuentaCorriente);
			this.Controls.Add(this.chkVoucher);
			this.Controls.Add(this.cmbFechaVoucherHasta);
			this.Controls.Add(this.chkDeposito);
			this.Controls.Add(this.cmbFechaDepDesde);
			this.Controls.Add(this.cmbFechaVoucherDesde);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox2);
			this.Name = "TarjetaImprime";
			this.Text = "Tarjeta Imprime";
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaVoucherDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaDepDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorrD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaDepHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaVoucherHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtReporte)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesdeV = (DateTime) cmbFechaVoucherDesde.Value;
			DateTime dtHastaV = (DateTime) cmbFechaVoucherHasta.Value;

			DateTime dtDesdeD = (DateTime) cmbFechaDepDesde.Value;
			DateTime dtHastaD = (DateTime) cmbFechaDepHasta.Value;

			string stFiltro = "1=1 ";
			if (!chkCuentaVoucher.Checked)
				stFiltro += " And {Tarjeta.idCuentaCorriente} = " + cmbCuentaCorriente.Value.ToString();

			if (!chkCuentaDep.Checked)
				stFiltro += " And {Tarjeta.idCuentaCorriente} = " + cmbCuentaCorrD.Value.ToString();

			if (!chkVoucher.Checked)
				stFiltro += " And {Tarjeta.Fecha} >= #" + dtDesdeV.ToString("MM/dd/yyyy")
					+ "# And {Tarjeta.Fecha} < #" + dtHastaV.AddDays(1).ToString("MM/dd/yyyy") + "#";
			if (!chkDeposito.Checked)
				stFiltro += " And {Tarjeta.FechaDep} >= #" + dtDesdeD.ToString("MM/dd/yyyy")
					+ "# And {Tarjeta.FechaDep} < #" + dtHastaD.AddDays(1).ToString("MM/dd/yyyy")+"#";
			string stReporte = txtReporte.Text.Trim();
			Reporte miRep = new Reporte(stReporte, stFiltro);
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
		}

		private void btCuentas_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
