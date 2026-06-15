using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for ChequeTarjeta.
	/// </summary>
	public class ChequeTCLot : DevExpress.XtraEditors.XtraForm
	{
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumVoucher;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutorizacion;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTarjeta;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuentaCheque;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;

		private System.Windows.Forms.Label lblAutorizacion;
		private System.Windows.Forms.Label lblVoucher;
		private System.Windows.Forms.Label lblBanco;
		private System.Windows.Forms.Label lblCtaCorriente;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label lblCaducidad;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private C1.Data.C1DataSet cdsBanco;
		private System.Windows.Forms.ErrorProvider errorProvider1;
		private System.Windows.Forms.Label lblCuenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtLote;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtReferencia;
		private System.Windows.Forms.Label lblPlazo;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbPlazo;
		private System.Windows.Forms.Label lblPlanes;
		private System.Windows.Forms.Label lblLote;
		private System.Windows.Forms.Label lblReferencia;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbPlanes;
		
		int IFormaDePago;
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
		private System.Windows.Forms.Button btGrabar;
		private System.Windows.Forms.Button btCancelar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		int IdPlazo = 0;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDesPrenda;
		public System.Windows.Forms.GroupBox gbTarjetas;
		private System.Windows.Forms.GroupBox gbPrendas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		public Infragistics.Win.UltraWinEditors.UltraOptionSet optPlan;
		public Infragistics.Win.UltraWinEditors.UltraOptionSet optTarjeta;
		string DesPrenda = "";

		public ChequeTCLot(int iFormaDePago, int idBanco, int idTarjeta, int idCuenta, string sNumero, DateTime dtFecha, string sCuentaCorriente, string sVoucher, string sAutorizacion, string sLote, string sReferencia, int idPlan, int idPlazo)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			IFormaDePago = iFormaDePago;
			IdBanco = idBanco;
			IdTarjeta = idTarjeta;
			IdCuenta = idCuenta;
			SNumero = sNumero;
			DTFecha = dtFecha;
			SCuentaCorriente = sCuentaCorriente;
			SVoucher = sVoucher;
			SAutorizacion = sAutorizacion;
			SLote = sLote;
			SReferencia = sReferencia;
			IdPlan = idPlan;
			IdPlazo = idPlazo;
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("TarjetasDeCredito", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlan");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(ChequeTCLot));
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			this.lblPlazo = new System.Windows.Forms.Label();
			this.txtNumVoucher = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblAutorizacion = new System.Windows.Forms.Label();
			this.lblVoucher = new System.Windows.Forms.Label();
			this.txtAutorizacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbTarjeta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblBanco = new System.Windows.Forms.Label();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsBanco = new C1.Data.C1DataSet();
			this.txtCuentaCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCtaCorriente = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCaducidad = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.errorProvider1 = new System.Windows.Forms.ErrorProvider();
			this.lblCuenta = new System.Windows.Forms.Label();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblLote = new System.Windows.Forms.Label();
			this.txtLote = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblReferencia = new System.Windows.Forms.Label();
			this.txtReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbPlazo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblPlanes = new System.Windows.Forms.Label();
			this.cmbPlanes = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btCancelar = new System.Windows.Forms.Button();
			this.btGrabar = new System.Windows.Forms.Button();
			this.txtDesPrenda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.gbTarjetas = new System.Windows.Forms.GroupBox();
			this.optPlan = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optTarjeta = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.gbPrendas = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtNumVoucher)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentaCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesPrenda)).BeginInit();
			this.gbTarjetas.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optPlan)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTarjeta)).BeginInit();
			this.gbPrendas.SuspendLayout();
			this.SuspendLayout();
			// 
			// lblPlazo
			// 
			this.lblPlazo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblPlazo.AutoSize = true;
			this.lblPlazo.BackColor = System.Drawing.Color.Transparent;
			this.lblPlazo.Enabled = false;
			this.lblPlazo.Location = new System.Drawing.Point(24, 200);
			this.lblPlazo.Name = "lblPlazo";
			this.lblPlazo.Size = new System.Drawing.Size(30, 16);
			this.lblPlazo.TabIndex = 166;
			this.lblPlazo.Text = "Plazo";
			this.lblPlazo.Visible = false;
			// 
			// txtNumVoucher
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumVoucher.Appearance = appearance1;
			this.txtNumVoucher.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumVoucher.Location = new System.Drawing.Point(168, 160);
			this.txtNumVoucher.MaxLength = 4;
			this.txtNumVoucher.Name = "txtNumVoucher";
			this.txtNumVoucher.Size = new System.Drawing.Size(120, 22);
			this.txtNumVoucher.TabIndex = 3;
			this.txtNumVoucher.Leave += new System.EventHandler(this.txtNumVoucher_Leave);
			this.txtNumVoucher.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumVoucher_KeyPress);
			this.txtNumVoucher.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumVoucher_Validating);
			this.txtNumVoucher.ValueChanged += new System.EventHandler(this.txtNumVoucher_ValueChanged);
			// 
			// lblAutorizacion
			// 
			this.lblAutorizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblAutorizacion.AutoSize = true;
			this.lblAutorizacion.BackColor = System.Drawing.Color.Transparent;
			this.lblAutorizacion.Enabled = false;
			this.lblAutorizacion.Location = new System.Drawing.Point(24, 128);
			this.lblAutorizacion.Name = "lblAutorizacion";
			this.lblAutorizacion.Size = new System.Drawing.Size(65, 16);
			this.lblAutorizacion.TabIndex = 162;
			this.lblAutorizacion.Text = "Autorización";
			this.lblAutorizacion.Visible = false;
			// 
			// lblVoucher
			// 
			this.lblVoucher.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblVoucher.AutoSize = true;
			this.lblVoucher.BackColor = System.Drawing.Color.Transparent;
			this.lblVoucher.Enabled = false;
			this.lblVoucher.Location = new System.Drawing.Point(56, 164);
			this.lblVoucher.Name = "lblVoucher";
			this.lblVoucher.Size = new System.Drawing.Size(78, 16);
			this.lblVoucher.TabIndex = 160;
			this.lblVoucher.Text = "N. Transacción";
			// 
			// txtAutorizacion
			// 
			this.txtAutorizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutorizacion.Appearance = appearance2;
			this.txtAutorizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutorizacion.Enabled = false;
			this.txtAutorizacion.Location = new System.Drawing.Point(128, 128);
			this.txtAutorizacion.MaxLength = 20;
			this.txtAutorizacion.Name = "txtAutorizacion";
			this.txtAutorizacion.Size = new System.Drawing.Size(200, 22);
			this.txtAutorizacion.TabIndex = 7;
			this.txtAutorizacion.Visible = false;
			this.txtAutorizacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAutorizacion_KeyPress);
			this.txtAutorizacion.ValueChanged += new System.EventHandler(this.txtAutorizacion_ValueChanged);
			// 
			// cmbTarjeta
			// 
			this.cmbTarjeta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTarjeta.Appearance = appearance3;
			this.cmbTarjeta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTarjeta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTarjeta.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 264;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTarjeta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTarjeta.DisplayMember = "Nombre";
			this.cmbTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTarjeta.Enabled = false;
			this.cmbTarjeta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTarjeta.Location = new System.Drawing.Point(88, 224);
			this.cmbTarjeta.Name = "cmbTarjeta";
			this.cmbTarjeta.Size = new System.Drawing.Size(264, 21);
			this.cmbTarjeta.TabIndex = 0;
			this.cmbTarjeta.ValueMember = "idTarjeta";
			this.cmbTarjeta.Visible = false;
			this.cmbTarjeta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTarjeta_KeyPress);
			this.cmbTarjeta.ValueChanged += new System.EventHandler(this.cmbTarjeta_ValueChanged);
			this.cmbTarjeta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTarjeta_InitializeLayout);
			this.cmbTarjeta.Enter += new System.EventHandler(this.cmbTarjeta_Enter);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			this.ultraDataSource1.Band.Key = "TarjetasDeCredito";
			// 
			// lblBanco
			// 
			this.lblBanco.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblBanco.AutoSize = true;
			this.lblBanco.BackColor = System.Drawing.Color.Transparent;
			this.lblBanco.Location = new System.Drawing.Point(24, 176);
			this.lblBanco.Name = "lblBanco";
			this.lblBanco.Size = new System.Drawing.Size(34, 16);
			this.lblBanco.TabIndex = 158;
			this.lblBanco.Text = "Banco";
			this.lblBanco.Visible = false;
			// 
			// cmbBanco
			// 
			this.cmbBanco.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance4;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataSource = this.ultraDataSource2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 224;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.Enabled = false;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(96, 176);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(256, 21);
			this.cmbBanco.TabIndex = 0;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.Visible = false;
			this.cmbBanco.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBanco_KeyPress);
			this.cmbBanco.ValueChanged += new System.EventHandler(this.cmbBanco_ValueChanged);
			this.cmbBanco.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBanco_InitializeLayout);
			this.cmbBanco.Enter += new System.EventHandler(this.cmbBanco_Enter);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
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
			// 
			// txtCuentaCheque
			// 
			this.txtCuentaCheque.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuentaCheque.Appearance = appearance5;
			this.txtCuentaCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuentaCheque.Enabled = false;
			this.txtCuentaCheque.Location = new System.Drawing.Point(128, 80);
			this.txtCuentaCheque.MaxLength = 15;
			this.txtCuentaCheque.Name = "txtCuentaCheque";
			this.txtCuentaCheque.Size = new System.Drawing.Size(200, 22);
			this.txtCuentaCheque.TabIndex = 5;
			this.txtCuentaCheque.Visible = false;
			this.txtCuentaCheque.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCuentaCheque_KeyPress);
			this.txtCuentaCheque.ValueChanged += new System.EventHandler(this.txtCuentaCheque_ValueChanged);
			// 
			// lblCtaCorriente
			// 
			this.lblCtaCorriente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblCtaCorriente.AutoSize = true;
			this.lblCtaCorriente.BackColor = System.Drawing.Color.Transparent;
			this.lblCtaCorriente.Enabled = false;
			this.lblCtaCorriente.Location = new System.Drawing.Point(24, 80);
			this.lblCtaCorriente.Name = "lblCtaCorriente";
			this.lblCtaCorriente.Size = new System.Drawing.Size(89, 16);
			this.lblCtaCorriente.TabIndex = 159;
			this.lblCtaCorriente.Text = "Cuenta Corriente";
			this.lblCtaCorriente.Visible = false;
			// 
			// lblNumero
			// 
			this.lblNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Enabled = false;
			this.lblNumero.Location = new System.Drawing.Point(24, 56);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 154;
			this.lblNumero.Text = "Número";
			this.lblNumero.Visible = false;
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance6;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(128, 56);
			this.txtNumero.MaxLength = 20;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(200, 22);
			this.txtNumero.TabIndex = 3;
			this.txtNumero.Visible = false;
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// lblCaducidad
			// 
			this.lblCaducidad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblCaducidad.AutoSize = true;
			this.lblCaducidad.Enabled = false;
			this.lblCaducidad.Location = new System.Drawing.Point(24, 32);
			this.lblCaducidad.Name = "lblCaducidad";
			this.lblCaducidad.Size = new System.Drawing.Size(89, 16);
			this.lblCaducidad.TabIndex = 151;
			this.lblCaducidad.Text = "Fecha Caducidad";
			this.lblCaducidad.Visible = false;
			// 
			// dtFecha
			// 
			this.dtFecha.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance7;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(128, 32);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 4;
			this.dtFecha.TrailingDaysVisible = false;
			this.dtFecha.Value = new System.DateTime(2025, 3, 19, 0, 0, 0, 0);
			this.dtFecha.Visible = false;
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			// 
			// errorProvider1
			// 
			this.errorProvider1.ContainerControl = this;
			// 
			// lblCuenta
			// 
			this.lblCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblCuenta.AutoSize = true;
			this.lblCuenta.BackColor = System.Drawing.Color.Transparent;
			this.lblCuenta.Location = new System.Drawing.Point(24, 224);
			this.lblCuenta.Name = "lblCuenta";
			this.lblCuenta.Size = new System.Drawing.Size(40, 16);
			this.lblCuenta.TabIndex = 168;
			this.lblCuenta.Text = "Tarjeta";
			this.lblCuenta.Visible = false;
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance8;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.Enabled = false;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(96, 224);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(256, 21);
			this.cmbCuenta.TabIndex = 2;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.Visible = false;
			this.cmbCuenta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbCuenta_KeyPress);
			this.cmbCuenta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCuenta_InitializeLayout);
			this.cmbCuenta.Enter += new System.EventHandler(this.cmbCuenta_Enter);
			// 
			// lblLote
			// 
			this.lblLote.AutoSize = true;
			this.lblLote.BackColor = System.Drawing.Color.Transparent;
			this.lblLote.Location = new System.Drawing.Point(56, 116);
			this.lblLote.Name = "lblLote";
			this.lblLote.Size = new System.Drawing.Size(86, 16);
			this.lblLote.TabIndex = 172;
			this.lblLote.Text = "Numero De Lote";
			// 
			// txtLote
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLote.Appearance = appearance9;
			this.txtLote.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLote.Location = new System.Drawing.Point(168, 113);
			this.txtLote.MaxLength = 4;
			this.txtLote.Name = "txtLote";
			this.txtLote.Size = new System.Drawing.Size(120, 22);
			this.txtLote.TabIndex = 2;
			this.txtLote.Leave += new System.EventHandler(this.txtLote_Leave);
			this.txtLote.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtLote_KeyPress);
			// 
			// lblReferencia
			// 
			this.lblReferencia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblReferencia.AutoSize = true;
			this.lblReferencia.BackColor = System.Drawing.Color.Transparent;
			this.lblReferencia.Enabled = false;
			this.lblReferencia.Location = new System.Drawing.Point(24, 152);
			this.lblReferencia.Name = "lblReferencia";
			this.lblReferencia.Size = new System.Drawing.Size(57, 16);
			this.lblReferencia.TabIndex = 174;
			this.lblReferencia.Text = "Referencia";
			this.lblReferencia.Visible = false;
			// 
			// txtReferencia
			// 
			this.txtReferencia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtReferencia.Appearance = appearance10;
			this.txtReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtReferencia.Enabled = false;
			this.txtReferencia.Location = new System.Drawing.Point(128, 152);
			this.txtReferencia.MaxLength = 15;
			this.txtReferencia.Name = "txtReferencia";
			this.txtReferencia.Size = new System.Drawing.Size(200, 22);
			this.txtReferencia.TabIndex = 9;
			this.txtReferencia.Visible = false;
			this.txtReferencia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtReferencia_KeyPress);
			// 
			// cmbPlazo
			// 
			this.cmbPlazo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPlazo.Appearance = appearance11;
			this.cmbPlazo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPlazo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlazo.DisplayMember = "Descripcion";
			this.cmbPlazo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPlazo.Enabled = false;
			this.cmbPlazo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlazo.Location = new System.Drawing.Point(96, 200);
			this.cmbPlazo.Name = "cmbPlazo";
			this.cmbPlazo.Size = new System.Drawing.Size(256, 21);
			this.cmbPlazo.TabIndex = 11;
			this.cmbPlazo.ValueMember = "idPlazo";
			this.cmbPlazo.Visible = false;
			this.cmbPlazo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbPlazo_KeyPress);
			this.cmbPlazo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbPlazo_InitializeLayout);
			this.cmbPlazo.Enter += new System.EventHandler(this.cmbPlazo_Enter);
			// 
			// lblPlanes
			// 
			this.lblPlanes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblPlanes.AutoSize = true;
			this.lblPlanes.BackColor = System.Drawing.Color.Transparent;
			this.lblPlanes.Enabled = false;
			this.lblPlanes.Location = new System.Drawing.Point(24, 256);
			this.lblPlanes.Name = "lblPlanes";
			this.lblPlanes.Size = new System.Drawing.Size(25, 16);
			this.lblPlanes.TabIndex = 176;
			this.lblPlanes.Text = "Plan";
			this.lblPlanes.Visible = false;
			// 
			// cmbPlanes
			// 
			this.cmbPlanes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPlanes.Appearance = appearance12;
			this.cmbPlanes.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPlanes.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlanes.DataSource = this.ultraDataSource3;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 264;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbPlanes.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPlanes.DisplayMember = "Descripcion";
			this.cmbPlanes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPlanes.Enabled = false;
			this.cmbPlanes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlanes.Location = new System.Drawing.Point(88, 256);
			this.cmbPlanes.Name = "cmbPlanes";
			this.cmbPlanes.Size = new System.Drawing.Size(264, 21);
			this.cmbPlanes.TabIndex = 1;
			this.cmbPlanes.ValueMember = "idPlan";
			this.cmbPlanes.Visible = false;
			this.cmbPlanes.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbPlanes_KeyPress);
			this.cmbPlanes.ValueChanged += new System.EventHandler(this.cmbPlanes_ValueChanged);
			this.cmbPlanes.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbPlanes_InitializeLayout);
			this.cmbPlanes.Enter += new System.EventHandler(this.cmbPlanes_Enter);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btCancelar.Image")));
			this.btCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btCancelar.Location = new System.Drawing.Point(180, 256);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(112, 24);
			this.btCancelar.TabIndex = 5;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.Image = ((System.Drawing.Image)(resources.GetObject("btGrabar.Image")));
			this.btGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btGrabar.Location = new System.Drawing.Point(40, 256);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(112, 24);
			this.btGrabar.TabIndex = 4;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// txtDesPrenda
			// 
			this.txtDesPrenda.AcceptsReturn = true;
			this.txtDesPrenda.AcceptsTab = true;
			this.txtDesPrenda.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDesPrenda.Appearance = appearance13;
			this.txtDesPrenda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDesPrenda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDesPrenda.Enabled = false;
			this.txtDesPrenda.Location = new System.Drawing.Point(8, 29);
			this.txtDesPrenda.Multiline = true;
			this.txtDesPrenda.Name = "txtDesPrenda";
			this.txtDesPrenda.Size = new System.Drawing.Size(328, 168);
			this.txtDesPrenda.TabIndex = 6;
			this.txtDesPrenda.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDesPrenda_KeyPress);
			// 
			// gbTarjetas
			// 
			this.gbTarjetas.Controls.Add(this.optPlan);
			this.gbTarjetas.Controls.Add(this.optTarjeta);
			this.gbTarjetas.Controls.Add(this.lblLote);
			this.gbTarjetas.Controls.Add(this.txtLote);
			this.gbTarjetas.Controls.Add(this.txtNumVoucher);
			this.gbTarjetas.Controls.Add(this.lblVoucher);
			this.gbTarjetas.Location = new System.Drawing.Point(8, 8);
			this.gbTarjetas.Name = "gbTarjetas";
			this.gbTarjetas.Size = new System.Drawing.Size(344, 200);
			this.gbTarjetas.TabIndex = 175;
			this.gbTarjetas.TabStop = false;
			this.gbTarjetas.Visible = false;
			// 
			// optPlan
			// 
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.optPlan.Appearance = appearance14;
			this.optPlan.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optPlan.CheckedIndex = 0;
			this.optPlan.Enabled = false;
			this.optPlan.ItemAppearance = appearance15;
			this.optPlan.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Crédito";
			valueListItem2.DataValue = 4;
			valueListItem2.DisplayText = "Débito";
			this.optPlan.Items.Add(valueListItem1);
			this.optPlan.Items.Add(valueListItem2);
			this.optPlan.ItemSpacingVertical = 10;
			this.optPlan.Location = new System.Drawing.Point(176, 65);
			this.optPlan.Name = "optPlan";
			this.optPlan.Size = new System.Drawing.Size(152, 24);
			this.optPlan.TabIndex = 1;
			this.optPlan.Text = "Crédito";
			// 
			// optTarjeta
			// 
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTarjeta.Appearance = appearance16;
			this.optTarjeta.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTarjeta.CheckedIndex = 0;
			this.optTarjeta.Enabled = false;
			this.optTarjeta.ItemAppearance = appearance17;
			this.optTarjeta.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Visa";
			valueListItem4.DataValue = 2;
			valueListItem4.DisplayText = "Mastercard";
			this.optTarjeta.Items.Add(valueListItem3);
			this.optTarjeta.Items.Add(valueListItem4);
			this.optTarjeta.ItemSpacingVertical = 10;
			this.optTarjeta.Location = new System.Drawing.Point(16, 65);
			this.optTarjeta.Name = "optTarjeta";
			this.optTarjeta.Size = new System.Drawing.Size(152, 24);
			this.optTarjeta.TabIndex = 0;
			this.optTarjeta.Text = "Visa";
			// 
			// gbPrendas
			// 
			this.gbPrendas.Controls.Add(this.txtDesPrenda);
			this.gbPrendas.Location = new System.Drawing.Point(8, 8);
			this.gbPrendas.Name = "gbPrendas";
			this.gbPrendas.Size = new System.Drawing.Size(344, 200);
			this.gbPrendas.TabIndex = 176;
			this.gbPrendas.TabStop = false;
			this.gbPrendas.Visible = false;
			// 
			// ChequeTCLot
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(362, 288);
			this.Controls.Add(this.gbTarjetas);
			this.Controls.Add(this.lblReferencia);
			this.Controls.Add(this.lblPlazo);
			this.Controls.Add(this.lblAutorizacion);
			this.Controls.Add(this.lblBanco);
			this.Controls.Add(this.lblCtaCorriente);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.lblCaducidad);
			this.Controls.Add(this.lblCuenta);
			this.Controls.Add(this.lblPlanes);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.cmbPlazo);
			this.Controls.Add(this.txtReferencia);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.txtAutorizacion);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.txtCuentaCheque);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbPlanes);
			this.Controls.Add(this.cmbTarjeta);
			this.Controls.Add(this.gbPrendas);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ChequeTCLot";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ChequeTCLot_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ChequeTCLot_Closing);
			this.Load += new System.EventHandler(this.ChequeTCLot_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNumVoucher)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentaCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesPrenda)).EndInit();
			this.gbTarjetas.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optPlan)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTarjeta)).EndInit();
			this.gbPrendas.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void ChequeTCLot_Load(object sender, System.EventArgs e)
		{	
			//this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsBanco, "Select idBanco, Nombre From Banco Where Caja = 1 and idBanco in (23, 21) Order By Nombre");
			this.cmbTarjeta.DataSource = Funcion.dvProcedimiento(cdsBanco, "Select idTarjeta, Nombre From TarjetasDeCredito Where idTarjeta < 3");
			this.cmbPlanes.DataSource = Funcion.dvProcedimiento(cdsBanco, "Select idPlan, Descripcion From Planes Where idPlan In (1, 4) ");

			if (IFormaDePago == 2)
			{
				this.Text = "CHEQUE";		
		
				#region Controles
				this.txtDesPrenda.Enabled = false;
				//this.lblTarjeta.Visible = false;
				this.cmbTarjeta.Visible = false;
				this.lblCaducidad.Visible = false;
				this.dtFecha.Visible = false;
				this.lblVoucher.Visible = false;
				this.txtNumVoucher.Visible = false;
				this.lblAutorizacion.Visible = false;
				this.txtAutorizacion.Visible = false;
				this.lblPlanes.Visible = false;
				this.cmbPlanes.Visible = false;
				this.lblPlazo.Visible = false;
				this.cmbPlazo.Visible = false;
				this.lblCuenta.Visible = false;
				this.cmbCuenta.Visible = false;
				this.lblLote.Visible = false;
				this.txtLote.Visible = false;
				this.lblReferencia.Visible = false;
				this.txtReferencia.Visible = false;
				this.lblNumero.Location = new Point(8,40);
				this.txtNumero.Location = new Point(112,40);
				this.lblCtaCorriente.Location = new Point(8,72);
				this.txtCuentaCheque.Location = new Point(112,72);
				#endregion Controles

				if (IdBanco > 0) this.cmbBanco.Value = IdBanco;
				if (SCuentaCorriente.Length > 0) this.txtCuentaCheque.Text = SCuentaCorriente;
				if (SNumero.Length > 0) this.txtNumero.Text = SNumero;

				this.Size = new Size(352, 168);				

				this.cmbBanco.Focus();
			}
			else if (IFormaDePago == 3 || IFormaDePago == 11)
			{
				this.Text = "TARJETA DE CRÉDITO - DÉBITO";

				this.gbTarjetas.Visible = true;
				//				this.dtFecha.CalendarInfo.MinDate = DateTime.Today;
				
				//if (IdBanco > 0) this.cmbBanco.Value = IdBanco;
				if (IdTarjeta > 0)	this.optTarjeta.Value = IdTarjeta;
				//if (SNumero.Length > 0) this.txtNumero.Text = SNumero;
				//this.dtFecha.Value = (DateTime)DTFecha;
				if (SVoucher.Length > 0) this.txtNumVoucher.Text = SVoucher;
				//if (SAutorizacion.Length > 0) this.txtAutorizacion.Text = SAutorizacion;
				if (SLote.Length > 0) this.txtLote.Text = SLote;
				//if (SReferencia.Length > 0) this.txtReferencia.Text = SReferencia;
				if (IdPlan > 0) this.optPlan.Value = IdPlan;
				//if (IdPlazo > 0) this.cmbPlazo.Value = IdPlazo;

//				if (IdTarjeta == 0)
//				{
					//this.cmbBanco.Enabled = false;
					this.optTarjeta.Enabled = true;
					this.optPlan.Enabled = true;
					this.txtLote.Enabled = true;
					//this.cmbPlazo.Enabled = false;

					this.optTarjeta.Focus();				
			}	
			else if (IFormaDePago == 7 || IFormaDePago == 13)
			{
				this.Text = "DEPOSITO - TRANSFERENCIA";

				#region Controles
				this.txtDesPrenda.Enabled = false;
				//this.lblTarjeta.Visible = false;
				this.lblBanco.Visible = false;
				this.cmbBanco.Visible = false;
				this.cmbTarjeta.Visible = false;
				this.lblCaducidad.Visible = false;
				this.dtFecha.Visible = false;
				this.lblVoucher.Visible = false;
				this.txtNumVoucher.Visible = false;
				this.lblAutorizacion.Visible = false;
				this.txtAutorizacion.Visible = false;
				this.cmbPlanes.Visible = false;
				this.lblPlazo.Visible = false;			
				this.lblCtaCorriente.Visible = false;
				this.txtCuentaCheque.Visible = false;

				this.lblCuenta.Visible = true;
				this.cmbCuenta.Visible = true;
				this.txtNumero.Enabled = true;
				#endregion Controles

				this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsBanco, "Exec ListaCuentasCierreCaja");
				this.cmbCuenta.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = true;
				this.cmbCuenta.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 224;

				if (IdCuenta > 0) this.cmbCuenta.Value = IdCuenta;
				if (SNumero.Length > 0) this.txtNumero.Text = SNumero;
				
				this.Size = new Size(352, 168);

				this.cmbCuenta.Focus();
			}
			else if (IFormaDePago == 11)
			{
				this.Text = "TARJETA DE DÉBITO";

				if (IdBanco > 0) this.cmbBanco.Value = IdBanco;
				if (IdTarjeta > 0)	this.cmbTarjeta.Value = IdTarjeta;
				if (SNumero.Length > 0) this.txtNumero.Text = SNumero;
				//this.dtFecha.Value = DateTime.Today;
				if (SVoucher.Length > 0) this.txtNumVoucher.Text = SVoucher;
				if (SAutorizacion.Length > 0) this.txtAutorizacion.Text = SAutorizacion;
				if (SLote.Length > 0) this.txtLote.Text = SLote;
				if (SReferencia.Length > 0) this.txtReferencia.Text = SReferencia;
				if (IdPlan > 0) this.cmbPlanes.Value = IdPlan;
				if (IdPlazo > 0) this.cmbPlazo.Value = IdPlazo;

				#region Controles
				this.txtDesPrenda.Enabled = false;
				this.lblCaducidad.Visible = false;
				this.dtFecha.Visible = false;
				this.lblVoucher.Visible = false;
				this.txtNumVoucher.Visible = false;
				this.lblAutorizacion.Visible = false;
				this.txtAutorizacion.Visible = false;
				this.lblCuenta.Visible = false;
				this.cmbCuenta.Visible = false;
				this.lblCtaCorriente.Visible = false;
				this.txtCuentaCheque.Visible = false;

				this.lblLote.Location = new Point(8, 104);
				this.txtLote.Location = new Point(48, 104);
				this.lblReferencia.Location = new Point(152, 104);
				this.txtReferencia.Location = new Point(224, 104);
				this.lblPlanes.Location = new Point(8, 136);
				this.cmbPlanes.Location = new Point(112, 136);
				this.lblPlazo.Location = new Point(8, 168);
				this.cmbPlazo.Location = new Point(112, 168);
				/*this.btGrabar.Location = new Point(45, 220);
				this.btCancelar.Location = new Point(181, 220);*/
				this.Size = new Size(352, 288);
				#endregion Controles
			}
			else if (IFormaDePago == 16)
			{
				this.Text = "PRENDA";

				this.gbPrendas.Visible = true;
				this.txtDesPrenda.Enabled = true;
				this.txtDesPrenda.Value = DesPrenda;
				this.txtDesPrenda.Focus();
			}
		}

		private void btTarjeta_Click(object sender, System.EventArgs e)
		{
			
		}

		private void rdgCorrDiferido_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			
		}

		private void rdgCorrDiferido_EditValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13)
			{
				if (this.txtNumero.Value != DBNull.Value)
				{
					if  (IFormaDePago == 2) this.txtCuentaCheque.Focus();
					else if (IFormaDePago == 3) this.dtFecha.Focus();
					else if (IFormaDePago == 7 || IFormaDePago == 16) this.btGrabar.Focus();
					else if (IFormaDePago == 11) this.txtLote.Focus();
				}
				//				else
				//				{
				//					MessageBox.Show("Ingrese el numero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				//					this.txtNumero.Focus();
				//				}
			}			
		}

		private void cmbTarjeta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTarjeta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (this.cmbTarjeta.Value != DBNull.Value)
				{
					this.txtNumero.Focus();
				}
				else
				{
					MessageBox.Show("Seleccione una tarjeta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.cmbTarjeta.Focus();
				}
			}
		}

		private void cmbBanco_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (this.cmbBanco.Value != DBNull.Value)
				{
					if (IFormaDePago == 3 || IFormaDePago == 11)this.cmbTarjeta.Focus();
					else if (IFormaDePago == 2)this.txtNumero.Focus();
				}
				else
				{
					MessageBox.Show("Escoja un banco", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.cmbBanco.Focus();
				}
			}
		}

		private void txtCuentaCheque_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13)
			{
				if (this.txtCuentaCheque.Value != DBNull.Value)
				{
					this.btGrabar.Focus();
				}
				else
				{
					MessageBox.Show("Escriba el número de la cuenta.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtCuentaCheque.Focus();
				}
			}
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.txtNumVoucher.Focus();
			}
		}

		private void txtNumVoucher_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13) this.btGrabar.Focus();
		}

		private void txtAutorizacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13)
			{
				this.txtLote.Focus();
				//				if (this.txt.Value != DBNull.Value)
				//				{
				//					this.rdgCorrDiferido.Focus();
				//				}
				//				else
				//				{
				//					MessageBox.Show("Escriba el número de autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				//					return;
				//				}
			}
		}

		private void rdgCorrDiferido_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
		}

		private void cmbPlan_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			
		}

		private void rdgInteres_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbPlan_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
		}

		private void rdgInteres_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private void txtCuentaCheque_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumVoucher_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtAutorizacion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btBanco_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ChequeTCLot_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
		}

		private void cmbCuenta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (this.cmbCuenta.Value != DBNull.Value)
				{
					this.txtNumero.Focus();
				}
				else
				{
					MessageBox.Show("Escoja una cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.cmbCuenta.Focus();
				}
			}
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Caducidad.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				e.Cancel = true;
			}
		}

		private void ChequeTCLot_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
			if (e.KeyCode == Keys.F12) this.btGrabar_Click(sender, e);
		}

		private void txtLote_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13) this.txtNumVoucher.Focus();
		}

		private void txtReferencia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13)
			{
				if (this.cmbPlanes.Enabled) this.cmbPlanes.Focus(); else this.btGrabar.Focus();
			}
		}

		private void ultraCombo1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBanco_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow != null)
			{
				this.cmbTarjeta.Value = System.DBNull.Value;
				this.cmbPlanes.Value = System.DBNull.Value;
				this.cmbPlazo.Value = System.DBNull.Value;
				
				string sSQL = string.Format("Exec BTPP {0}, 0, 0, 0", (int)this.cmbBanco.Value);
				this.cmbTarjeta.DataSource = Funcion.dvProcedimiento(cdsBanco, sSQL);
				this.cmbTarjeta.DisplayLayout.Bands[0].Columns["idTarjeta"].Hidden = true;
				this.cmbTarjeta.DisplayLayout.Bands[0].Columns["Nombre"].Width = 224;
			}
		}

		private void cmbPlazo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (this.cmbPlazo.Value != DBNull.Value)
				{
					this.btGrabar.Focus();
				}
				else
				{
					MessageBox.Show("Seleccione el plazo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.cmbPlazo.Focus();
				}
			}
		}

		private void cmbTarjeta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow != null)
			{
				if (this.cmbTarjeta.ActiveRow != null)
				{
					//					this.cmbPlanes.Value = System.DBNull.Value;
					//					this.cmbPlazo.Value = System.DBNull.Value;
					//
					//					string sSQL = string.Format("Exec BTPP {0}, {1}, 0, 1", (int)this.cmbBanco.Value, (int)this.cmbTarjeta.Value);
					//					this.cmbPlanes.DataSource = Funcion.dvProcedimiento(cdsBanco, sSQL);
					//					this.cmbPlanes.DisplayLayout.Bands[0].Columns["idPlan"].Hidden = true;
					//					this.cmbPlanes.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 224;
				}
			}
		}

		
		private void cmbPlanes_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbPlanes_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow != null) 
			{
				if (this.cmbTarjeta.ActiveRow != null) 
				{
					if (this.cmbPlanes.ActiveRow != null) 
					{
						this.cmbPlazo.Value = System.DBNull.Value;

						string sSQL = string.Format("Exec BTPP {0}, {1}, {2}, 2", (int)this.cmbBanco.Value, (int)this.cmbTarjeta.Value, (int)this.cmbPlanes.Value);
						this.cmbPlazo.DataSource = Funcion.dvProcedimiento(cdsBanco, sSQL);
						this.cmbPlazo.DisplayLayout.Bands[0].Columns["idPlazo"].Hidden = true;
						this.cmbPlazo.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 224;
					}
				}
			}
		}

		private void cmbPlanes_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (this.cmbPlanes.Value != DBNull.Value)
				{
					this.cmbPlazo.Focus();
				}
				else
				{
					MessageBox.Show("Seleccione un plan", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.cmbPlanes.Focus();
				}
			}
		}

		private void cmbBanco_Enter(object sender, System.EventArgs e)
		{
			this.cmbBanco.ToggleDropdown();
		}

		private void cmbCuenta_Enter(object sender, System.EventArgs e)
		{
			this.cmbCuenta.ToggleDropdown();
		}

		private void cmbTarjeta_Enter(object sender, System.EventArgs e)
		{
			this.cmbTarjeta.ToggleDropdown();
		}

		private void cmbPlanes_Enter(object sender, System.EventArgs e)
		{
			this.cmbPlanes.ToggleDropdown();
		}

		private void cmbPlazo_Enter(object sender, System.EventArgs e)
		{
			this.cmbPlazo.ToggleDropdown();
		}

		private void cmbCuenta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{	
			if (!this.Validate())
			{
				return;
			}

			#region Validacion
			#region Cheque 
			if (IFormaDePago == 2)
			{				
				if (this.cmbBanco.ActiveRow == null)
				{
					MessageBox.Show("Seleccione el banco", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbBanco.Focus();
					return;
				}
				if (this.txtCuentaCheque.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el número de cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCuentaCheque.Focus();
					return;					
				}
				if (this.txtNumero.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el número de cheque", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;					
				}
			}
			#endregion Cheque

			#region Tarjeta de Crédito/Debito
			if (IFormaDePago == 3 || IFormaDePago == 11)
			{
				//				if (this.cmbBanco.ActiveRow == null)
				//				{
				//					MessageBox.Show("Seleccione el Banco", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//					this.cmbBanco.Focus();
				//					return;										
				//				}

				if (this.optTarjeta.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la Tarjeta de Crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTarjeta.Focus();
					return;					
				}				

				//				if (this.dtFecha.Value == System.DBNull.Value)
				//				{					
				//					MessageBox.Show("Ingrese la Fecha de Caducidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//					this.dtFecha.Focus();
				//					return;										
				//				}

				//				DateTime Fecha = (DateTime) this.dtFecha.Value;
				//				int Mes = DateTime.Now.Month;
				//				int Anio = DateTime.Now.Year;
				//				int Mes1 = Fecha.Month;
				//				int Anio1 = Fecha.Year;
				//
				//				if ((Mes1 < Mes && Anio1 <= Anio) || (Mes1 > Mes && Anio1 < Anio)) //Fecha.Date.CompareTo(DateTime.Now)==1)
				//				{	
				//					MessageBox.Show("La Tarjeta esta caducada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//					this.dtFecha.Focus();
				//					return;							
				//				}

				//				if (this.txtNumero.Text.ToString().Length == 0)
				//				{
				//					MessageBox.Show("Ingrese el Número de Tarjeta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//					this.txtNumero.Focus();
				//					return;												
				//				}

				if (this.txtLote.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Lote de Voucher", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtLote.Focus();
					return;
				}

				//				if (this.txtReferencia.ToString().Length == 0)
				//				{
				//					MessageBox.Show("Ingrese la Referencia del Voucher", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//					this.txtReferencia.Focus();
				//					return;					
				//				}

				if (this.optPlan.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione el Plan", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbPlanes.Focus();
					return;							
				}

				//				if (this.cmbPlazo.ActiveRow == null)
				//				{
				//					MessageBox.Show("Seleccione el Plazo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//					this.cmbPlazo.Focus();
				//					return;							
				//				}
				if (this.txtNumVoucher.Text == "")
				{
					MessageBox.Show("Ingrese el número de voucher", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);									
					this.txtNumVoucher.Focus();
					return;
				}
				//				if (this.txtAutorizacion.Text == "")
				//				{
				//					this.errorProvider1.SetError(this.txtAutorizacion, "Ingrese la autorización del voucher");
				//					bOk = false;
				//				}
			}
			#endregion Tarjeta de crédito

			#region Transferencia - Deposito
			if (IFormaDePago == 7 || IFormaDePago == 13)
			{				
				if (this.cmbCuenta.ActiveRow == null)
				{
					MessageBox.Show("Seleccione una Cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbCuenta.Focus();
					return;
				}
				if (this.txtNumero.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Número de Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;					
				}
			}
			#endregion Transferencia - Deposito

			if (IFormaDePago == 16)
			{
				if (this.txtDesPrenda.Text.Length <= 20)
				{
					MessageBox.Show("Ingrese el concepto de la prenda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;	
				}
			}
			#endregion Validacion

			this.DialogResult = DialogResult.OK;			
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;			
		}

		private void cmbBanco_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbPlazo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtDesPrenda_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btGrabar.Focus();
		}

		private void txtLote_Leave(object sender, System.EventArgs e)
		{
			string texto = txtLote.Text.Trim();

			bool sonDigitos = true;
			for (int i = 0; i < texto.Length; i++)
			{
				if (!Char.IsDigit(texto[i]))
				{
					sonDigitos = false;
					break;
				}
			}

			if (sonDigitos && texto.Length > 0 && texto.Length < 4)
			{
				txtLote.Text = texto.PadLeft(4, '0');
			}
		}

		private void txtNumVoucher_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			#region Validación de voucher duplicado

			string sLote = txtLote.Text.Trim();
			string sNumVoucher = txtNumVoucher.Text.Trim();

			string loteConsulta = sLote;
			if (sLote.Length == 4 && sLote.StartsWith("0"))
			{
				loteConsulta = sLote.Substring(1); 
			}

			if ((sLote != null && sLote != "") && (sNumVoucher != null && sNumVoucher != ""))
			{
				string sSQLValida = string.Format("Exec ValidaLoteExistente '{0}','{1}'", loteConsulta, sNumVoucher);
				int contador = Convert.ToInt32(Funcion.iEscalarSQL(cdsBanco, sSQLValida));

				if (contador > 0)
				{
					MessageBox.Show("Ya existe un voucher registrado con ese número para el lote ingresado.","Duplicado",	MessageBoxButtons.OK,MessageBoxIcon.Warning);

					txtNumVoucher.Text = "";
					e.Cancel = true;
					return;
				}
			}

			#endregion
		}

		private void txtNumVoucher_Leave(object sender, System.EventArgs e)
		{
			string texto = txtNumVoucher.Text.Trim();

			bool sonDigitos = true;
			for (int i = 0; i < texto.Length; i++)
			{
				if (!Char.IsDigit(texto[i]))
				{
					sonDigitos = false;
					break;
				}
			}

			if (sonDigitos && texto.Length > 0 && texto.Length < 4)
			{
				txtNumVoucher.Text = texto.PadLeft(4, '0');
			}
		}
	}
}

