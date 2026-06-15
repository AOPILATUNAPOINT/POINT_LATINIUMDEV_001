using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmInformacionFormasDePago.
	/// </summary>
	public class frmInformacionFormasDePago : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsBanco;
		private System.Windows.Forms.Label lblPlanes;
		private System.Windows.Forms.Label lblReferencia;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btCancelar;
		private System.Windows.Forms.Button btGrabar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbPlanes;
		private System.Windows.Forms.Label lblLote;
		private System.Windows.Forms.Label lblCuenta;
		private System.Windows.Forms.Label lblPlazo;
		private System.Windows.Forms.Label lblAutorizacion;
		private System.Windows.Forms.Label lblVoucher;
		private System.Windows.Forms.Label lblTarjeta;
		private System.Windows.Forms.Label lblBanco;
		private System.Windows.Forms.Label lblCtaCorriente;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label lblCaducidad;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbPlazo;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtReferencia;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtLote;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumVoucher;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutorizacion;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTarjeta;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuentaCheque;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmInformacionFormasDePago()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

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
		int IdPlazo = 0;

		public frmInformacionFormasDePago(int iFormaDePago, int idBanco, int idTarjeta, int idCuenta, string sNumero, DateTime dtFecha, string sCuentaCorriente, string sVoucher, string sAutorizacion, string sLote, string sReferencia, int idPlan, int idPlazo)
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmInformacionFormasDePago));
			this.cdsBanco = new C1.Data.C1DataSet();
			this.lblPlanes = new System.Windows.Forms.Label();
			this.lblReferencia = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btCancelar = new System.Windows.Forms.Button();
			this.btGrabar = new System.Windows.Forms.Button();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbPlanes = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblLote = new System.Windows.Forms.Label();
			this.lblCuenta = new System.Windows.Forms.Label();
			this.lblPlazo = new System.Windows.Forms.Label();
			this.lblAutorizacion = new System.Windows.Forms.Label();
			this.lblVoucher = new System.Windows.Forms.Label();
			this.lblTarjeta = new System.Windows.Forms.Label();
			this.lblBanco = new System.Windows.Forms.Label();
			this.lblCtaCorriente = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.lblCaducidad = new System.Windows.Forms.Label();
			this.cmbPlazo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtLote = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNumVoucher = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtAutorizacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbTarjeta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtCuentaCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumVoucher)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentaCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
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
			// lblPlanes
			// 
			this.lblPlanes.AutoSize = true;
			this.lblPlanes.BackColor = System.Drawing.Color.Transparent;
			this.lblPlanes.Location = new System.Drawing.Point(10, 229);
			this.lblPlanes.Name = "lblPlanes";
			this.lblPlanes.Size = new System.Drawing.Size(39, 16);
			this.lblPlanes.TabIndex = 202;
			this.lblPlanes.Text = "Planes";
			// 
			// lblReferencia
			// 
			this.lblReferencia.AutoSize = true;
			this.lblReferencia.BackColor = System.Drawing.Color.Transparent;
			this.lblReferencia.Location = new System.Drawing.Point(153, 197);
			this.lblReferencia.Name = "lblReferencia";
			this.lblReferencia.Size = new System.Drawing.Size(59, 16);
			this.lblReferencia.TabIndex = 201;
			this.lblReferencia.Text = "Referencia";
			// 
			// ultraDataSource1
			// 
			this.ultraDataSource1.Band.Key = "TarjetasDeCredito";
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.Location = new System.Drawing.Point(181, 301);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(91, 23);
			this.btCancelar.TabIndex = 190;
			this.btCancelar.Text = "&Cancelar";
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.btGrabar.Location = new System.Drawing.Point(77, 301);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(91, 23);
			this.btGrabar.TabIndex = 189;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbPlanes
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPlanes.Appearance = appearance1;
			this.cmbPlanes.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPlanes.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlanes.DisplayMember = "Descripcion";
			this.cmbPlanes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPlanes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlanes.Location = new System.Drawing.Point(113, 229);
			this.cmbPlanes.Name = "cmbPlanes";
			this.cmbPlanes.Size = new System.Drawing.Size(224, 21);
			this.cmbPlanes.TabIndex = 187;
			this.cmbPlanes.ValueMember = "idPlan";
			this.cmbPlanes.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbPlanes_KeyPress);
			this.cmbPlanes.ValueChanged += new System.EventHandler(this.cmbPlanes_ValueChanged);
			// 
			// lblLote
			// 
			this.lblLote.AutoSize = true;
			this.lblLote.BackColor = System.Drawing.Color.Transparent;
			this.lblLote.Location = new System.Drawing.Point(9, 197);
			this.lblLote.Name = "lblLote";
			this.lblLote.Size = new System.Drawing.Size(26, 16);
			this.lblLote.TabIndex = 200;
			this.lblLote.Text = "Lote";
			// 
			// lblCuenta
			// 
			this.lblCuenta.AutoSize = true;
			this.lblCuenta.BackColor = System.Drawing.Color.Transparent;
			this.lblCuenta.Location = new System.Drawing.Point(9, 37);
			this.lblCuenta.Name = "lblCuenta";
			this.lblCuenta.Size = new System.Drawing.Size(41, 16);
			this.lblCuenta.TabIndex = 199;
			this.lblCuenta.Text = "Cuenta";
			this.lblCuenta.Visible = false;
			// 
			// lblPlazo
			// 
			this.lblPlazo.AutoSize = true;
			this.lblPlazo.BackColor = System.Drawing.Color.Transparent;
			this.lblPlazo.Location = new System.Drawing.Point(9, 261);
			this.lblPlazo.Name = "lblPlazo";
			this.lblPlazo.Size = new System.Drawing.Size(33, 16);
			this.lblPlazo.TabIndex = 198;
			this.lblPlazo.Text = "Plazo";
			// 
			// lblAutorizacion
			// 
			this.lblAutorizacion.AutoSize = true;
			this.lblAutorizacion.BackColor = System.Drawing.Color.Transparent;
			this.lblAutorizacion.Location = new System.Drawing.Point(9, 165);
			this.lblAutorizacion.Name = "lblAutorizacion";
			this.lblAutorizacion.Size = new System.Drawing.Size(67, 16);
			this.lblAutorizacion.TabIndex = 197;
			this.lblAutorizacion.Text = "Autorización";
			// 
			// lblVoucher
			// 
			this.lblVoucher.AutoSize = true;
			this.lblVoucher.BackColor = System.Drawing.Color.Transparent;
			this.lblVoucher.Location = new System.Drawing.Point(9, 133);
			this.lblVoucher.Name = "lblVoucher";
			this.lblVoucher.Size = new System.Drawing.Size(46, 16);
			this.lblVoucher.TabIndex = 196;
			this.lblVoucher.Text = "Voucher";
			// 
			// lblTarjeta
			// 
			this.lblTarjeta.AutoSize = true;
			this.lblTarjeta.BackColor = System.Drawing.Color.Transparent;
			this.lblTarjeta.Location = new System.Drawing.Point(9, 37);
			this.lblTarjeta.Name = "lblTarjeta";
			this.lblTarjeta.Size = new System.Drawing.Size(39, 16);
			this.lblTarjeta.TabIndex = 192;
			this.lblTarjeta.Text = "Tarjeta";
			// 
			// lblBanco
			// 
			this.lblBanco.AutoSize = true;
			this.lblBanco.BackColor = System.Drawing.Color.Transparent;
			this.lblBanco.Location = new System.Drawing.Point(9, 5);
			this.lblBanco.Name = "lblBanco";
			this.lblBanco.Size = new System.Drawing.Size(36, 16);
			this.lblBanco.TabIndex = 194;
			this.lblBanco.Text = "Banco";
			// 
			// lblCtaCorriente
			// 
			this.lblCtaCorriente.AutoSize = true;
			this.lblCtaCorriente.BackColor = System.Drawing.Color.Transparent;
			this.lblCtaCorriente.Location = new System.Drawing.Point(9, 101);
			this.lblCtaCorriente.Name = "lblCtaCorriente";
			this.lblCtaCorriente.Size = new System.Drawing.Size(90, 16);
			this.lblCtaCorriente.TabIndex = 195;
			this.lblCtaCorriente.Text = "Cuenta Corriente";
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(9, 69);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 193;
			this.lblNumero.Text = "Número";
			// 
			// lblCaducidad
			// 
			this.lblCaducidad.AutoSize = true;
			this.lblCaducidad.Location = new System.Drawing.Point(9, 101);
			this.lblCaducidad.Name = "lblCaducidad";
			this.lblCaducidad.Size = new System.Drawing.Size(93, 16);
			this.lblCaducidad.TabIndex = 191;
			this.lblCaducidad.Text = "Fecha Caducidad";
			// 
			// cmbPlazo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPlazo.Appearance = appearance2;
			this.cmbPlazo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPlazo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlazo.DisplayMember = "Descripcion";
			this.cmbPlazo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPlazo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlazo.Location = new System.Drawing.Point(113, 261);
			this.cmbPlazo.Name = "cmbPlazo";
			this.cmbPlazo.Size = new System.Drawing.Size(224, 21);
			this.cmbPlazo.TabIndex = 188;
			this.cmbPlazo.ValueMember = "idPlazo";
			this.cmbPlazo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbPlazo_KeyPress);
			// 
			// txtReferencia
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtReferencia.Appearance = appearance3;
			this.txtReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtReferencia.Location = new System.Drawing.Point(225, 197);
			this.txtReferencia.MaxLength = 15;
			this.txtReferencia.Name = "txtReferencia";
			this.txtReferencia.Size = new System.Drawing.Size(88, 21);
			this.txtReferencia.TabIndex = 186;
			this.txtReferencia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtReferencia_KeyPress);
			// 
			// txtLote
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLote.Appearance = appearance4;
			this.txtLote.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLote.Location = new System.Drawing.Point(49, 197);
			this.txtLote.MaxLength = 15;
			this.txtLote.Name = "txtLote";
			this.txtLote.Size = new System.Drawing.Size(88, 21);
			this.txtLote.TabIndex = 185;
			this.txtLote.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLote_KeyDown);
			this.txtLote.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtLote_KeyPress);
			// 
			// cmbCuenta
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance5;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(113, 37);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(224, 21);
			this.cmbCuenta.TabIndex = 179;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.Visible = false;
			this.cmbCuenta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbCuenta_KeyPress);
			// 
			// txtNumVoucher
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumVoucher.Appearance = appearance6;
			this.txtNumVoucher.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumVoucher.Location = new System.Drawing.Point(113, 133);
			this.txtNumVoucher.MaxLength = 20;
			this.txtNumVoucher.Name = "txtNumVoucher";
			this.txtNumVoucher.Size = new System.Drawing.Size(200, 21);
			this.txtNumVoucher.TabIndex = 183;
			this.txtNumVoucher.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumVoucher_KeyPress);
			// 
			// txtAutorizacion
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutorizacion.Appearance = appearance7;
			this.txtAutorizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutorizacion.Location = new System.Drawing.Point(113, 165);
			this.txtAutorizacion.MaxLength = 20;
			this.txtAutorizacion.Name = "txtAutorizacion";
			this.txtAutorizacion.Size = new System.Drawing.Size(200, 21);
			this.txtAutorizacion.TabIndex = 184;
			this.txtAutorizacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAutorizacion_KeyPress);
			// 
			// cmbTarjeta
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTarjeta.Appearance = appearance8;
			this.cmbTarjeta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTarjeta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTarjeta.DisplayMember = "Nombre";
			this.cmbTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTarjeta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTarjeta.Location = new System.Drawing.Point(113, 37);
			this.cmbTarjeta.Name = "cmbTarjeta";
			this.cmbTarjeta.Size = new System.Drawing.Size(224, 21);
			this.cmbTarjeta.TabIndex = 178;
			this.cmbTarjeta.ValueMember = "idTarjeta";
			this.cmbTarjeta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTarjeta_KeyPress);
			this.cmbTarjeta.ValueChanged += new System.EventHandler(this.cmbTarjeta_ValueChanged);
			// 
			// cmbBanco
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance9;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 224;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(113, 5);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(224, 21);
			this.cmbBanco.TabIndex = 177;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBanco_KeyPress);
			this.cmbBanco.ValueChanged += new System.EventHandler(this.cmbBanco_ValueChanged);
			// 
			// txtCuentaCheque
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuentaCheque.Appearance = appearance10;
			this.txtCuentaCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuentaCheque.Location = new System.Drawing.Point(113, 101);
			this.txtCuentaCheque.MaxLength = 15;
			this.txtCuentaCheque.Name = "txtCuentaCheque";
			this.txtCuentaCheque.Size = new System.Drawing.Size(200, 21);
			this.txtCuentaCheque.TabIndex = 182;
			this.txtCuentaCheque.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCuentaCheque_KeyPress);
			// 
			// txtNumero
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance11;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(113, 69);
			this.txtNumero.MaxLength = 20;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(200, 21);
			this.txtNumero.TabIndex = 180;
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			// 
			// dtFecha
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance12;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(113, 101);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 181;
			this.dtFecha.TrailingDaysVisible = false;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			// 
			// frmInformacionFormasDePago
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(346, 328);
			this.ControlBox = false;
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.cmbPlanes);
			this.Controls.Add(this.lblLote);
			this.Controls.Add(this.lblCuenta);
			this.Controls.Add(this.lblPlazo);
			this.Controls.Add(this.lblAutorizacion);
			this.Controls.Add(this.lblVoucher);
			this.Controls.Add(this.lblTarjeta);
			this.Controls.Add(this.lblBanco);
			this.Controls.Add(this.lblCtaCorriente);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.lblCaducidad);
			this.Controls.Add(this.cmbPlazo);
			this.Controls.Add(this.txtReferencia);
			this.Controls.Add(this.txtLote);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.txtNumVoucher);
			this.Controls.Add(this.txtAutorizacion);
			this.Controls.Add(this.cmbTarjeta);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.txtCuentaCheque);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.lblPlanes);
			this.Controls.Add(this.lblReferencia);
			this.Controls.Add(this.btCancelar);
			this.KeyPreview = true;
			this.Name = "frmInformacionFormasDePago";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Tarjeta de Crédito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmInformacionFormasDePago_KeyDown);
			this.Load += new System.EventHandler(this.frmInformacionFormasDePago_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumVoucher)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentaCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmInformacionFormasDePago_Load(object sender, System.EventArgs e)
		{
			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsBanco, "Select idBanco, Nombre From Banco Where Caja = 1 Order By Nombre");

			if (IFormaDePago == 2)
			{
				this.Text = "CHEQUE";		
		
				#region Controles
				this.lblTarjeta.Visible = false;
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
			else if (IFormaDePago == 3)
			{
				this.Text = "TARJETA DE CRÉDITO";

				this.dtFecha.CalendarInfo.MinDate = DateTime.Today;

				#region Controles
				this.lblCtaCorriente.Visible = false;
				this.txtCuentaCheque.Visible = false;
				this.lblCuenta.Visible = false;
				this.cmbCuenta.Visible = false;
				#endregion Controles

				if (IdBanco > 0) this.cmbBanco.Value = IdBanco;
				if (IdTarjeta > 0)	this.cmbTarjeta.Value = IdTarjeta;
				if (SNumero.Length > 0) this.txtNumero.Text = SNumero;
				this.dtFecha.Value = (DateTime)DTFecha;
				if (SVoucher.Length > 0) this.txtNumVoucher.Text = SVoucher;
				if (SAutorizacion.Length > 0) this.txtAutorizacion.Text = SAutorizacion;
				if (SLote.Length > 0) this.txtLote.Text = SLote;
				if (SReferencia.Length > 0) this.txtReferencia.Text = SReferencia;
				if (IdPlan > 0) this.cmbPlanes.Value = IdPlan;
				if (IdPlazo > 0) this.cmbPlazo.Value = IdPlazo;

				if (IdBanco > 0)
				{
					this.cmbBanco.Enabled = false;
					this.cmbTarjeta.Enabled = false;
					this.cmbPlanes.Enabled = false;
					this.cmbPlazo.Enabled = false;

					this.txtNumero.Focus();
				}
				else this.cmbBanco.Focus();
			}	
			else if (IFormaDePago == 7 || IFormaDePago == 13)
			{
				this.Text = "TRANSFERENCIA - DEPOSITOS";

				#region Controles
				this.lblTarjeta.Visible = false;
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
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			
			if (e.KeyChar == 13)
			{
				if  (IFormaDePago == 2) this.txtCuentaCheque.Focus();
				else if (IFormaDePago == 3) this.dtFecha.Focus();
				else if (IFormaDePago == 7) this.btGrabar.Focus();
				else if (IFormaDePago == 11) this.txtLote.Focus();
			}	
		}

		private void cmbTarjeta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNumero.Focus();
		}

		private void cmbBanco_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (IFormaDePago == 3 || IFormaDePago == 11) this.cmbTarjeta.Focus();
				else if (IFormaDePago == 2) this.txtNumero.Focus();
			}
		}

		private void txtCuentaCheque_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.btGrabar.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNumVoucher.Focus();
		}

		private void txtNumVoucher_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13) this.txtAutorizacion.Focus();
		}

		private void txtAutorizacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13) this.txtLote.Focus();
		}

		private void cmbCuenta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNumero.Focus();
		}

		private void frmInformacionFormasDePago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
			if (e.KeyCode == Keys.F12) this.btGrabar_Click(sender, e);
		}

		private void txtLote_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
		}

		private void txtReferencia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13) 
			{
				if (this.cmbPlanes.Enabled) this.cmbPlanes.Focus(); else this.btGrabar.Focus();
			}
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
			if (e.KeyChar == 13) this.btGrabar.Focus();
		}

		private void cmbTarjeta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow != null)
			{
				if (this.cmbTarjeta.ActiveRow != null)
				{
					this.cmbPlanes.Value = System.DBNull.Value;
					this.cmbPlazo.Value = System.DBNull.Value;

					string sSQL = string.Format("Exec BTPP {0}, {1}, 0, 1", (int)this.cmbBanco.Value, (int)this.cmbTarjeta.Value);
					this.cmbPlanes.DataSource = Funcion.dvProcedimiento(cdsBanco, sSQL);
					this.cmbPlanes.DisplayLayout.Bands[0].Columns["idPlan"].Hidden = true;
					this.cmbPlanes.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 224;
				}
			}
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
			if (e.KeyChar == 13) this.cmbPlazo.Focus();
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtLote_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13) this.txtReferencia.Focus();
		}
	}
}
