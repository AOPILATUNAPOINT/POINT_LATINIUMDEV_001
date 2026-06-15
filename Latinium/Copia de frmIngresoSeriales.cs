using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmIngresoSeriales.
	/// </summary>
	public class frmIngresoSeriales : DevExpress.XtraEditors.XtraForm
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnLlegaron;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerie;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		int IdCompra;
		public int IdTipoFactura = 0;
		string SNumero = "";
		int IdDetCompra;		
		double DCantidad = 0; 
		int IdArticulo = 0;
		string SCodigo = "";
		string SArticulo = "";
		int IBodega = 0;
		private C1.Data.C1DataSet cdsSeteoF;		
		int IEstado = 0;
		bool bSerialNoExiste = false;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.Misc.UltraLabel lblContador;
		private Infragistics.Win.Misc.UltraLabel lblCodigo;
		private Infragistics.Win.Misc.UltraLabel lblArticulo;
		private Infragistics.Win.Misc.UltraLabel lblNumero;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDesde;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtHasta;
		bool ModSerial = false;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnGenerarSeriales;
		private System.Windows.Forms.Button btnGeneraSeriales;
		private Acceso miAcceso;

		public frmIngresoSeriales(int idCompra, int idTipoFactura, string sNumero, int idDetCompra, double dCantidad, int idArticulo, string sCodigo, string sArticulo, int iBodega, int iEstado)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			IdCompra = idCompra;
			IdTipoFactura = idTipoFactura;
			SNumero = sNumero;
			IdDetCompra = idDetCompra;
			DCantidad = dCantidad;
			IdArticulo = idArticulo;	
			SCodigo = sCodigo;
			SArticulo = sArticulo;
			IBodega = iBodega;
			IEstado = iEstado;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasDelSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasDelSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblCodigo = new Infragistics.Win.Misc.UltraLabel();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblArticulo = new Infragistics.Win.Misc.UltraLabel();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblNumero = new Infragistics.Win.Misc.UltraLabel();
			this.spnLlegaron = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.txtSerie = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblContador = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.txtDesde = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.txtHasta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnGenerarSeriales = new System.Windows.Forms.Button();
			this.btnGeneraSeriales = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnLlegaron)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "Serie";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 66;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 58;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 47;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 270;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 500;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 46;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 82;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8.5F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 112);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(791, 216);
			this.ultraGrid1.TabIndex = 2;
			this.ultraGrid1.DoubleClick += new System.EventHandler(this.ultraGrid1_DoubleClick);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			this.ultraGrid1.Click += new System.EventHandler(this.ultraGrid1_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// lblCodigo
			// 
			this.lblCodigo.AutoSize = true;
			this.lblCodigo.Location = new System.Drawing.Point(8, 12);
			this.lblCodigo.Name = "lblCodigo";
			this.lblCodigo.Size = new System.Drawing.Size(38, 15);
			this.lblCodigo.TabIndex = 18;
			this.lblCodigo.Text = "Código";
			// 
			// txtArticulo
			// 
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance7;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Enabled = false;
			this.txtArticulo.Location = new System.Drawing.Point(56, 32);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(392, 22);
			this.txtArticulo.TabIndex = 17;
			// 
			// lblArticulo
			// 
			this.lblArticulo.AutoSize = true;
			this.lblArticulo.Location = new System.Drawing.Point(8, 36);
			this.lblArticulo.Name = "lblArticulo";
			this.lblArticulo.Size = new System.Drawing.Size(42, 15);
			this.lblArticulo.TabIndex = 16;
			this.lblArticulo.Text = "Artículo";
			// 
			// txtCodigo
			// 
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance8;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(56, 8);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(392, 22);
			this.txtCodigo.TabIndex = 15;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.Location = new System.Drawing.Point(8, 68);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 15);
			this.lblNumero.TabIndex = 157;
			this.lblNumero.Text = "Número";
			// 
			// spnLlegaron
			// 
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnLlegaron.Appearance = appearance9;
			this.spnLlegaron.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnLlegaron.Enabled = false;
			this.spnLlegaron.Location = new System.Drawing.Point(56, 64);
			this.spnLlegaron.MinValue = -2147483647;
			this.spnLlegaron.Name = "spnLlegaron";
			this.spnLlegaron.PromptChar = ' ';
			this.spnLlegaron.Size = new System.Drawing.Size(96, 22);
			this.spnLlegaron.TabIndex = 156;
			this.spnLlegaron.Click += new System.EventHandler(this.spnLlegaron_Click);
			this.spnLlegaron.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.spnLlegaron_KeyPress);
			this.spnLlegaron.Validating += new System.ComponentModel.CancelEventHandler(this.spnLlegaron_Validating);
			this.spnLlegaron.Enter += new System.EventHandler(this.spnLlegaron_Enter);
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(168, 68);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(32, 15);
			this.ultraLabel3.TabIndex = 155;
			this.ultraLabel3.Text = "Serial";
			// 
			// txtSerie
			// 
			this.txtSerie.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerie.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerie.Location = new System.Drawing.Point(208, 64);
			this.txtSerie.Name = "txtSerie";
			this.txtSerie.Size = new System.Drawing.Size(240, 22);
			this.txtSerie.TabIndex = 0;
			this.txtSerie.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerie_KeyPress);
			this.txtSerie.ValueChanged += new System.EventHandler(this.txtSerie_ValueChanged);
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
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.Location = new System.Drawing.Point(544, 67);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 0);
			this.lblContador.TabIndex = 158;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(464, 12);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(35, 15);
			this.ultraLabel1.TabIndex = 160;
			this.ultraLabel1.Text = "Desde";
			this.ultraLabel1.Visible = false;
			// 
			// txtDesde
			// 
			appearance10.ForeColor = System.Drawing.Color.Black;
			this.txtDesde.Appearance = appearance10;
			this.txtDesde.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDesde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDesde.Location = new System.Drawing.Point(520, 8);
			this.txtDesde.Name = "txtDesde";
			this.txtDesde.Size = new System.Drawing.Size(272, 22);
			this.txtDesde.TabIndex = 6;
			this.txtDesde.Visible = false;
			this.txtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDesde_KeyDown);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(464, 36);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(48, 15);
			this.ultraLabel2.TabIndex = 162;
			this.ultraLabel2.Text = "Cantidad";
			this.ultraLabel2.Visible = false;
			// 
			// txtHasta
			// 
			appearance11.ForeColor = System.Drawing.Color.Black;
			this.txtHasta.Appearance = appearance11;
			this.txtHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHasta.Location = new System.Drawing.Point(520, 32);
			this.txtHasta.Name = "txtHasta";
			this.txtHasta.Size = new System.Drawing.Size(120, 22);
			this.txtHasta.TabIndex = 161;
			this.txtHasta.Visible = false;
			this.txtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtHasta_KeyDown);
			this.txtHasta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtHasta_KeyPress);
			this.txtHasta.ValueChanged += new System.EventHandler(this.txtHasta_ValueChanged);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(8, 336);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(72, 23);
			this.btnAceptar.TabIndex = 164;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(88, 336);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(78, 23);
			this.btnCancelar.TabIndex = 165;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 96);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(808, 8);
			this.groupBox1.TabIndex = 166;
			this.groupBox1.TabStop = false;
			// 
			// btnGenerarSeriales
			// 
			this.btnGenerarSeriales.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGenerarSeriales.CausesValidation = false;
			this.btnGenerarSeriales.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGenerarSeriales.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGenerarSeriales.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerarSeriales.Location = new System.Drawing.Point(208, 64);
			this.btnGenerarSeriales.Name = "btnGenerarSeriales";
			this.btnGenerarSeriales.Size = new System.Drawing.Size(104, 23);
			this.btnGenerarSeriales.TabIndex = 672;
			this.btnGenerarSeriales.Text = "Generar Seriales";
			this.btnGenerarSeriales.Visible = false;
			this.btnGenerarSeriales.Click += new System.EventHandler(this.btnGenerarSeriales_Click);
			// 
			// btnGeneraSeriales
			// 
			this.btnGeneraSeriales.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGeneraSeriales.CausesValidation = false;
			this.btnGeneraSeriales.Enabled = false;
			this.btnGeneraSeriales.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGeneraSeriales.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGeneraSeriales.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGeneraSeriales.Location = new System.Drawing.Point(688, 32);
			this.btnGeneraSeriales.Name = "btnGeneraSeriales";
			this.btnGeneraSeriales.Size = new System.Drawing.Size(104, 23);
			this.btnGeneraSeriales.TabIndex = 673;
			this.btnGeneraSeriales.Text = "Generar";
			this.btnGeneraSeriales.Visible = false;
			this.btnGeneraSeriales.Click += new System.EventHandler(this.btnGeneraSeriales_Click);
			// 
			// frmIngresoSeriales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(807, 368);
			this.ControlBox = false;
			this.Controls.Add(this.btnGeneraSeriales);
			this.Controls.Add(this.btnGenerarSeriales);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.txtHasta);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtDesde);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.spnLlegaron);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.txtSerie);
			this.Controls.Add(this.lblCodigo);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.lblArticulo);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.ultraGrid1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmIngresoSeriales";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Registro de Seriales";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmIngresoSeriales_KeyDown);
			this.Load += new System.EventHandler(this.frmIngresoSeriales_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnLlegaron)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iBod = 0;
		int iBodegaTransf = 0;

		private void frmIngresoSeriales_Load(object sender, System.EventArgs e)
		{
			Funcion.Resolucion(this, this.Height, this.Width);

			miAcceso = new Acceso(cdsSeteoF, "0301");

			string stAudita = "Exec AuditaCrear 53, 6, 'Serie'"; 
			Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);

			if (IdTipoFactura == 9) iBodegaTransf = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Bodega From CompraTransf Where idCompra1 = {0}", IdCompra));
			
			if (IdTipoFactura == 6 || IdTipoFactura == 11 || IdTipoFactura == 44) iBod = 1;
			else iBod = IBodega;

			this.txtCodigo.Text = SCodigo;
			this.txtArticulo.Text = SArticulo;

			if (IEstado != 15)			
			{
				string sSQL = string.Format("Exec ConsultaSeriales {0}, {1}, {2}", IdDetCompra, IdTipoFactura, IdCompra);
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			}

			this.spnLlegaron.Value = DCantidad;			
			int iFilas = this.ultraGrid1.Rows.Count;
			this.spnLlegaron.Value = DCantidad - iFilas;
			this.lblContador.Text = "SERIALES: " + iFilas + " DE " + DCantidad;
			
			#region Para Orden de Compra verifica si genera las seriales
			if (IdTipoFactura == 2 || IdTipoFactura == 4 || IdTipoFactura == 36)
			{					
				if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select SeriePropiaAlterna From Articulo Where idArticulo = {0}", IdArticulo), true)) 
				{					
					this.btnGenerarSeriales.Visible = true;
					this.ultraLabel3.Visible = false;
					this.txtSerie.Visible = false;
				}
				
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idImportacionTipo, 0) From Compra Where idCompra = {0}", IdCompra)) > 0)
				{			
					if (Funcion.bEscalarSQL(cdsSeteoF, "Select GeneraSerial From SeteoF", true))
					{
						#region Genera Seriales 
						this.btnGenerarSeriales.Visible = false;
						this.ultraLabel3.Visible = true;
						//this.txtSerie.Visible = true;
						if ((int)this.spnLlegaron.Value == 0) this.btnGeneraSeriales.Enabled = false;
						
						this.ultraLabel1.Visible = true;
						this.ultraLabel2.Visible = true;

						this.txtDesde.Visible = true;
						this.txtHasta.Visible = true;
						this.btnGeneraSeriales.Enabled = true;
						this.btnGeneraSeriales.Visible = true;
						#endregion Genera Seriales
					}					
				}	
			}											
			#endregion Para Orden de Compra verifica si genera seriales			
			
			this.txtSerie.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void spnLlegaron_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			int iFilas = this.ultraGrid1.Rows.Count;
			int CantGenerar = (int)this.spnLlegaron.Value;
			int c = (int)DCantidad - (int)iFilas;

			if (CantGenerar > (DCantidad - iFilas))
			{
				MessageBox.Show("La cantidad de seriales a ingresar no puede ser mayor a " + c + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
			}
		}

		private void spnLlegaron_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
			{
				if (this.txtSerie.Visible && !this.btnGenerarSeriales.Visible) this.txtSerie.Focus();
				else if (!this.txtSerie.Visible && this.btnGenerarSeriales.Visible) this.btnGenerarSeriales.Focus();
			}
		}

		private void txtSerie_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{				
			if (Char.IsControl(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsLetterOrDigit(e.KeyChar))
			{
				e.Handled = false;
			}				
			else
			{
				e.Handled = true;
			}

			if (e.KeyChar == 13)
			{
				int IdSerial = 0;
				int IdDetalleSerial = 0;

				#region Valida ingreso serial
				if (this.txtSerie.Text.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Ingrese la serial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerie.Text = "";
					this.txtSerie.Focus();
					return;
				}
				#endregion Valida ingreso serial
				
				#region Valida Serial Existente en el sistema para ordenes de compra
				if (IdTipoFactura == 2 || IdTipoFactura == 36)
				{
					string sSQL = string.Format("Exec ValidaSerialExisteODC '{0}', {1}", this.txtSerie.Text.ToString(), IdArticulo);
					bool bExisteODC = Funcion.bEscalarSQL(cdsSeteoF, sSQL, true);

					if (Funcion.bEscalarSQL(cdsSeteoF, sSQL, true))
					{
						MessageBox.Show("La Serial ya esta registrada en el Sistema", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						ModSerial = false;
						this.txtSerie.Text = "";
						this.txtSerie.Focus();
						return;
					}
				}
				#endregion Valida Serial Existente en el sistema para ordenes de compra

				#region Cantidad de seriales registradas
				if (!ModSerial && IdTipoFactura != 9 && DCantidad == this.ultraGrid1.Rows.Count)
				{
					MessageBox.Show("Ya ha registrado las seriales para este producto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					ModSerial = false;
					this.txtSerie.Text = "";
					this.txtSerie.Focus();
					return;
				}
				#endregion Cantidad de seriales registradas

				#region Valida Serial existente en la base de datos 
				int Cont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValExistenciaSerial '{0}', {1}, {2}, {3}", this.txtSerie.Text.ToString().Trim(), IdArticulo, IdTipoFactura, IBodega));
				if(Cont > 0)  
				{
					MessageBox.Show("La Serial ya esta registrada en el sistema.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					ModSerial = false;
					this.txtSerie.Text = "";
					this.txtSerie.Focus();
					return;
				}				
				#endregion Valida Serial existente en la base de datos

				#region Valida serial existente en el grid
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					string sSerial = dr.Cells["Serial"].Value.ToString();
					if (this.txtSerie.Text.ToString() == sSerial)
					{
						MessageBox.Show("Ya registro esta serial.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						ModSerial = false;
						this.txtSerie.Text = "";
						this.txtSerie.Focus();
						return;
					}
				}
				#endregion Valida serial existente en el grid

				if (IdTipoFactura == 8)
				{
					string sSQL = string.Format("Exec IRMValidaCantidadMovSeriales '{0}', {1}", this.txtSerie.Text.ToString(), IdCompra);
					if (!Funcion.bEscalarSQL(cdsSeteoF, sSQL, true))
					{
						MessageBox.Show("La serial tiene mas de dos movimientos a locales no puede ser transferida\n\nConsulte con Auditoria", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtSerie.Text = "";
						this.txtSerie.Focus();
						return;
					}
				}

				if (IdTipoFactura == 2 || IdTipoFactura == 4 || IdTipoFactura == 36)
				{
					#region Orden de Compra
					if (!ModSerial) 
					{
						this.ultraGrid1.Rows.Band.AddNew();
						this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = 0;
						this.ultraGrid1.ActiveRow.Cells["idDetalleSerial"].Value = 0;
						this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = 0;
						this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
						this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();
						this.ultraGrid1.ActiveRow.Cells["Estado"].Value = 9;	
					}
					else this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();						
					#endregion Orden de Compra
				}
				else if (IdTipoFactura == 5)
				{
					IdSerial = 0;
					IdDetalleSerial = 0;

					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Consignacion From Compra Where idTipoFactura = 5 And idCompra = {0}", IdCompra), true))
					{
						#region Notas de Credito Consignaciones
						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaSerialDevolucionConsignacion {0}, {1}, '{2}'", IdCompra, IdArticulo, this.txtSerie.Text.ToString().Trim())) == 0)
						{
							MessageBox.Show(string.Format("La Serial '{0}' no esta registrada en las consignaciones de este cliente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop));
							this.txtSerie.Text = "";
							this.txtSerie.Focus();
							return;
						}

						string sSQLIdSerial = string.Format("Exec BuscaIdSerialDevolucionConsignacion '{0}', {1}, {2}", this.txtSerie.Text.ToString(), IdArticulo, IdCompra);
						SqlDataReader drIdSerial = Funcion.rEscalarSQL(cdsSeteoF, sSQLIdSerial, true);
						drIdSerial.Read();
						if (drIdSerial.HasRows)
						{
							IdSerial = drIdSerial.GetInt32(0);
							IdDetalleSerial = drIdSerial.GetInt32(1);
						}
						drIdSerial.Close();
						
						if (!ModSerial) 
						{															
							this.ultraGrid1.Rows.Band.AddNew();
							this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = IdSerial;
							this.ultraGrid1.ActiveRow.Cells["idDetalleSerial"].Value = 0;
							this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;
							this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
							this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();
							this.ultraGrid1.ActiveRow.Cells["Estado"].Value = 18;
						}
						else 
						{
							this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = IdSerial;	
							this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;													
							this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();							
						}
						#endregion Notas de Credito Consignaciones
					}
					else
					{ 
						#region Notas de Credito
						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaSerialNotaDeCredito {0}, {1}, '{2}'", IdCompra, IdArticulo, this.txtSerie.Text.ToString().Trim())) == 0)
						{
							MessageBox.Show(string.Format("La Serial '{0}' no esta registrada en la Factura Relacionada a esta Nota de Credito", this.txtSerie.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtSerie.Text = "";
							this.txtSerie.Focus();
							return;
						}
						
						string sSQLIdSerial = string.Format("Exec BuscaIdSerialNotaDeCredito '{0}', {1}, {2}", this.txtSerie.Text.ToString(), IdArticulo, IdCompra);
						SqlDataReader drIdSerial = Funcion.rEscalarSQL(cdsSeteoF, sSQLIdSerial, true);
						drIdSerial.Read();
						if (drIdSerial.HasRows)
						{
							IdSerial = drIdSerial.GetInt32(0);
							IdDetalleSerial = drIdSerial.GetInt32(1);
						}
						drIdSerial.Close();

						if (!ModSerial) 
						{															
							this.ultraGrid1.Rows.Band.AddNew();
							this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = IdSerial;
							this.ultraGrid1.ActiveRow.Cells["idDetalleSerial"].Value = 0;
							this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;
							this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
							this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();
							this.ultraGrid1.ActiveRow.Cells["Estado"].Value = 17;
						}
						else 
						{
							this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = IdSerial;	
							this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;													
							this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();							
						}
						#endregion Notas de Credito 
					}
				}
				else if (IdTipoFactura == 6)
				{
					#region Notas de Credito Proveedor
					IdSerial = 0;
					IdDetalleSerial = 0;

					string sSQL = string.Format("Exec ValidaSerialDevCompra {0}, '{1}', {2}", IdCompra, this.txtSerie.Text.ToString(), IdArticulo);
					int iValida = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
					if (iValida == 0)
					{
						MessageBox.Show("La serial no pertenece a la Factura que hace referencia la Nota de crédito seleccionada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtSerie.Text = "";
						this.txtSerie.Focus();
						return;
					}

					string sSQLIdSerial = string.Format("Exec BuscaIdSerialDevCmp '{0}', {1}, {2}", (string) this.txtSerie.Text, IdArticulo, IdCompra);
					SqlDataReader drIdSerial = Funcion.rEscalarSQL(cdsSeteoF, sSQLIdSerial, true);
					drIdSerial.Read();
					if (drIdSerial.HasRows)
					{
						IdSerial = drIdSerial.GetInt32(0);
						IdDetalleSerial = drIdSerial.GetInt32(1);
					}
					drIdSerial.Close();
					
					if (!ModSerial) 
					{													
						this.ultraGrid1.Rows.Band.AddNew();
						this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = IdSerial;
						this.ultraGrid1.ActiveRow.Cells["idDetalleSerial"].Value = 0;
						this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;
						this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
						this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();
						this.ultraGrid1.ActiveRow.Cells["Estado"].Value = 11;
					}
					else 
					{
						this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = IdSerial;
						this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;
						this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();							
					}
					#endregion Notas de Credito Proveedor
				}
				else  if (IdTipoFactura == 9)
				{
					#region Recepcion de mercaderia
					int IEstRec = 0;												
					string sMensaje = "";						

					string sSQLRecTransf = string.Format("Exec ValidaSerialesEnRecepcionMercaderia '{0}', {1}, {2}, {3}", this.txtSerie.Text.ToString(), IdArticulo, iBod, IdCompra);
					SqlDataReader drIdSerial = Funcion.rEscalarSQL(cdsSeteoF, sSQLRecTransf, true);
					drIdSerial.Read();
					if (drIdSerial.HasRows)
					{
						IdSerial = drIdSerial.GetInt32(0);
						IdDetalleSerial = drIdSerial.GetInt32(1);
						IEstRec = drIdSerial.GetInt32(2);
						sMensaje = drIdSerial.GetString(3);
					}
					drIdSerial.Close();

					if (IEstRec != 1)
					{
						string sSQLSS = string.Format("Exec RegistroSeguimientoSeriales {0}, {1}, '{2}', {3}, {4}, '{5}'",
							IdCompra, 9, this.txtSerie.Text.ToString(), IdArticulo, iBod, sMensaje);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLSS, true);

						MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtSerie.Text = "";
						this.txtSerie.Focus();
						return;
					}

					if (!ModSerial)
					{
						this.ultraGrid1.Rows.Band.AddNew();					
						this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = IdSerial;
						this.ultraGrid1.ActiveRow.Cells["idDetalleSerial"].Value = 0;
						this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;
						this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
						this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();
						this.ultraGrid1.ActiveRow.Cells["NotasDelSistema"].Value = sMensaje;
						this.ultraGrid1.ActiveRow.Cells["Estado"].Value = IEstRec;
					}
					else
					{
						this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = IdSerial;						
						this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;						
						this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();
						this.ultraGrid1.ActiveRow.Cells["NotasDelSistema"].Value = sMensaje;
						this.ultraGrid1.ActiveRow.Cells["Estado"].Value = IEstRec;
					}
					#endregion Recepcion de mercaderia
				}
				else if (IdTipoFactura == 51 && IEstado == 15)
				{
					#region Notas de Entrega 
					IdSerial = 0;
					IdDetalleSerial = 0;

					string sSQL = string.Format("Exec ValidaSerialNotaDeEntrega {0}, {1}, {2}, '{3}'", IdCompra, IdArticulo, iBod, this.txtSerie.Text.ToString());
					int iValida = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
					if (iValida == 0)
					{
						MessageBox.Show("La serial no pertenece a la Nota de entrega seleccionada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtSerie.Text = "";
						this.txtSerie.Focus();
						return;
					}

					string sSQLIdSerial = string.Format("Exec BuscaIdSerialNotaDeEntrega {0}, {1}, {2}, '{3}'", IdCompra, IdArticulo, iBod, this.txtSerie.Text.ToString());
					SqlDataReader drIdSerial = Funcion.rEscalarSQL(cdsSeteoF, sSQLIdSerial, true);
					drIdSerial.Read();
					if (drIdSerial.HasRows)
					{
						IdSerial = drIdSerial.GetInt32(0);
						IdDetalleSerial = drIdSerial.GetInt32(1);
					}
					drIdSerial.Close();

					if (!ModSerial)
					{
						this.ultraGrid1.Rows.Band.AddNew();
						this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = IdSerial;
						this.ultraGrid1.ActiveRow.Cells["idDetalleSerial"].Value = 0;
						this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;
						this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
						this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();	
						this.ultraGrid1.ActiveRow.Cells["Estado"].Value = 1;
					}
					else
					{
						this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = IdSerial;						
						this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;
						this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();							
					}
					#endregion Notas de Entrega 
				}
				else 
				{
					#region Valida si la Serial Existe o si Esta relacionada al Articulo Seleccionado
					string sSQLVSA = string.Format("Exec ValidaSerialArticulo '{0}', {1}", this.txtSerie.Text.ToString(), IdArticulo);
					string sMensajeVSA = Funcion.sEscalarSQL(cdsSeteoF, sSQLVSA);
						
					if (sMensajeVSA.Length > 0)
					{
						MessageBox.Show(sMensajeVSA, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtSerie.Text = "";
						this.txtSerie.Focus();
						return;
					}
					#endregion Valida si la Serial Existe o si Esta relacionada al Articulo Seleccionado	

					#region Valida Proceso
					int iCont = 0;
					string sMens = "";
					string sSQLValSerEsp = string.Format("Exec ValidaSerialOrdCmpEnTransito '{0}', {1}, {2}, {3}", this.txtSerie.Text.ToString(), IdArticulo, iBod, IdTipoFactura);
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLValSerEsp, true);
					dr.Read();
					if (dr.HasRows)
					{
						iCont = dr.GetInt32(0);
						sMens = dr.GetString(1);
					}
					dr.Close();
														
					if (iCont > 0)
					{
						MessageBox.Show(sMens, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtSerie.Text = "";
						this.txtSerie.Focus();
						return;
					}
					#endregion Valida Proceso 

					#region Carga Grid de Seriales
					string sSQL = string.Format("Exec CuentaSerialArticulo '{0}', {1}, {2}", this.txtSerie.Text.ToString(), IdArticulo, iBod);
					int iCtaSerial = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
        
					if (iCtaSerial == 1)
					{	
						IdSerial = 0;
						IdDetalleSerial = 0;
						string sSQLIdSerial = string.Format("Exec BuscaIdSerial '{0}', {1}, {2}", this.txtSerie.Text.ToString(), IdArticulo, iBod);
						SqlDataReader drIdSerial = Funcion.rEscalarSQL(cdsSeteoF, sSQLIdSerial, true);
						drIdSerial.Read();
						if (drIdSerial.HasRows)
						{
							IdSerial = drIdSerial.GetInt32(0);
							IdDetalleSerial = drIdSerial.GetInt32(1);
						}
						drIdSerial.Close();
							
						if (!ModSerial)
						{													
							this.ultraGrid1.Rows.Band.AddNew();
							this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = IdSerial;
							this.ultraGrid1.ActiveRow.Cells["idDetalleSerial"].Value = 0;
							this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;
							this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
							this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();
							this.ultraGrid1.ActiveRow.Cells["Estado"].Value = 0;
						}
						else
						{
							this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = IdSerial;							
							this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;							
							this.ultraGrid1.ActiveRow.Cells["Serial"].Value = this.txtSerie.Text.ToString();	
							this.ultraGrid1.ActiveRow.Cells["Estado"].Value = 0;
						}
					}
					else
					{
							MessageBox.Show("La serial no existe o no pertenece al artículo seleccionado, Comuniquese con Auditoria", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtSerie.Text = "";
							this.txtSerie.Focus();
							return;
					}					
					#endregion Carga Grid de Seriales
				}
							
				this.txtSerie.Text = "";
				ModSerial = false;
				
				int iFilas = this.ultraGrid1.Rows.Count;
				this.spnLlegaron.Value = DCantidad - iFilas;
				this.lblContador.Text = "SERIALES: " + iFilas + " DE " + DCantidad;
				
				if (iFilas == DCantidad) this.btnAceptar.Focus(); else this.txtSerie.Focus();
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!miAcceso.BEliminaSeriales)
			{
				MessageBox.Show("No tiene acceso a Eliminar Seriales", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}

			if (IdTipoFactura == 2 || IdTipoFactura == 8 || IdTipoFactura == 9)
			{
				int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From Compra Where idCompra = {0}", IdCompra));

				if (iEstado == 3)
				{
					if (!miAcceso.BEliminaModificaSeriales)
					{
						MessageBox.Show("No tiene acceso a Eliminar Seriales el documento esta procesado\n\nComuniquese con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
						return;
					}
				}

				if (iEstado == 9)
				{
					MessageBox.Show("Para Orden de Compra con registro de factura no puede eliminar seriales\n\nComuniquese con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			FuncionesInventario.EliminaFilasGridsSeriales(sender, e, IdTipoFactura, IdCompra, 1, cdsSeteoF);
		}

		private void ultraGrid1_Click(object sender, System.EventArgs e)
		{
			
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);			
		}

		private void txtSerie_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void spnLlegaron_Click(object sender, System.EventArgs e)
		{
			this.spnLlegaron.SelectAll();
		}

		private void spnLlegaron_Enter(object sender, System.EventArgs e)
		{
			this.spnLlegaron.SelectAll();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmIngresoSeriales_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();
		}

		private void ultraGrid1_DoubleClick(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.ActiveRow != null) 
			{
				string Serial = ultraGrid1.ActiveRow.Cells["Serial"].Value.ToString();

				string sSQL = string.Format("Exec ValSerialBorrar {0}, {1}, {2}", 
					(int)this.ultraGrid1.ActiveRow.Cells["idSerial"].Value, IdTipoFactura, (int)this.ultraGrid1.ActiveRow.Cells["idDetCompra"].Value);
				if (!Funcion.bEscalarSQL(cdsSeteoF, sSQL, true)) return;

				if (!miAcceso.BModificaSeriales)
				{
					MessageBox.Show("No tiene acceso a modificar seriales", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit = 1 If (Select Estado From Compra Where idCompra = {0}) In (3, 9) Set @Res = 0 Select @Res ", IdCompra), true))
				{
					MessageBox.Show("El documento ya esta procesado no puede modificar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if ((int)this.ultraGrid1.ActiveRow.Cells["idDetalleSerial"].Value > 0)
				{
					if (miAcceso.BModificaSeriales)
					{			
						if (DialogResult.Yes ==	MessageBox.Show(string.Format("¿Desea Modificar la Serial '{0}'?", Serial), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							this.txtSerie.Text = Serial; 
							ModSerial = true;
						}
						else ModSerial = false;
					}
				}
				else if (DialogResult.Yes ==	MessageBox.Show(string.Format("¿Desea Modificar la Serial '{0}'?", Serial), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					this.txtSerie.Text = Serial; 
					ModSerial = true;
				}
				else ModSerial = false;
			}
		}

		private void txtHasta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtHasta.Focus();
		}

		private void txtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGeneraSeriales_Click(sender, e);
		}

		private void txtHasta_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if ((int)this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese las seriales del Artículo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				this.txtSerie.Focus();
				return;
			}

			if (DialogResult.Yes == MessageBox.Show("Antes de Guardar Revise si las Seriales Ingresadas son las Correctas \n\n Una Vez Guardado No Puede Modificar \n\n\n ¿Esta Seguro de Guardar?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{
				#region Estados
				int iEstado = 0;
				
				if (IdTipoFactura == 5) iEstado = 18;
				if (IdTipoFactura == 6) iEstado = 11;
				if (IdTipoFactura == 8) iEstado = 3;
				if (IdTipoFactura == 11 || IdTipoFactura == 13) iEstado = 2;
				if (IdTipoFactura == 44) iEstado = 7;
				if (IdTipoFactura == 51 && IEstado != 15) iEstado = 14;
				if (IdTipoFactura == 51 && IEstado == 15) iEstado = 1;				
				#endregion Estados

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

						#region Guarda Seriales y Movimientos
						if (IdTipoFactura == 2 || IdTipoFactura == 4 || IdTipoFactura == 36)
						{
							int iEstadoIngreso = 9;
							if (IdTipoFactura == 36) iEstadoIngreso = 1;

							#region Orden de compra - Compra
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drSeriales in this.ultraGrid1.Rows.All)
							{
								int idSerial = 0;
								if (drSeriales.Cells["idSerial"].Value != DBNull.Value) idSerial = (int)drSeriales.Cells["idSerial"].Value;
 
								string sSQLGSeriales = string.Format("Exec GuardaSeriales {0}, {1}, '{2}', {3}, {4}", 
									idSerial, (int)drSeriales.Cells["idArticulo"].Value, drSeriales.Cells["Serial"].Value.ToString(), iEstadoIngreso, iBod);
								oCmdActualiza.CommandText = sSQLGSeriales;
								int NIdSerial = (int)oCmdActualiza.ExecuteScalar();

								string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, 0, {2}, '{3}', 1, '', 0, 0, {4}, 0", 
									NIdSerial, IdDetCompra, iBod, drSeriales.Cells["Observaciones"].Value.ToString(), (int)drSeriales.Cells["idDetalleSerial"].Value);
								oCmdActualiza.CommandText = sSQLGDetalleSeriales;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Orden de compra - Compra
						}
						else
						{
							#region Pedidos - Notas de Credito - Ordenes de Produccion - Transferencias
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid1.Rows.All)
							{
								#region Valida Salida de Mercaderia en otros movimientos
								if (IdTipoFactura == 8 || IdTipoFactura == 11 || IdTipoFactura == 13 || IdTipoFactura == 44)
								{
									if ((int)dr1.Cells["idDetalleSerial"].Value < 1)
									{
										string sSQL = string.Format("Exec CuentaSerialArticulo '{0}', {1}, {2}", dr1.Cells["Serial"].Value.ToString(), IdArticulo, iBod);
										oCmdActualiza.CommandText = sSQL;
										int iCtaSerial = (int)oCmdActualiza.ExecuteScalar();
        
										if (iCtaSerial < 1)
										{	
											MessageBox.Show(string.Format("La serial '{0}' ya fue registrada en otro movimiento", dr1.Cells["Serial"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
											this.ultraGrid1.ActiveRow = dr1;
											this.ultraGrid1.ActiveRow.Selected = true;

											oTransaction.Rollback();

											return;
										}
									}
								}
								#endregion Valida Salida de Mercaderia en otros movimientos

								int idSerial = 0;
								int idOrigen = 0;
								int Origen = 0;
								if (IdTipoFactura == 7) Origen = 1;					
									
								if (dr1.Cells["idSerial"].Value != DBNull.Value) idSerial = (int)dr1.Cells["idSerial"].Value;						
								if (dr1.Cells["idOrigen"].Value != DBNull.Value) idOrigen = (int)dr1.Cells["idOrigen"].Value;

								string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, {2}, {3}, '{4}', 1, '{5}', {6}, {7}, {8}", 
									idSerial, IdDetCompra, idOrigen, iBod, dr1.Cells["Observaciones"].Value.ToString(), dr1.Cells["NotasDelSistema"].Value.ToString(), (int)dr1.Cells["Estado"].Value, Origen, (int)dr1.Cells["idDetalleSerial"].Value);
								oCmdActualiza.CommandText = sSQLGDetalleSeriales;
								oCmdActualiza.ExecuteNonQuery();
							
								if (IdTipoFactura == 9) if ((int)dr1.Cells["Estado"].Value == 3) iEstado = 1; else iEstado = (int)dr1.Cells["Estado"].Value;
										
								string sSQLEstado = string.Format("Update Seriales Set Estado = {0}, Bodega = {2} Where idSerial = {1}", iEstado, idSerial, iBod);
								oCmdActualiza.CommandText = sSQLEstado;
								oCmdActualiza.ExecuteNonQuery();									
							}
							#endregion Pedidos - Notas de Credito - Ordenes de Produccion - Transferencias
						}
						#endregion Guarda Seriales y Movimientos

						#region Actualiza el estado del articulo (Estado Serial)
						string sSQLEstadoDC = string.Format("Exec ERMActualizaEstadoSerialEnDetCompra {0}, {1}", IdDetCompra, this.ultraGrid1.Rows.Count);
						oCmdActualiza.CommandText = sSQLEstadoDC;
						oCmdActualiza.ExecuteNonQuery();							
						#endregion Actualiza el estado del articulo (Estado Serial)

						#region Procesar el articulo
						if (IdTipoFactura != 8 && IdTipoFactura != 9)
						{								
							string sSQLProcesa = string.Format("Exec EstadoOrdenDeCompra {0}, {1}", IdCompra, IdTipoFactura);
							oCmdActualiza.CommandText = sSQLProcesa;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Procesar el articulo

						oTransaction.Commit();

						this.DialogResult = DialogResult.OK;
					}
					catch (Exception ex)
					{
						try
						{
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						}
						catch (Exception ex2)
						{
							MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);							
						}
					}
					finally
					{
						oConexion.Close();
					}
				}
			}			
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;			
		}

		private void btnGenerarSeriales_Click(object sender, System.EventArgs e)
		{
			int iFilas = this.ultraGrid1.Rows.Count;
			int CantGenerar = (int)this.spnLlegaron.Value;
			int c = (int)DCantidad - (int)iFilas;
			
			if (CantGenerar > (DCantidad - iFilas))
			{
				MessageBox.Show("La cantidad de seriales a generar no puede ser mayor a " + c + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if (iFilas < DCantidad)
			{				
				if (this.ultraGrid1.Rows.Count > 0) 
					this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaSeriales 0, 0, 0");	

				string sSQL = string.Format("Exec GeneraSeriales {0}, {1}, {2}, {3}, {4}", 
					IdArticulo, IdCompra, IdDetCompra, (int)this.spnLlegaron.Value, IdTipoFactura);
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
				
				this.spnLlegaron.Value = DCantidad - iFilas;
				this.lblContador.Text = "SERIALES: " + this.ultraGrid1.Rows.Count + " DE " + DCantidad;				
				this.btnAceptar.Focus();
			}
			else if (iFilas == DCantidad) MessageBox.Show("Las seriales de este artículo ya fueron registradas", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btnGeneraSeriales_Click(object sender, System.EventArgs e)
		{
			int IFilas = this.ultraGrid1.Rows.Count;

			if (this.txtDesde.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba la Serial Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDesde.Focus();
				return;
			}
			if (this.txtHasta.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba la Cantidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtHasta.Focus();
				return;
			}
			if (IFilas == DCantidad)
			{
				MessageBox.Show("Ya ha registrado las seriales de este Artículo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDesde.Focus();
				return;
			}
			if ((int.Parse(this.txtHasta.Text.ToString()) + this.ultraGrid1.Rows.Count) > DCantidad)
			{
				MessageBox.Show(string.Format("No puede ingresar mas de {0}", (DCantidad - this.ultraGrid1.Rows.Count)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtHasta.Focus();
				return;
			}
							
				int iCont = 0;
				int iContS = 0;
				string sMensajeGrid = "La(s) Serial(es) siguientes ya estan registradas : \n\n";
				string sMensajeBDD = "\n\n\n La(s) Serial(es) siguientes ya estan registradas en la Base de Datos : \n\n";

				string sSQL = string.Format("Exec GeneraSerialesOrdenDeCompra '{0}', '{1}', {2}", this.txtDesde.Text.ToString(), this.txtHasta.Text, IdArticulo);
				SqlDataReader drGenera = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
				while (drGenera.Read())
				{	
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					{						
						if (drGenera.GetString(3) == dr.Cells["Serial"].Value.ToString())
						{
							iCont++;
							sMensajeGrid = sMensajeGrid + drGenera.GetString(3) + "\n";
						}
					}
							
					if (iCont == 0)
					{
						if (drGenera.GetInt32(5) == - 1)
						{
							iCont++;
							sMensajeBDD = sMensajeBDD + drGenera.GetString(3) + "\n";
						}
						else
						{
							this.ultraGrid1.Rows.Band.AddNew();
							this.ultraGrid1.ActiveRow.Cells["idSerial"].Value = drGenera.GetInt32(0);
							this.ultraGrid1.ActiveRow.Cells["idDetalleSerial"].Value = drGenera.GetInt32(1);
							this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
							this.ultraGrid1.ActiveRow.Cells["Serial"].Value = drGenera.GetString(3);
							this.ultraGrid1.ActiveRow.Cells["Estado"].Value = 9;

							iContS ++;

							int iFilas = this.ultraGrid1.Rows.Count;
							this.spnLlegaron.Value = DCantidad - iFilas;
							this.lblContador.Text = "SERIALES: " + iFilas + " DE " + DCantidad;
						}
					}
				}
				drGenera.Close();

				if (iContS != int.Parse(this.txtHasta.Text.ToString())) MessageBox.Show(sMensajeGrid + sMensajeBDD,  "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				
				IFilas = this.ultraGrid1.Rows.Count;
				if (IFilas == DCantidad) this.btnAceptar.Focus();
				else this.txtSerie.Focus();
		}
	}
	
}

