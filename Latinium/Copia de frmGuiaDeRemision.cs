using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions.Shared;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmGuiaDeRemision.
	/// </summary>
	public class frmGuiaDeRemision : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblNumero;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label lblComprobante;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.GroupBox gbBotones;
		public System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtIniTraslado;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFinTraslado;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactura;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTraslado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDestino;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtRazonSocialDestinatario;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtRucDestinatario;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtRucTransportista;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtPlaca;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGuiaRemision;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblEstado;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumComprobante;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutorizacion;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEmision;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public bool bNuevo = false;
		private System.Windows.Forms.Label label18;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtOrigen;
		bool bEdicion = false;

		int IdCompra = 0;
		int IdTipoFactura = 0;
		int IBodega = 0;
		bool BAutomatico = false;

		public frmGuiaDeRemision(int idCompra, int idTipoFactura, int iBodega, bool bAutomatico)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			IdCompra = idCompra;
			IdTipoFactura = idTipoFactura;
			IBodega = iBodega;
			BAutomatico = bAutomatico;

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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmGuiaDeRemision));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMotivoTraslado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMotivoTraslado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.dtIniTraslado = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFinTraslado = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtAutorizacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.dtFechaEmision = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtDestino = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtRucDestinatario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRazonSocialDestinatario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label15 = new System.Windows.Forms.Label();
			this.txtRucTransportista = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.txtPlaca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTraslado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtIdGuiaRemision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.txtOrigen = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIniTraslado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFinTraslado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEmision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRucDestinatario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRazonSocialDestinatario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRucTransportista)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTraslado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGuiaRemision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).BeginInit();
			this.SuspendLayout();
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
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(792, 11);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 192;
			this.lblNumero.Text = "Número";
			// 
			// txtNumero
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(856, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 22);
			this.txtNumero.TabIndex = 3;
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumero_Validating);
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(77, 16);
			this.label6.TabIndex = 194;
			this.label6.Text = "Inicio Traslado";
			// 
			// dtIniTraslado
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtIniTraslado.Appearance = appearance2;
			this.dtIniTraslado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtIniTraslado.DateButtons.Add(dateButton1);
			this.dtIniTraslado.Enabled = false;
			this.dtIniTraslado.Format = "dd/MM/yyyy";
			this.dtIniTraslado.Location = new System.Drawing.Point(104, 8);
			this.dtIniTraslado.Name = "dtIniTraslado";
			this.dtIniTraslado.NonAutoSizeHeight = 23;
			this.dtIniTraslado.Size = new System.Drawing.Size(112, 21);
			this.dtIniTraslado.SpinButtonsVisible = true;
			this.dtIniTraslado.TabIndex = 1;
			this.dtIniTraslado.Value = ((object)(resources.GetObject("dtIniTraslado.Value")));
			this.dtIniTraslado.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtIniTraslado_KeyPress);
			this.dtIniTraslado.Validating += new System.ComponentModel.CancelEventHandler(this.dtIniTraslado_Validating);
			this.dtIniTraslado.Validated += new System.EventHandler(this.dtIniTraslado_Validated);
			this.dtIniTraslado.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtIniTraslado_BeforeDropDown);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(256, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(91, 16);
			this.label1.TabIndex = 196;
			this.label1.Text = "Termina Traslado";
			// 
			// dtFinTraslado
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFinTraslado.Appearance = appearance3;
			this.dtFinTraslado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFinTraslado.DateButtons.Add(dateButton2);
			this.dtFinTraslado.Enabled = false;
			this.dtFinTraslado.Format = "dd/MM/yyyy";
			this.dtFinTraslado.Location = new System.Drawing.Point(368, 8);
			this.dtFinTraslado.Name = "dtFinTraslado";
			this.dtFinTraslado.NonAutoSizeHeight = 23;
			this.dtFinTraslado.Size = new System.Drawing.Size(112, 21);
			this.dtFinTraslado.SpinButtonsVisible = true;
			this.dtFinTraslado.TabIndex = 2;
			this.dtFinTraslado.Value = ((object)(resources.GetObject("dtFinTraslado.Value")));
			this.dtFinTraslado.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFinTraslado_KeyPress);
			this.dtFinTraslado.Validating += new System.ComponentModel.CancelEventHandler(this.dtFinTraslado_Validating);
			this.dtFinTraslado.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFinTraslado_BeforeDropDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.label2.Location = new System.Drawing.Point(8, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(196, 17);
			this.label2.TabIndex = 197;
			this.label2.Text = "DATOS COMPROBANTE DE VENTA";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 67);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(26, 16);
			this.label8.TabIndex = 227;
			this.label8.Text = "Tipo";
			// 
			// cmbTipoFactura
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoFactura.Appearance = appearance4;
			this.cmbTipoFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFactura.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipoFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoFactura.DisplayMember = "Nombre";
			this.cmbTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFactura.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoFactura.Enabled = false;
			this.cmbTipoFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFactura.Location = new System.Drawing.Point(56, 64);
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(144, 21);
			this.cmbTipoFactura.TabIndex = 4;
			this.cmbTipoFactura.ValueMember = "idTipoFactura";
			this.cmbTipoFactura.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTipoFactura_KeyPress);
			this.cmbTipoFactura.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoFactura_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(504, 67);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(69, 16);
			this.lblComprobante.TabIndex = 231;
			this.lblComprobante.Text = "Autorización ";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(224, 67);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(133, 16);
			this.label3.TabIndex = 230;
			this.label3.Text = "Número del Comprobante";
			// 
			// txtAutorizacion
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutorizacion.Appearance = appearance5;
			this.txtAutorizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutorizacion.Enabled = false;
			this.txtAutorizacion.Location = new System.Drawing.Point(592, 64);
			this.txtAutorizacion.Name = "txtAutorizacion";
			this.txtAutorizacion.Size = new System.Drawing.Size(104, 22);
			this.txtAutorizacion.TabIndex = 6;
			// 
			// txtNumComprobante
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumComprobante.Appearance = appearance6;
			this.txtNumComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumComprobante.Enabled = false;
			this.txtNumComprobante.Location = new System.Drawing.Point(376, 64);
			this.txtNumComprobante.Name = "txtNumComprobante";
			this.txtNumComprobante.Size = new System.Drawing.Size(104, 22);
			this.txtNumComprobante.TabIndex = 5;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(736, 67);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(91, 16);
			this.label4.TabIndex = 233;
			this.label4.Text = "Fecha de Emisión";
			// 
			// dtFechaEmision
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEmision.Appearance = appearance7;
			this.dtFechaEmision.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaEmision.DateButtons.Add(dateButton3);
			this.dtFechaEmision.Enabled = false;
			this.dtFechaEmision.Format = "dd/MM/yyyy";
			this.dtFechaEmision.Location = new System.Drawing.Point(848, 64);
			this.dtFechaEmision.Name = "dtFechaEmision";
			this.dtFechaEmision.NonAutoSizeHeight = 23;
			this.dtFechaEmision.Size = new System.Drawing.Size(112, 21);
			this.dtFechaEmision.SpinButtonsVisible = true;
			this.dtFechaEmision.TabIndex = 7;
			this.dtFechaEmision.Value = ((object)(resources.GetObject("dtFechaEmision.Value")));
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 99);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(102, 16);
			this.label5.TabIndex = 235;
			this.label5.Text = "Motivo del Traslado";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(504, 128);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(42, 16);
			this.label9.TabIndex = 239;
			this.label9.Text = "Destino";
			// 
			// txtDestino
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDestino.Appearance = appearance8;
			this.txtDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDestino.Enabled = false;
			this.txtDestino.Location = new System.Drawing.Point(568, 128);
			this.txtDestino.Multiline = true;
			this.txtDestino.Name = "txtDestino";
			this.txtDestino.Size = new System.Drawing.Size(392, 56);
			this.txtDestino.TabIndex = 10;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.label10.Location = new System.Drawing.Point(8, 200);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(222, 17);
			this.label10.TabIndex = 240;
			this.label10.Text = "IDENTIFICACIÓN DEL DESTINATARIO";
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(504, 227);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(58, 16);
			this.label11.TabIndex = 244;
			this.label11.Text = "R.U.C./C.I.";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(8, 227);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(67, 16);
			this.label12.TabIndex = 243;
			this.label12.Text = "Razón Social";
			// 
			// txtRucDestinatario
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRucDestinatario.Appearance = appearance9;
			this.txtRucDestinatario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRucDestinatario.Enabled = false;
			this.txtRucDestinatario.Location = new System.Drawing.Point(592, 224);
			this.txtRucDestinatario.Name = "txtRucDestinatario";
			this.txtRucDestinatario.Size = new System.Drawing.Size(104, 22);
			this.txtRucDestinatario.TabIndex = 12;
			// 
			// txtRazonSocialDestinatario
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRazonSocialDestinatario.Appearance = appearance10;
			this.txtRazonSocialDestinatario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRazonSocialDestinatario.Enabled = false;
			this.txtRazonSocialDestinatario.Location = new System.Drawing.Point(96, 224);
			this.txtRazonSocialDestinatario.Name = "txtRazonSocialDestinatario";
			this.txtRazonSocialDestinatario.Size = new System.Drawing.Size(376, 22);
			this.txtRazonSocialDestinatario.TabIndex = 11;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Font = new System.Drawing.Font("Tahoma", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.label13.Location = new System.Drawing.Point(8, 256);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(230, 17);
			this.label13.TabIndex = 245;
			this.label13.Text = "IDENTIFICACIÓN DEL TRANSPORTISTA";
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(8, 283);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(44, 16);
			this.label14.TabIndex = 247;
			this.label14.Text = "Nombre";
			// 
			// cmbResponsable
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance11;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResponsable.DataSource = this.ultraDataSource4;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 376;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.Enabled = false;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(96, 280);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.Size = new System.Drawing.Size(376, 21);
			this.cmbResponsable.TabIndex = 13;
			this.cmbResponsable.ValueMember = "idPersonal";
			this.cmbResponsable.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbResponsable_KeyPress);
			this.cmbResponsable.Validating += new System.ComponentModel.CancelEventHandler(this.cmbResponsable_Validating);
			this.cmbResponsable.ValueChanged += new System.EventHandler(this.cmbResponsable_ValueChanged);
			this.cmbResponsable.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbResponsable_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(504, 283);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(58, 16);
			this.label15.TabIndex = 249;
			this.label15.Text = "R.U.C./C.I.";
			// 
			// txtRucTransportista
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRucTransportista.Appearance = appearance12;
			this.txtRucTransportista.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRucTransportista.Enabled = false;
			this.txtRucTransportista.Location = new System.Drawing.Point(592, 280);
			this.txtRucTransportista.Name = "txtRucTransportista";
			this.txtRucTransportista.Size = new System.Drawing.Size(104, 22);
			this.txtRucTransportista.TabIndex = 14;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(728, 283);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(30, 16);
			this.label16.TabIndex = 251;
			this.label16.Text = "Placa";
			// 
			// txtPlaca
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPlaca.Appearance = appearance13;
			this.txtPlaca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPlaca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPlaca.Enabled = false;
			this.txtPlaca.Location = new System.Drawing.Point(784, 280);
			this.txtPlaca.MaxLength = 7;
			this.txtPlaca.Name = "txtPlaca";
			this.txtPlaca.Size = new System.Drawing.Size(104, 22);
			this.txtPlaca.TabIndex = 15;
			this.txtPlaca.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPlaca_KeyPress);
			this.txtPlaca.Validating += new System.ComponentModel.CancelEventHandler(this.txtPlaca_Validating);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Font = new System.Drawing.Font("Tahoma", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.label17.Location = new System.Drawing.Point(8, 312);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(103, 17);
			this.label17.TabIndex = 253;
			this.label17.Text = "OBSERVACIONES";
			// 
			// txtNotas
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance14;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 336);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(952, 64);
			this.txtNotas.TabIndex = 16;
			this.txtNotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotas_KeyPress);
			// 
			// gbBotones
			// 
			this.gbBotones.Location = new System.Drawing.Point(0, 408);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(968, 8);
			this.gbBotones.TabIndex = 275;
			this.gbBotones.TabStop = false;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(88, 418);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 19;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 418);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 18;
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
			this.btnAnular.Location = new System.Drawing.Point(169, 418);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 22;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(249, 418);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 23;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// cmbTraslado
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTraslado.Appearance = appearance15;
			this.cmbTraslado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTraslado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTraslado.DataSource = this.ultraDataSource2;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 352;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbTraslado.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTraslado.DisplayMember = "Motivo";
			this.cmbTraslado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTraslado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTraslado.Enabled = false;
			this.cmbTraslado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTraslado.Location = new System.Drawing.Point(128, 96);
			this.cmbTraslado.Name = "cmbTraslado";
			this.cmbTraslado.Size = new System.Drawing.Size(352, 21);
			this.cmbTraslado.TabIndex = 276;
			this.cmbTraslado.ValueMember = "idMotivoTraslado";
			// 
			// ultraDataSource3
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// txtIdGuiaRemision
			// 
			this.txtIdGuiaRemision.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdGuiaRemision.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGuiaRemision.Enabled = false;
			this.txtIdGuiaRemision.Location = new System.Drawing.Point(936, 418);
			this.txtIdGuiaRemision.Name = "txtIdGuiaRemision";
			this.txtIdGuiaRemision.PromptChar = ' ';
			this.txtIdGuiaRemision.Size = new System.Drawing.Size(16, 22);
			this.txtIdGuiaRemision.TabIndex = 277;
			this.txtIdGuiaRemision.Visible = false;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(912, 418);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtIdCompra.TabIndex = 278;
			this.txtIdCompra.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(888, 418);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 279;
			this.txtEstado.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(704, 224);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 280;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(504, 11);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(30, 16);
			this.label18.TabIndex = 283;
			this.label18.Text = "Local";
			// 
			// cmbBodega
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance16;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource3;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 352;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(592, 8);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(176, 21);
			this.cmbBodega.TabIndex = 282;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBodega_KeyPress);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(8, 128);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(37, 16);
			this.label7.TabIndex = 285;
			this.label7.Text = "Origen";
			// 
			// txtOrigen
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOrigen.Appearance = appearance17;
			this.txtOrigen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOrigen.Enabled = false;
			this.txtOrigen.Location = new System.Drawing.Point(88, 128);
			this.txtOrigen.Multiline = true;
			this.txtOrigen.Name = "txtOrigen";
			this.txtOrigen.Size = new System.Drawing.Size(392, 56);
			this.txtOrigen.TabIndex = 284;
			// 
			// frmGuiaDeRemision
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(968, 448);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtOrigen);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtIdGuiaRemision);
			this.Controls.Add(this.cmbTraslado);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.txtPlaca);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.txtRucTransportista);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.cmbResponsable);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtRucDestinatario);
			this.Controls.Add(this.txtRazonSocialDestinatario);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.txtDestino);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtFechaEmision);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtAutorizacion);
			this.Controls.Add(this.txtNumComprobante);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.cmbTipoFactura);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFinTraslado);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtIniTraslado);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.txtNumero);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmGuiaDeRemision";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Guía De Remisión";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGuiaDeRemision_KeyDown);
			this.Load += new System.EventHandler(this.frmGuiaDeRemision_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIniTraslado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFinTraslado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEmision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRucDestinatario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRazonSocialDestinatario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRucTransportista)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTraslado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGuiaRemision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		public int iBodegaDestino = 0;

		private void frmGuiaDeRemision_Load(object sender, System.EventArgs e)
		{
			#region Listas
			this.cmbTipoFactura.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoFactura, Nombre From CompraNumero Where GuiaRemision = 1 Order By idTipoFactura Asc");			
			this.cmbTraslado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idMotivoTraslado, Motivo From MotivosDeTraslado Where Activo = 1");		
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));			
			#endregion Listas
			
			bool bImpreso = false;

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From MaestroGuiaRemision Where idCompra = {0} And idTipoFactura = {1} And Estado = 0", IdCompra, IdTipoFactura)) > 0)
			{
				int idPersonal = 0;				
				string sSQLGRemision = string.Format("Exec ConsultaGuiaDeRemision {0}, {1}", IdCompra, IdTipoFactura);
				SqlDataReader drConsulta = Funcion.rEscalarSQL(cdsSeteoF, sSQLGRemision, true);
				drConsulta.Read();

				if (drConsulta.HasRows)
				{
					this.txtIdGuiaRemision.Value = drConsulta.GetInt32(0);
					this.txtIdCompra.Value = drConsulta.GetInt32(1);
					this.dtIniTraslado.Value = drConsulta.GetDateTime(2);
					this.dtFinTraslado.Value = drConsulta.GetDateTime(3);
					this.cmbBodega.Value = drConsulta.GetInt32(4);
					this.txtNumero.Text = drConsulta.GetString(5);
					
					if (drConsulta.GetInt32(6) == 8) this.cmbTipoFactura.Value = 42;
					else this.cmbTipoFactura.Value = drConsulta.GetInt32(6);

					this.cmbTraslado.Value = drConsulta.GetInt32(7); 
					idPersonal = drConsulta.GetInt32(8);
					this.txtPlaca.Text = drConsulta.GetString(9);
					this.txtNotas.Text = drConsulta.GetString(10);
					this.txtNumComprobante.Text = drConsulta.GetString(11);
					this.dtFechaEmision.Value = drConsulta.GetDateTime(12);
					this.txtOrigen.Text = drConsulta.GetString(13);
					this.txtDestino.Text = drConsulta.GetString(14);

					if (IdTipoFactura != 8) this.txtAutorizacion.Text = drConsulta.GetString(15);

					this.txtRazonSocialDestinatario.Text = drConsulta.GetString(16);
					this.txtRucDestinatario.Text = drConsulta.GetString(17);
					iBodegaDestino = drConsulta.GetInt32(18);
					bImpreso = drConsulta.GetBoolean(19);
				}
				drConsulta.Close();
				
				this.cmbResponsable.DataSource = FuncionesProcedimientos.Lista_Choferes(Convert.ToDateTime(this.dtIniTraslado.Value), (int)this.cmbBodega.Value, cdsSeteoF);
				this.cmbResponsable.Value = idPersonal;

				bNuevo = false;
				bEdicion = true;

				//if (!bImpreso) 
				this.btnImprimir.Enabled = true;
				this.btnAnular.Enabled = true;
			}
			else
			{
				bNuevo = true;
				bEdicion = true;

				this.txtIdCompra.Value = IdCompra;
				FuncionesProcedimientos.FechasDocumentos(this.dtIniTraslado);
				FuncionesProcedimientos.FechasDocumentos(this.dtFinTraslado);				
				
				FuncionesProcedimientos.RetornaFechaServidor(this.dtIniTraslado, cdsSeteoF, false);
				FuncionesProcedimientos.RetornaFechaServidor(this.dtFinTraslado, cdsSeteoF, false);

				this.dtFinTraslado.CalendarInfo.MinDate = (DateTime)this.dtIniTraslado.Value;
				this.dtFinTraslado.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtIniTraslado.Value).AddDays(7);

				this.cmbBodega.Value = IBodega;
				if (IdTipoFactura == 8) this.cmbTipoFactura.Value = 42;
				else this.cmbTipoFactura.Value = IdTipoFactura;

				if (IdTipoFactura == 1) this.cmbTraslado.Value = 1;
				if (IdTipoFactura == 8) this.cmbTraslado.Value = 5;
				if (IdTipoFactura == 44) this.cmbTraslado.Value = 4;

				string sSQLDatos = string.Format("Exec DatosParaGuiaRemision {0}, {1}", IdCompra, IdTipoFactura);
				SqlDataReader drDatos = Funcion.rEscalarSQL(cdsSeteoF, sSQLDatos, true);
				drDatos.Read();
				if (drDatos.HasRows)
				{
					this.txtNumComprobante.Text = drDatos.GetString(0);
					this.dtFechaEmision.Value = drDatos.GetDateTime(1);
					this.txtOrigen.Text = drDatos.GetString(2); 
					this.txtDestino.Text = drDatos.GetString(3); 
					
					if (IdTipoFactura != 8) this.txtAutorizacion.Text = drDatos.GetString(4); 

					this.txtRazonSocialDestinatario.Text = drDatos.GetString(5); 
					this.txtRucDestinatario.Text = drDatos.GetString(6); 
					iBodegaDestino = drDatos.GetInt32(7); 
				}
				drDatos.Close();

				this.cmbResponsable.DataSource = FuncionesProcedimientos.Lista_Choferes(Convert.ToDateTime(this.dtIniTraslado.Value), (int)this.cmbBodega.Value, cdsSeteoF);
			}
			
			if (!bImpreso)
			{
				if (!BAutomatico) this.txtNumero.Enabled = true;

				this.dtIniTraslado.Enabled = true;
				this.dtFinTraslado.Enabled = true;
				this.cmbResponsable.Enabled = true;
				this.txtPlaca.Enabled = true;
				this.txtNotas.Enabled = true;
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		public void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ParameterFields paramFields = new ParameterFields ();

			ParameterField Compra = new ParameterField ();
			ParameterDiscreteValue IdCompra = new ParameterDiscreteValue ();
			//Compra.ParameterFieldName = "@idCompra";			
			Compra.ParameterFieldName = "@idCPA";
			//IdCompra.Value = int.Parse(this.txtIdCompra.Value.ToString());
			IdCompra.Value = int.Parse(this.txtIdGuiaRemision.Value.ToString());
			Compra.CurrentValues.Add (IdCompra);
			paramFields.Add (Compra);

//			ParameterField Factura = new ParameterField ();
//			ParameterDiscreteValue TipoFactura = new ParameterDiscreteValue ();
//			Factura.ParameterFieldName = "@idTipofactura";
//			TipoFactura.Value = IdTipoFactura;// int.Parse(IdTipoFactura.ToString());
//			Factura.CurrentValues.Add (TipoFactura);
//			paramFields.Add (Factura);

			Reporte miRep = new Reporte("Doc_GR.rpt", "");
			//Reporte miRep = new Reporte("GuiaDeRemision.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Guia de Remisión");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
			
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update MaestroGuiaRemision Set Impreso = 1 Where idMaestroGuiaRemision = {0}", (int)this.txtIdGuiaRemision.Value));

			Cursor = Cursors.Default;
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Esta seguro de eliminar este documento.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) return;

			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update MaestroGuiaRemision Set Estado = 1 Where idMaestroGuiaRemision = {0}", (int)this.txtIdGuiaRemision.Value));

			MessageBox.Show("Documento anulado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void dtIniTraslado_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFinTraslado.Focus();
		}

		private void dtFinTraslado_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
			{
				if (BAutomatico) this.cmbResponsable.Focus();
				else this.txtNumero.Focus();
			}
		}

		private void dtIniTraslado_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.dtIniTraslado.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la fecha de inicio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtIniTraslado.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void dtIniTraslado_Validated(object sender, System.EventArgs e)
		{
			
		}

		private void dtFinTraslado_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.dtFinTraslado.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la fecha final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFinTraslado.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void cmbTipoFactura_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtFinTraslado_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbBodega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
		}

		private void cmbTipoFactura_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
		}

		private void cmbResponsable_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtPlaca.Focus();
		}

		private void txtPlaca_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericosPasaporte(sender, e);
			if (e.KeyChar == 13) this.txtNotas.Focus();
		}

		private void txtNotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnGuardar.Focus();
		}

		private void cmbResponsable_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.cmbResponsable.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Transportista", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbResponsable.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void txtPlaca_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
//			if (this.txtPlaca.Text.ToString().Length == 0)
//			{
//				MessageBox.Show("Escriba el número de placa", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.txtPlaca.Focus();
//				e.Cancel = true;
//				return;
//			}
		}

		private void cmbResponsable_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbResponsable.ActiveRow == null) return;
			this.txtRucTransportista.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Cedula From Personal Where idPersonal = {0}", (int)this.cmbResponsable.Value));
		}

		private void frmGuiaDeRemision_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void dtIniTraslado_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!BAutomatico)
			{
				if (this.txtNumero.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Numero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}

				char Pad = '0';
				this.txtNumero.Text = this.txtNumero.Value.ToString().PadLeft(9, Pad);

				string sSQLVal = string.Format("Exec ValidaNumeracionGuiaDeRemision '{0}', {1}", this.txtNumero.Text.ToString(), (int)this.cmbBodega.Value);
				bool bRes = Funcion.bEscalarSQL(cdsSeteoF, sSQLVal, true);
				if (bRes)
				{
					MessageBox.Show("El Numero " + this.txtNumero.Text.ToString() + " ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}
			}

			if (this.dtIniTraslado.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la fecha de inicio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtIniTraslado.Focus();
				return;
			}

			if (this.dtFinTraslado.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la fecha final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFinTraslado.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtFinTraslado.Value) < Convert.ToDateTime(this.dtIniTraslado.Value))
			{
				MessageBox.Show("La fecha de inicio no puede ser posterior a la fecha final de traslado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtIniTraslado.Focus();
				return;
			}
			
			if (this.cmbResponsable.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un transportista", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbResponsable.Focus();
				return;
			}

			if (!Validacion.vbValidaPlaca(this.txtPlaca)) return;

			this.DialogResult = DialogResult.OK;
		}

		private void cmbResponsable_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13) this.cmbResponsable.Focus();
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!BAutomatico)
			{
				if (this.txtNumero.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Numero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}

				char Pad = '0';
				this.txtNumero.Text = this.txtNumero.Value.ToString().PadLeft(9, Pad);

				string sSQLVal = string.Format("Exec ValidaNumeracionGuiaDeRemision '{0}', {1}", this.txtNumero.Text.ToString(), (int)this.cmbBodega.Value);
				bool bRes = Funcion.bEscalarSQL(cdsSeteoF, sSQLVal, true);
				if (bRes)
				{
					MessageBox.Show("El Numero " + this.txtNumero.Text.ToString() + " ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}
			}
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}

