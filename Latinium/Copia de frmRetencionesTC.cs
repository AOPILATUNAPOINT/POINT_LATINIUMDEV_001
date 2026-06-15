using System;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.IO;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmRetenciones.
	/// </summary>
	public class frmRetencionesTC : DevExpress.XtraEditors.XtraForm
	{
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutFactura;
		private System.Windows.Forms.Label label31;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblNumero;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerieFact;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdRetencion;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label11;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.ComponentModel.IContainer components;

		Infragistics.Win.UltraWinGrid.UltraGrid UGridCompra;
		private Acceso miAcceso;
		bool bEliminaAlValidar = false;
		int IdCompra = 0;
		int IdCliente = 0;
		string Factura = "";
		decimal DIVA = 0;
		decimal DTotal = 0;
		DateTime DTFecha = DateTime.Today;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetenido;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetenidoIVA;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetenidoRenta;
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtCaducidad;
		bool BNuevo = false;
		private System.Windows.Forms.Button btnFe;
		bool BCxP = false;
	  bool Proveedor = false;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtFactura;
		CultureInfo us = new CultureInfo("en-US");
		
		public frmRetencionesTC()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRetencionesTC));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVARenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencionIVARenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVARenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencionIVARenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			this.txtAutFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSerieFact = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtIdRetencion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtRetenido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtRetenidoIVA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtRetenidoRenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.label7 = new System.Windows.Forms.Label();
			this.dtCaducidad = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnFe = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenidoIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenidoRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtCaducidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).BeginInit();
			this.SuspendLayout();
			// 
			// txtAutFactura
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutFactura.Appearance = appearance1;
			this.txtAutFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutFactura.Enabled = false;
			this.txtAutFactura.Location = new System.Drawing.Point(568, 40);
			this.txtAutFactura.MaxLength = 37;
			this.txtAutFactura.Name = "txtAutFactura";
			this.txtAutFactura.Size = new System.Drawing.Size(248, 22);
			this.txtAutFactura.TabIndex = 2;
			this.txtAutFactura.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAutFactura_KeyDown);
			this.txtAutFactura.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAutFactura_KeyPress);
			this.txtAutFactura.Validating += new System.ComponentModel.CancelEventHandler(this.txtAutFactura_Validating);
			this.txtAutFactura.ValueChanged += new System.EventHandler(this.txtAutFactura_ValueChanged);
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(488, 43);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(65, 16);
			this.label31.TabIndex = 298;
			this.label31.Text = "Autorización";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(29, 16);
			this.label3.TabIndex = 297;
			this.label3.Text = "Serie";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(248, 43);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 296;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance2;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(344, 40);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(128, 22);
			this.txtNumero.TabIndex = 1;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumero_Validating);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// txtSerieFact
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerieFact.Appearance = appearance3;
			this.txtSerieFact.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerieFact.Enabled = false;
			this.txtSerieFact.Location = new System.Drawing.Point(80, 40);
			this.txtSerieFact.MaxLength = 6;
			this.txtSerieFact.Name = "txtSerieFact";
			this.txtSerieFact.Size = new System.Drawing.Size(64, 22);
			this.txtSerieFact.TabIndex = 0;
			this.txtSerieFact.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSerieFact_KeyDown);
			this.txtSerieFact.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerieFact_KeyPress);
			this.txtSerieFact.Validating += new System.ComponentModel.CancelEventHandler(this.txtSerieFact_Validating);
			this.txtSerieFact.ValueChanged += new System.EventHandler(this.txtSerieFact_ValueChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 72);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 293;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance4;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(80, 72);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 3;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn5.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
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
			// txtIdRetencion
			// 
			this.txtIdRetencion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdRetencion.Enabled = false;
			this.txtIdRetencion.Location = new System.Drawing.Point(496, 336);
			this.txtIdRetencion.Name = "txtIdRetencion";
			this.txtIdRetencion.PromptChar = ' ';
			this.txtIdRetencion.Size = new System.Drawing.Size(16, 22);
			this.txtIdRetencion.TabIndex = 326;
			this.txtIdRetencion.Visible = false;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(512, 336);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtIdCompra.TabIndex = 325;
			this.txtIdCompra.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(544, 336);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 327;
			this.txtEstado.Visible = false;
			// 
			// txtRetenido
			// 
			this.txtRetenido.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetenido.Appearance = appearance5;
			this.txtRetenido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetenido.Enabled = false;
			this.txtRetenido.FormatString = "#,##0.00";
			this.txtRetenido.Location = new System.Drawing.Point(720, 376);
			this.txtRetenido.Name = "txtRetenido";
			this.txtRetenido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetenido.PromptChar = ' ';
			this.txtRetenido.Size = new System.Drawing.Size(96, 22);
			this.txtRetenido.TabIndex = 330;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(656, 376);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(29, 16);
			this.label2.TabIndex = 331;
			this.label2.Text = "Total";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(656, 312);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(32, 16);
			this.label4.TabIndex = 332;
			this.label4.Text = "I.V.A.";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRetenidoIVA
			// 
			this.txtRetenidoIVA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetenidoIVA.Appearance = appearance6;
			this.txtRetenidoIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetenidoIVA.Enabled = false;
			this.txtRetenidoIVA.FormatString = "#,##0.00";
			this.txtRetenidoIVA.Location = new System.Drawing.Point(720, 312);
			this.txtRetenidoIVA.Name = "txtRetenidoIVA";
			this.txtRetenidoIVA.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetenidoIVA.PromptChar = ' ';
			this.txtRetenidoIVA.Size = new System.Drawing.Size(96, 22);
			this.txtRetenidoIVA.TabIndex = 333;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(656, 344);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(33, 16);
			this.label5.TabIndex = 334;
			this.label5.Text = "Renta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRetenidoRenta
			// 
			this.txtRetenidoRenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetenidoRenta.Appearance = appearance7;
			this.txtRetenidoRenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetenidoRenta.Enabled = false;
			this.txtRetenidoRenta.FormatString = "#,##0.00";
			this.txtRetenidoRenta.Location = new System.Drawing.Point(720, 344);
			this.txtRetenidoRenta.Name = "txtRetenidoRenta";
			this.txtRetenidoRenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetenidoRenta.PromptChar = ' ';
			this.txtRetenidoRenta.Size = new System.Drawing.Size(96, 22);
			this.txtRetenidoRenta.TabIndex = 335;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(328, 64);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 336;
			// 
			// label11
			// 
			this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(8, 312);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(37, 16);
			this.label11.TabIndex = 338;
			this.label11.Text = "Notas:";
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(528, 336);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsiento.TabIndex = 342;
			this.txtIdAsiento.Visible = false;
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
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance8;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 336);
			this.txtNotas.MaxLength = 600;
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(440, 64);
			this.txtNotas.TabIndex = 4;
			this.txtNotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotas_KeyPress);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 408);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(826, 8);
			this.groupBox1.TabIndex = 344;
			this.groupBox1.TabStop = false;
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(248, 72);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(56, 16);
			this.label7.TabIndex = 348;
			this.label7.Text = "Caducidad";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtCaducidad
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtCaducidad.Appearance = appearance9;
			this.dtCaducidad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtCaducidad.DateButtons.Add(dateButton2);
			this.dtCaducidad.Enabled = false;
			this.dtCaducidad.Format = "dd/MM/yyyy";
			this.dtCaducidad.Location = new System.Drawing.Point(344, 72);
			this.dtCaducidad.Name = "dtCaducidad";
			this.dtCaducidad.NonAutoSizeHeight = 23;
			this.dtCaducidad.Size = new System.Drawing.Size(112, 21);
			this.dtCaducidad.SpinButtonsVisible = true;
			this.dtCaducidad.TabIndex = 347;
			this.dtCaducidad.Value = ((object)(resources.GetObject("dtCaducidad.Value")));
			this.dtCaducidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtCaducidad_KeyDown);
			this.dtCaducidad.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtCaducidad_KeyPress);
			this.dtCaducidad.Validating += new System.ComponentModel.CancelEventHandler(this.dtCaducidad_Validating);
			this.dtCaducidad.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtCaducidad_BeforeDropDown);
			// 
			// btnFe
			// 
			this.btnFe.CausesValidation = false;
			this.btnFe.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFe.Location = new System.Drawing.Point(656, 424);
			this.btnFe.Name = "btnFe";
			this.btnFe.TabIndex = 628;
			this.btnFe.Text = "&FE";
			this.btnFe.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnFe.Visible = false;
			this.btnFe.Click += new System.EventHandler(this.btnFe_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 104);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(832, 8);
			this.groupBox2.TabIndex = 629;
			this.groupBox2.TabStop = false;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(246, 424);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 757;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(488, 424);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 756;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(166, 424);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 755;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(84, 424);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 754;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click_1);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 424);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 753;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(744, 424);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 749;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Visible = false;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click_1);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(326, 424);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 750;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click_1);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(406, 424);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(78, 23);
			this.btnCancelar.TabIndex = 748;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.ForeColor = System.Drawing.Color.Firebrick;
			this.label8.Location = new System.Drawing.Point(567, 424);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(0, 26);
			this.label8.TabIndex = 751;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(488, 11);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(38, 16);
			this.label9.TabIndex = 859;
			this.label9.Text = "Cliente";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 11);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(71, 16);
			this.lblIdentificacion.TabIndex = 857;
			this.lblIdentificacion.Text = "Identificación";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(248, 11);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(86, 16);
			this.lblCliente.TabIndex = 856;
			this.lblCliente.Text = "N. Identificación";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance10;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(568, 8);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(248, 22);
			this.txtNombre.TabIndex = 858;
			// 
			// cmbTipoRuc
			// 
			appearance11.FontData.Name = "Tahoma";
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance11;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 233;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(80, 8);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 22);
			this.cmbTipoRuc.TabIndex = 854;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoRuc_KeyDown);
			this.cmbTipoRuc.ValueChanged += new System.EventHandler(this.cmbTipoRuc_ValueChanged);
			// 
			// txtRuc
			// 
			this.txtRuc.AcceptsTab = true;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance12;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(344, 8);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(128, 22);
			this.txtRuc.TabIndex = 855;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// txtidCliente
			// 
			this.txtidCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtidCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.Location = new System.Drawing.Point(560, 336);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(16, 22);
			this.txtidCliente.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidCliente.TabIndex = 860;
			this.txtidCliente.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance13.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance13;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 119;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 51;
			ultraGridColumn5.Header.Caption = "Renta";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 40;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 168;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance14;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Width = 90;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Width = 60;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 399;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance15;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Width = 100;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance16;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Width = 100;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance17;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance18;
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 9F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ForeColorDisabled = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance20;
			appearance21.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance21.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance23;
			this.ultraGrid1.Enabled = false;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 120);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(810, 184);
			this.ultraGrid1.TabIndex = 305;
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.ultraGrid1_BeforeRowInsert);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// txtFactura
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFactura.Appearance = appearance24;
			this.txtFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFactura.Enabled = false;
			this.txtFactura.Location = new System.Drawing.Point(808, 72);
			this.txtFactura.MaxLength = 9;
			this.txtFactura.Name = "txtFactura";
			this.txtFactura.Size = new System.Drawing.Size(8, 22);
			this.txtFactura.TabIndex = 861;
			this.txtFactura.Visible = false;
			// 
			// frmRetencionesTC
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(826, 456);
			this.Controls.Add(this.txtFactura);
			this.Controls.Add(this.txtidCliente);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnFe);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label31);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtCaducidad);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtRetenidoRenta);
			this.Controls.Add(this.txtRetenidoIVA);
			this.Controls.Add(this.txtRetenido);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdRetencion);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtAutFactura);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtSerieFact);
			this.Controls.Add(this.dtFecha);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmRetencionesTC";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Retención";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRetenciones_KeyDown);
			this.Load += new System.EventHandler(this.frmRetenciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenidoIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenidoRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtCaducidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		bool bNuevo = false;
		bool bEdicion = false;
			private bool bActualiza = true;
	
		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
		
		private void Suma()
		{
			this.txtRetenidoIVA.Value = 0.00m;
			this.txtRetenidoRenta.Value = 0.00m;
			this.txtRetenido.Value = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				dr.Cells["Valor"].Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select CONVERT(DECIMAL(18, 2), {0})", Convert.ToDecimal(dr.Cells["Valor"].Value)));					
			}
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (!(bool) dr.Cells["IVARenta"].Value) this.txtRetenidoIVA.Value = Convert.ToDecimal(this.txtRetenidoIVA.Value) + (decimal)dr.Cells["Valor"].Value;
				if ((bool) dr.Cells["IVARenta"].Value) this.txtRetenidoRenta.Value = Convert.ToDecimal(this.txtRetenidoRenta.Value) + (decimal)dr.Cells["Valor"].Value;				
			}

			this.txtRetenido.Value = decimal.Round(Convert.ToDecimal(this.txtRetenidoIVA.Value), 2) + decimal.Round(Convert.ToDecimal(this.txtRetenidoRenta.Value), 2);
		}


		private void frmRetenciones_Load(object sender, System.EventArgs e)
		{
			string stTipo = "040601";					
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga 1");

			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;						
			this.dtFecha.CalendarInfo.MinDate = DateTime.Parse("01/11/2017");
		}

		private void txtSerieFact_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtNumero.Focus();
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtAutFactura.Focus();
		}

		private void txtAutFactura_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.dtCaducidad.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNotas.Focus();
		}

		private void txtNotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private void txtSerieFact_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtSerieFact.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número de serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerieFact.Focus();
				e.Cancel = true;				
			}
			else if (this.txtSerieFact.Text.Length < 6)
			{
				MessageBox.Show("El número de serie es de 6 digitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerieFact.Focus();
				e.Cancel = true;				
			}
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtNumero.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				e.Cancel = true;				
			}
			else
			{
				char Pad = '0';
				this.txtNumero.Text = this.txtNumero.Text.ToString().PadLeft(9, Pad);

				string sSQL = string.Format("Select COUNT(*) From Retenciones Where Autorizacion = '{0}' And Serie = '{1}' And Numero = '{2}' and idRetencion <> {3}", this.txtAutFactura.Text.ToString(), this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString(), (int)this.txtIdRetencion.Value);
				if (Funcion.iEscalarSQL(cdsSeteoF, sSQL) > 0)
				{
					MessageBox.Show("Ya existe este numero de documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					e.Cancel = true;			
				}
			}
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			char Pad = '0';
			this.txtNumero.Text = this.txtNumero.Text.ToString().PadLeft(9, Pad);
		}

		private void txtAutFactura_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtAutFactura.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el número de autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAutFactura.Focus();
				e.Cancel = true;				
			}
			else if (this.txtAutFactura.Text.Length < 3 && this.txtAutFactura.Text.Length > 37)
			{
				MessageBox.Show("Longitud de autorización debe estar entre 3 y 37 caracteres.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAutFactura.Focus();
				e.Cancel = true;				
			}
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				e.Cancel = true;				
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Retenciones en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				e.Cancel = true;					
			}
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (this.txtSerieFact.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número de serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerieFact.Focus();			
			}
			else if (this.txtSerieFact.Text.Length < 6)
			{
				MessageBox.Show("El número de serie es de 6 digitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerieFact.Focus();				
			}
			else if (this.txtNumero.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();				
			}
			else if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();				
			}			
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Retenciones en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();			
			}
			else if (this.txtAutFactura.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el número de autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAutFactura.Focus();				
			}
			else if (this.txtAutFactura.Text.Length < 10 && this.txtAutFactura.Text.Length > 37)
			{
				MessageBox.Show("Longitud de autorización debe estar entre 10 y 37 caracteres.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAutFactura.Focus();				
			}
			else if (this.dtCaducidad.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la fecha de caducidad.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtCaducidad.Focus();				
			}
			else
			{
				#region Filas Vacias
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (dr.Cells["idRetencionIVARenta"].Value == System.DBNull.Value)
					{
						bEliminaAlValidar = true;
						dr.Delete();
					}
				}
				#endregion Filas Vacias

				this.DialogResult = DialogResult.OK;
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;

			#region Impresion
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
			pfRet.ParameterFieldName = "@idCompra";
			idCompraRet.Value = (int)this.txtIdCompra.Value;
			pfRet.CurrentValues.Add (idCompraRet);
			paramFields.Add (pfRet);
				
			Reporte miRepor = new Reporte("Retenciones.rpt", "");
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("Retenciones");
			miRepor.crVista.ParameterFieldInfo = paramFields;
			#endregion Impresion
				
			miRepor.Show();			
						
			Cursor = Cursors.Default;
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value != 2)
			{
				if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No puede Anular esta Retención, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();				
				}
				else if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Anular esta retención?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				{
					string sSQLAnul = string.Format("Exec AnulaRetencion {0}, {1}, 0, 4", (int)this.txtIdRetencion.Value, (int)this.txtIdAsiento.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAnul, true);

					MessageBox.Show("Retención Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
					this.txtEstado.Value = 2;
					this.lblEstado.Text = "ANULADO";
					this.btnAnular.Text = "Eliminar";
				}
			}
			else
			{
				if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No puede Eliminar esta Retención, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();				
				}
				else if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Eliminar esta retención?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				{			
					string sSQLAnul = string.Format("Exec AnulaRetencion {0}, {1}, 1, 4", (int)this.txtIdRetencion.Value, (int)this.txtIdAsiento.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAnul, true);

					MessageBox.Show("Retención Eliminada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
					this.txtEstado.Value = 2;
					this.lblEstado.Text = "";

					this.btnConsultar_Click(sender, e);
				}
			}
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{	
			DateTime dtFechaAsi = (DateTime)this.dtFecha.Value;

			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Crear Asientos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if ((int)this.txtEstado.Value == 2) 
			{
				MessageBox.Show("No puede crear un Asiento de una Retención Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				MessageBox.Show(string.Format("No Puede Crear Asiento de esta Retención, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFechaAsi.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}						
			else
			{
				Cursor = Cursors.WaitCursor;
				int idFormaPago = 0;
				string sNombre = "";
				int idProyecto = 0;
				string sRUC = "";
				DateTime dtFechaRet = (DateTime)this.dtFecha.Value;

				string sSQL = string.Format("Select c.idFormaPago, cl.Nombre, c.idProyecto, cl.Ruc From Compra c Inner Join Cliente cl On cl.idCliente = c.idCliente Where c.idCompra = {0}", IdCompra);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					idFormaPago = dr.GetInt32(0);
					sNombre = dr.GetString(1);
					idProyecto = dr.GetInt32(2);
					sRUC = dr.GetString(3);
				}
				dr.Close();

				string sSQLAsRet = string.Format("Exec CreaAsientoRetencion {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', {8}, {9}, 0, 0, '{10}'",
					(int)this.txtIdAsiento.Value, (int)this.txtIdRetencion.Value, IdCompra, idFormaPago, sNombre, this.txtNumero.Text.ToString(),
					this.txtFactura.Text.ToString(), dtFechaRet.ToString("yyyyMMdd"), idProyecto, Convert.ToDecimal(this.txtRetenido.Value), sRUC);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAsRet, true);
				Cursor = Cursors.Default;		
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Eliminar las filas seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				{	
					for (int i = 0; i < e.Rows.Length; i++)
					{			
						if ((int)e.Rows[i].Cells["idDetalleRetencion"].Value > 0)
						{	
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete from DetalleRetenciones where idDetalleRetencion = {0}", (int)e.Rows[i].Cells["idDetalleRetencion"].Value));

							e.DisplayPromptMsg = false;
						}
						else e.DisplayPromptMsg = false;
					
						this.btnCancelar.Enabled = false;
					}
				}
				else e.Cancel = true;
			}
			else e.Cancel = true;
		}

		private void dtCaducidad_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFecha.Focus();
		}

		private void dtCaducidad_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.dtCaducidad.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la fecha de caducidad.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtCaducidad.Focus();
				e.Cancel = true;				
			}
		}

		private void frmRetenciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdRetencion.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) 
						if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
																		
					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoF, "040601");

						this.Consultar((int)this.txtIdRetencion.Value );					
					}
				}
				

				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape)
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void btnFe_Click(object sender, System.EventArgs e)
		{
			#region Variables
			int iAmbiente = 0;
			int iTipoEmision = 0;
			string sRazonSocial = "";
			string sNombreComercial = "";
			string sRUC = "";
			string sClaveAcceso = "";
			string sCodDoc = "";
			string sEstab = "";
			string sPtoEmi = "";
			string sDirMatriz = "";
			string sDirEstablecimiento = "";			
			int iResolucion = 0;
			string sCodDocClte = "";
			string sNombreProv = "";
			string sRUCProv = "";
			DateTime dtFechaDoc = DateTime.Today;
			string sNumeroDoc = "";			
			#endregion Variables
						
			DateTime dtFecha = (DateTime)this.dtFecha.Value;
			
			#region Consulta Informacion Tributaria
			string sSQL = string.Format("Exec FEInformacionTributaria {0}, {1}, {2}, '{3}', '{4}', {5}", 1, 1, 1, dtFecha.ToString("ddMMyyyy"), this.txtNumero.Text.ToString(), IdCliente);
			SqlDataReader drInfoTributaria = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drInfoTributaria.Read();
			if (drInfoTributaria.HasRows)
			{				
				iAmbiente = (int)drInfoTributaria.GetValue(0);
				iTipoEmision = (int)drInfoTributaria.GetValue(1);
				sRazonSocial = drInfoTributaria.GetValue(2).ToString();
				sNombreComercial = drInfoTributaria.GetValue(3).ToString();
				sRUC = drInfoTributaria.GetValue(4).ToString();
				sClaveAcceso = drInfoTributaria.GetValue(5).ToString();
				sCodDoc = drInfoTributaria.GetValue(6).ToString();
				sEstab = drInfoTributaria.GetValue(7).ToString();
				sPtoEmi = drInfoTributaria.GetValue(8).ToString();
				sDirMatriz = drInfoTributaria.GetValue(9).ToString();
				sDirEstablecimiento = drInfoTributaria.GetValue(10).ToString();
				iResolucion = (int)drInfoTributaria.GetValue(11);
				sCodDocClte = drInfoTributaria.GetValue(12).ToString();
				sNombreProv = drInfoTributaria.GetValue(13).ToString();
				sRUCProv = drInfoTributaria.GetValue(14).ToString();
			}
			drInfoTributaria.Close();
			#endregion Consulta Informacion Tributaria

			#region XML
			#region Encabezado
			string stArchivo = "RETPRUEBA.XML";
			XmlTextWriter XMLFactura = new XmlTextWriter(stArchivo, Encoding.GetEncoding("UTF-8"));
			XMLFactura.Formatting = Formatting.Indented;
			XMLFactura.WriteStartDocument(true);
			XMLFactura.WriteStartElement("comprobanteRetencion");
			#endregion Encabezado
			
			#region Informacion Tributaria
			XMLFactura.WriteStartElement("infoTributaria");

			XMLFactura.WriteElementString("ambiente", iAmbiente.ToString());
			XMLFactura.WriteElementString("tipoEmision", iTipoEmision.ToString());
			XMLFactura.WriteElementString("razonSocial", sRazonSocial);
			XMLFactura.WriteElementString("nombreComercial", sNombreComercial);
			XMLFactura.WriteElementString("ruc", sRUC);
			XMLFactura.WriteElementString("claveAcceso", sClaveAcceso);
			XMLFactura.WriteElementString("codDoc", sCodDoc);
			XMLFactura.WriteElementString("estab", sEstab);
			XMLFactura.WriteElementString("ptoEmi", sPtoEmi);
			XMLFactura.WriteElementString("secuencial", this.txtNumero.Text.ToString());
			XMLFactura.WriteElementString("dirMatriz", sDirMatriz);

			XMLFactura.WriteEndElement();
			#endregion Informacion Tributaria

			#region Informacion Retencion
			XMLFactura.WriteStartElement("infoCompRetencion");
			
			XMLFactura.WriteElementString("fechaEmision", dtFecha.ToString("dd/MM/yyyy"));
			XMLFactura.WriteElementString("dirEstablecimiento", sDirEstablecimiento);
			XMLFactura.WriteElementString("contribuyenteEspecial", iResolucion.ToString());
			XMLFactura.WriteElementString("obligadoContabilidad", "SI");
			XMLFactura.WriteElementString("tipoIdentificacionSujetoRetenido", sCodDocClte);
			XMLFactura.WriteElementString("razonSocialSujetoRetenido", sNombreProv);
			XMLFactura.WriteElementString("identificacionSujetoRetenido", sRUCProv);			
			XMLFactura.WriteElementString("periodoFiscal", dtFecha.ToString("MM/yyyy"));		
			
			XMLFactura.WriteEndElement();
			#endregion Informacion Retencion

			#region Impuestos
			string sSQLInfoDoc = string.Format("Select SerieFactura + Numero, Fecha From Compra Where idCompra = {0} And idTipoFactura = 4", (int)this.txtIdCompra.Value);
			SqlDataReader drInfoDoc = Funcion.rEscalarSQL(cdsSeteoF, sSQLInfoDoc, true);
			drInfoDoc.Read();
			if (drInfoDoc.HasRows)
			{					 
				sNumeroDoc = drInfoDoc.GetValue(0).ToString();
				dtFechaDoc = (DateTime)drInfoDoc.GetValue(1);
			}
			drInfoDoc.Close();

			XMLFactura.WriteStartElement("Impuestos");
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				int iCodigo = 2;
				string sCodigoRetencion = "";
        if ((bool)dr.Cells["IVARenta"].Value) iCodigo = 1;
				if (iCodigo == 2)
				{
					if (dr.Cells["Codigo"].Value.ToString() == "721") sCodigoRetencion = "1";
					if (dr.Cells["Codigo"].Value.ToString() == "723") sCodigoRetencion = "2";
					if (dr.Cells["Codigo"].Value.ToString() == "725") sCodigoRetencion = "3";
				}
				else sCodigoRetencion = dr.Cells["Codigo"].Value.ToString();

				decimal dBaseImponible = (decimal)dr.Cells["Base"].Value;
				decimal dPorcentaje = (decimal)dr.Cells["Porcentaje"].Value; 
				decimal dValorRetenido = (decimal)dr.Cells["Valor"].Value; 

				XMLFactura.WriteStartElement("Impuesto");
	
				XMLFactura.WriteElementString("código", iCodigo.ToString());
				XMLFactura.WriteElementString("codigoRetencion", sCodigoRetencion);
				XMLFactura.WriteElementString("baseImponible", dBaseImponible.ToString("0.00", us));
				XMLFactura.WriteElementString("porcentajeRetener", dPorcentaje.ToString("0.00", us));
				XMLFactura.WriteElementString("valorRetenido", dValorRetenido.ToString("0.00", us));
				XMLFactura.WriteElementString("codDocSustento", "01");
				XMLFactura.WriteElementString("numDocSustento", sNumeroDoc);
				XMLFactura.WriteElementString("fechaEmisionDocSustento", dtFechaDoc.ToString("dd/MM/yyyy"));

				XMLFactura.WriteEndElement();
			}

			XMLFactura.WriteEndElement();
			#endregion Impuestos

			XMLFactura.WriteEndElement(); 
			XMLFactura.WriteEndDocument();

			XMLFactura.Flush();
			XMLFactura.Close();
			#endregion XML
		}

		private void txtSerieFact_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtAutFactura_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void dtCaducidad_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}
		
		

		private void btnAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				if (miAcceso.CrearAsiento)
				{
					if ((int) this.txtIdAsiento.Value == 0) MessageBox.Show("No Existe un Asiento Creado de esta Transacción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					else
					{					
						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0} And Borrar = 0) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtIdAsiento.Value), true))
						{
							Cursor = Cursors.WaitCursor;
							frmAsientos Asiento = new frmAsientos((int) this.txtIdAsiento.Value);
							Asiento.MdiParent = MdiParent;
							Asiento.Show();
							Cursor = Cursors.Default;
						}
						else if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0} And Borrar = 1) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtIdAsiento.Value), true))
						{
							MessageBox.Show("El Asiento de esta Retención esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						}
						else if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0}) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtIdAsiento.Value), true))
						{
							MessageBox.Show("El Asiento de esta Retención fue Eliminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						}
					}
				}
			}						
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
				this.Close();		
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region Desabilita campos
			this.cmbTipoRuc.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtSerieFact.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtAutFactura.Enabled  = false;
			this.dtCaducidad.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtFactura.Enabled = false;
			this.txtNotas.Enabled = false;
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			#endregion Desabilita campos

			#region vacia campos
			this.cmbTipoRuc.Value = System.DBNull.Value;
			this.txtRuc.Text = "";
			this.txtSerieFact.Text = "";
			this.txtNumero.Text= "";
			this.txtAutFactura.Text = "";
			this.dtCaducidad.Value = System.DBNull.Value;
			this.dtFecha.Value = System.DBNull.Value;
			this.txtFactura.Text = "";
			this.txtNotas.Text= "";
			this.txtNombre.Text= "";
			this.lblEstado.Text ="";

			this.txtidCliente.Value =0;
			this.txtIdCompra.Value = 0;
			this.txtIdRetencion.Value = 0;
			this.txtIdAsiento.Value = 0;
			this.txtEstado.Value = 0;
			this.txtRetenidoIVA.Value = 0;
			this.txtRetenidoRenta.Value = 0;
			this.txtRetenido.Value = 0;
			Proveedor = false;
			string sSQLCRD = string.Format("Exec ConsultaDetalleRetencion {0}", 0);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLCRD);

			#endregion vacia campos

			//			if (miAcceso.Nuevo)
			this.btnNuevo.Enabled = true;
			//			if (miAcceso.Buscar)
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;			
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;	
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			this.dtFecha.Value = DateTime.Today;
			#region Habilita campos
			this.cmbTipoRuc.Enabled = true;
			this.txtRuc.Enabled = true;
			this.txtSerieFact.Enabled = true;
			this.txtNumero.Enabled = true;
			this.txtAutFactura.Enabled  = true;
			this.dtCaducidad.Enabled = true;
			this.dtFecha.Enabled = true;
			this.txtFactura.Enabled = true;
			this.txtNotas.Enabled = true;
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
			#endregion Habilita campos

			#region controles
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;	
			this.btnGuardar.Enabled = true;
			#endregion controles

			bNuevo = true;
			bEdicion = true;

			this.cmbTipoRuc.Value = 4;
		}

		private void cmbTipoRuc_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbTipoRuc.ActiveRow != null) 
				{
					int idCliente = 0;
					this.txtRuc.Text = "";
					this.txtNombre.Text = "";
					//				 this.txtRuc.Enabled = true;

					if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 6) this.txtRuc.MaxLength = 13;
					if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;

					#region Consumidor Final
					if ((int)this.cmbTipoRuc.Value == 11) 
					{						
						this.txtRuc.Enabled = false;
						this.txtRuc.MaxLength = 13;
											
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select cr.idCliente, cl.Ruc, cl.Nombre From Cliente cl Inner Join ClienteTipoRuc cr On cr.idCliente = cl.idCliente Where cr.idTipoRuc = 11", true);
						dr.Read();
						if (dr.HasRows)
						{
							idCliente = dr.GetInt32(0);
							this.txtRuc.Value = dr.GetString(1);
							this.txtNombre.Value = dr.GetString(2);
						}
						dr.Close();
					}
					#endregion Consumidor Final					
				}
			}
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!bEdicion) return;
			
			if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;
			
			//if (this.txtRuc.Text.ToString().Length == 0) return;

			#region extrae informacion del cliente 
			int idCliente= Funcion.iEscalarSQL(cdsSeteoF,string.Format("If Exists(Select idCliente From Cliente Where Ruc= '{0}' And Proveedor = 0) Select idCliente From Cliente Where Ruc= '{0}' And Proveedor = 0 Else Select CONVERT(Int, 0)", this.txtRuc.Text.ToString()));
			
			if( idCliente > 0)
			{
				this.txtidCliente.Value = idCliente;
				this.txtNombre.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Nombre From cliente Where idCliente = {0} ", idCliente));
				this.txtRuc.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Ruc From cliente Where idCliente = {0} ", idCliente));
			}
			else 
			{
				using (frmClientesProveedores cliente  = new frmClientesProveedores(false,1,idCliente))
				{
					if (DialogResult.OK == cliente.ShowDialog())
					{
						this.txtidCliente.Value = cliente.txtidCliente.Value;
						this.txtNombre.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Nombre From cliente Where idCliente = {0} ", (int)this.txtidCliente.Value));
						this.txtRuc.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Ruc From cliente Where idCliente = {0} ", (int)this.txtidCliente.Value));						
					}
				}
			
			}
			#endregion Extrae informacion del cliente 
		}

    int icodigo = 0;
		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
//			if (e.Cell.Column.ToString() == "Codigo")
//			{
//				
//				if (e.Cell.Row.Cells["Codigo"].Value.ToString().Length < 1)
//				{
//					e.Cell.Row.Cells["Codigo"].Value = e.Cell.OriginalValue;
//				}
//			}
		
			if (e.Cell.Column.ToString() == "Codigo")
			{
				icodigo = 0;

				decimal dPorcentaje = 0.00m;
	
				string sSQLInfo = string.Format("Exec RetencionesCuenta '{0}',{1}", e.Cell.Row.Cells["Codigo"].Value.ToString(), (bool)e.Cell.Row.Cells["IVARenta"].Value);//bIva);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLInfo, true);
				dr.Read();
				if (dr.HasRows)
				{          
					e.Cell.Row.Cells["idRetencionIVARenta"].Value = dr.GetInt32(0);
					dPorcentaje = dr.GetDecimal(1);					
					e.Cell.Row.Cells["Detalle"].Value = dr.GetString(2);
					icodigo = 1;
				}				
				dr.Close();
				
				e.Cell.Row.Cells["Porcentaje"].Value = dPorcentaje;
			}
			if (!bActualiza) return;	

			if (e.Cell.Column.ToString() == "Codigo")
			{
				bActualiza = false;
				 if (icodigo == 0) { e.Cell.Row.Cells["Codigo"].Value = e.Cell.OriginalValue; }
				bActualiza = true;
	
			}

			if (e.Cell.Column.ToString() == "Porcentaje")
			{
				if (e.Cell.Row.Cells["Base"].Value != System.DBNull.Value)
				{         
					e.Cell.Row.Cells["Valor"].Value = (Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Base"].Value), 6))/100;			

					Suma();
				}				
			}

			if (e.Cell.Column.ToString() == "Base")
			{
				if (e.Cell.Row.Cells["Porcentaje"].Value != System.DBNull.Value)
				{         
					e.Cell.Row.Cells["Valor"].Value = (Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Base"].Value), 6))/100;			

					Suma();
				}				
			}
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if(e.Cell.Column.ToString()== "IVARenta")
			{
				if((int)e.Cell.Row.Cells["idRetencionIVARenta"].Value > 0)
				{
					e.Cancel=true;
				}
			}
		}

		private void ultraGrid1_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
//			if (bEdicion)
//			{
//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
//				{
//					dr.Cells["idDetalleRetencion"].Value = 0;
//					dr.Cells["idRetencion"].Value = 0; 
//					dr.Cells["IVARenta"].Value = false; 
//					dr.Cells["idRetencionIVARenta"].Value = 0; 
//					dr.Cells["Porcentaje"].Value = 0; 
//					dr.Cells["Base"].Value = 0; 
//					dr.Cells["idDetalleRetencion"].Value = 0;
//					dr.Cells["Valor"].Value = 0; 
//					dr.Cells["idDetalleRetencion"].Value = 0;
// 
//				}
//			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idRetencionIVARenta"].Value = 0; 
				e.Row.Cells["idDetalleRetencion"].Value = 0;
				e.Row.Cells["idRetencion"].Value = 0; 
				e.Row.Cells["IVARenta"].Value = false;				
				e.Row.Cells["Porcentaje"].Value = 0; 
				e.Row.Cells["Base"].Value = 0; 
				e.Row.Cells["Valor"].Value = 0; 
				e.Row.Cells["Codigo"].Value = ""; 
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region validacion
			//if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione una cuenta")) return;
			if (!Validacion.vbTexto(this.txtRuc, 10, 13, "Por favor ingrese C.I/RUC ")) return;
			//if (!Validacion.vbTexto(this.txtNombre, 10, 13, "Ingrese correctamente la identificacion")) return;

			if (!Validacion.vbTexto(this.txtSerieFact, 6, 6, "Serie")) return;

			if (!Validacion.vbTexto(this.txtNumero, 9, 9, "Numero")) return;
			
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione una Fecha", true, "Venta", cdsSeteoF)) return;
			
			if (!Validacion.vbTexto(this.txtAutFactura, 10, 48, "Autorizacion")) return;

			if (!Validacion.vbFechaEnDocumentos(this.dtCaducidad, "Seleccione la Fecha de caducidad", false, "", cdsSeteoF)) return;

			if (this.dtCaducidad.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la fecha de caducidad.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtCaducidad.Focus();	
				return;
			}
			
			#region Filas Vacias
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idRetencionIVARenta"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}
			}
			#endregion Filas Vacias
			
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese el detalle del documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
      
			int iCantidad =0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDTi in this.ultraGrid1.Rows.All)
			{
				if((Decimal)drDTi.Cells["Base"].Value == 0.00m)	iCantidad++;
				this.ultraGrid1.ActiveRow = drDTi;
				this.ultraGrid1.ActiveRow.Selected = true;				
			}

			if(iCantidad > 0) 
			{
				MessageBox.Show("No Puede Ingresar La Base en 0" , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
				//				ugDetalleFactura.ActiveRow = drDTi;
				ultraGrid1.ActiveRow.Selected = true;
				return;
			}

			if( (Decimal) this.txtRetenido.Value == 0.00m)
			{
				MessageBox.Show("El Documento No puede ser 0" , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
				return;
			}
			#endregion validacion
					
			#region Guardar Maestro
			string sSQLl = string.Format("Exec GuardaRetencionesTarjetas {0},{1},{2},{3},{4},'{5}','{6}','{7}','{8}','{9}','{10}',{11},{12},{13},{14}",
        
				(int)this.txtIdRetencion.Value,
				(int) this.txtIdCompra.Value,
				(int) this.txtIdAsiento.Value,
				Proveedor,
				(int) this.txtEstado.Value,
				this.txtSerieFact.Text,
				this.txtNumero.Text,
				this.txtAutFactura.Text,
				Convert.ToDateTime(this.dtCaducidad.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),
				this.txtNotas.Text,
				(Decimal) this.txtRetenidoIVA.Value,
				(Decimal) this.txtRetenidoRenta.Value,
				(Decimal) this.txtRetenido.Value,
				(int) this.txtidCliente.Value);
      
			this.txtIdRetencion.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQLl);
			#endregion Guardar Maestro

			#region Guarda detalle Maestro
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.ultraGrid1.Rows.All)
			{				
				string sSQLDetalle = string.Format("Exec GuardaDetalleRetiro  {0},{1},{2},{3},{4},{5},{6}", 
					(int)drr.Cells["idDetalleRetencion"].Value,
					(int)this.txtIdRetencion.Value,
					(bool)drr.Cells["IVARenta"].Value,
					(int)drr.Cells["idRetencionIVARenta"].Value, 
					(decimal)drr.Cells["Base"].Value, 
					(decimal)drr.Cells["Valor"].Value, 
					(decimal)drr.Cells["Porcentaje"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			}	
			#endregion Guarda detalle Maestro

			#region bloqueo controles
			this.cmbTipoRuc.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtSerieFact.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtAutFactura.Enabled  = false;
			this.dtCaducidad.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtFactura.Enabled = false;
			this.txtNotas.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

			if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
			if (miAcceso.Anular)this.btnAnular.Enabled = true;
			if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
			if (miAcceso.Editar)this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

			
			bNuevo = false;
			bEdicion = false;
			#endregion bloqueo controles
		}

		private void btnImprimir_Click_1(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;

			#region Impresion
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
			pfRet.ParameterFieldName = "@idCompra";
			idCompraRet.Value = (int)this.txtIdCompra.Value;
			pfRet.CurrentValues.Add (idCompraRet);
			paramFields.Add (pfRet);
				
			Reporte miRepor = new Reporte("Retenciones.rpt", "");
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("Retenciones");
			miRepor.crVista.ParameterFieldInfo = paramFields;
			#endregion Impresion
				
			miRepor.Show();			
						
			Cursor = Cursors.Default;
		}

		private void btnAnular_Click_1(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value != 2)
			{
				if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No puede Anular esta Retención, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();				
				}
				else if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Anular esta retención?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				{
					string sSQLAnul = string.Format("Exec AnulaRetencion {0}, {1}, 0, 4", (int)this.txtIdRetencion.Value, (int)this.txtIdAsiento.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAnul, true);

					MessageBox.Show("Retención Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
					this.txtEstado.Value = 2;
					this.lblEstado.Text = "ANULADO";
					this.btnAnular.Text = "Eliminar";
				}
			}
			else
			{
				if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No puede Eliminar esta Retención, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();				
				}
				else if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Eliminar esta retención?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				{			
					string sSQLAnul = string.Format("Exec AnulaRetencion {0}, {1}, 1, 4", (int)this.txtIdRetencion.Value, (int)this.txtIdAsiento.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAnul, true);

					MessageBox.Show("Retención Eliminada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
					this.txtEstado.Value = 2;
					this.lblEstado.Text = "";

					this.btnConsultar_Click(sender, e);
				}
			}
		}

		private void btnConsultar_Click_1(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaRetenciones reten = new frmConsultaRetenciones())
			{
				if (DialogResult.OK == reten.ShowDialog())
				{
					this.Consultar((int)reten.ultraGrid1.ActiveRow.Cells["idRetencion"].Value);
				}				
			}	
		}

		
		private void Consultar (int IdRetencion )
		{
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled =false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;

			#region Consulta Individual
			string ssql= string.Format ("Exec ConsultaIndividualRetencion {0}",IdRetencion); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
       
      this.txtIdRetencion.Value = dr.GetInt32 (0);
			if (dr.GetValue(1) != System.DBNull.Value) this.cmbTipoRuc.Value =  dr.GetInt32(1);
			if (dr.GetValue(2) != System.DBNull.Value) this.txtRuc.Text = dr.GetString (2);
			if (dr.GetValue(3) != System.DBNull.Value) this.txtNombre.Text = dr.GetString (3);
			if (dr.GetValue(4) != System.DBNull.Value) this.txtSerieFact.Text = dr.GetString (4);
			if (dr.GetValue(5) != System.DBNull.Value) this.txtNumero.Text = dr.GetString (5);
			if (dr.GetValue(6) != System.DBNull.Value) this.txtAutFactura.Text = dr.GetString (6);
			if (dr.GetValue(7) != System.DBNull.Value) this.dtCaducidad.Value = dr.GetDateTime (7);
			if (dr.GetValue(8) != System.DBNull.Value) this.dtFecha.Value = dr.GetDateTime (8);
//			if (dr.GetValue(9) != System.DBNull.Value) this.txtFactura.Text = dr.GetString (9);
			if (dr.GetValue(10) != System.DBNull.Value) this.txtNotas.Text = dr.GetString (10);
			if (dr.GetValue(11) != System.DBNull.Value) this.txtIdCompra.Value = dr.GetInt32 (11);
			if (dr.GetValue(12) != System.DBNull.Value) this.txtIdAsiento.Value = dr.GetInt32 (12);
			if (dr.GetValue(13) != System.DBNull.Value) this.txtEstado.Value = dr.GetInt32 (13);
			if (dr.GetValue(14) != System.DBNull.Value) this.txtidCliente.Value = dr.GetInt32 (14);
			if (dr.GetValue(15) != System.DBNull.Value) this.txtRetenidoIVA.Value = dr.GetDecimal(15);
			if (dr.GetValue(16) != System.DBNull.Value) this.txtRetenidoRenta.Value = dr.GetDecimal(16);
			if (dr.GetValue(17) != System.DBNull.Value) this.txtRetenido.Value = dr.GetDecimal(17);
			if (dr.GetValue(18) != System.DBNull.Value) Proveedor = dr.GetBoolean(18);
			dr.Close();

			string sSQLCRD = string.Format("Exec ConsultaDetalleRetencion {0}", (int)this.txtIdRetencion.Value);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLCRD);
			#endregion Consulta Individual

			if( (int)this.txtEstado.Value != 2)
			{
				if (miAcceso.Anular)this.btnAnular.Enabled = true;
				if (miAcceso.Editar)this.btnEditar.Enabled = true;
				if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
				
			}
			else this.lblEstado.Text = "ANULADO";
			
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
			this.btnCancelar.Enabled = true;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consultar((int)this.txtIdRetencion.Value);

			if((int)this.txtEstado.Value ==2)
			{
				MessageBox.Show("El documento ya fue anulado ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}		
			
			if (!miAcceso.Editar )
			{
				MessageBox.Show("No tiene acceso a Editar " + Text.Trim(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}


			#region habilita campos
//			this.cmbTipoRuc.Enabled = true;
//			this.txtRuc.Enabled = true;
			this.txtSerieFact.Enabled = true;
			this.txtNumero.Enabled = true;
			this.txtAutFactura.Enabled  = true;
			this.dtCaducidad.Enabled = true;
			this.dtFecha.Enabled = true;
			this.txtFactura.Enabled = true;
			this.txtNotas.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);		
			#endregion habilita campos

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnGuardar.Enabled = true;

			bNuevo = false;
			bEdicion = true;
		}

		private void txtSerieFact_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNumero.Focus();
		}

		private void cmbTipoRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtRuc.Focus();
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSerieFact.Focus();
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtAutFactura.Focus();
		}

		private void txtAutFactura_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtFecha.Focus();
		}

		private void dtCaducidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtCaducidad.Focus();
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);	
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			this.Suma();
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Base", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Valor", 2);
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				this.dtCaducidad.CalendarInfo.MinDate = DateTime.Today;
			else this.dtCaducidad.CalendarInfo.MinDate = DateTime.Parse("01/11/2017");
		}
	
	}
}

