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
	public class frmRetenciones : DevExpress.XtraEditors.XtraForm
	{
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutFactura;
		private System.Windows.Forms.Label label31;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblNumero;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerieFact;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdRetencion;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnAnular;
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
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFactura;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalFactura;
		private System.Windows.Forms.Button btnAsiento;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetenido;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetenidoIVA;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetenidoRenta;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtCaducidad;
		bool BNuevo = false;
		private System.Windows.Forms.Button btnFe;
		bool BCxP = false;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.GroupBox groupBox2;
		CultureInfo us = new CultureInfo("en-US");

		public frmRetenciones(int idCompra, string factura, decimal dTotal, bool bCxP)
		{
			InitializeComponent();
			
			IdCompra = idCompra;
			Factura = factura;
			DTotal = dTotal;
			BCxP = bCxP;
		}

		int IdComprobante = 0;
		
		public frmRetenciones(Infragistics.Win.UltraWinGrid.UltraGrid uGridCompra, int idCompra, int idCliente, string factura, DateTime dtFecha, decimal dIVA, decimal dTotal, bool bNuevo, int idComprobante)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			UGridCompra = uGridCompra;
			IdCompra = idCompra;
			IdCliente = idCliente;
			Factura = factura;
			DTFecha = dtFecha;
			DIVA = dIVA;
			DTotal = dTotal;
			BNuevo = bNuevo;
			IdComprobante = idComprobante;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRetenciones));
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVARenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencionIVARenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVARenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencionIVARenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.txtAutFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSerieFact = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.txtFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdRetencion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.txtRetenido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtRetenidoIVA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtRetenidoRenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnImprimir = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.dtCaducidad = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnFe = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenidoIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenidoRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtCaducidad)).BeginInit();
			this.SuspendLayout();
			// 
			// txtAutFactura
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutFactura.Appearance = appearance1;
			this.txtAutFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutFactura.Enabled = false;
			this.txtAutFactura.Location = new System.Drawing.Point(392, 8);
			this.txtAutFactura.MaxLength = 37;
			this.txtAutFactura.Name = "txtAutFactura";
			this.txtAutFactura.Size = new System.Drawing.Size(232, 22);
			this.txtAutFactura.TabIndex = 2;
			this.txtAutFactura.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAutFactura_KeyPress);
			this.txtAutFactura.Validating += new System.ComponentModel.CancelEventHandler(this.txtAutFactura_Validating);
			this.txtAutFactura.ValueChanged += new System.EventHandler(this.txtAutFactura_ValueChanged);
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(312, 11);
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
			this.label3.Location = new System.Drawing.Point(8, 11);
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
			this.lblNumero.Location = new System.Drawing.Point(144, 11);
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
			this.txtNumero.Location = new System.Drawing.Point(200, 8);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(96, 22);
			this.txtNumero.TabIndex = 1;
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
			this.txtSerieFact.Location = new System.Drawing.Point(56, 8);
			this.txtSerieFact.MaxLength = 6;
			this.txtSerieFact.Name = "txtSerieFact";
			this.txtSerieFact.Size = new System.Drawing.Size(64, 22);
			this.txtSerieFact.TabIndex = 0;
			this.txtSerieFact.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerieFact_KeyPress);
			this.txtSerieFact.Validating += new System.ComponentModel.CancelEventHandler(this.txtSerieFact_Validating);
			this.txtSerieFact.ValueChanged += new System.EventHandler(this.txtSerieFact_ValueChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 40);
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
			this.dtFecha.Location = new System.Drawing.Point(56, 38);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 3;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance5.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance5;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 74;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance6;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 103;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 102;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 381;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance7;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 101;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance8;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 102;
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
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 10F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance11;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(810, 160);
			this.ultraGrid1.TabIndex = 305;
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
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
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(210, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(41, 16);
			this.label1.TabIndex = 307;
			this.label1.Text = "Factura";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFactura
			// 
			this.txtFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFactura.Appearance = appearance15;
			this.txtFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFactura.Enabled = false;
			this.txtFactura.Location = new System.Drawing.Point(266, 37);
			this.txtFactura.Name = "txtFactura";
			this.txtFactura.Size = new System.Drawing.Size(112, 22);
			this.txtFactura.TabIndex = 306;
			// 
			// txtIdRetencion
			// 
			this.txtIdRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdRetencion.Enabled = false;
			this.txtIdRetencion.Location = new System.Drawing.Point(496, 280);
			this.txtIdRetencion.Name = "txtIdRetencion";
			this.txtIdRetencion.PromptChar = ' ';
			this.txtIdRetencion.Size = new System.Drawing.Size(16, 22);
			this.txtIdRetencion.TabIndex = 326;
			this.txtIdRetencion.Visible = false;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(512, 280);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtIdCompra.TabIndex = 325;
			this.txtIdCompra.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(544, 280);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 327;
			this.txtEstado.Visible = false;
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(402, 40);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(70, 16);
			this.lblSubTotal.TabIndex = 328;
			this.lblSubTotal.Text = "Total Factura";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRetenido
			// 
			this.txtRetenido.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetenido.Appearance = appearance16;
			this.txtRetenido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetenido.Enabled = false;
			this.txtRetenido.FormatString = "#,##0.00";
			this.txtRetenido.Location = new System.Drawing.Point(720, 320);
			this.txtRetenido.Name = "txtRetenido";
			this.txtRetenido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetenido.PromptChar = ' ';
			this.txtRetenido.Size = new System.Drawing.Size(96, 22);
			this.txtRetenido.TabIndex = 330;
			// 
			// txtTotalFactura
			// 
			this.txtTotalFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalFactura.Appearance = appearance17;
			this.txtTotalFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalFactura.Enabled = false;
			this.txtTotalFactura.FormatString = "#,##0.00";
			this.txtTotalFactura.Location = new System.Drawing.Point(490, 37);
			this.txtTotalFactura.Name = "txtTotalFactura";
			this.txtTotalFactura.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalFactura.PromptChar = ' ';
			this.txtTotalFactura.Size = new System.Drawing.Size(80, 22);
			this.txtTotalFactura.TabIndex = 329;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(656, 320);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(29, 16);
			this.label2.TabIndex = 331;
			this.label2.Text = "Total";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(656, 256);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(32, 16);
			this.label4.TabIndex = 332;
			this.label4.Text = "I.V.A.";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRetenidoIVA
			// 
			this.txtRetenidoIVA.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetenidoIVA.Appearance = appearance18;
			this.txtRetenidoIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetenidoIVA.Enabled = false;
			this.txtRetenidoIVA.FormatString = "#,##0.00";
			this.txtRetenidoIVA.Location = new System.Drawing.Point(720, 256);
			this.txtRetenidoIVA.Name = "txtRetenidoIVA";
			this.txtRetenidoIVA.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetenidoIVA.PromptChar = ' ';
			this.txtRetenidoIVA.Size = new System.Drawing.Size(96, 22);
			this.txtRetenidoIVA.TabIndex = 333;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(656, 288);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(33, 16);
			this.label5.TabIndex = 334;
			this.label5.Text = "Renta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRetenidoRenta
			// 
			this.txtRetenidoRenta.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetenidoRenta.Appearance = appearance19;
			this.txtRetenidoRenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetenidoRenta.Enabled = false;
			this.txtRetenidoRenta.FormatString = "#,##0.00";
			this.txtRetenidoRenta.Location = new System.Drawing.Point(720, 288);
			this.txtRetenidoRenta.Name = "txtRetenidoRenta";
			this.txtRetenidoRenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetenidoRenta.PromptChar = ' ';
			this.txtRetenidoRenta.Size = new System.Drawing.Size(96, 22);
			this.txtRetenidoRenta.TabIndex = 335;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(570, 368);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 336;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(8, 256);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(37, 16);
			this.label11.TabIndex = 338;
			this.label11.Text = "Notas:";
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Enabled = false;
			this.btnAceptar.Image = ((System.Drawing.Image)(resources.GetObject("btnAceptar.Image")));
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(8, 368);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(72, 23);
			this.btnAceptar.TabIndex = 5;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(85, 368);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 6;
			this.btnConsultar.Text = "&Cancelar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(335, 368);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 9;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(528, 280);
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
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance20;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 280);
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
			this.groupBox1.Location = new System.Drawing.Point(0, 352);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(826, 8);
			this.groupBox1.TabIndex = 344;
			this.groupBox1.TabStop = false;
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.ImageList = this.imageList1;
			this.btnAsiento.Location = new System.Drawing.Point(168, 368);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
			this.btnAsiento.TabIndex = 7;
			this.btnAsiento.Text = "&Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			this.btnAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnAsiento_MouseDown);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(253, 368);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 8;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(640, 11);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(56, 16);
			this.label7.TabIndex = 348;
			this.label7.Text = "Caducidad";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtCaducidad
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtCaducidad.Appearance = appearance21;
			this.dtCaducidad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtCaducidad.DateButtons.Add(dateButton2);
			this.dtCaducidad.Enabled = false;
			this.dtCaducidad.Format = "dd/MM/yyyy";
			this.dtCaducidad.Location = new System.Drawing.Point(704, 9);
			this.dtCaducidad.Name = "dtCaducidad";
			this.dtCaducidad.NonAutoSizeHeight = 23;
			this.dtCaducidad.Size = new System.Drawing.Size(112, 21);
			this.dtCaducidad.SpinButtonsVisible = true;
			this.dtCaducidad.TabIndex = 347;
			this.dtCaducidad.Value = ((object)(resources.GetObject("dtCaducidad.Value")));
			this.dtCaducidad.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtCaducidad_KeyPress);
			this.dtCaducidad.Validating += new System.ComponentModel.CancelEventHandler(this.dtCaducidad_Validating);
			this.dtCaducidad.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtCaducidad_BeforeDropDown);
			// 
			// btnFe
			// 
			this.btnFe.CausesValidation = false;
			this.btnFe.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFe.Location = new System.Drawing.Point(736, 368);
			this.btnFe.Name = "btnFe";
			this.btnFe.TabIndex = 628;
			this.btnFe.Text = "&FE";
			this.btnFe.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnFe.Visible = false;
			this.btnFe.Click += new System.EventHandler(this.btnFe_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 72);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(832, 8);
			this.groupBox2.TabIndex = 629;
			this.groupBox2.TabStop = false;
			// 
			// frmRetenciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(826, 400);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnFe);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label31);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtCaducidad);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtRetenidoRenta);
			this.Controls.Add(this.txtRetenidoIVA);
			this.Controls.Add(this.txtRetenido);
			this.Controls.Add(this.txtTotalFactura);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdRetencion);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtFactura);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtAutFactura);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtSerieFact);
			this.Controls.Add(this.dtFecha);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmRetenciones";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Retención";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRetenciones_KeyDown);
			this.Load += new System.EventHandler(this.frmRetenciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenidoIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenidoRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtCaducidad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

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

		bool bAutomatico = false;
		bool bFE = false;

		private void frmRetenciones_Load(object sender, System.EventArgs e)
		{
			#region Factura De Compra
			string stTipo = "040601";
									
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			
			bool bRetIVAProv = false;
			this.dtFecha.CalendarInfo.MinDate = DTFecha;
			this.txtFactura.Text = Factura;
			this.txtTotalFactura.Value = DTotal;

			bRetIVAProv = Funcion.bEscalarSQL(cdsSeteoF, string.Format("If Exists(SELECT ISNULL(ContribEspecial, 0) FROM Cliente WHERE idCliente = {0} And Proveedor = 1) SELECT ISNULL(ContribEspecial, 0) FROM Cliente WHERE idCliente = {0} And Proveedor = 1 Else Select Convert(Bit, 0)", IdCliente), true);
			bFE = Funcion.bEscalarSQL(cdsSeteoF, "Select FacturaElectronica From SeteoF", false);
	
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Retenciones Where Estado <> 2 And ClienteProveedor = 1 And idCompra = {0}", IdCompra)) == 0)
			{
				this.dtFecha.Value = DTFecha;
				this.dtFecha.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
				this.dtFecha.CalendarInfo.MaxDate = (DateTime)this.dtFecha.Value;

				bAutomatico = Funcion.bEscalarSQL (cdsSeteoF, "Select Automatico From CompraNumero Where idTipoFactura = 10", false);

				#region Parametros por default
				this.txtIdRetencion.Value = 0;
				this.txtIdCompra.Value = IdCompra;
				this.txtIdAsiento.Value = 0;
				this.txtEstado.Value = 0;
				if (!bFE)
				{
					SqlDataReader drParametros = Funcion.rEscalarSQL(cdsSeteoF, "Select SriSerieRet, SriAutRet, FechaRCad From Seteo", true);
					drParametros.Read();
					if (drParametros.HasRows)
					{
						this.txtSerieFact.Text = drParametros.GetString(0);
						this.txtAutFactura.Text = drParametros.GetString(1);
						this.dtCaducidad.Value = drParametros.GetDateTime(2);
					}
					drParametros.Close();
				}				
				#endregion Parametros por default
				
				#region Variables Retencion Articulo
				int idArticulo = 0;
				int idTipoGrupo = 0;
				decimal dImpuesto = 0;
				decimal dSubtotal = 0;				
				decimal dDescuento = 0;
				decimal dRenta = 0;
				decimal dIVA = 0;
				int idRetencionRenta = 0;
				int idRetencionIVA = 0;
				decimal sPorc = 0;
				string sCodigo = "";
				string sDetalle = "";
				decimal dBaseRenta = 0;
				decimal dBaseIVA = 0;
				decimal dRentaRetenida = 0;
				decimal dIVARetenido = 0;				
				int iContR = 0;
				int iContI = 0;
				decimal dBaseIVAProv = 0;
				#endregion Variables Retencion Articulo
				
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in UGridCompra.Rows.All)
				{
					#region Inicializa Variables
					idArticulo = 0;
					dImpuesto = 0.00m;
					dSubtotal = 0.00m;
					dDescuento = 0.00m;
					idRetencionRenta = 0;
					idRetencionIVA = 0;
					sPorc = 0.00m;
					sCodigo = "";
					sDetalle = "";
					dBaseRenta = 0.00m;
					dBaseIVA = 0.00m;
					dRentaRetenida = 0.00m;
					dIVARetenido = 0.00m;
					iContR = 0;
					iContI = 0;
					dIVA = 0.00m;
					idTipoGrupo = 0;
					#endregion Inicializa Variables

					#region Asignación de valores
					idArticulo = (int)dr.Cells["idArticulo"].Value;
					dImpuesto = (decimal)dr.Cells["Impuesto"].Value;
					dSubtotal = (decimal)dr.Cells["Subtotal"].Value;
					dDescuento = (decimal)dr.Cells["DescuentoArt"].Value;
					dRenta = dSubtotal - dDescuento;
					dIVA = ((dSubtotal - dDescuento) * dImpuesto)/100.00m;
					if (UGridCompra.Rows.Count == 1) dIVA = decimal.Round(dIVA, 2);
					#endregion Asignación de valores

					if (IdComprobante == 1 || IdComprobante == 3)
					{
						#region I.V.A.
//						if (!bRetIVAProv)
//						{
//							int idRetIVAProveedor = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idRetIva From Cliente Where Proveedor = 1 And idCliente = {0}", IdCliente));
//							if (idRetIVAProveedor < 2)
//							{
								if (dImpuesto > 0.00m)
								{								
//									idTipoGrupo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idTipoGrupo From Articulo Where idArticulo = {0}", idArticulo));
//									if (idTipoGrupo != 1)
									//									{ 	@idArticulo Int, @Base Float, @Fecha Datetime, @idTipoFactura Int, @idCliente Int

										string sSQLIVA = string.Format("Exec RetencionArticuloIVA {0}, {1}, '{2}', 4, {3}", idArticulo, dIVA, DTFecha.ToString("yyyyMMdd"), IdCliente);
										SqlDataReader drIVA = Funcion.rEscalarSQL(cdsSeteoF, sSQLIVA, true);
										drIVA.Read();
										if (drIVA.HasRows)
										{
											idRetencionIVA = drIVA.GetInt32(0);
											sPorc = drIVA.GetDecimal(1);
											sCodigo = drIVA.GetString(2);
											sDetalle = drIVA.GetString(3);
											dBaseIVA = drIVA.GetDecimal(4);	
											dIVARetenido = drIVA.GetDecimal(5);	
										}
										drIVA.Close();

										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr2 in this.ultraGrid1.Rows.All)
										{
											if ((int)dr2.Cells["idRetencionIVARenta"].Value == idRetencionIVA && !(bool) dr2.Cells["IVARenta"].Value)	
											{
												dr2.Cells["Base"].Value = Convert.ToDecimal(dr2.Cells["Base"].Value) + dBaseIVA;
												dr2.Cells["Valor"].Value = Convert.ToDecimal(dr2.Cells["Valor"].Value) + dIVARetenido;
												iContI ++;
											}
										}

										if (idRetencionIVA > 0 && iContI == 0)
										{
											this.ultraGrid1.Rows.Band.AddNew();
											this.ultraGrid1.ActiveRow.Cells["idDetalleRetencion"].Value = 0;
											this.ultraGrid1.ActiveRow.Cells["idRetencion"].Value = 0;
											this.ultraGrid1.ActiveRow.Cells["IVARENTA"].Value = false;
											this.ultraGrid1.ActiveRow.Cells["idRetencionIVARenta"].Value = idRetencionIVA;
											this.ultraGrid1.ActiveRow.Cells["Porcentaje"].Value = sPorc;
											this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = sCodigo;
											this.ultraGrid1.ActiveRow.Cells["Detalle"].Value = sDetalle;
											this.ultraGrid1.ActiveRow.Cells["Base"].Value = dBaseIVA;
											this.ultraGrid1.ActiveRow.Cells["Valor"].Value = dIVARetenido;	
										}
									}
									else
									{
										dBaseIVAProv = dBaseIVAProv + (dSubtotal - dDescuento);									
									}
//								}
							//}
//						}
						#endregion I.V.A.
					}

					#region Renta
					string sSQLRenta = string.Format("Exec RetencionArticuloRenta {0}, {1}", idArticulo, dRenta);
					SqlDataReader drRenta = Funcion.rEscalarSQL(cdsSeteoF, sSQLRenta, true);
					//drRenta.Read();
					while (drRenta.Read())
						//					if (drRenta.HasRows)
					{
						idRetencionRenta = drRenta.GetInt32(0);
						sPorc = drRenta.GetDecimal(1);
						sCodigo = drRenta.GetString(2);
						sDetalle = drRenta.GetString(3);
						dBaseRenta = drRenta.GetDecimal(4);
						dRentaRetenida = drRenta.GetDecimal(5);	
					

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid1.Rows.All)
						{
							if ((int)dr1.Cells["idRetencionIVARenta"].Value == idRetencionRenta && (bool) dr1.Cells["IVARenta"].Value)	
							{
								dr1.Cells["Base"].Value = Convert.ToDecimal(dr1.Cells["Base"].Value) + dBaseRenta;
								dr1.Cells["Valor"].Value = Convert.ToDecimal(dr1.Cells["Valor"].Value) + dRentaRetenida;
								iContR ++;
							}
						}

						if (iContR == 0 && idRetencionRenta > 0)
						{
							this.ultraGrid1.Rows.Band.AddNew();
							this.ultraGrid1.ActiveRow.Cells["idDetalleRetencion"].Value = 0;
							this.ultraGrid1.ActiveRow.Cells["idRetencion"].Value = 0;
							this.ultraGrid1.ActiveRow.Cells["IVARENTA"].Value = true;
							this.ultraGrid1.ActiveRow.Cells["idRetencionIVARenta"].Value = idRetencionRenta;
							this.ultraGrid1.ActiveRow.Cells["Porcentaje"].Value = sPorc;
							this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = sCodigo;
							this.ultraGrid1.ActiveRow.Cells["Detalle"].Value = sDetalle;
							this.ultraGrid1.ActiveRow.Cells["Base"].Value = dBaseRenta;
							this.ultraGrid1.ActiveRow.Cells["Valor"].Value = dRentaRetenida;	
						}
					}
					drRenta.Close();
					#endregion Renta
				}

				#region IVA Proveedor
//				if (DIVA > 0)
//				{
//					//								decimal dBaseValorIVA = Math.Round(dBaseIVAProv * 12 /100, 2);
//					string sSQL = string.Format("Exec RetencionIVAProveedor {0}, {1}", IdCliente, DIVA);
//					SqlDataReader drReader = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
//					drReader.Read();
//					if (drReader.HasRows)
//					{
//						this.ultraGrid1.Rows.Band.AddNew();
//						this.ultraGrid1.ActiveRow.Cells["idDetalleRetencion"].Value = 0;
//						this.ultraGrid1.ActiveRow.Cells["idRetencion"].Value = 0;
//						this.ultraGrid1.ActiveRow.Cells["IVARENTA"].Value = false;
//						this.ultraGrid1.ActiveRow.Cells["idRetencionIVARenta"].Value = drReader.GetInt32(0);
//						this.ultraGrid1.ActiveRow.Cells["Porcentaje"].Value = drReader.GetDecimal(1);
//						this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = drReader.GetString(2);
//						this.ultraGrid1.ActiveRow.Cells["Detalle"].Value = drReader.GetString(3);
//						this.ultraGrid1.ActiveRow.Cells["Base"].Value = drReader.GetDecimal(4);
//						this.ultraGrid1.ActiveRow.Cells["Valor"].Value = drReader.GetDecimal(5);	
//					}
//					drReader.Close();
//				}	
				#endregion IVA Proveedor
								
				#region Retencion Proveedor
//				if (IdComprobante == 3)
//				{
					if (!bRetIVAProv)
					{
						if (DIVA > 0.00m)
						{
							//decimal dBaseValorIVA = Math.Round(DIVA * 12 /100, 2);
							string sSQL = string.Format("Exec RetencionIVAProveedor {0}, {1}", IdCliente, DIVA);// dBaseValorIVA);
							SqlDataReader drReader = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
							drReader.Read();
							if (drReader.HasRows)
							{
								this.ultraGrid1.Rows.Band.AddNew();
								this.ultraGrid1.ActiveRow.Cells["idDetalleRetencion"].Value = 0;
								this.ultraGrid1.ActiveRow.Cells["idRetencion"].Value = 0;
								this.ultraGrid1.ActiveRow.Cells["IVARENTA"].Value = false;
								this.ultraGrid1.ActiveRow.Cells["idRetencionIVARenta"].Value = drReader.GetInt32(0);
								this.ultraGrid1.ActiveRow.Cells["Porcentaje"].Value = drReader.GetDecimal(1);
								this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = drReader.GetString(2);
								this.ultraGrid1.ActiveRow.Cells["Detalle"].Value = drReader.GetString(3);
								this.ultraGrid1.ActiveRow.Cells["Base"].Value = drReader.GetDecimal(4);
								this.ultraGrid1.ActiveRow.Cells["Valor"].Value = drReader.GetDecimal(5);	
							}
							drReader.Close();
						}	
					}
//				}
				#endregion Retencion Proveedor

				this.ultraGrid1.Rows.Band.AddNew();

				Suma();

				this.btnAceptar.Enabled = true;
			}
			else
			{
				string sSQLCRE = string.Format("Exec ConsultaRetencion {0}, 1", IdCompra);
				SqlDataReader drConsulta = Funcion.rEscalarSQL(cdsSeteoF, sSQLCRE, true);
				drConsulta.Read();
				if (drConsulta.HasRows)
				{
					this.txtIdRetencion.Value = drConsulta.GetInt32(0);
					this.txtIdCompra.Value = drConsulta.GetInt32(1);
					this.txtIdAsiento.Value = drConsulta.GetInt32(2);
					this.txtEstado.Value = drConsulta.GetInt32(4);
					this.txtSerieFact.Text = drConsulta.GetString(5);
					this.txtNumero.Text = drConsulta.GetString(6);
					this.txtAutFactura.Text = drConsulta.GetString(7);
					this.dtCaducidad.Value = drConsulta.GetDateTime(8);
					this.dtFecha.Value = drConsulta.GetDateTime(9);
					this.txtNotas.Text = drConsulta.GetString(10);
					this.txtRetenidoIVA.Value = drConsulta.GetDecimal(11);
					this.txtRetenidoRenta.Value = drConsulta.GetDecimal(12);
					this.txtRetenido.Value = drConsulta.GetDecimal(13);

					if ((int)this.txtEstado.Value == 2)
					{
						this.lblEstado.Text = "ANULADO";
						this.btnAnular.Text = "Eliminar";
					}
				}
				drConsulta.Close();

				this.dtFecha.Value = DTFecha;

				this.dtFecha.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
				this.dtFecha.CalendarInfo.MaxDate = (DateTime)this.dtFecha.Value; 

				if (!BNuevo || BCxP)
				{
					string sSQLCRD = string.Format("Exec ConsultaDetalleRetencion {0}", (int)this.txtIdRetencion.Value);
					this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLCRD);
				}
				else
				{
					#region Variables Retencion Articulo
					int idArticulo = 0;
					int idTipoGrupo = 0;
					decimal dImpuesto = 0.00m;
					decimal dSubtotal = 0.00m;				
					decimal dDescuento = 0.00m;
					decimal dRenta = 0.00m;
					decimal dIVA = 0.00m;
					int idRetencionRenta = 0;
					int idRetencionIVA = 0;
					decimal sPorc = 0.00m;
					string sCodigo = "";
					string sDetalle = "";
					decimal dBaseRenta = 0.00m;
					decimal dBaseIVA = 0.00m;
					decimal dRentaRetenida = 0.00m;
					decimal dIVARetenido = 0.00m;				
					int iContR = 0;
					int iContI = 0;
					decimal dBaseIVAProv = 0.00m;
					#endregion Variables Retencion Articulo
					
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in UGridCompra.Rows.All)
					{
						#region Inicializa Variables
						idArticulo = 0;
						dImpuesto = 0.00m;
						dSubtotal = 0.00m;
						dDescuento = 0.00m;
						idRetencionRenta = 0;
						idRetencionIVA = 0;
						sPorc = 0.00m;
						sCodigo = "";
						sDetalle = "";
						dBaseRenta = 0.00m;
						dBaseIVA = 0.00m;
						dRentaRetenida = 0.00m;
						dIVARetenido = 0.00m;
						iContR = 0;
						iContI = 0;
						dIVA = 0.00m;
						idTipoGrupo = 0;
						#endregion Inicializa Variables

						#region Asignación de valores
						idArticulo = (int)dr.Cells["idArticulo"].Value;
						dImpuesto = (decimal)dr.Cells["Impuesto"].Value;
						dSubtotal = (decimal)dr.Cells["Subtotal"].Value;
						dDescuento = (decimal)dr.Cells["DescuentoArt"].Value;
						dRenta = dSubtotal - dDescuento;
						dIVA = Math.Round(((dSubtotal - dDescuento) * dImpuesto)/100.00m, 2);
						#endregion Asignación de valores

						if (IdComprobante == 1 || IdComprobante == 3)
						{
							#region I.V.A.
//							if (!bRetIVAProv)
//							{
								//								int idRetIVAProveedor = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idRetIva From Cliente Where Proveedor = 1 And idCliente = {0}", IdCliente));
								//								if (idRetIVAProveedor < 2)
								//								{
								if (dImpuesto > 0)
								{
									//										idTipoGrupo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idTipoGrupo From Articulo Where idArticulo = {0}", idArticulo));
									//										if (idTipoGrupo != 1)
									//										{@idArticulo Int, @Base Float, @Fecha Datetime, @idTipoFactura Int, @idCliente Int
									string sSQLIVA = string.Format("Exec RetencionArticuloIVA {0}, {1}, '{2}', 4, {3}", idArticulo, dIVA, DTFecha.ToString("yyyyMMdd"), IdCliente);
									SqlDataReader drIVA = Funcion.rEscalarSQL(cdsSeteoF, sSQLIVA, true);
									drIVA.Read();
									if (drIVA.HasRows)
									{
										idRetencionIVA = drIVA.GetInt32(0);
										sPorc = drIVA.GetDecimal(1);
										sCodigo = drIVA.GetString(2);
										sDetalle = drIVA.GetString(3);
										dBaseIVA = drIVA.GetDecimal(4);	
										dIVARetenido = drIVA.GetDecimal(5);	
									}
									drIVA.Close();

									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr2 in this.ultraGrid1.Rows.All)
									{
										if ((int)dr2.Cells["idRetencionIVARenta"].Value == idRetencionIVA && !(bool) dr2.Cells["IVARenta"].Value)	
										{
											dr2.Cells["Base"].Value = Convert.ToDecimal(dr2.Cells["Base"].Value) + dBaseIVA;
											dr2.Cells["Valor"].Value = Convert.ToDecimal(dr2.Cells["Valor"].Value) + dIVARetenido;
											iContI ++;
										}
									}

									if (idRetencionIVA > 0 && iContI == 0)
									{
										this.ultraGrid1.Rows.Band.AddNew();
										this.ultraGrid1.ActiveRow.Cells["idDetalleRetencion"].Value = 0;
										this.ultraGrid1.ActiveRow.Cells["idRetencion"].Value = 0;
										this.ultraGrid1.ActiveRow.Cells["IVARENTA"].Value = false;
										this.ultraGrid1.ActiveRow.Cells["idRetencionIVARenta"].Value = idRetencionIVA;
										this.ultraGrid1.ActiveRow.Cells["Porcentaje"].Value = sPorc;
										this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = sCodigo;
										this.ultraGrid1.ActiveRow.Cells["Detalle"].Value = sDetalle;
										this.ultraGrid1.ActiveRow.Cells["Base"].Value = dBaseIVA;
										this.ultraGrid1.ActiveRow.Cells["Valor"].Value = dIVARetenido;	
									}
								}
								else
								{
									dBaseIVAProv = dBaseIVAProv + (dSubtotal - dDescuento);
								}
								//									}
								//								}
//							}
							#endregion I.V.A.
						}

						#region Renta
						string sSQLRenta = string.Format("Exec RetencionArticuloRenta {0}, {1}", idArticulo, dRenta);
						SqlDataReader drRenta = Funcion.rEscalarSQL(cdsSeteoF, sSQLRenta, true);
						//drRenta.Read();
						//if (drRenta.HasRows)
						while (drRenta.Read())
						{
							idRetencionRenta = drRenta.GetInt32(0);
							sPorc = drRenta.GetDecimal(1);
							sCodigo = drRenta.GetString(2);
							sDetalle = drRenta.GetString(3);
							dBaseRenta = drRenta.GetDecimal(4);
							dRentaRetenida = drRenta.GetDecimal(5);	
						
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid1.Rows.All)
							{
								if ((int)dr1.Cells["idRetencionIVARenta"].Value == idRetencionRenta && (bool) dr1.Cells["IVARenta"].Value)	
								{
									dr1.Cells["Base"].Value = (decimal)dr1.Cells["Base"].Value + dBaseRenta;
									dr1.Cells["Valor"].Value = (decimal)dr1.Cells["Valor"].Value + dRentaRetenida;
									iContR ++;
								}
							}

							if (iContR == 0 && idRetencionRenta > 0)
							{
								this.ultraGrid1.Rows.Band.AddNew();
								this.ultraGrid1.ActiveRow.Cells["idDetalleRetencion"].Value = 0;
								this.ultraGrid1.ActiveRow.Cells["idRetencion"].Value = 0;
								this.ultraGrid1.ActiveRow.Cells["IVARENTA"].Value = true;
								this.ultraGrid1.ActiveRow.Cells["idRetencionIVARenta"].Value = idRetencionRenta;
								this.ultraGrid1.ActiveRow.Cells["Porcentaje"].Value = sPorc;
								this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = sCodigo;
								this.ultraGrid1.ActiveRow.Cells["Detalle"].Value = sDetalle;
								this.ultraGrid1.ActiveRow.Cells["Base"].Value = dBaseRenta;
								this.ultraGrid1.ActiveRow.Cells["Valor"].Value = dRentaRetenida;	
							}
						}
						drRenta.Close();
						#endregion Renta
					}

					#region IVA Proveedor
//					if (DIVA > 0)
//					{
//						//								decimal dBaseValorIVA = Math.Round(dBaseIVAProv * 12 /100, 2);
//						string sSQL = string.Format("Exec RetencionIVAProveedor {0}, {1}", IdCliente, DIVA);
//						SqlDataReader drReader = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
//						drReader.Read();
//						if (drReader.HasRows)
//						{
//							this.ultraGrid1.Rows.Band.AddNew();
//							this.ultraGrid1.ActiveRow.Cells["idDetalleRetencion"].Value = 0;
//							this.ultraGrid1.ActiveRow.Cells["idRetencion"].Value = 0;
//							this.ultraGrid1.ActiveRow.Cells["IVARENTA"].Value = false;
//							this.ultraGrid1.ActiveRow.Cells["idRetencionIVARenta"].Value = drReader.GetInt32(0);
//							this.ultraGrid1.ActiveRow.Cells["Porcentaje"].Value = drReader.GetDecimal(1);
//							this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = drReader.GetString(2);
//							this.ultraGrid1.ActiveRow.Cells["Detalle"].Value = drReader.GetString(3);
//							this.ultraGrid1.ActiveRow.Cells["Base"].Value = drReader.GetDecimal(4);
//							this.ultraGrid1.ActiveRow.Cells["Valor"].Value = drReader.GetDecimal(5);	
//						}
//						drReader.Close();
//					}
					#endregion IVA Proveedor

					#region Retencion Proveedor
					if (!bRetIVAProv)
					{
						if (DIVA > 0)
						{
							//								decimal dBaseValorIVA = Math.Round(dBaseIVAProv * 12 /100, 2);
							string sSQL = string.Format("Exec RetencionIVAProveedor {0}, {1}", IdCliente, DIVA);
							SqlDataReader drReader = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
							drReader.Read();
							if (drReader.HasRows)
							{
								this.ultraGrid1.Rows.Band.AddNew();
								this.ultraGrid1.ActiveRow.Cells["idDetalleRetencion"].Value = 0;
								this.ultraGrid1.ActiveRow.Cells["idRetencion"].Value = 0;
								this.ultraGrid1.ActiveRow.Cells["IVARENTA"].Value = false;
								this.ultraGrid1.ActiveRow.Cells["idRetencionIVARenta"].Value = drReader.GetInt32(0);
								this.ultraGrid1.ActiveRow.Cells["Porcentaje"].Value = drReader.GetDecimal(1);
								this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = drReader.GetString(2);
								this.ultraGrid1.ActiveRow.Cells["Detalle"].Value = drReader.GetString(3);
								this.ultraGrid1.ActiveRow.Cells["Base"].Value = drReader.GetDecimal(4);
								this.ultraGrid1.ActiveRow.Cells["Valor"].Value = drReader.GetDecimal(5);	
							}
							drReader.Close();
						}	
					}
					#endregion Retencion Proveedor

					this.ultraGrid1.Rows.Band.AddNew();

					Suma();

					this.btnAceptar.Enabled = true;
				}

				if ((int)this.txtEstado.Value != 2)
				{
					if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
					if (miAcceso.Anular) this.btnAnular.Enabled = true;
				}
				else
				{
					if (miAcceso.Eliminar) this.btnAnular.Enabled = true;
				}
			}
			#endregion Factura De Compra

			if (!BCxP)
			{
				if (!BNuevo)
				{
					if (miAcceso.Editar)
					{
						if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
						{
							if (!bAutomatico)
							{													
								this.txtNumero.Enabled = true;
								this.txtNumero.Focus();
							}
							this.txtSerieFact.Enabled = true;
							if (!bFE)
							{	
								this.txtAutFactura.Enabled = true;
								this.dtCaducidad.Enabled = true;
							}
							this.dtFecha.Enabled = true;
							this.txtNotas.Enabled = true;

							this.btnAceptar.Enabled = true;
						}
					}
				}
				else
				{
					if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
					{											
						this.txtSerieFact.Enabled = true;
						if (!bAutomatico)
						{
							this.txtNumero.Focus();
							this.txtNumero.Enabled = true;
						}
						if (!bFE)
						{
							this.txtAutFactura.Enabled = true;
							this.dtCaducidad.Enabled = true;
						}
						this.dtFecha.Enabled = true;
						this.txtNotas.Enabled = true;

						this.btnAceptar.Enabled = true;
					}
				}
			}

			this.txtNumero.Focus();
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
			if (e.KeyChar == 13) this.btnAceptar.Focus();
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
				return;
			}
			if (this.txtSerieFact.Text.Length < 6)
			{
				MessageBox.Show("El número de serie es de 6 digitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerieFact.Focus();				
				return;
			}

			if (this.txtNumero.Text.ToString().Length == 0 && !bAutomatico)
			{
				MessageBox.Show("Ingrese el número del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();				
				return;
			}
			
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha del documento", true, "Compra", cdsSeteoF)) return;
						
			if (!bFE)
			{
				if (this.txtAutFactura.Text.Length == 0)
				{
					MessageBox.Show("Ingrese el número de autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtAutFactura.Focus();				
					return;
				}
				if (this.txtAutFactura.Text.Length < 10 && this.txtAutFactura.Text.Length > 37)
				{
					MessageBox.Show("Longitud de autorización debe estar entre 10 y 37 caracteres.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtAutFactura.Focus();				
					return;
				}
				if (this.dtCaducidad.Value == DBNull.Value)
				{
					MessageBox.Show("Ingrese la fecha de caducidad.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtCaducidad.Focus();				
					return;
				}
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

			if (bAutomatico && (int)this.txtIdRetencion.Value == 0)
				this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, "Exec NumeracionLocales 10, 1, 1");

			this.DialogResult = DialogResult.OK;			
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
			if (!bEliminaAlValidar) e.Cancel = true;				
			else e.DisplayPromptMsg = false;
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
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
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
	}
}

