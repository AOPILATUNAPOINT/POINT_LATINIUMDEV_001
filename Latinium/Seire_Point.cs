using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Summary description for Seire_Point.
	/// </summary>
	public class Seire_Point : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsCompraS;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactura;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinGrid.UltraGrid grdSerie;
		private C1.Data.C1DataSet cdsCompra;
		private Infragistics.Win.Misc.UltraButton btSerie;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private DevExpress.XtraEditors.ComboBoxEdit cmbEstados;
		private Infragistics.Win.Misc.UltraButton btnImprimir;
		private Infragistics.Win.Misc.UltraButton btnGenerarSeriales;
		private Infragistics.Win.Misc.UltraButton btnActualizar;
		private Infragistics.Win.Misc.UltraButton btnProcesar;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodegaDestino;
		private Infragistics.Win.Misc.UltraButton btnLimpiaBodDest;
		private System.Windows.Forms.CheckBox chkProcesado;
		private Infragistics.Win.Misc.UltraButton btnGuiaDeRemision;
		private Infragistics.Win.Misc.UltraButton btnISG;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroCosto;

		private Acceso miAcceso;

		public Seire_Point()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Código");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LLegada");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hora");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Código");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("LLegada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			this.grdSerie = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btSerie = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.cdsCompraS = new C1.Data.C1DataSet();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.label14 = new System.Windows.Forms.Label();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbEstados = new DevExpress.XtraEditors.ComboBoxEdit();
			this.btnImprimir = new Infragistics.Win.Misc.UltraButton();
			this.btnGenerarSeriales = new Infragistics.Win.Misc.UltraButton();
			this.btnActualizar = new Infragistics.Win.Misc.UltraButton();
			this.btnProcesar = new Infragistics.Win.Misc.UltraButton();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbBodegaDestino = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnLimpiaBodDest = new Infragistics.Win.Misc.UltraButton();
			this.chkProcesado = new System.Windows.Forms.CheckBox();
			this.btnGuiaDeRemision = new Infragistics.Win.Misc.UltraButton();
			this.btnISG = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.grdSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraS)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).BeginInit();
			this.SuspendLayout();
			// 
			// grdSerie
			// 
			this.grdSerie.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdSerie.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdSerie.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdSerie.DisplayLayout.Appearance = appearance1;
			this.grdSerie.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 127;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 148;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 149;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 61;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 180;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 320;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Header.Caption = "Estado";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 100;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellMultiLine = Infragistics.Win.DefaultableBoolean.True;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 278;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand1.Header.Caption = "";
			ultraGridBand1.HeaderVisible = true;
			this.grdSerie.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdSerie.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.grdSerie.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdSerie.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.grdSerie.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdSerie.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.grdSerie.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerie.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerie.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerie.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.grdSerie.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdSerie.Location = new System.Drawing.Point(1, 328);
			this.grdSerie.Name = "grdSerie";
			this.grdSerie.Size = new System.Drawing.Size(960, 336);
			this.grdSerie.TabIndex = 9;
			this.grdSerie.DoubleClick += new System.EventHandler(this.grdSerie_DoubleClick);
			this.grdSerie.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.grdSerie_KeyPress);
			this.grdSerie.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdSerie_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// btSerie
			// 
			this.btSerie.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btSerie.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSerie.Enabled = false;
			this.btSerie.Location = new System.Drawing.Point(536, 72);
			this.btSerie.Name = "btSerie";
			this.btSerie.Size = new System.Drawing.Size(72, 22);
			this.btSerie.TabIndex = 98;
			this.btSerie.Text = "Se&riales";
			this.btSerie.Click += new System.EventHandler(this.btSerie_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(888, 72);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(72, 22);
			this.btCancelar.TabIndex = 97;
			this.btCancelar.Text = "&Salir";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.BackColor = System.Drawing.Color.Transparent;
			this.lblCliente.Location = new System.Drawing.Point(352, 11);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(64, 16);
			this.lblCliente.TabIndex = 142;
			this.lblCliente.Text = "Transacción";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(41, 16);
			this.label1.TabIndex = 144;
			this.label1.Text = "Bodega";
			// 
			// cdsCompraS
			// 
			this.cdsCompraS.BindingContextCtrl = this;
			this.cdsCompraS.DataLibrary = "LibFacturacion";
			this.cdsCompraS.DataLibraryUrl = "";
			this.cdsCompraS.DataSetDef = "dsCompraTabla";
			this.cdsCompraS.EnforceConstraints = false;
			this.cdsCompraS.FillOnRequest = false;
			this.cdsCompraS.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompraS.Name = "cdsCompraS";
			this.cdsCompraS.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraS.SchemaDef = null;
			this.cdsCompraS.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompraS_BeforeFill);
			// 
			// cmbTipoFactura
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoFactura.Appearance = appearance10;
			this.cmbTipoFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFactura.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompraS, "CompraNumero.idTipoFactura"));
			this.cmbTipoFactura.DisplayLayout.Override.NullText = "";
			this.cmbTipoFactura.DisplayMember = "Nombre";
			this.cmbTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFactura.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFactura.Location = new System.Drawing.Point(440, 8);
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(208, 21);
			this.cmbTipoFactura.TabIndex = 1;
			this.cmbTipoFactura.ValueMember = "IdTipoFactura";
			this.cmbTipoFactura.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTipoFactura_KeyPress);
			this.cmbTipoFactura.ValueChanged += new System.EventHandler(this.cmbTipoFactura_ValueChanged);
			this.cmbTipoFactura.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoFactura_InitializeLayout);
			// 
			// cmbCentroCosto
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCentroCosto.Appearance = appearance11;
			this.cmbCentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroCosto.DisplayMember = "Nombre";
			this.cmbCentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentroCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroCosto.Location = new System.Drawing.Point(80, 8);
			this.cmbCentroCosto.Name = "cmbCentroCosto";
			this.cmbCentroCosto.Size = new System.Drawing.Size(232, 21);
			this.cmbCentroCosto.TabIndex = 0;
			this.cmbCentroCosto.ValueMember = "Bodega";
			this.cmbCentroCosto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbCentroCosto_KeyPress);
			this.cmbCentroCosto.ValueChanged += new System.EventHandler(this.cmbCentroCosto_ValueChanged);
			this.cmbCentroCosto.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCentroCosto_InitializeLayout);
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
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.EnforceConstraints = false;
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(8, 75);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(37, 16);
			this.label14.TabIndex = 151;
			this.label14.Text = "Buscar";
			// 
			// txtArticulo
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance12;
			this.txtArticulo.Location = new System.Drawing.Point(80, 72);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(232, 22);
			this.txtArticulo.TabIndex = 7;
			this.txtArticulo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtArticulo_KeyPress);
			this.txtArticulo.ValueChanged += new System.EventHandler(this.txtArticulo_ValueChanged);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance13.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance13;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 120;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Número";
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 105;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Width = 79;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 262;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 4;
			ultraGridColumn13.Width = 341;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance14;
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn14.Width = 76;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance15;
			ultraGridColumn15.Format = "HH:mm";
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Width = 76;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 59;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			ultraGridBand2.Header.Caption = "";
			ultraGridBand2.HeaderVisible = true;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance17;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 10F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance18;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance19.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance19.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(1, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(960, 216);
			this.ultraGrid1.TabIndex = 153;
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			this.ultraGrid1.Click += new System.EventHandler(this.ultraGrid1_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn11.DataType = typeof(System.DateTime);
			ultraDataColumn11.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn14.DataType = typeof(System.DateTime);
			ultraDataColumn14.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn15.DataType = typeof(System.DateTime);
			ultraDataColumn15.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn16.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// dtDesde
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance22;
			this.dtDesde.AutoSelect = false;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(440, 40);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 5;
			this.dtDesde.Value = new System.DateTime(2013, 6, 18, 0, 0, 0, 0);
			this.dtDesde.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtDesde_KeyPress);
			this.dtDesde.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtDesde_BeforeDropDown);
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// dtHasta
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance23;
			this.dtHasta.AutoSelect = false;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(648, 40);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 6;
			this.dtHasta.Value = new System.DateTime(2013, 6, 18, 0, 0, 0, 0);
			this.dtHasta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtHasta_KeyPress);
			this.dtHasta.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtHasta_BeforeDropDown);
			this.dtHasta.ValueChanged += new System.EventHandler(this.dtHasta_ValueChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(352, 43);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(35, 16);
			this.label2.TabIndex = 156;
			this.label2.Text = "Desde";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(584, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(33, 16);
			this.label3.TabIndex = 157;
			this.label3.Text = "Hasta";
			// 
			// cmbEstados
			// 
			this.cmbEstados.Location = new System.Drawing.Point(752, 8);
			this.cmbEstados.Name = "cmbEstados";
			// 
			// cmbEstados.Properties
			// 
			this.cmbEstados.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbEstados.Size = new System.Drawing.Size(160, 20);
			this.cmbEstados.TabIndex = 2;
			this.cmbEstados.SelectedIndexChanged += new System.EventHandler(this.cmbEstados_SelectedIndexChanged);
			this.cmbEstados.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbEstados_KeyPress);
			this.cmbEstados.SelectedValueChanged += new System.EventHandler(this.cmbEstados_SelectedValueChanged);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Location = new System.Drawing.Point(808, 72);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(72, 22);
			this.btnImprimir.TabIndex = 165;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGenerarSeriales
			// 
			this.btnGenerarSeriales.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerarSeriales.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnGenerarSeriales.Enabled = false;
			this.btnGenerarSeriales.Location = new System.Drawing.Point(616, 72);
			this.btnGenerarSeriales.Name = "btnGenerarSeriales";
			this.btnGenerarSeriales.Size = new System.Drawing.Size(104, 22);
			this.btnGenerarSeriales.TabIndex = 166;
			this.btnGenerarSeriales.Text = "&Generar Seriales";
			this.btnGenerarSeriales.Click += new System.EventHandler(this.btnGenerarSeriales_Click);
			// 
			// btnActualizar
			// 
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnActualizar.Location = new System.Drawing.Point(344, 72);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(72, 22);
			this.btnActualizar.TabIndex = 167;
			this.btnActualizar.Text = "&Actualizar";
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// btnProcesar
			// 
			this.btnProcesar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnProcesar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnProcesar.Location = new System.Drawing.Point(728, 72);
			this.btnProcesar.Name = "btnProcesar";
			this.btnProcesar.Size = new System.Drawing.Size(72, 22);
			this.btnProcesar.TabIndex = 168;
			this.btnProcesar.Text = "&Procesar";
			this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(688, 11);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(38, 16);
			this.label4.TabIndex = 170;
			this.label4.Text = "Estado";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 43);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(42, 16);
			this.label5.TabIndex = 172;
			this.label5.Text = "Destino";
			// 
			// cmbBodegaDestino
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodegaDestino.Appearance = appearance24;
			this.cmbBodegaDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodegaDestino.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodegaDestino.DisplayMember = "Nombre";
			this.cmbBodegaDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodegaDestino.Enabled = false;
			this.cmbBodegaDestino.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaDestino.Location = new System.Drawing.Point(80, 40);
			this.cmbBodegaDestino.Name = "cmbBodegaDestino";
			this.cmbBodegaDestino.Size = new System.Drawing.Size(232, 21);
			this.cmbBodegaDestino.TabIndex = 3;
			this.cmbBodegaDestino.ValueMember = "Bodega";
			this.cmbBodegaDestino.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBodegaDestino_KeyPress);
			this.cmbBodegaDestino.ValueChanged += new System.EventHandler(this.cmbBodegaDestino_ValueChanged);
			this.cmbBodegaDestino.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodegaDestino_InitializeLayout);
			// 
			// btnLimpiaBodDest
			// 
			this.btnLimpiaBodDest.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnLimpiaBodDest.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnLimpiaBodDest.Enabled = false;
			this.btnLimpiaBodDest.Location = new System.Drawing.Point(312, 40);
			this.btnLimpiaBodDest.Name = "btnLimpiaBodDest";
			this.btnLimpiaBodDest.Size = new System.Drawing.Size(24, 21);
			this.btnLimpiaBodDest.TabIndex = 4;
			this.btnLimpiaBodDest.Text = "...";
			this.btnLimpiaBodDest.Click += new System.EventHandler(this.btnLimpiaBodDest_Click);
			// 
			// chkProcesado
			// 
			this.chkProcesado.Location = new System.Drawing.Point(792, 40);
			this.chkProcesado.Name = "chkProcesado";
			this.chkProcesado.Size = new System.Drawing.Size(80, 16);
			this.chkProcesado.TabIndex = 182;
			this.chkProcesado.Text = "En Espera";
			this.chkProcesado.Visible = false;
			// 
			// btnGuiaDeRemision
			// 
			this.btnGuiaDeRemision.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuiaDeRemision.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnGuiaDeRemision.Enabled = false;
			this.btnGuiaDeRemision.Location = new System.Drawing.Point(424, 72);
			this.btnGuiaDeRemision.Name = "btnGuiaDeRemision";
			this.btnGuiaDeRemision.Size = new System.Drawing.Size(104, 22);
			this.btnGuiaDeRemision.TabIndex = 183;
			this.btnGuiaDeRemision.Text = "&Guía de Remisión";
			this.btnGuiaDeRemision.Click += new System.EventHandler(this.btnGuiaDeRemision_Click);
			// 
			// btnISG
			// 
			this.btnISG.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnISG.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnISG.Enabled = false;
			this.btnISG.Location = new System.Drawing.Point(768, 40);
			this.btnISG.Name = "btnISG";
			this.btnISG.Size = new System.Drawing.Size(160, 22);
			this.btnISG.TabIndex = 184;
			this.btnISG.Text = "&Imprimir Seriales Generadas";
			this.btnISG.Click += new System.EventHandler(this.btnISG_Click);
			// 
			// Seire_Point
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(962, 668);
			this.Controls.Add(this.btnISG);
			this.Controls.Add(this.btnGuiaDeRemision);
			this.Controls.Add(this.chkProcesado);
			this.Controls.Add(this.btnLimpiaBodDest);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.cmbBodegaDestino);
			this.Controls.Add(this.btnProcesar);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.btnGenerarSeriales);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.cmbCentroCosto);
			this.Controls.Add(this.cmbTipoFactura);
			this.Controls.Add(this.btSerie);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.grdSerie);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "Seire_Point";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Entrega y recepción de mercadería";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Seire_Point_KeyDown);
			this.Load += new System.EventHandler(this.Seire_Point_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraS)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

	
		CultureInfo us = new CultureInfo("en-US");

		private void Seire_Point_Load(object sender, System.EventArgs e)
		{
      Funcion.Resolucion(this, this.Height, this.Width);

			miAcceso = new Acceso(cdsSeteoF, "0301");

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.Today; // DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
		
			string sSQL = string.Format("Exec ListaBodegas {0}, 48", MenuLatinium.IdUsuario);
			this.cmbCentroCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			
			cmbCentroCosto.DisplayLayout.Bands[0].Columns["Bodega"].Hidden = true;
			cmbCentroCosto.DisplayLayout.Bands[0].Columns["Codigo"].Hidden = true;
			cmbCentroCosto.DisplayLayout.Bands[0].Columns["Nombre"].Width = 232;
			cmbCentroCosto.DisplayMember = "Nombre";
			cmbCentroCosto.ValueMember = "Bodega";

			sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodega > 0)
			{
				this.cmbCentroCosto.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				this.cmbCentroCosto.Enabled = false;
			}

			this.cmbBodegaDestino.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Codigo, Nombre, Bodega From Bodega Where Transferencia = 1 Order By Bodega");
			this.cmbBodegaDestino.DisplayLayout.Bands[0].Columns["Bodega"].Hidden = true;
			this.cmbBodegaDestino.DisplayLayout.Bands[0].Columns["Codigo"].Hidden = true;
			this.cmbBodegaDestino.DisplayLayout.Bands[0].Columns["Nombre"].Width = 232;
			this.cmbBodegaDestino.DisplayMember = "Nombre";
			this.cmbBodegaDestino.ValueMember = "Bodega";
		}
		
		private void Informacion()
		{
			DateTime dDesde = (DateTime)this.dtDesde.Value;
			DateTime dHasta = (DateTime)this.dtHasta.Value;

			#region Asigna valor a Estado de facturas
			if (this.cmbEstados.Text.ToString() == "PENDIENTE" && ((int)this.cmbTipoFactura.Value != 11 || (int)this.cmbTipoFactura.Value != 13 || (int)this.cmbTipoFactura.Value != 44))  Estado = 1;
			if (this.cmbEstados.Text.ToString() == "PENDIENTE" && ((int)this.cmbTipoFactura.Value == 11 || (int)this.cmbTipoFactura.Value == 13 || (int)this.cmbTipoFactura.Value == 44))  Estado = 5;
			if (this.cmbEstados.Text.ToString() == "PENDIENTE" && ((int)this.cmbTipoFactura.Value == 5))  Estado = 5;
			if (this.cmbEstados.Text.ToString() == "PROCESADO" && ((int)this.cmbTipoFactura.Value == 5))  Estado = 3;
			if (this.cmbEstados.Text.ToString() == "PROCESADO") Estado = 3;
			if (this.cmbEstados.Text.ToString() == "TERMINADO") Estado = 4;
			if (this.cmbEstados.Text.ToString() == "LIQUIDADO") Estado = 7;
			if (this.cmbEstados.Text.ToString() == "FACTURADO") Estado = 9;
			if (this.cmbEstados.Text.ToString() == "FACTURADO SIN MERCADERIA") Estado = 12;
			if (this.cmbEstados.Text.ToString() == "ENTREGADO") Estado = 15;
			if (this.cmbEstados.Text.ToString() == "DEVUELTO") Estado = 16;			
			#endregion Asigna valor a Estado de facturas

			#region Cabecera
			int iBodegaDestino = 0;
			if (this.cmbBodegaDestino.ActiveRow != null) iBodegaDestino = (int)this.cmbBodegaDestino.Value;
			string stExec = string.Format ("Exec EntregaRecepcionMercaderia {0}, {1}, {2}, '{3}', '{4}', '{5}', {6}", (int)cmbTipoFactura.Value, (int)cmbCentroCosto.Value, Estado, this.txtArticulo.Text.ToString(), dDesde.ToString("yyyyMMdd"), dHasta.ToString("yyyyMMdd"), iBodegaDestino);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
			#endregion Cabecera
			
			this.ultraGrid1.DisplayLayout.Bands[0].Header.Caption = this.ultraGrid1.Rows.Count + " Registro(s) encontrado(s).";

			#region Encabezado Transferencias
			if ((int)this.cmbTipoFactura.Value == 8 || (int)this.cmbTipoFactura.Value == 9)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Código"].Header.Caption = "Origen";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Nombre"].Header.Caption = "Destino";
			}
			else 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Código"].Header.Caption = "Código";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Nombre"].Header.Caption = "Nombre";
			}
			#endregion Encabezado Transferencias

			#region Detalle
			string sSQL = string.Format("Exec DetalleEntregaRecepcionMercaderia 0, 0");
			this.grdSerie.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			#endregion Detalle
			
			if (this.ultraGrid1.Rows.Count > 0) this.btSerie.Enabled = true; else this.btSerie.Enabled = false;

			if ((int)this.cmbTipoFactura.Value == 7) 
			{
				this.btnGenerarSeriales.Enabled = true;
				this.btSerie.Enabled = false; 
				if (this.cmbEstados.Text.ToString().Trim() == "PENDIENTE")
				{
					this.btnGenerarSeriales.Text = "&Generar Seriales";
					this.btnISG.Enabled = false;
				}
				else
				{
					this.btnGenerarSeriales.Text = "&Ver Seriales";
					this.btnISG.Enabled = true;
				}
			}
			else if ((int)this.cmbTipoFactura.Value == 2)
			{
				this.btnISG.Enabled = true;
			}
			else 
			{
				this.btnGenerarSeriales.Text = "&Generar Seriales";
				this.btnGenerarSeriales.Enabled = false;
				this.btSerie.Enabled = true; 
				this.btnISG.Enabled = false;
			}

			if ((int)this.cmbTipoFactura.Value == 9 || (int)this.cmbTipoFactura.Value == 8)
			{
				if (this.cmbEstados.Text.ToString().Trim() == "PENDIENTE")
				{
					this.btnProcesar.Enabled = true;
					this.btnImprimir.Enabled = false;
					this.btnGuiaDeRemision.Enabled = false;
				}
				else if (this.cmbEstados.Text.ToString().Trim() == "PROCESADO") 
				{
					this.btnProcesar.Enabled = false;
					this.btnImprimir.Enabled = true;
					this.btnGuiaDeRemision.Enabled = true;
				}
			}
			else
			{
				this.btnProcesar.Enabled = false;
			}
		}

		private void cmbTipoFactura_ValueChanged(object sender, System.EventArgs e)
		{
			if (cmbCentroCosto.Value == null)
			{
				cmbCentroCosto.Value = 0;
			}
			
			if (this.cmbTipoFactura.Value	== null || this.cmbTipoFactura.Text.ToString().Length == 0) return;

			if ((int)this.cmbTipoFactura.Value == 9) this.cmbBodegaDestino.Value = System.DBNull.Value;

			#region Carga Combo Estados de facturas
			this.cmbEstados.Properties.Items.Clear();
			if ((int)this.cmbTipoFactura.Value == 2 || (int)this.cmbTipoFactura.Value == 11 || (int)this.cmbTipoFactura.Value == 13) /* ORDEN DE COMPRA - PEDIDO - RESERVA */
			{
				this.cmbEstados.Properties.Items.Add("PENDIENTE");
				this.cmbEstados.Properties.Items.Add("PROCESADO");
				this.cmbEstados.Properties.Items.Add("FACTURADO");
				
				this.cmbEstados.Text = "PENDIENTE";
			}			
			else if ((int)this.cmbTipoFactura.Value == 44) /* CONSIGNACIÓN */
			{
				this.cmbEstados.Properties.Items.Add("PENDIENTE");				
				this.cmbEstados.Properties.Items.Add("PROCESADO");
				this.cmbEstados.Properties.Items.Add("LIQUIDADO");

				this.cmbEstados.Text = "PENDIENTE";
			}
			else if ((int)this.cmbTipoFactura.Value == 51) /* NOTA DE ENTREGA */
			{
				this.cmbEstados.Properties.Items.Add("PENDIENTE");				
				this.cmbEstados.Properties.Items.Add("ENTREGADO");
				this.cmbEstados.Properties.Items.Add("DEVUELTO");

				this.cmbEstados.Text = "PENDIENTE";
			}
			else  if ((int)this.cmbTipoFactura.Value == 7) /* ORDÉN DE PRODUCCIÓN OBSEQUIOS */ 
			{
				this.cmbEstados.Properties.Items.Add("PENDIENTE");				
				this.cmbEstados.Properties.Items.Add("TERMINADO");

				this.cmbEstados.Text = "PENDIENTE";
			}
			else /* ORDÉN DE PRODUCCIÓN COMBOS */ 
			{
				this.cmbEstados.Properties.Items.Add("PENDIENTE");				
				this.cmbEstados.Properties.Items.Add("PROCESADO");

				this.cmbEstados.Text = "PENDIENTE";
			}
			#endregion Carga Combo Estados de facturas

			Informacion();

//			if ((int)this.cmbTipoFactura.Value == 9 || (int)this.cmbTipoFactura.Value == 8) this.btnImprimir.Enabled = true;
//			else this.btnImprimir.Enabled = false;
				
			if ((int)this.cmbTipoFactura.Value == 8)
			{
				this.cmbBodegaDestino.Enabled = true; 
				this.btnLimpiaBodDest.Enabled = true;
			}
			else
			{				
				this.cmbBodegaDestino.Enabled = false;
				this.btnLimpiaBodDest.Enabled = false;
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void cdsCompraS_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompraS.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		int Estado = 1;

		public void cmbCentroCosto_ValueChanged(object sender, System.EventArgs e)
		{
			int IBodega = 0;
			this.cmbTipoFactura.Value = 0;

			if (this.cmbCentroCosto.Text.ToString().Length > 0) IBodega = (int)this.cmbCentroCosto.Value;
			this.dtDesde.CalendarInfo.MaxDate = (DateTime)this.dtHasta.Value;
			this.dtHasta.CalendarInfo.MinDate = (DateTime)this.dtDesde.Value;

			cmbTipoFactura.DataSource = Funcion.dvProcedimiento(cdsCompraS, string.Format("Exec ListaTiposFacturaSerial {0}", IBodega));
			cmbTipoFactura.DisplayLayout.Bands[0].Columns["idTipoFactura"].Hidden = true;
			cmbTipoFactura.DisplayLayout.Bands[0].Columns["Nombre"].Width = 208;
			cmbTipoFactura.DisplayMember = "Nombre";
			cmbTipoFactura.ValueMember  = "idTipoFactura";
		}

		private void chkProcesado_CheckedChanged(object sender, System.EventArgs e)
		{
			if ((string) this.cmbTipoFactura.Text == "")
			{
				MessageBox.Show("Seleccione un tipo de factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				//this.chkProcesado.CheckedValue = false;	
				return;
			}

			if ((string) this.cmbCentroCosto.Text == "")
			{
				MessageBox.Show("Seleccione un Centro de costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				//this.chkProcesado.CheckedValue = false;
				return;
			}			
			
			DateTime dDesde = (DateTime)this.dtDesde.Value;
			DateTime dHasta = (DateTime)this.dtHasta.Value;

			if ((int)cmbTipoFactura.Value == 13)
			{
				if (this.chkProcesado.Checked) Estado = 3;
				else Estado = 5;
			}
			else
			{
				if (this.chkProcesado.Checked) Estado = 3;
				else Estado = 1;
			}

			string stExec = string.Format ("Exec EntregaRecepcionMercaderia {0}, {1}, {2}, '{3}', '{4}', '{5}'", (int)cmbTipoFactura.Value, (int)cmbCentroCosto.Value, Estado, this.txtArticulo.Text.ToString(), dDesde.ToString("yyyyMMdd"), dHasta.ToString("yyyyMMdd"));
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);

			this.ultraGrid1.DisplayLayout.Bands[0].Header.Caption = this.ultraGrid1.Rows.Count + " Registro(s) encontrado(s).";

			#region Detalle
			string sSQL = string.Format("Exec DetalleEntregaRecepcionMercaderia 0, 0");
			this.grdSerie.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			#endregion Detalle

			if (this.ultraGrid1.Rows.Count > 0) this.btSerie.Enabled = true;
			else this.btSerie.Enabled = false;
		}

		private void Mensaje()
		{
			int iContPendiente = 0;
			int iContProcesado = 0;

      this.grdSerie.DisplayLayout.Bands[0].Header.Caption = this.grdSerie.Rows.Count + " ITEMS";

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerie.Rows.All)
			{
				if (dr.Cells["EstadoSerial"].Value.ToString() == "PROCESADO")
				{
					dr.Appearance.BackColor = Color.LightBlue;
					dr.Appearance.BackColor2 = Color.LightBlue;				
				}

				if (dr.Cells["EstadoSerial"].Value.ToString() == "PENDIENTE") iContPendiente++;
				
				if (dr.Cells["EstadoSerial"].Value.ToString() == "PROCESADO") iContProcesado++;
			}

			this.grdSerie.DisplayLayout.Bands[0].Header.Caption = this.grdSerie.DisplayLayout.Bands[0].Header.Caption + " - " + iContPendiente.ToString() + " PENDIENTES " + iContProcesado + " PROCESADOS";
		}

		private void btSerie_Click(object sender, System.EventArgs e)
		{
			try
			{
				#region Validacion
				if (this.ultraGrid1.Rows.Count == 0)
				{
					MessageBox.Show("No existen transacciones con estos parámetros de búsqueda.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if (grdSerie.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if (grdSerie.ActiveRow.Cells["Articulo"].Value == DBNull.Value)
				{
					MessageBox.Show("Seleccione una fila.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if ((int)grdSerie.ActiveRow.Cells["Cantidad"].Value == 0)
				{
					MessageBox.Show("Cantidad en 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if (!miAcceso.BVerSerialesProcesadas && (this.cmbEstados.Text == "PROCESADO" || this.cmbEstados.Text == "FACTURADO"))
				{
					MessageBox.Show("La " + this.cmbTipoFactura.Text.ToString() + " ya esta Procesada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
						
				string sSQLValRB = string.Format("Select COUNT(*) From RegistroBloqueaERMercaderia Where Bloqueo = 1 And idDetCompra = {0}", (int) grdSerie.ActiveRow.Cells["idDetCompra"].Value);
				if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValRB, true) > 0)
				{
					string sSQLCB = string.Format("Select Top 1 Usuario From RegistroBloqueaERMercaderia Where Bloqueo = 1 And idDetCompra = {0} Order By Fecha Desc", (int) grdSerie.ActiveRow.Cells["idDetCompra"].Value);
					string sUsuario = Funcion.sEscalarSQL(cdsSeteoF, sSQLCB);
												
					MessageBox.Show(string.Format("No puede ingresar a Registro de Seriales de este Articulo Hasta que el usuario '{0}' cierre esta Pantalla", sUsuario), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				
					return;				
				}
				#endregion Validacion

				#region Variables
				int IdDetCompra = (int) grdSerie.ActiveRow.Cells["idDetCompra"].Value;
				int IdCompra = (int) grdSerie.ActiveRow.Cells["idCompra"].Value;
				int IdTipoFactura = (int)this.cmbTipoFactura.Value;// grdSerie.ActiveRow.Cells["idTipoFactura"].Value;
				int Cantidad = (int) grdSerie.ActiveRow.Cells["Cantidad"].Value;
				string Codigo = (string) grdSerie.ActiveRow.Cells["Codigo"].Value;
				string Articulo = (string) grdSerie.ActiveRow.Cells["Articulo"].Value;			
				int IdArt = (int) grdSerie.ActiveRow.Cells["idArticulo"].Value;
				#endregion Variables

				if (IdTipoFactura == 7) IdTipoFactura = 8;

				#region Valida Pagos en orden compra de Importacion Mixta
				//			if (IdTipoFactura == 2)
				//			{
				//				int idGrupoCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idGrupoCliente From Cliente cl Inner Join Compra c On c.idCliente = cl.idCliente Where c.idCompra = {0}", IdCompra));
				//				if (idGrupoCliente == 3)
				//				{
				//					double dTotal = Funcion.dEscalarSQL(cdsSeteoF, string.Format("Select Total From Compra Where idCompra = {0}", IdCompra));
				//					double dSaldo = Funcion.dEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(SUM(Pago), 0) From Pago Where idCompra = {0}", IdCompra));
				//
				//					if (dSaldo < dTotal)
				//					{
				//						MessageBox.Show("El documento no puede ser procesado tiene saldo pendiente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//						return;
				//					}
				//				}
				//			}
				#endregion Valida Pagos en orden compra de Importacion Mixta

				if (IdArt < 1)
				{
					MessageBox.Show("Seleccione el Articulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
				}
				else
				{
					#region Variables
					int dCantidad = (int) grdSerie.ActiveRow.Cells["Cantidad"].Value;
					string sNum = "";
					int iCantidad = int.Parse(dCantidad.ToString("0", us));
					int iDisponible = 0;
					int iBodega = (int) this.cmbCentroCosto.Value;

					bool bUsado = true;
					bool ManejaSerial = false;
					ManejaSerial = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", IdArt), true);
					#endregion Variables

					#region Registro de Mercaderia
					if (ManejaSerial)
					{
						string sSQLB = string.Format("Exec BloqueaERMercaderia 1, {0}", IdDetCompra);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLB, true);
	
						#region Registro Productos con Seriales
						using (frmIngresoSeriales Serial = new frmIngresoSeriales(IdCompra, IdTipoFactura, sNum, IdDetCompra, Cantidad, IdArt, Codigo, Articulo, iBodega, Estado))
						{							
							if (DialogResult.OK == Serial.ShowDialog())
							{
								#region Seriales de importación mixta
								if (IdTipoFactura == 2 && (int)this.ultraGrid1.ActiveRow.Cells["Estado"].Value == 12)
								{
									string sSQLGenSer =  string.Format("Exec SerialesImpMixOCCompra {0}", IdCompra);
									Funcion.EjecutaSQL(cdsSeteoF, sSQLGenSer, true);
								}
								#endregion Seriales de importación mixta

								this.cmbTipoFactura_ValueChanged(sender, e);

								string sSQLDB = string.Format("Exec BloqueaERMercaderia 0, {0}", IdDetCompra);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLDB, true);

								string sSQLEstado = string.Format("Select Estado From Compra Where idCompra = {0}", IdCompra);
								int Est = Funcion.iEscalarSQL(cdsSeteoF, sSQLEstado);
								if (Est == 3 || Est == 13) IdCompra = 0;

								string sSQL = string.Format("Exec DetalleEntregaRecepcionMercaderia {0}, {1}", IdCompra, IdTipoFactura);
								this.grdSerie.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

								Mensaje();
							}
							else
							{
								string sSQLDB = string.Format("Exec BloqueaERMercaderia 0, {0}", IdDetCompra);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLDB, true);
							}
						}
						#endregion Registro Productos con Seriales
					}
					else
					{
						#region Registro Productos sin Seriales
						using (ArticulosSinSerial Serial = new ArticulosSinSerial(Cantidad, IdArt, Codigo, Articulo, IdCompra, sNum, iDisponible, IdTipoFactura, IdDetCompra, bUsado, iBodega, (bool)this.chkProcesado.Checked, this.grdSerie.ActiveRow.Cells["Notas"].Value.ToString()))
						{												
							if (DialogResult.OK == Serial.ShowDialog())
							{
								this.cmbTipoFactura_ValueChanged(sender, e);
							
								string sSQLEstado = string.Format("Select Estado From Compra Where idCompra = {0}", IdCompra);
								int Est = Funcion.iEscalarSQL(cdsSeteoF, sSQLEstado);
								if (Est == 3) IdCompra = 0;

								string sSQL = string.Format("Exec DetalleEntregaRecepcionMercaderia {0}, {1}", IdCompra, IdTipoFactura);
								this.grdSerie.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

								Mensaje();
							}
						}
						#endregion Registro Productos sin Seriales
					}
					#endregion Registro de Mercaderia
				}
			}
			catch(Exception oExc)
			{
				MessageBox.Show(oExc.ToString() + "Error del Sistema Comuniquese Con Sistemas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
			}
		}	

		private void txtArticulo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if ((string) this.cmbTipoFactura.Text == "")
				{
					MessageBox.Show("Seleccione un tipo de factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if ((string) this.cmbCentroCosto.Text == "")
				{
					MessageBox.Show("Seleccione un Centro de costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
			
				Informacion();			
			}
		}

		private void grdSerie_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbCentroCosto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
		}

		private void txtArticulo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void Seire_Point_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();			
		}

		private void DetalleDocumento()
		{
			if (ultraGrid1.ActiveRow == null) return;
			
			this.grdSerie.DisplayLayout.Bands[0].Header.Caption = "";

			int idCompra = (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value;

			string sSQL = string.Format("Exec DetalleEntregaRecepcionMercaderia_V1 {0}, {1}", idCompra, (int)this.cmbTipoFactura.Value);
			this.grdSerie.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			Mensaje();
			/*
			this.grdSerie.DisplayLayout.Bands[0].Header.Caption = this.grdSerie.Rows.Count + " ITEMS";

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerie.Rows.All)
			{
				if (dr.Cells["EstadoSerial"].Value.ToString() == "PROCESADO")
				{
					dr.Appearance.BackColor = Color.LightBlue;
					dr.Appearance.BackColor2 = Color.LightBlue;				
				}

				if (dr.Cells["EstadoSerial"].Value.ToString() == "PENDIENTE") iContPendiente++;
				
				if (dr.Cells["EstadoSerial"].Value.ToString() == "PROCESADO") iContProcesado++;
			}

			this.grdSerie.DisplayLayout.Bands[0].Header.Caption = this.grdSerie.DisplayLayout.Bands[0].Header.Caption + " - " + iContPendiente.ToString() + " PENDIENTES " + iContProcesado + " PROCESADOS";*/
		}

		private void ultraGrid1_Click(object sender, System.EventArgs e)
		{
			//DetalleDocumento();
		}

		private void cmbTipoFactura_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtDesde_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbTipoFactura.Text.ToString() == "") return;
			if (this.cmbCentroCosto.Text.ToString() == "") return;
						
			this.dtHasta.CalendarInfo.MinDate = (DateTime)this.dtDesde.Value;

			Informacion();
		}

		private void dtHasta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbTipoFactura.Text.ToString() == "") return;
			if (this.cmbCentroCosto.Text.ToString() == "") return;
						
			this.dtDesde.CalendarInfo.MaxDate = (DateTime)this.dtHasta.Value;

			Informacion();
		}

		private void dtHasta_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbEstados_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if ((string) this.cmbTipoFactura.Text == "")
			{
				MessageBox.Show("Seleccione un tipo de factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if ((string) this.cmbCentroCosto.Text == "")
			{
				MessageBox.Show("Seleccione un Centro de costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			Informacion();

//			if ((int)this.cmbTipoFactura.Value == 8 && this.cmbEstados.Text.ToString().Trim() == "PROCESADO") this.btnGuiaDeRemision.Enabled = true;
//			else this.btnGuiaDeRemision.Enabled = false;
		}

		private void grdSerie_DoubleClick(object sender, System.EventArgs e)
		{
			if ((int)this.cmbTipoFactura.Value == 7) return;

			this.btSerie_Click(sender, e);
		}

		private void cmbEstados_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			DetalleDocumento();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Trasnferencia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (this.ultraGrid1.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una Trasnferencia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.Cursor = Cursors.WaitCursor;

			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldCompra = new ParameterField ();
			ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
			paramFieldCompra.ParameterFieldName = "@idCompra";
			ValIdCompra.Value = int.Parse(this.ultraGrid1.ActiveRow.Cells["idCompra"].Value.ToString());
			paramFieldCompra.CurrentValues.Add (ValIdCompra);
			paramFields.Add (paramFieldCompra);

			string sReporte = "";
			string sTitulo = "";
			if ((int)this.cmbTipoFactura.Value == 9) 
			{
				sReporte = "ReporteTransferencias.rpt";
				sTitulo = "Resumen de Recepcion de Transferencia";
			}
			else if ((int)this.cmbTipoFactura.Value == 8)
			{
				sReporte = "RepTransfSalida.rpt";
				sTitulo = "Resumen de Envio de Transferencia";
			}

			Reporte miRep = new Reporte(sReporte, "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo(sTitulo);
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();

			this.Cursor = Cursors.Default;
		}

		private void btnGenerarSeriales_Click(object sender, System.EventArgs e)
		{
			if ((int)this.cmbTipoFactura.Value != 7)
			{
				MessageBox.Show("Se genera seriales solo para Ordenes de producción.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

//			if (Estado == 1)
//			{
//				MessageBox.Show("Se genera seriales solo para Ordenes de producción procesadas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				return;
//			}

			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No existe Ordenes de Producción, intente con otros parametros de busqueda.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (this.ultraGrid1.ActiveRow == null || this.ultraGrid1.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una Orden de Produccion.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (this.cmbEstados.Text.ToString().Trim() == "PENDIENTE") 
			{
				if (DialogResult.No == MessageBox.Show("¿Seguro de generar las seriales para esta Orden de Produccion?, \n\n Este Proceso es irreversible.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			
				string sSQL = string.Format("Exec CreaIBDeOP {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			
				MessageBox.Show("Proceso exitoso.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				
				frmSerialesOPObs Seriales = new frmSerialesOPObs((int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value, "");
        Seriales.ShowDialog();

				Informacion();
			}
			else
			{
				frmSerialesOPObs Seriales = new frmSerialesOPObs((int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value, "");
        Seriales.ShowDialog();
			}
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			if (this.cmbCentroCosto.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una transacción.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCentroCosto.Focus();
				return;
			}

			Informacion();
		}

		private void btnProcesar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay transferencias pendientes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtArticulo.Focus();
				return;
			}

			if (this.ultraGrid1.ActiveRow == null || this.ultraGrid1.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ultraGrid1.Focus();
				return;
			}

			if ((int)this.grdSerie.Rows.Count == 0)
			{
				MessageBox.Show("Seleccione una transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ultraGrid1.Focus();
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerie.Rows.All)
			{
				if ((int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value != (int)dr.Cells["idCompra"].Value)
				{
					MessageBox.Show("Seleccione una transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.Focus();
					return;
				}
			}

			if ((int)this.cmbTipoFactura.Value != 8 && (int)this.cmbTipoFactura.Value != 9)
			{
				MessageBox.Show("Solo para Transferencias.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipoFactura.Focus();
				return;
			}

			if (this.cmbEstados.Text.ToString().Trim() != "PENDIENTE")
			{
				MessageBox.Show("Solo puede procesar transferencias pendientes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbEstados.Focus();
				return;
			}

			#region Valida TF Entrada Cantidades
			if ((int)this.cmbTipoFactura.Value == 9)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerie.Rows.All)
				{
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
					{
						int iCantidadValidaEntrada = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleSeriales Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));
						int iCantidadValida = (int)dr.Cells["Cantidad"].Value;

						if (iCantidadValidaEntrada < iCantidadValida && !miAcceso.BProcesarSerialesFaltantes)
						{
							MessageBox.Show(string.Format("Ingrese todas las Seriales del Articulo : '{0}'", dr.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.grdSerie.ActiveRow = dr;
							this.grdSerie.ActiveRow.Selected = true;
							return;
						}

						if (iCantidadValidaEntrada == 0 && !miAcceso.BProcesarSerialesFaltantes)
						{
							MessageBox.Show(string.Format("No puede procesar en Cero. '{0}'", dr.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.grdSerie.ActiveRow = dr;
							this.grdSerie.ActiveRow.Selected = true;
							return;
						}
					}
					else
					{
						int iContS = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Declare @Res Int = 0 If Exists(Select idDetCompra From Serie Where idArticulo = {0} And idDetCompra = {1}) Select @Res = ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1} Select @Res", (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetCompra"].Value));
						
						int iCantidadValida = (int)dr.Cells["Cantidad"].Value;

						if (iContS < iCantidadValida && !miAcceso.BProcesarSerialesFaltantes)
						{
							MessageBox.Show(string.Format("No puede procesar una Cantidad menor a la Enviada '{0}'", dr.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.grdSerie.ActiveRow = dr;
							this.grdSerie.ActiveRow.Selected = true;
							return;
						}

						if (iContS == 0)
						{
							MessageBox.Show(string.Format("No puede procesar en Cero. '{0}'", dr.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.grdSerie.ActiveRow = dr;
							this.grdSerie.ActiveRow.Selected = true;
							return;
						}
					}
				}
			}
			#endregion Valida TF Entrada Cantidades

			#region Valida Que no se procese en cero la TF de Salida
			if ((int)this.cmbTipoFactura.Value == 8)
			{
				int iCantidadValidaSalida = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerie.Rows.All)
				{
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
					{
						iCantidadValidaSalida = iCantidadValidaSalida + Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleSeriales Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));						
					}
					else
					{
						iCantidadValidaSalida = iCantidadValidaSalida + Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1}", (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetCompra"].Value));						
					}
				}

				if (iCantidadValidaSalida == 0)
				{
					MessageBox.Show("No puede procesar toda la Transaccion en Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					return;
				}
			}
			#endregion Valida Que no se procese en cero la TF de Salida

			#region Valida Cantidades en TF Salida
			if ((int)this.cmbTipoFactura.Value == 8)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerie.Rows.All)
				{
					int iContValida = 0;
					int iCantidadValida = 0;

					if (dr.Cells["Notas"].Value.ToString().Length == 0)
					{
						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
						{
							iContValida	= Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleSeriales Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));
				
							iCantidadValida = (int)dr.Cells["Cantidad"].Value;

							if (iContValida < iCantidadValida) 
							{
								string sSQLValOP = string.Format("Declare @idOrigen Int = 0 If Exists(Select ISNULL(idCompraTransf, 0) From CompraTransf Where Bodega = 1 And idCompra = {0}) Select @idOrigen = ISNULL(idOrigen, 0) From CompraTransf Where Bodega = 1 And idCompra = {0} Select @idOrigen", (int)dr.Cells["idCompra"].Value);
								if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValOP) > 0)
								{
									MessageBox.Show("En una Orden de Produccion no puede hacer ajustes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									this.grdSerie.ActiveRow = dr;
									this.grdSerie.ActiveRow.Selected = true;
									return;
								}
								else
								{
									MessageBox.Show(string.Format("El articulo '{0}' tiene {1} seriales registradas de {2}, \n\npara procesar escriba una observación de 20 caracteres", dr.Cells["Codigo"].Value.ToString(), iContValida, iCantidadValida), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									this.grdSerie.ActiveRow = dr;
									this.grdSerie.ActiveRow.Selected = true;
									return;
								}								
							}
						}
						else
						{
							iContValida = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Declare @Res Int = 0 If Exists(Select idDetCompra From Serie Where idArticulo = {0} And idDetCompra = {1}) Select @Res = ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1} Select @Res", (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetCompra"].Value));

							iCantidadValida = (int)dr.Cells["Cantidad"].Value;

							if (iContValida < iCantidadValida) 
							{
								string sSQLValOP = string.Format("Declare @idOrigen Int = 0 If Exists(Select ISNULL(idCompraTransf, 0) From CompraTransf Where Bodega = 1 And idCompra = {0}) Select @idOrigen = ISNULL(idOrigen, 0) From CompraTransf Where Bodega = 1 And idCompra = {0} Select @idOrigen", (int)dr.Cells["idCompra"].Value);
								if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValOP) > 0)
								{
									MessageBox.Show("En una Orden de Produccion no puede hacer ajustes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									this.grdSerie.ActiveRow = dr;
									this.grdSerie.ActiveRow.Selected = true;
									return;
								}
								else
								{
									MessageBox.Show(string.Format("Del articulo '{0}' ha registrado {1} de {2}, \n para procesar escriba una observación", dr.Cells["Codigo"].Value.ToString(), iContValida, iCantidadValida), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									this.grdSerie.ActiveRow = dr;
									this.grdSerie.ActiveRow.Selected = true;
									return;
								}
							}
						}
					}
				}
			}
			#endregion Valida Cantidades en TF Salida

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From Compra Where idCompra = {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value)) == 3)
			{
				MessageBox.Show("La Transferencia ya fue Procesada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.btnActualizar_Click(sender, e);
				return;
			}
			#endregion Validacion
			
			if (DialogResult.No == MessageBox.Show("¿Esta seguro de procesar la transferencia seleccionada?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) return;

			int iCont = 0;
			double iCantidad = 0;
			string sMensaje = "";

			this.Cursor = Cursors.WaitCursor;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerie.Rows.All)
			{
				if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
				{
					iCont	= Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleSeriales Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));
				
					iCantidad = (int)dr.Cells["Cantidad"].Value;
				
					if (iCont < iCantidad) sMensaje = sMensaje + "\n\n" + dr.Cells["Codigo"].Value.ToString() + " tiene " + iCont.ToString() + " seriales registradas de " + iCantidad.ToString(); 
				}
				else
				{
					iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1}", (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetCompra"].Value));

					iCantidad = (int)dr.Cells["Cantidad"].Value;
				
					if (iCont < iCantidad) sMensaje = sMensaje + "\n\n" + dr.Cells["Codigo"].Value.ToString() + " tiene " + iCont.ToString() + " productos registrados de " + iCantidad.ToString(); 
				}
			}
			
			this.Cursor = Cursors.Default;
				
			if (sMensaje.Length > 0 && (int)this.cmbTipoFactura.Value == 8)
			{
				if (!miAcceso.BProcesarSerialesFaltantes) 
				{
					MessageBox.Show("Hay Articulos que no se han cargado todas las Seriales \n\n Ingrese todas las Seriales o Consulte al Administrador", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Stop);
					return;
				}

				if (DialogResult.No == MessageBox.Show("La transferencia seleccionada tiene articulos que no se han cargado todas las seriales. \n\n\n\n" + sMensaje + "\n\n\n\n ¿Esta seguro de procesarla?", 
					"Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) return;
			}
			else if (sMensaje.Length > 0 && (int)this.cmbTipoFactura.Value == 9)
			{
				if (!miAcceso.BProcesarSerialesFaltantes) 
				{
					MessageBox.Show("Hay Articulos que no se han cargado todas las Seriales\n\nIngrese todas las Seriales para procesar", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Stop);
					return;
				}
			}
			
			if ((int)this.cmbTipoFactura.Value == 8)
			{							
				bool bCreaGR = true;

				if ((int)this.cmbCentroCosto.Value == 1) bCreaGR = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit = 1 If (Select Bodega1 From CompraTransf Where idCompra = {0}) In (5, 12, 18, 20, 23, 36) Set @Res = 0 Select @Res", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value), true);	
				
				if (bCreaGR)
				{
					bool bAutomatico = Funcion.bEscalarSQL(cdsSeteoF, string.Format("SELECT AutGuiaRemision FROM Bodega Where Bodega = {0}", (int)this.cmbCentroCosto.Value), true);

					using (frmGuiaDeRemision GuiaDeRemision = new frmGuiaDeRemision((int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value, (int)this.cmbTipoFactura.Value, (int)this.cmbCentroCosto.Value, bAutomatico))
					{
						if (DialogResult.OK == GuiaDeRemision.ShowDialog())
						{	
							this.Cursor = Cursors.WaitCursor;

							DateTime dtInicio = (DateTime)GuiaDeRemision.dtIniTraslado.Value;
							DateTime dtFin = (DateTime)GuiaDeRemision.dtFinTraslado.Value;
											
							if (bAutomatico) if (GuiaDeRemision.bNuevo) GuiaDeRemision.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 30, {0}, 0", (int)this.cmbCentroCosto.Value));

							#region Guarda Guia de Remision y Procesa Transferencia
							using (frmNumeracionGuiaRemision Numero = new frmNumeracionGuiaRemision(GuiaDeRemision.txtNumero.Text))
							{
								#region Guarda Guia Remision
								string sSQL = string.Format("Exec GGuiaRemision {0}, {1}, '{2}', '{3}', {4}, '{5}', {6}, {7}, {8}, '{9}', '{10}', {11}, '', 0", 
									GuiaDeRemision.txtIdGuiaRemision.Value, GuiaDeRemision.txtIdCompra.Value, 
									dtInicio.ToString("yyyyMMdd"), dtFin.ToString("yyyyMMdd"), 
									(int)this.cmbCentroCosto.Value, 
									GuiaDeRemision.txtNumero.Text.ToString(), 
									(int)this.cmbTipoFactura.Value,
									(int)GuiaDeRemision.cmbTraslado.Value, 
									(int)GuiaDeRemision.cmbResponsable.Value, 
									GuiaDeRemision.txtPlaca.Text.ToString(), 
									GuiaDeRemision.txtNotas.Text.ToString().Trim(),
									GuiaDeRemision.iBodegaDestino);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL);
								#endregion Guarda Guia Remision
							
								#region Procesa Transferencia
								string sSQLProc = string.Format("Exec ProcesarTransferencias {0}, {1}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value, (int)this.cmbTipoFactura.Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLProc);
							
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerie.Rows.All)
								{
									if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
									{
										iCont	= Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleSeriales Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));
									}
									else
									{
										iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1}", (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetCompra"].Value));
									}

									string sSQLAjuste = string.Format("Exec ActualizaEnvioRecepcionTransferencia {0}, {1}, {2}, '{3}'", (int)dr.Cells["idDetCompra"].Value, (int)this.cmbTipoFactura.Value, iCont, dr.Cells["Notas"].Value.ToString());
									Funcion.EjecutaSQL(cdsSeteoF, sSQLAjuste);
								}
								#endregion Procesa Transferencia		

								this.btnImprimir_Click(sender, e);			

								GuiaDeRemision.btnImprimir_Click(sender, e);

								#region Documento Electronico
								/*
														on Variables
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
														string sCodDocClte = "04";			
														DateTime dtFechaIni = (DateTime)GuiaDeRemision.dtIniTraslado.Value;
														DateTime dtFechaFin = (DateTime)GuiaDeRemision.dtFinTraslado.Value;
														DateTime dtFechaEmision = (DateTime)GuiaDeRemision.dtFechaEmision.Value;
														string sNumeroDocOrigen = GuiaDeRemision.txtNumComprobante.Text.ToString();
														egion Variables
														if (GuiaDeRemision.txtRucTransportista.Text.ToString().Length == 10) sCodDocClte = "05";

														on Consulta Informacion Tributaria
														string sSQLDE = string.Format("Exec FEInformacionTributaria {0}, {1}, {2}, '{3}', '{4}', {5}", 1, 1, (int)GuiaDeRemision.cmbBodega.Value, dtFechaIni.ToString("ddMMyyyy"), GuiaDeRemision.txtNumero.Text.ToString(), 0);
														SqlDataReader drInfoTributaria = Funcion.rEscalarSQL(cdsSeteoF, sSQLDE, true);
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
														}
														drInfoTributaria.Close();
														egion Consulta Informacion Tributaria

														on XML
														on Encabezado
														string stArchivo = "GRPRUEBA.XML";
														XmlTextWriter XMLFactura = new XmlTextWriter(stArchivo, Encoding.GetEncoding("UTF-8"));
														XMLFactura.Formatting = Formatting.Indented;
														XMLFactura.WriteStartDocument(true);
														XMLFactura.WriteStartElement("guiaRemision");
														egion Encabezado
			
														on Informacion Tributaria
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
														XMLFactura.WriteElementString("secuencial", GuiaDeRemision.txtNumero.Text.ToString());
														XMLFactura.WriteElementString("dirMatriz", sDirMatriz);

														XMLFactura.WriteEndElement();
														egion Informacion Tributaria

														on Informacion Guia De Remision
														XMLFactura.WriteStartElement("infoGuiaRemision>");
			
														XMLFactura.WriteElementString("dirEstablecimiento", sDirEstablecimiento);
														XMLFactura.WriteElementString("dirPartida", sDirEstablecimiento);
														XMLFactura.WriteElementString("razonSocialTransportista", GuiaDeRemision.cmbResponsable.Text.ToString());
														XMLFactura.WriteElementString("tipoIdentificacionTransportista", sCodDocClte);
														XMLFactura.WriteElementString("rucTransportista", GuiaDeRemision.txtRucTransportista.Text.ToString());
														XMLFactura.WriteElementString("obligadoContabilidad", "SI");
														XMLFactura.WriteElementString("fechaIniTransporte", dtFechaIni.ToString("dd/MM/yyyy"));
														XMLFactura.WriteElementString("fechaFinTransporte", dtFechaFin.ToString("dd/MM/yyyy"));
														XMLFactura.WriteElementString("placa", GuiaDeRemision.txtPlaca.Text.ToString());

														XMLFactura.WriteEndElement();		
														egion Informacion Guia De Remision

														on Destinatario
														if ((int)this.cmbTipoFactura.Value == 1)sNumeroDocOrigen = sEstab + "-" + sPtoEmi + "-" + GuiaDeRemision.txtNumComprobante.Text.ToString();
														XMLFactura.WriteStartElement("destinatarios");

														XMLFactura.WriteStartElement("destinatario");

														XMLFactura.WriteElementString("identificacionDestinatario", GuiaDeRemision.txtRucDestinatario.Text.ToString());
														XMLFactura.WriteElementString("razonSocialDestinatario", GuiaDeRemision.txtRazonSocialDestinatario.Text.ToString());
														XMLFactura.WriteElementString("dirDestinatario", GuiaDeRemision.cmbResponsable.Text.ToString());
														XMLFactura.WriteElementString("motivoTraslado", GuiaDeRemision.cmbTraslado.Text.ToString());
														//XMLFactura.WriteElementString("docAduaneroUnico", this.txtRucTransportista.Text.ToString());
														XMLFactura.WriteElementString("codEstabDestino", "SI");
														XMLFactura.WriteElementString("ruta", GuiaDeRemision.txtDestino.Text.ToString());
			
														if ((int)this.cmbTipoFactura.Value == 1) XMLFactura.WriteElementString("codDocSustento", "01");
														if ((int)this.cmbTipoFactura.Value != 8) XMLFactura.WriteElementString("numDocSustento", sNumeroDocOrigen);			
														if ((int)this.cmbTipoFactura.Value == 1)	XMLFactura.WriteElementString("numAutDocSustento", GuiaDeRemision.txtAutorizacion.Text.ToString());
														if ((int)this.cmbTipoFactura.Value != 8)	XMLFactura.WriteElementString("fechaEmisionDocSustento", dtFechaEmision.ToString("dd/MM/yyyy"));

														on Detalles
			
														XMLFactura.WriteStartElement("detalles");
														foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerie.Rows.All)
														{
															double dCantidad = (double)dr.Cells["Cantidad"].Value;

															XMLFactura.WriteStartElement("Detalle");					
						
															XMLFactura.WriteElementString("codigoInterno", dr.Cells["Codigo"].Value.ToString());
															XMLFactura.WriteElementString("descripcion", dr.Cells["Articulo"].Value.ToString());
															XMLFactura.WriteElementString("cantidad", dCantidad.ToString("0.00", us));
											
															XMLFactura.WriteEndElement();
														}
														XMLFactura.WriteEndElement();			
														egion Detalles

														XMLFactura.WriteEndElement();
														XMLFactura.WriteEndElement();
														egion Destinatario

														XMLFactura.WriteEndElement(); 
														XMLFactura.WriteEndDocument();

														XMLFactura.Flush();
														XMLFactura.Close();
														egion XML
														*/
								#endregion Documento Electronico

								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec MailControlTransferencias {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value), true);

								DateTime dtFechaTR = (DateTime)this.ultraGrid1.ActiveRow.Cells["Fecha"].Value;
								string sSQLActConteo = string.Format("Exec ActualizaTransferidoConteo '{0}', {1}, {2}", dtFechaTR.ToString("yyyyMMdd"), (int)this.cmbCentroCosto.Value, (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLActConteo, true);
	
								MessageBox.Show("Transferencia procesada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

								this.grdSerie.DisplayLayout.Bands[0].Header.Caption = "";

								this.cmbTipoFactura_ValueChanged(sender, e);					

								this.Cursor = Cursors.Default;
							}
							#endregion Guarda Guia de Remision y Procesa Transferencia
						}
					}
				}
				else
				{
					this.Cursor = Cursors.WaitCursor;

					#region Procesa Transferencia
					string sSQLProc = string.Format("Exec ProcesarTransferencias {0}, {1}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value, (int)this.cmbTipoFactura.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLProc);
							
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerie.Rows.All)
					{
						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
						{
							iCont	= Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleSeriales Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));
						}
						else
						{
							iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1}", (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetCompra"].Value));
						}

						string sSQLAjuste = string.Format("Exec ActualizaEnvioRecepcionTransferencia {0}, {1}, {2}, '{3}'", (int)dr.Cells["idDetCompra"].Value, (int)this.cmbTipoFactura.Value, iCont, dr.Cells["Notas"].Value.ToString());
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAjuste);
					}
					#endregion Procesa Transferencia		

					this.btnImprimir_Click(sender, e);					

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec MailControlTransferencias {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value), true);

					DateTime dtFechaTR = (DateTime)this.ultraGrid1.ActiveRow.Cells["Fecha"].Value;
					string sSQLActConteo = string.Format("Exec ActualizaTransferidoConteo '{0}', {1}, {2}", dtFechaTR.ToString("yyyyMMdd"), (int)this.cmbCentroCosto.Value, (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLActConteo, true);
	
					MessageBox.Show("Transferencia procesada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.grdSerie.DisplayLayout.Bands[0].Header.Caption = "";

					this.cmbTipoFactura_ValueChanged(sender, e);					

					this.Cursor = Cursors.Default;
				}
			}
			else if ((int)this.cmbTipoFactura.Value == 9) 
			{
				string sSQLProc = string.Format("Exec ProcesarTransferencias {0}, {1}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value, (int)this.cmbTipoFactura.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLProc);

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerie.Rows.All)
				{
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
					{
						iCont	= Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleSeriales Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));
					}
					else
					{
						iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1}", (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetCompra"].Value));
					}

					string sSQLAjuste = string.Format("Exec ActualizaEnvioRecepcionTransferencia {0}, {1}, {2}, '{3}'", (int)dr.Cells["idDetCompra"].Value, (int)this.cmbTipoFactura.Value, iCont, dr.Cells["Notas"].Value.ToString());
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAjuste);
				}
			
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NovedadesTransferencias {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value), true);

				MessageBox.Show("Transferencia procesada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				int idCompraIMR = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ERMIBRecepcionIncompleta {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value));

				if (idCompraIMR > 0)
				{
					MessageBox.Show("Ajuste de ingreso de mercaderia no recibida generado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}

				this.btnImprimir_Click(sender, e);
			
				this.cmbTipoFactura_ValueChanged(sender, e);					

				this.grdSerie.DisplayLayout.Bands[0].Header.Caption = "";					
			}			
		}

		private void cmbBodegaDestino_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodegaDestino_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbCentroCosto.ActiveRow == null) return;

			Informacion();
		}

		private void btnLimpiaBodDest_Click(object sender, System.EventArgs e)
		{
			this.cmbBodegaDestino.Value = 0;
		}

		private void cmbCentroCosto_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbTipoFactura.Focus();
		}

		private void cmbTipoFactura_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbEstados.Focus();
		}

		private void cmbEstados_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (this.cmbBodegaDestino.Enabled) this.cmbBodegaDestino.Focus();
				else this.dtDesde.Focus();
			}
		}

		private void dtDesde_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtHasta.Focus();
		}

		private void cmbBodegaDestino_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtDesde.Focus();
		}

		private void dtHasta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtArticulo.Focus();
		}

		private void grdSerie_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) if ((int)this.cmbTipoFactura.Value == 7) return; else this.btSerie_Click(sender, e);
			Funcion.CaracteresEspeciales(sender, e);
		}

		private void btnGuiaDeRemision_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay transferencias procesadas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtArticulo.Focus();
				return;
			}

			if (this.ultraGrid1.ActiveRow == null || this.ultraGrid1.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ultraGrid1.Focus();
				return;
			}

			if ((int)this.cmbTipoFactura.Value != 8)
			{
				MessageBox.Show("Solo para Transferencias.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipoFactura.Focus();
				return;
			}

//			if (this.cmbEstados.Text.ToString().Trim() == "PENDIENTE")
//			{
//				MessageBox.Show("Solo para transferencias procesadas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.cmbEstados.Focus();
//				return;
//			}
			#endregion Validacion

			bool bAutomatico = Funcion.bEscalarSQL(cdsSeteoF, string.Format("SELECT AutGuiaRemision FROM Bodega Where Bodega = {0}", (int)this.cmbCentroCosto.Value), true);

			using (frmGuiaDeRemision GuiaDeRemision = new frmGuiaDeRemision((int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value, (int)this.cmbTipoFactura.Value, (int)this.cmbCentroCosto.Value, bAutomatico))
			{
				if (DialogResult.OK == GuiaDeRemision.ShowDialog())
				{	
					this.Cursor = Cursors.WaitCursor;

					DateTime dtInicio = (DateTime)GuiaDeRemision.dtIniTraslado.Value;
					DateTime dtFin = (DateTime)GuiaDeRemision.dtFinTraslado.Value;

					if ((int)this.cmbCentroCosto.Value == 95) this.cmbCentroCosto.Value = 156;
					
					if (bAutomatico) if (GuiaDeRemision.bNuevo) GuiaDeRemision.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 30, {0}, 0", (int)this.cmbCentroCosto.Value));

					string sSQL = string.Format("Exec GGuiaRemision {0}, {1}, '{2}', '{3}', {4}, '{5}', {6}, {7}, {8}, '{9}', '{10}', {11}", 
						GuiaDeRemision.txtIdGuiaRemision.Value, GuiaDeRemision.txtIdCompra.Value, 
						dtInicio.ToString("yyyyMMdd"), dtFin.ToString("yyyyMMdd"), 
						(int)this.cmbCentroCosto.Value, 
						GuiaDeRemision.txtNumero.Text.ToString(), 
						(int)this.cmbTipoFactura.Value,
						(int)GuiaDeRemision.cmbTraslado.Value, 
						(int)GuiaDeRemision.cmbResponsable.Value, 
						GuiaDeRemision.txtPlaca.Text.ToString(), 
						GuiaDeRemision.txtNotas.Text.ToString().Trim(),
						GuiaDeRemision.iBodegaDestino);

					Funcion.EjecutaSQL(cdsSeteoF, sSQL);

					#region Documento Electronico
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
					string sCodDocClte = "04";			
					DateTime dtFechaIni = (DateTime)GuiaDeRemision.dtIniTraslado.Value;
					DateTime dtFechaFin = (DateTime)GuiaDeRemision.dtFinTraslado.Value;
					DateTime dtFechaEmision = (DateTime)GuiaDeRemision.dtFechaEmision.Value;
					string sNumeroDocOrigen = GuiaDeRemision.txtNumComprobante.Text.ToString();
					#endregion Variables
					if (GuiaDeRemision.txtRucTransportista.Text.ToString().Length == 10) sCodDocClte = "05";

					#region Consulta Informacion Tributaria
					string sSQLDE = string.Format("Exec FEInformacionTributaria {0}, {1}, {2}, '{3}', '{4}', {5}", 1, 1, (int)GuiaDeRemision.cmbBodega.Value, dtFechaIni.ToString("ddMMyyyy"), GuiaDeRemision.txtNumero.Text.ToString(), 0);
					SqlDataReader drInfoTributaria = Funcion.rEscalarSQL(cdsSeteoF, sSQLDE, true);
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
					}
					drInfoTributaria.Close();
					#endregion Consulta Informacion Tributaria

					#region XML
					#region Encabezado
					string stArchivo = "GRPRUEBA.XML";
					XmlTextWriter XMLFactura = new XmlTextWriter(stArchivo, Encoding.GetEncoding("UTF-8"));
					XMLFactura.Formatting = Formatting.Indented;
					XMLFactura.WriteStartDocument(true);
					XMLFactura.WriteStartElement("guiaRemision");
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
					XMLFactura.WriteElementString("secuencial", GuiaDeRemision.txtNumero.Text.ToString());
					XMLFactura.WriteElementString("dirMatriz", sDirMatriz);

					XMLFactura.WriteEndElement();
					#endregion Informacion Tributaria

					#region Informacion Guia De Remision
					XMLFactura.WriteStartElement("infoGuiaRemision>");
			
					XMLFactura.WriteElementString("dirEstablecimiento", sDirEstablecimiento);
					XMLFactura.WriteElementString("dirPartida", sDirEstablecimiento);
					XMLFactura.WriteElementString("razonSocialTransportista", GuiaDeRemision.cmbResponsable.Text.ToString());
					XMLFactura.WriteElementString("tipoIdentificacionTransportista", sCodDocClte);
					XMLFactura.WriteElementString("rucTransportista", GuiaDeRemision.txtRucTransportista.Text.ToString());
					XMLFactura.WriteElementString("obligadoContabilidad", "SI");
					XMLFactura.WriteElementString("fechaIniTransporte", dtFechaIni.ToString("dd/MM/yyyy"));
					XMLFactura.WriteElementString("fechaFinTransporte", dtFechaFin.ToString("dd/MM/yyyy"));
					XMLFactura.WriteElementString("placa", GuiaDeRemision.txtPlaca.Text.ToString());

					XMLFactura.WriteEndElement();		
					#endregion Informacion Guia De Remision

					#region Destinatario
					if ((int)this.cmbTipoFactura.Value == 1)sNumeroDocOrigen = sEstab + "-" + sPtoEmi + "-" + GuiaDeRemision.txtNumComprobante.Text.ToString();
					XMLFactura.WriteStartElement("destinatarios");

					XMLFactura.WriteStartElement("destinatario");

					XMLFactura.WriteElementString("identificacionDestinatario", GuiaDeRemision.txtRucDestinatario.Text.ToString());
					XMLFactura.WriteElementString("razonSocialDestinatario", GuiaDeRemision.txtRazonSocialDestinatario.Text.ToString());
					XMLFactura.WriteElementString("dirDestinatario", GuiaDeRemision.cmbResponsable.Text.ToString());
					XMLFactura.WriteElementString("motivoTraslado", GuiaDeRemision.cmbTraslado.Text.ToString());
					//XMLFactura.WriteElementString("docAduaneroUnico", this.txtRucTransportista.Text.ToString());
					XMLFactura.WriteElementString("codEstabDestino", "SI");
					XMLFactura.WriteElementString("ruta", GuiaDeRemision.txtDestino.Text.ToString());
			
					if ((int)this.cmbTipoFactura.Value == 1) XMLFactura.WriteElementString("codDocSustento", "01");
					if ((int)this.cmbTipoFactura.Value != 8) XMLFactura.WriteElementString("numDocSustento", sNumeroDocOrigen);			
					if ((int)this.cmbTipoFactura.Value == 1)	XMLFactura.WriteElementString("numAutDocSustento", GuiaDeRemision.txtAutorizacion.Text.ToString());
					if ((int)this.cmbTipoFactura.Value != 8)	XMLFactura.WriteElementString("fechaEmisionDocSustento", dtFechaEmision.ToString("dd/MM/yyyy"));

					#region Detalles
			
					XMLFactura.WriteStartElement("detalles");
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerie.Rows.All)
					{
						int dCantidad = (int)dr.Cells["Cantidad"].Value;

						XMLFactura.WriteStartElement("Detalle");					
						
						XMLFactura.WriteElementString("codigoInterno", dr.Cells["Codigo"].Value.ToString());
						XMLFactura.WriteElementString("descripcion", dr.Cells["Articulo"].Value.ToString());
						XMLFactura.WriteElementString("cantidad", dCantidad.ToString("0.00", us));
											
						XMLFactura.WriteEndElement();
					}
					XMLFactura.WriteEndElement();			
					#endregion Detalles

					XMLFactura.WriteEndElement();
					XMLFactura.WriteEndElement();
					#endregion Destinatario

					XMLFactura.WriteEndElement(); 
					XMLFactura.WriteEndDocument();

					XMLFactura.Flush();
					XMLFactura.Close();
					#endregion XML
					#endregion Documento Electronico

					GuiaDeRemision.btnImprimir_Click(sender, e);

					this.Cursor = Cursors.Default;
				}
			}
		}

		private void btnISG_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una Orden de Compra", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtArticulo.Focus();
				return;
			}

			int idCompra = (int)grdSerie.ActiveRow.Cells["idCompra"].Value;

			if ((int)this.cmbTipoFactura.Value == 2)
			{			
				if (this.grdSerie.ActiveRow.Cells["idDetCompra"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Producto de la Orden de Compra Seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtArticulo.Focus();
					return;
				}

				bool SerialPropia = false;
				SerialPropia = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select SeriePropiaAlterna From Articulo Where idArticulo = {0}", (int)this.grdSerie.ActiveRow.Cells["idArticulo"].Value), true);
				if (!SerialPropia)
				{
					MessageBox.Show("El Artículo Seleccionado No Genera Serial de POINT", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtArticulo.Focus();
					return;
				}

				idCompra = (int)grdSerie.ActiveRow.Cells["idCompra"].Value;
			}
			
			string sSQL = string.Format("Exec GuardaTempoSeriales {0}, {1}", idCompra, (int)this.cmbTipoFactura.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			string sDir = Funcion.sEscalarSQL(cdsSeteoF, "Select ADLabelJoy From SeteoF");

			Process.Start(sDir, "");
		}
	}
}

