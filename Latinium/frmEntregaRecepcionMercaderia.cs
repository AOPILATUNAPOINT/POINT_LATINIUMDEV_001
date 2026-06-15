using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using Latinium.Services;
using Latinium.Models;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmEntregaRecepcionMercaderia.
	/// </summary>
	public class frmEntregaRecepcionMercaderia : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodegaDestino;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactura;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.ContextMenu mnuImprimir;
		private System.Windows.Forms.MenuItem mnuImprimirTransferencias;
		private System.Windows.Forms.MenuItem mnuImprimirSerialesGeneradas;
		private System.Windows.Forms.Button btnGenerarSeriales;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.ComboBox cmbEstados;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridArticulos;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridTransacciones;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Button btnIngresoSeriales;
		private System.Windows.Forms.Label lblContadorTransacciones;
		private System.Windows.Forms.Button btnGuiaDeRemision;
		private System.Windows.Forms.Button btnProcesar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label label;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button btnRecepcionBodega;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmEntregaRecepcionMercaderia()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmEntregaRecepcionMercaderia));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoFacturacionDescripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadFacturada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPorFacturar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPendiente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRecibida");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoFacturacionDescripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadFacturada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadPorFacturar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadPendiente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadRecibida");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LLegada");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hora");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("LLegada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label5 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label = new System.Windows.Forms.Label();
			this.cmbBodegaDestino = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.mnuImprimir = new System.Windows.Forms.ContextMenu();
			this.mnuImprimirTransferencias = new System.Windows.Forms.MenuItem();
			this.mnuImprimirSerialesGeneradas = new System.Windows.Forms.MenuItem();
			this.btnGuiaDeRemision = new System.Windows.Forms.Button();
			this.btnGenerarSeriales = new System.Windows.Forms.Button();
			this.btnIngresoSeriales = new System.Windows.Forms.Button();
			this.btnProcesar = new System.Windows.Forms.Button();
			this.uGridArticulos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridTransacciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstados = new System.Windows.Forms.ComboBox();
			this.label6 = new System.Windows.Forms.Label();
			this.lblContadorTransacciones = new System.Windows.Forms.Label();
			this.button1 = new System.Windows.Forms.Button();
			this.btnRecepcionBodega = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridArticulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTransacciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
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
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(5, 58);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(43, 16);
			this.label5.TabIndex = 182;
			this.label5.Text = "Destino";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(181, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(34, 16);
			this.label3.TabIndex = 181;
			this.label3.Text = "Hasta";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(5, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 180;
			this.label2.Text = "Desde";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(5, 34);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 179;
			this.label1.Text = "Bodega";
			// 
			// label
			// 
			this.label.AutoSize = true;
			this.label.BackColor = System.Drawing.Color.Transparent;
			this.label.Location = new System.Drawing.Point(317, 34);
			this.label.Name = "label";
			this.label.Size = new System.Drawing.Size(62, 16);
			this.label.TabIndex = 178;
			this.label.Text = "Documento";
			// 
			// cmbBodegaDestino
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodegaDestino.Appearance = appearance1;
			this.cmbBodegaDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodegaDestino.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodegaDestino.DataSource = this.ultraDataSource4;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodegaDestino.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodegaDestino.DisplayMember = "Nombre";
			this.cmbBodegaDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodegaDestino.Enabled = false;
			this.cmbBodegaDestino.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaDestino.Location = new System.Drawing.Point(61, 56);
			this.cmbBodegaDestino.Name = "cmbBodegaDestino";
			this.cmbBodegaDestino.Size = new System.Drawing.Size(240, 21);
			this.cmbBodegaDestino.TabIndex = 175;
			this.cmbBodegaDestino.ValueMember = "Bodega";
			this.cmbBodegaDestino.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodegaDestino_KeyDown);
			this.cmbBodegaDestino.ValueChanged += new System.EventHandler(this.cmbBodegaDestino_ValueChanged);
			this.cmbBodegaDestino.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodegaDestino_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// dtHasta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance2;
			this.dtHasta.AutoSelect = false;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(237, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 177;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			this.dtHasta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtHasta_KeyPress);
			this.dtHasta.Validated += new System.EventHandler(this.dtHasta_Validated);
			this.dtHasta.ValueChanged += new System.EventHandler(this.dtHasta_ValueChanged);
			// 
			// dtDesde
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance3;
			this.dtDesde.AutoSelect = false;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(61, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 176;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
			this.dtDesde.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtDesde_KeyPress);
			this.dtDesde.Validated += new System.EventHandler(this.dtDesde_Validated);
			this.dtDesde.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtDesde_BeforeDropDown);
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// cmbBodega
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance4;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource4;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 60;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 250;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 50;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(61, 32);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(240, 21);
			this.cmbBodega.TabIndex = 173;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// cmbTipoFactura
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoFactura.Appearance = appearance5;
			this.cmbTipoFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFactura.DataSource = this.ultraDataSource3;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 250;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbTipoFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipoFactura.DisplayLayout.Override.NullText = "";
			this.cmbTipoFactura.DisplayMember = "Nombre";
			this.cmbTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFactura.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFactura.Location = new System.Drawing.Point(397, 32);
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(240, 21);
			this.cmbTipoFactura.TabIndex = 174;
			this.cmbTipoFactura.ValueMember = "idTipoFactura";
			this.cmbTipoFactura.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoFactura_KeyDown);
			this.cmbTipoFactura.ValueChanged += new System.EventHandler(this.cmbTipoFactura_ValueChanged);
			this.cmbTipoFactura.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoFactura_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(317, 58);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 16);
			this.label4.TabIndex = 666;
			this.label4.Text = "Buscar";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance6;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(397, 56);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(240, 21);
			this.txtBuscar.TabIndex = 665;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(645, 54);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(77, 24);
			this.btnVer.TabIndex = 667;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(608, 84);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 669;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// mnuImprimir
			// 
			this.mnuImprimir.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuImprimirTransferencias,
																																								this.mnuImprimirSerialesGeneradas});
			this.mnuImprimir.Popup += new System.EventHandler(this.mnuImprimir_Popup);
			// 
			// mnuImprimirTransferencias
			// 
			this.mnuImprimirTransferencias.Index = 0;
			this.mnuImprimirTransferencias.Text = "Transferencias";
			this.mnuImprimirTransferencias.Click += new System.EventHandler(this.mnuImprimirTransferencias_Click);
			// 
			// mnuImprimirSerialesGeneradas
			// 
			this.mnuImprimirSerialesGeneradas.Index = 1;
			this.mnuImprimirSerialesGeneradas.Text = "Seriales Generadas";
			this.mnuImprimirSerialesGeneradas.Click += new System.EventHandler(this.mnuImprimirSerialesGeneradas_Click);
			// 
			// btnGuiaDeRemision
			// 
			this.btnGuiaDeRemision.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuiaDeRemision.CausesValidation = false;
			this.btnGuiaDeRemision.Enabled = false;
			this.btnGuiaDeRemision.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuiaDeRemision.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuiaDeRemision.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuiaDeRemision.Location = new System.Drawing.Point(373, 84);
			this.btnGuiaDeRemision.Name = "btnGuiaDeRemision";
			this.btnGuiaDeRemision.Size = new System.Drawing.Size(104, 23);
			this.btnGuiaDeRemision.TabIndex = 670;
			this.btnGuiaDeRemision.Text = "Guia De Remisión";
			this.btnGuiaDeRemision.Click += new System.EventHandler(this.btnGuiaDeRemision_Click);
			// 
			// btnGenerarSeriales
			// 
			this.btnGenerarSeriales.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGenerarSeriales.CausesValidation = false;
			this.btnGenerarSeriales.Enabled = false;
			this.btnGenerarSeriales.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGenerarSeriales.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGenerarSeriales.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerarSeriales.Location = new System.Drawing.Point(125, 84);
			this.btnGenerarSeriales.Name = "btnGenerarSeriales";
			this.btnGenerarSeriales.Size = new System.Drawing.Size(104, 23);
			this.btnGenerarSeriales.TabIndex = 671;
			this.btnGenerarSeriales.Text = "Generar Seriales";
			this.btnGenerarSeriales.Click += new System.EventHandler(this.btnGenerarSeriales_Click);
			// 
			// btnIngresoSeriales
			// 
			this.btnIngresoSeriales.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnIngresoSeriales.CausesValidation = false;
			this.btnIngresoSeriales.Enabled = false;
			this.btnIngresoSeriales.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnIngresoSeriales.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnIngresoSeriales.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnIngresoSeriales.Location = new System.Drawing.Point(5, 84);
			this.btnIngresoSeriales.Name = "btnIngresoSeriales";
			this.btnIngresoSeriales.Size = new System.Drawing.Size(112, 23);
			this.btnIngresoSeriales.TabIndex = 672;
			this.btnIngresoSeriales.Text = "Ingreso de Seriales";
			this.btnIngresoSeriales.Click += new System.EventHandler(this.btnIngresoSeriales_Click);
			// 
			// btnProcesar
			// 
			this.btnProcesar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnProcesar.CausesValidation = false;
			this.btnProcesar.Enabled = false;
			this.btnProcesar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnProcesar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnProcesar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnProcesar.Location = new System.Drawing.Point(237, 84);
			this.btnProcesar.Name = "btnProcesar";
			this.btnProcesar.Size = new System.Drawing.Size(128, 23);
			this.btnProcesar.TabIndex = 673;
			this.btnProcesar.Text = "Procesar Transferencia";
			this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
			// 
			// uGridArticulos
			// 
			this.uGridArticulos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridArticulos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridArticulos.DataSource = this.ultraDataSource1;
			appearance7.BackColor = System.Drawing.Color.White;
			this.uGridArticulos.DisplayLayout.Appearance = appearance7;
			this.uGridArticulos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 97;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 113;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 149;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance8;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 35;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 4;
			ultraGridColumn13.Width = 107;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn14.Width = 129;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance9;
			ultraGridColumn15.Header.Caption = "Estado";
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Width = 71;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.CellMultiLine = Infragistics.Win.DefaultableBoolean.True;
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn16.Width = 136;
			ultraGridColumn17.Header.Caption = "Estado Fact.";
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Width = 105;
			ultraGridColumn18.Header.Caption = "Cnt. Facturada";
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridColumn18.Width = 69;
			ultraGridColumn19.Header.Caption = "Cnt. Por Facturar";
			ultraGridColumn19.Header.VisiblePosition = 9;
			ultraGridColumn19.Width = 74;
			ultraGridColumn20.Header.Caption = "Cnt. Pendiente";
			ultraGridColumn20.Header.VisiblePosition = 11;
			ultraGridColumn20.Width = 73;
			ultraGridColumn21.Header.Caption = "Cnt. Recibida";
			ultraGridColumn21.Header.VisiblePosition = 10;
			ultraGridColumn21.Width = 94;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn21});
			ultraGridBand4.Header.Caption = "";
			ultraGridBand4.HeaderVisible = true;
			this.uGridArticulos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridArticulos.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.uGridArticulos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridArticulos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridArticulos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridArticulos.DisplayLayout.Override.CardAreaAppearance = appearance11;
			this.uGridArticulos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8.75F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridArticulos.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.uGridArticulos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridArticulos.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridArticulos.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridArticulos.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.uGridArticulos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridArticulos.Location = new System.Drawing.Point(0, 288);
			this.uGridArticulos.Name = "uGridArticulos";
			this.uGridArticulos.Size = new System.Drawing.Size(914, 242);
			this.uGridArticulos.TabIndex = 674;
			this.uGridArticulos.DoubleClick += new System.EventHandler(this.uGridArticulos_DoubleClick);
			this.uGridArticulos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridArticulos_KeyPress);
			this.uGridArticulos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridArticulos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn18});
			// 
			// uGridTransacciones
			// 
			this.uGridTransacciones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridTransacciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTransacciones.DataSource = this.ultraDataSource2;
			appearance16.BackColor = System.Drawing.Color.White;
			this.uGridTransacciones.DisplayLayout.Appearance = appearance16;
			this.uGridTransacciones.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 105;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.Caption = "Número";
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 111;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Width = 73;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.VisiblePosition = 3;
			ultraGridColumn25.Width = 277;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.VisiblePosition = 4;
			ultraGridColumn26.Width = 277;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn27.CellAppearance = appearance17;
			ultraGridColumn27.Header.VisiblePosition = 5;
			ultraGridColumn27.Width = 73;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn28.CellAppearance = appearance18;
			ultraGridColumn28.Format = "HH:mm";
			ultraGridColumn28.Header.VisiblePosition = 6;
			ultraGridColumn28.Width = 82;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 7;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 59;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			ultraGridBand5.Header.Caption = "";
			this.uGridTransacciones.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTransacciones.DisplayLayout.Override.ActiveRowAppearance = appearance19;
			this.uGridTransacciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridTransacciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridTransacciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance20.BackColor = System.Drawing.Color.Transparent;
			this.uGridTransacciones.DisplayLayout.Override.CardAreaAppearance = appearance20;
			this.uGridTransacciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance21.FontData.BoldAsString = "True";
			appearance21.FontData.Name = "Arial";
			appearance21.FontData.SizeInPoints = 8.75F;
			appearance21.ForeColor = System.Drawing.Color.White;
			appearance21.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTransacciones.DisplayLayout.Override.HeaderAppearance = appearance21;
			this.uGridTransacciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance22.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance22.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransacciones.DisplayLayout.Override.RowAlternateAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransacciones.DisplayLayout.Override.RowSelectorAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransacciones.DisplayLayout.Override.SelectedRowAppearance = appearance24;
			this.uGridTransacciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTransacciones.Location = new System.Drawing.Point(0, 112);
			this.uGridTransacciones.Name = "uGridTransacciones";
			this.uGridTransacciones.Size = new System.Drawing.Size(914, 174);
			this.uGridTransacciones.TabIndex = 675;
			this.uGridTransacciones.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridTransacciones_BeforeSelectChange);
			this.uGridTransacciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridTransacciones_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn21.DataType = typeof(System.DateTime);
			ultraDataColumn21.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn24.DataType = typeof(System.DateTime);
			ultraDataColumn24.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn25.DataType = typeof(System.DateTime);
			ultraDataColumn25.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn26.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// cmbEstados
			// 
			this.cmbEstados.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbEstados.Location = new System.Drawing.Point(709, 32);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(200, 21);
			this.cmbEstados.TabIndex = 676;
			this.cmbEstados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstados_KeyDown);
			this.cmbEstados.SelectedIndexChanged += new System.EventHandler(this.cmbEstados_SelectedIndexChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(653, 34);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(39, 16);
			this.label6.TabIndex = 677;
			this.label6.Text = "Estado";
			// 
			// lblContadorTransacciones
			// 
			this.lblContadorTransacciones.AutoSize = true;
			this.lblContadorTransacciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
			this.lblContadorTransacciones.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContadorTransacciones.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContadorTransacciones.Location = new System.Drawing.Point(688, 88);
			this.lblContadorTransacciones.Name = "lblContadorTransacciones";
			this.lblContadorTransacciones.Size = new System.Drawing.Size(0, 17);
			this.lblContadorTransacciones.TabIndex = 678;
			this.lblContadorTransacciones.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button1
			// 
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.CausesValidation = false;
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(893, 76);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(7, 23);
			this.button1.TabIndex = 679;
			this.button1.Visible = false;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// btnRecepcionBodega
			// 
			this.btnRecepcionBodega.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnRecepcionBodega.CausesValidation = false;
			this.btnRecepcionBodega.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnRecepcionBodega.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnRecepcionBodega.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRecepcionBodega.Location = new System.Drawing.Point(488, 84);
			this.btnRecepcionBodega.Name = "btnRecepcionBodega";
			this.btnRecepcionBodega.Size = new System.Drawing.Size(104, 23);
			this.btnRecepcionBodega.TabIndex = 680;
			this.btnRecepcionBodega.Text = "Recepción Bodega";
			this.btnRecepcionBodega.Click += new System.EventHandler(this.btnRecepcionBodega_Click);
			// 
			// frmEntregaRecepcionMercaderia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(913, 530);
			this.Controls.Add(this.btnRecepcionBodega);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.lblContadorTransacciones);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.btnProcesar);
			this.Controls.Add(this.btnIngresoSeriales);
			this.Controls.Add(this.btnGenerarSeriales);
			this.Controls.Add(this.btnGuiaDeRemision);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.cmbBodegaDestino);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbTipoFactura);
			this.Controls.Add(this.uGridArticulos);
			this.Controls.Add(this.uGridTransacciones);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmEntregaRecepcionMercaderia";
			this.Text = "Entrega y Recepción Mercaderia";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmEntregaRecepcionMercaderia_KeyDown);
			this.Load += new System.EventHandler(this.frmEntregaRecepcionMercaderia_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridArticulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTransacciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		int Estado = 1;

		private void frmEntregaRecepcionMercaderia_Load(object sender, System.EventArgs e)
		{
			if(MenuLatinium.stUsuario == "AOPILATUÑA")
			{
			 this.button1.Visible = true;
			}

			Funcion.Resolucion(this, this.Height, this.Width);

			miAcceso = new Acceso(cdsSeteoF, "0301");

			this.dtDesde.Value = DateTime.Today.AddDays(-30);//"01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.Today;

			string sSQL = string.Format("Exec ListaBodegas {0}, 48", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodega > 0) this.cmbBodega.Value = iBodega; else this.cmbBodega.Enabled = true;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void mnuImprimirTransferencias_Click(object sender, System.EventArgs e)
		{
			ImprimeTF();
		}

		private void LimpiaControles()
		{
			this.cmbTipoFactura.Value = System.DBNull.Value;
			this.cmbBodegaDestino.Value = System.DBNull.Value;
			this.cmbEstados.Text = "";
			this.cmbEstados.Items.Clear();
			this.txtBuscar.Text = "";
			this.lblContadorTransacciones.Text = "";

			this.uGridTransacciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec EntregaRecepcionMercaderia 0, 0, 0, '', '20010101', '20010101', 0");
			this.uGridArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DetalleEntregaRecepcionMercaderia 0");			
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			this.LimpiaControles();

			int IBodega = 0;
						
			if (this.cmbBodega.ActiveRow != null) IBodega = (int)this.cmbBodega.Value;

			this.cmbTipoFactura.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaTiposFacturaSerial {0}", IBodega));

			if (this.cmbBodega.ActiveRow != null && this.dtHasta.Value != System.DBNull.Value)
			{
				bool bConsignacion = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Consignacion From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);

				int iBodegaOrigen = (int)this.cmbBodega.Value;
				if (bConsignacion)
					iBodegaOrigen = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select Local from AsignaBodegasLocales Where Bodega = {0} And Consignacion = 1", (int)this.cmbBodega.Value));

				this.cmbBodegaDestino.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaDeBodegasActivas '{0}', {1}, {2}", 
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value, bConsignacion));
			}
		}

		private void ConfigurarBotonPorFacturar()
		{
			int iTipoFactura = (int)this.cmbTipoFactura.Value;

			//---------------------------------------------------------
			// SOLO ORDEN DE COMPRA (ID 2) PUEDE USAR RECEPCIÓN BODEGA
			//---------------------------------------------------------
			if (iTipoFactura == 2)
			{
				this.btnRecepcionBodega.Enabled = true;
				return;
			}

			//-----------------------------------------------------
			// CUALQUIER OTRO TIPO SE BLOQUEA
			//-----------------------------------------------------
			this.btnRecepcionBodega.Enabled = false;
		}

		private void cmbTipoFactura_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null) return;
			
			if (this.cmbTipoFactura.ActiveRow	== null) return;		 

			this.cmbBodegaDestino.Value = System.DBNull.Value;
			this.cmbBodegaDestino.Enabled = false;

			

			#region Carga Combo Estados de facturas											
			this.cmbEstados.Items.Clear();								
											
			if ((int)this.cmbTipoFactura.Value != 2 )								
			{								
				this.cmbEstados.Items.Add("PENDIENTE");							
			}								
			if ((int)this.cmbTipoFactura.Value == 2 )								
			{								
				//				this.cmbEstados.Items.Add("PROCESADO");		
				this.cmbEstados.Items.Add("RECEPCIÓN BODEGA");	
				this.cmbEstados.Items.Add("FACTURADO");							
				this.cmbEstados.Items.Add("APROBADO");							
			}								
											
			if ( (int)this.cmbTipoFactura.Value == 11 || (int)this.cmbTipoFactura.Value == 13) /* ORDEN DE COMPRA - PEDIDO - RESERVA */								
			{								
				this.cmbEstados.Items.Add("PROCESADO");							
				this.cmbEstados.Items.Add("FACTURADO");							
											
				this.cmbEstados.Text = "PENDIENTE";							
			}								
			else if ((int)this.cmbTipoFactura.Value == 44) /* CONSIGNACIÓN */								
			{								
				this.cmbEstados.Items.Add("PROCESADO");							
				this.cmbEstados.Items.Add("LIQUIDADO");							
											
				this.cmbEstados.Text = "PENDIENTE";							
			}								
			else if ((int)this.cmbTipoFactura.Value == 51) /* NOTA DE ENTREGA */								
			{								
				this.cmbEstados.Items.Add("ENTREGADO");							
				this.cmbEstados.Items.Add("DEVUELTO");							
											
				this.btnImprimir.Enabled = true;							
				this.cmbEstados.Text = "PENDIENTE";							
			}								
			else if ((int)this.cmbTipoFactura.Value == 7) /* ORDÉN DE PRODUCCIÓN OBSEQUIOS */ 								
			{								
				this.cmbEstados.Items.Add("TERMINADO");							
											
				this.cmbEstados.Text = "PENDIENTE";							
			}								
			else /* ORDÉN DE PRODUCCIÓN COMBOS */ 								
			{								
				this.cmbEstados.Items.Add("PROCESADO");							
											
				this.cmbEstados.Text = "PENDIENTE";							
			}								
			#endregion Carga Combo Estados de facturas								

			
			if ((int)this.cmbTipoFactura.Value == 8) 
				this.cmbBodegaDestino.Enabled = true;

			//-----------------------------------------------------
			// ConfigurarBotonPorFacturar
			//-----------------------------------------------------
			this.ConfigurarBotonPorFacturar();
		}

		private void cmbEstados_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null) return;
			if (this.cmbTipoFactura.ActiveRow == null) return;
			
			this.Informacion();
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value == System.DBNull.Value) return;

			this.dtHasta.CalendarInfo.MinDate = (DateTime)this.dtDesde.Value;

			if (this.cmbEstados.Text.ToString().Length == 0) return;
			if (this.cmbTipoFactura.ActiveRow == null) return;
			if (this.cmbBodega.ActiveRow == null) return;

			this.Informacion();
		}

		private void dtHasta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtHasta.Value == System.DBNull.Value) return;

			this.dtDesde.CalendarInfo.MaxDate = (DateTime)this.dtHasta.Value;

			if (this.cmbEstados.Text.ToString().Length == 0) return;
			if (this.cmbTipoFactura.ActiveRow == null) return;
			if (this.cmbBodega.ActiveRow == null) return;					

			this.Informacion();
		}

		private void cmbBodegaDestino_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null) return;
			if (this.cmbTipoFactura.ActiveRow == null) return;
			if (this.cmbEstados.Text.ToString().Length == 0) return;

			this.Informacion();
		}

		private void Informacion()
		{
			#region Asigna valor a Estado de facturas

			int Estado = 0;

			if (this.cmbEstados.Text.ToString() == "PENDIENTE" &&
				((int)this.cmbTipoFactura.Value != 11 ||
				(int)this.cmbTipoFactura.Value != 13 ||
				(int)this.cmbTipoFactura.Value != 44))
				Estado = 1;

			if (this.cmbEstados.Text.ToString() == "PENDIENTE" &&
				((int)this.cmbTipoFactura.Value == 11 ||
				(int)this.cmbTipoFactura.Value == 13 ||
				(int)this.cmbTipoFactura.Value == 44))
				Estado = 5;

			if (this.cmbEstados.Text.ToString() == "PENDIENTE" &&
				(int)this.cmbTipoFactura.Value == 5)
				Estado = 5;

			if (this.cmbEstados.Text.ToString() == "PROCESADO")
				Estado = 3;

			if (this.cmbEstados.Text.ToString() == "RECEPCIÓN BODEGA")
				Estado = 28;

			if (this.cmbEstados.Text.ToString() == "TERMINADO")
				Estado = 4;

			if (this.cmbEstados.Text.ToString() == "LIQUIDADO")
				Estado = 7;

			if (this.cmbEstados.Text.ToString() == "FACTURADO")
				Estado = 9;

			if (this.cmbEstados.Text.ToString() == "APROBADO" &&
				(int)this.cmbTipoFactura.Value == 2)
				Estado = 5;

			if (this.cmbEstados.Text.ToString() == "FACTURADO SIN MERCADERIA")
				Estado = 12;

			if (this.cmbEstados.Text.ToString() == "ENTREGADO")
				Estado = 15;

			if (this.cmbEstados.Text.ToString() == "DEVUELTO")
				Estado = 16;

			#endregion


			#region Cabecera
			int iBodegaDestino = 0;
			if (this.cmbBodegaDestino.ActiveRow != null) iBodegaDestino = (int)this.cmbBodegaDestino.Value;

			string stExec = string.Format ("Exec EntregaRecepcionMercaderia {0}, {1}, {2}, '{3}', '{4}', '{5}', {6}", 
				(int)cmbTipoFactura.Value, (int)cmbBodega.Value, Estado, this.txtBuscar.Text.ToString(), 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), iBodegaDestino);
			FuncionesProcedimientos.dsGeneral(stExec, this.uGridTransacciones);			

			#region Mostrar ocultar Columnas facturacion

			int tipoFactura = 0;

			if (this.cmbTipoFactura.Value != null &&
				this.cmbTipoFactura.Value != DBNull.Value)
			{
				tipoFactura = Convert.ToInt32(this.cmbTipoFactura.Value);
			}

			this.uGridTransacciones.BeginUpdate();

			if (tipoFactura == 2)
			{
				ConfigurarVistaFacturacion();
			}
			else
			{
				ConfigurarVistaNormal();
			}

			// cierre de edición (IMPORTANTE en Infragistics viejo)
			this.uGridTransacciones.PerformAction(UltraGridAction.ExitEditMode);

			// repaint real en .NET 1.1
			this.uGridTransacciones.Refresh();

			Application.DoEvents();

			this.uGridTransacciones.EndUpdate();

			#endregion

			#endregion Cabecera
			
			this.lblContadorTransacciones.Text = this.uGridTransacciones.Rows.Count + " REGISTRO(S) ENCONTRADOS(S)";

			#region Encabezado Columnas Codigo y Nombre 
			if ((int)this.cmbTipoFactura.Value == 8 || (int)this.cmbTipoFactura.Value == 9)
			{
				this.uGridTransacciones.DisplayLayout.Bands[0].Columns["Codigo"].Header.Caption = "Origen";
				this.uGridTransacciones.DisplayLayout.Bands[0].Columns["Nombre"].Header.Caption = "Destino";
			}
			else 
			{
				this.uGridTransacciones.DisplayLayout.Bands[0].Columns["Codigo"].Header.Caption = "Codigo";
				this.uGridTransacciones.DisplayLayout.Bands[0].Columns["Nombre"].Header.Caption = "Nombre";
			}
			#endregion Encabezado Columnas Codigo y Nombre

			#region Limpia Detalle
			string sSQL = string.Format("Exec DetalleEntregaRecepcionMercaderia 0");
			this.uGridArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			#endregion Limpia Detalle
			
			#region Controles
			if (this.uGridTransacciones.Rows.Count > 0) this.btnIngresoSeriales.Enabled = true; else this.btnIngresoSeriales.Enabled = false;

			if ((int)this.cmbTipoFactura.Value == 7) 
			{
				this.btnGenerarSeriales.Enabled = true;
				this.btnIngresoSeriales.Enabled = false;

				if (this.cmbEstados.Text.ToString().Trim() == "PENDIENTE")
				{
					this.btnGenerarSeriales.Text = "&Generar Seriales";
					this.mnuImprimirSerialesGeneradas.Enabled = false;
				}
				else
				{
					this.btnGenerarSeriales.Text = "&Ver Seriales";
					this.btnImprimir.Enabled = true;
					this.mnuImprimirSerialesGeneradas.Enabled = true;
				}
			}
			else if ((int)this.cmbTipoFactura.Value == 2) 
			{
				this.btnImprimir.Enabled = true;
				this.mnuImprimirSerialesGeneradas.Enabled = true;
			}
			else 
			{
				this.btnGenerarSeriales.Text = "&Generar Seriales";
				this.btnGenerarSeriales.Enabled = false;
				this.btnIngresoSeriales.Enabled = true; 
				this.mnuImprimirSerialesGeneradas.Enabled = false;
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
			else if ((int)this.cmbTipoFactura.Value == 51)
			{
				if (this.cmbEstados.Text.ToString().Trim() == "PENDIENTE")
				{
					this.btnProcesar.Enabled = true;
					this.btnImprimir.Enabled = true;					
				}
				else if (this.cmbEstados.Text.ToString().Trim() == "ENTREGADO") 
				{
					this.btnProcesar.Enabled = true;					
				}
				else if (this.cmbEstados.Text.ToString().Trim() == "DEVUELTO") 
				{
					this.btnProcesar.Enabled = false;
				}
			}
			else this.btnProcesar.Enabled = false;
			#endregion Controles
		}

		#region Vista ODC Para facturacion parcial
		private void ConfigurarVistaFacturacion()
		{
			UltraGridBand band = this.uGridArticulos.DisplayLayout.Bands[0];

			// Mostrar columnas nuevas de facturación
			if (band.Columns.Exists("EstadoFacturacionDescripcion"))
				band.Columns["EstadoFacturacionDescripcion"].Hidden = false;

			if (band.Columns.Exists("CantidadFacturada"))
				band.Columns["CantidadFacturada"].Hidden = false;

			if (band.Columns.Exists("CantidadPorFacturar"))
				band.Columns["CantidadPorFacturar"].Hidden = false;

			if (band.Columns.Exists("CantidadPendiente"))
				band.Columns["CantidadPendiente"].Hidden = false;

			//  Opcional: ocultar columnas antiguas si no se usan en este tipo
			// (solo si realmente quieres reemplazar la vista)
			// band.Columns["Codigo"].Hidden = true;
			// band.Columns["Nombre"].Hidden = true;
		}

		private void ConfigurarVistaNormal()
		{
			UltraGridBand band = this.uGridArticulos.DisplayLayout.Bands[0];

			// Ocultar columnas de facturación
			if (band.Columns.Exists("EstadoFacturacionDescripcion"))
				band.Columns["EstadoFacturacionDescripcion"].Hidden = true;

			if (band.Columns.Exists("CantidadFacturada"))
				band.Columns["CantidadFacturada"].Hidden = true;

			if (band.Columns.Exists("CantidadPorFacturar"))
				band.Columns["CantidadPorFacturar"].Hidden = true;

			if (band.Columns.Exists("CantidadPendiente"))
				band.Columns["CantidadPendiente"].Hidden = true;
		}
		#endregion

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null) return;
			if (this.cmbTipoFactura.ActiveRow == null) return;
			if (this.cmbEstados.Text.ToString().Length == 0) return;

			this.Informacion();
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipoFactura_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodegaDestino_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void Mensaje()
		{
			int iContPendiente = 0;
			int iContProcesado = 0;

			this.uGridArticulos.DisplayLayout.Bands[0].Header.Caption = this.uGridArticulos.Rows.Count + " ITEMS";

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridArticulos.Rows.All)
			{
				if (dr.Cells["EstadoSerial"].Value.ToString() == "PROCESADO")
				{
					dr.Appearance.BackColor = Color.LightBlue;
					dr.Appearance.BackColor2 = Color.LightBlue;				
				}

				if (dr.Cells["EstadoSerial"].Value.ToString() == "PENDIENTE") iContPendiente++;
				
				if (dr.Cells["EstadoSerial"].Value.ToString() == "PROCESADO") iContProcesado++;
			}

			this.uGridArticulos.DisplayLayout.Bands[0].Header.Caption = this.uGridArticulos.DisplayLayout.Bands[0].Header.Caption + " - " + iContPendiente.ToString() + " PENDIENTES " + iContProcesado + " PROCESADOS";
		}

		private void ListaDetalle()
		{
			if (this.uGridTransacciones.ActiveRow == null) return;
			
			this.uGridArticulos.DisplayLayout.Bands[0].Header.Caption = "";
			
			string sSQL = string.Format("Exec DetalleEntregaRecepcionMercaderia {0}", 
				(int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value);
			this.uGridArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			Mensaje();
		}

		private void uGridTransacciones_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			ListaDetalle();
		}

		private void frmEntregaRecepcionMercaderia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnGenerarSeriales_Click(object sender, System.EventArgs e)
		{
			try
			{
				if ((int)this.cmbTipoFactura.Value != 7)
				{
					MessageBox.Show("Se genera seriales solo para Ordenes de producción.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				
				if (this.uGridTransacciones.Rows.Count == 0)
				{
					MessageBox.Show("No existen Ordenes de Producción, intente con otros parametros de busqueda.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (this.uGridTransacciones.ActiveRow == null || this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione una Orden de Produccion.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (this.cmbEstados.Text.ToString().Trim() == "PENDIENTE") 
				{
					if (DialogResult.No == MessageBox.Show("¿Seguro de generar las seriales para esta Orden de Produccion?, \n\n Este Proceso es irreversible.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			
					#region Generar
					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
					{
						oConexion.Open();

						SqlTransaction oTransaction = null;
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandType = CommandType.Text;
						oCmdActualiza.CommandTimeout = 120;

						try
						{
							oTransaction = oConexion.BeginTransaction();
							oCmdActualiza.Transaction = oTransaction;

							string sSQL = string.Format("Exec CreaIBDeOP {0}", (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value);
							oCmdActualiza.CommandText = sSQL;
							oCmdActualiza.ExecuteNonQuery();
							          
							oTransaction.Commit();

							oConexion.Close();

							Informacion();	

							MessageBox.Show("Series Generadas Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);							

							Cursor = Cursors.Default;															
						}
						catch(Exception ex)
						{
								oTransaction.Rollback();
								oConexion.Close();
								MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);						
						}
						finally
						{
							oConexion.Close();
						}
					}				
					#endregion Generar	
				}
				else
				{
					frmSerialesOPObs Seriales = new frmSerialesOPObs((int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value, this.uGridTransacciones.ActiveRow.Cells["Numero"].Value.ToString());
					Seriales.ShowDialog();
				}
			}
			catch(Exception oExc)
			{
				MessageBox.Show(oExc.ToString() + " Error del Sistema Comuniquese Con Sistemas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
			}
		}

		private bool GuiaDeRemision()
		{
			bool bGrabo = false;

			bool bAutomatico = Funcion.bEscalarSQL(cdsSeteoF, string.Format("SELECT AutGuiaRemision FROM Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);

			int iBodega = Funcion.iEscalarSQL (cdsSeteoF, string.Format("Select Local From AsignaBodegasLocales a Inner Join Bodega b On b.Bodega = a.Local Where b.Factura = 1 And b.Principal = 1 And a.Bodega = {0}", (int)this.cmbBodega.Value));

			using (frmGuiaDeRemision GuiaDeRemision = new frmGuiaDeRemision((int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value, (int)this.cmbTipoFactura.Value, iBodega, bAutomatico))
			{
				if (DialogResult.OK == GuiaDeRemision.ShowDialog())
				{	
					this.Cursor = Cursors.WaitCursor;

					if (bAutomatico) 
						if (GuiaDeRemision.bNuevo) 
							if (iBodega == 95) iBodega = 156;

							GuiaDeRemision.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 30, {0}, 0", iBodega));

					#region Guarda Guia de Remision
					using (frmNumeracionGuiaRemision Numero = new frmNumeracionGuiaRemision(GuiaDeRemision.txtNumero.Text))
					{						
						string sSQL = string.Format("Exec GGuiaRemision {0}, {1}, '{2}', '{3}', {4}, '{5}', {6}, {7}, {8}, '{9}', '{10}', {11}, '', 0", 
							GuiaDeRemision.txtIdGuiaRemision.Value, GuiaDeRemision.txtIdCompra.Value, 
							Convert.ToDateTime(GuiaDeRemision.dtIniTraslado.Value).ToString("yyyyMMdd"), Convert.ToDateTime(GuiaDeRemision.dtFinTraslado.Value).ToString("yyyyMMdd"), 
							iBodega, GuiaDeRemision.txtNumero.Text.ToString(), (int)this.cmbTipoFactura.Value, (int)GuiaDeRemision.cmbTraslado.Value, 
							(int)GuiaDeRemision.cmbResponsable.Value, GuiaDeRemision.txtPlaca.Text.ToString(), GuiaDeRemision.txtNotas.Text.ToString().Trim(), GuiaDeRemision.iBodegaDestino);
						GuiaDeRemision.txtIdGuiaRemision.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
						
						if (GuiaDeRemision.bNuevo)
						{							
							string sClaveAcceso = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec FE_ClaveDeAcceso {0}, 6", (int)GuiaDeRemision.txtIdGuiaRemision.Value));

							string cadInvertida = Facturacion.invertirCadena(sClaveAcceso.Substring(0, 48));
							string ValM11 = Facturacion.Modulo11(cadInvertida).ToString().Trim();
							string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update MaestroGuiaRemision Set Autorizacion = '{0}' Where idMaestroGuiaRemision = {1}", sClaveSRI, (int)GuiaDeRemision.txtIdGuiaRemision.Value));
						}


						//GuiaDeRemision.btnImprimir_Click(sender, e);				

						this.Cursor = Cursors.Default;
					}
					#endregion Guarda Guia de Remision

					bGrabo = true;
				}
				else bGrabo = false;
			}

			return bGrabo;
		}

		private void ImprimeTF()
		{
			if (this.uGridTransacciones.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Transferencia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una Transferencia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.Cursor = Cursors.WaitCursor;

			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldCompra = new ParameterField ();
			ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
			paramFieldCompra.ParameterFieldName = "@idCompra";
			ValIdCompra.Value = int.Parse(this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value.ToString());
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

		private void ImprimeGuiaDeRemision(int idCompra)
		{
			int idGuiaRemision = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idMaestroGuiaRemision From MaestroGuiaRemision Where idCompra = {0}", idCompra));

			Cursor = Cursors.WaitCursor;
			ParameterFields paramFields = new ParameterFields ();

			ParameterField Compra = new ParameterField ();
			ParameterDiscreteValue IdCompra = new ParameterDiscreteValue ();
			//Compra.ParameterFieldName = "@idCompra";
			Compra.ParameterFieldName = "@idCPA";
			IdCompra.Value = idGuiaRemision;
			Compra.CurrentValues.Add (IdCompra);
			paramFields.Add (Compra);

//			ParameterField Factura = new ParameterField ();
//			ParameterDiscreteValue TipoFactura = new ParameterDiscreteValue ();
//			Factura.ParameterFieldName = "@idTipofactura";
//			TipoFactura.Value = (int)this.cmbTipoFactura.Value;
//			Factura.CurrentValues.Add (TipoFactura);
//			paramFields.Add (Factura);

//			Reporte miRep = new Reporte("GuiaDeRemision.rpt", "");
			Reporte miRep = new Reporte("Doc_GR.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Guia de Remisión");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
			
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update MaestroGuiaRemision Set Impreso = 1 Where idMaestroGuiaRemision = {0}", idGuiaRemision));
		}

		/// <summary>
		/// Llama al servicio de transferencia con el código dado
		/// y maneja la respuesta.
		/// </summary>
		/// <param name="codigoTransf">Código de la transferencia</param>
		private void btnProcesar_Click(object sender, System.EventArgs e)
		{
			/*
			if ((int)this.cmbTipoFactura.Value == 51)
			{
			#region Procesa Acta de Entrega
				int iEstado = 1;
				if (this.cmbEstados.Text.ToString() == "ENTREGADO") iEstado = 15;
				
				string sSQLProc = string.Format("Exec ProcesarActasDeEntrega {0}, {1}", 
					(int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value, iEstado);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLProc);
							
				int iCont = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridArticulos.Rows.All)
				{
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
					{
						iCont	= Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleSeriales Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));
					}
					else
					{
						iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1}) Select ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1} Else Select 0", (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetCompra"].Value));
					}

					string sSQLAjuste = string.Format("Exec ActualizaEnvioRecepcionTransferencia {0}, {1}, {2}, '{3}'", 
						(int)dr.Cells["idDetCompra"].Value, (int)this.cmbTipoFactura.Value, iCont, dr.Cells["Notas"].Value.ToString());
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAjuste);
				}
			#endregion Procesa Acta de Entrega

				return;
			}
*/
			#region Validacion
			#region Parametros
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
			#endregion Parametros

			#region Documentos
			if (this.uGridTransacciones.Rows.Count == 0)
			{
				MessageBox.Show("No hay transferencias pendientes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBuscar.Focus();
				return;
			}

			if (this.uGridTransacciones.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (Funcion.ValidaBloqueo((int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value, 1, cdsSeteoF)) return;	

			if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Borrar From Compra Where idCompra = {0}", (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value), true))
			{
				MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.Informacion();
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From Compra Where idCompra = {0}", (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value)) == 3)
			{
				MessageBox.Show("La Transferencia ya fue Procesada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.btnVer_Click(sender, e);
				return;
			}
			#endregion Documentos

			#region Articulos
			if ((int)this.uGridArticulos.Rows.Count == 0)
			{
				MessageBox.Show("Seleccione una transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridArticulos.Rows.All)
			{
				if ((int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value != (int)dr.Cells["idCompra"].Value)
				{
					MessageBox.Show("Seleccione una transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					return;
				}
			}
			#endregion Articulos

			#region Valida que no se procese en cero
			int iCantidadValidaSalida = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridArticulos.Rows.All)
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
				MessageBox.Show("No puede procesar toda la Transferencia en Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				return;
			}			
			#endregion Valida que no se procese en cero
			
			#region Valida Cantidades
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridArticulos.Rows.All)
			{
				bool bValidaCantidadSeriales = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ERMValidaCantidadSerialesRegistradas {0}, {1}, {2}", 
					(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["Cantidad"].Value), true);
					
				if (!bValidaCantidadSeriales)
				{					
					if (!miAcceso.BProcesarSerialesFaltantes)
					{
						MessageBox.Show(string.Format("Ingrese todas las Seriales del Articulo : '{0}'", dr.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridArticulos.ActiveRow = dr;
						this.uGridArticulos.ActiveRow.Selected = true;
						return;
					}
					else
					{
						#region Validar Notas
						if (dr.Cells["Notas"].Value.ToString().Length == 0)
						{
							MessageBox.Show(string.Format("Del articulo : '{0}' no se ha registrado todas las seriales \n\nIngrese el motivo por el cual no se transfiere completo", dr.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.uGridArticulos.ActiveRow = dr;
							this.uGridArticulos.ActiveRow.Selected = true;
							return;
						}
						if (dr.Cells["Notas"].Value.ToString().Length < 20)
						{
							MessageBox.Show(string.Format("Del articulo : '{0}' no se ha registrado todas las seriales \n\nDebe ingresar el motivo con un minimo de 20 caracteres", dr.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.uGridArticulos.ActiveRow = dr;
							this.uGridArticulos.ActiveRow.Selected = true;
							return;
						}
						#endregion Validar Notas
					}
				}
			}
			#endregion Valida Cantidades			
			#endregion Validacion
			
			#region Valida Existencia Invenatrio enero
			int iVSerial = 0;
			int iVSerialVal = 0;
			string sVSerial = "";
			int icountS = 0;

			if ((int)this.cmbTipoFactura.Value == 8 )
			{
				if (MenuLatinium.stIdDB == 1) /*POINT*/
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridArticulos.Rows.All)
					{
						iVSerial = 0;
						iVSerialVal = 0;
						sVSerial ="";
						icountS = 0;

						icountS = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT (1) From dbo.DetalleSeriales Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));						
						if(icountS > 0)
						{
							iVSerial = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idSerial From dbo.DetalleSeriales Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));						
							if (iVSerial > 0)
							{
								iVSerialVal = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaSerialesConteo] {0}", iVSerial));	

								if ( iVSerialVal <= 0)
								{
									sVSerial = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaSerialesConteoMensaje] {0}", iVSerial));

									MessageBox.Show("" + sVSerial + "\n\n  No existe por favor comunicarse con el área de \n\n Auditoria Ext: 2011 / 2010 ", "Point Technology");
									return;
								}
							}
						}
					}
				}
			}
			#endregion Valida Existencia Invenatrio enero

			if (DialogResult.No == MessageBox.Show("¿Esta seguro de procesar la transferencia seleccionada?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{
				this.Informacion();
				return;
			}

			#region Mensaje de confirmacion
			int iCont = 0;
			double iCantidad = 0;
			string sMensaje = "";
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridArticulos.Rows.All)
			{
				if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
				{
					iCont	= Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleSeriales Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));
				
					iCantidad = (int)dr.Cells["Cantidad"].Value;
				
					if (iCont < iCantidad) sMensaje = sMensaje + "\n\n" + dr.Cells["Codigo"].Value.ToString() + " tiene " + iCont.ToString() + " seriales registradas de " + iCantidad.ToString(); 
				}
				else
				{
					iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1}) Select ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1} Else Select 0", (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetCompra"].Value));

					iCantidad = (int)dr.Cells["Cantidad"].Value;
				
					if (iCont < iCantidad) sMensaje = sMensaje + "\n\n" + dr.Cells["Codigo"].Value.ToString() + " tiene " + iCont.ToString() + " productos registrados de " + iCantidad.ToString(); 
				}
			}
			#endregion Mensaje de confirmacion			
				
			#region Muestra mensaje de confirmacion validando si hay permiso de procesar faltantes
			if (sMensaje.Length > 0)
			{
				if (!miAcceso.BProcesarSerialesFaltantes) 
				{
					MessageBox.Show("Hay Articulos que no se han cargado todas las Seriales \n\n Ingrese todas las Seriales o Consulte al Administrador", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Stop);
					return;
				}

				if (DialogResult.No == MessageBox.Show("La transferencia seleccionada tiene articulos que no se han cargado todas las seriales. \n\n\n\n" + sMensaje + "\n\n\n\n ¿Esta seguro de procesarla?", 
					"Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) return;
			}
			#endregion Muestra mensaje de confirmacion validando si hay permiso de procesar faltantes

			this.Cursor = Cursors.WaitCursor;

			int idBloqueaTransacciones = 0;
			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			#region Guia de remisión en transferencia de salida
			bool bCreaGuiaDeRemision = false;
			if ((int)this.cmbTipoFactura.Value == 8)
			{
				string sSQLVGR = string.Format("Exec ERMValidaGuiaDeremision {0}, {1}", (int)this.cmbBodega.Value, (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value);
				bCreaGuiaDeRemision = Funcion.bEscalarSQL(cdsSeteoF, sSQLVGR, true);
								
				if (bCreaGuiaDeRemision) 
				{
					if (!this.GuiaDeRemision())
					{
						string sSQLDesbloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value);
						Funcion.iEscalarSQL(cdsSeteoF, sSQLDesbloqueo, true);

						idBloqueaTransacciones = 0;
						this.Informacion();
						
						this.Cursor = Cursors.Default;

						return;
					}
				}
			}
			#endregion Guia de remisión en transferencia de salida

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;
						
				try
				{							
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Procesa Transferencia
					oCmdActualiza.CommandText = string.Format("Exec ProcesarTransferencias {0}, {1}", (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value, (int)this.cmbTipoFactura.Value);
					oCmdActualiza.ExecuteNonQuery();

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridArticulos.Rows.All)
					{
						oCmdActualiza.CommandText = string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value);
						bool bManejaSerial = (bool)oCmdActualiza.ExecuteScalar();

						if (bManejaSerial)
						{
							oCmdActualiza.CommandText = string.Format("Select Count(*) From DetalleSeriales Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value);
							iCont	= (int)oCmdActualiza.ExecuteScalar();
						}
						else
						{
							oCmdActualiza.CommandText = string.Format("If Exists(Select ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1}) Select ISNULL(Cantidad, 0) From Serie Where idArticulo = {0} And idDetCompra = {1} Else Select 0", (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetCompra"].Value);
							iCont = (int)oCmdActualiza.ExecuteScalar();
						}

						oCmdActualiza.CommandText = string.Format("Exec ActualizaEnvioRecepcionTransferencia {0}, {1}, {2}, '{3}'", 
							(int)dr.Cells["idDetCompra"].Value, (int)this.cmbTipoFactura.Value, iCont, dr.Cells["Notas"].Value.ToString());
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Procesa Transferencia

					#region Mail de control - Actualiza cantidad transferido en conteo de inventario
					oCmdActualiza.CommandText = string.Format("Exec MailControlTransferencias {0}", (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value);
					oCmdActualiza.ExecuteNonQuery();

					DateTime dtFechaTR = (DateTime)this.uGridTransacciones.ActiveRow.Cells["Fecha"].Value;
					oCmdActualiza.CommandText = string.Format("Exec ActualizaTransferidoConteo '{0}', {1}, {2}", 
						dtFechaTR.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value);
					oCmdActualiza.ExecuteNonQuery();
					#endregion Mail de control - Actualiza cantidad transferido en conteo de inventario

					#region Asiento
//					oCmdActualiza.CommandText = string.Format("Exec AsientoTransferencias {0}, {1}", 
//						(int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value, (int)this.cmbTipoFactura.Value);
//					oCmdActualiza.ExecuteNonQuery();
					#endregion Asiento

					#region Inventario
					oCmdActualiza.CommandText = string.Format("Exec GuardaMovimientosInventario {0}, {1}", 
						(int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value, (int)this.cmbTipoFactura.Value);
					oCmdActualiza.ExecuteNonQuery();
					#endregion Inventario
			
					this.uGridArticulos.DisplayLayout.Bands[0].Header.Caption = "";
								
					//Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NovedadesTransferencias {0}", (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value), true);
					
					#region Retorna Mercaderia no recibida en transferencia de ingreso al origen
					if ((int)this.cmbTipoFactura.Value == 9 && MenuLatinium.stDirFacturacion == "POINT")
					{
						oCmdActualiza.CommandText = string.Format("Exec ERMIBRecepcionIncompleta {0}", (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value);
            int idCompraIMR = (int)oCmdActualiza.ExecuteScalar();

						if (idCompraIMR > 0)
						{
							#region Inventario
							oCmdActualiza.CommandText = string.Format("Exec GuardaMovimientosInventario {0}, 9", idCompraIMR);
							oCmdActualiza.ExecuteNonQuery();
							#endregion Inventario

							MessageBox.Show("Ajuste de ingreso de mercaderia no recibida generado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
							#region Imprime el reporte
							string stFiltro = "{Compra.idCompra} = " + idCompraIMR.ToString();					
							using (Reporte miReporte = new Reporte("EgresosDeBodega.Rpt", stFiltro))
							{						
								miReporte.ShowDialog();				
							}
							#endregion Imprime el reporte		
							Cursor = Cursors.Default;
						}
					}
					#endregion Retorna Mercaderia no recibida en transferencia de ingreso al origen
					
					oCmdActualiza.CommandText = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value);
					oCmdActualiza.ExecuteNonQuery();

					idBloqueaTransacciones = 0;

					oTransaction.Commit();

					if ((int)this.cmbTipoFactura.Value == 8 && MenuLatinium.stIdDB == 1) // Solo salida, Solo POINT
					{
						string codigoTransf = this.uGridTransacciones.ActiveRow.Cells["Numero"].Value.ToString();
    
						// Llamada simple al servicio (ya no ensuciamos el front con lógica de logs ni API)
						TransferenciaService service = new TransferenciaService();
						service.EjecutarProcesamiento(codigoTransf);
					}

					MessageBox.Show("Transferencia procesada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				
					#region Imprime Transferencia
					ImprimeTF();

					if ((int)this.cmbTipoFactura.Value == 8) 
						if (bCreaGuiaDeRemision) 
							ImprimeGuiaDeRemision((int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value);
					#endregion Imprime Transferencia

					this.Informacion();
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
				finally
				{
					oConexion.Close();
				}
			}
			this.Cursor = Cursors.Default;
		}

		
		
		private void dtDesde_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void uGridTransacciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnIngresoSeriales_Click(object sender, System.EventArgs e)
		{
			try
			{

				#region Codigo Anterior
//				#region Validacion
//				int idRegistroBloqueaERMercaderia = 0;
//				int idBloqueaTransacciones = 0;
//
//				#region Transacciones
//				if (this.uGridTransacciones.Rows.Count == 0)
//				{
//					MessageBox.Show("No existen documentos con estos parámetros de búsqueda.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					return;
//				}
//
//				if (this.uGridTransacciones.ActiveRow == null)
//				{
//					MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					return;
//				}
//
//				if (Funcion.ValidaBloqueo((int)this.uGridArticulos.ActiveRow.Cells["idCompra"].Value, 1, cdsSeteoF)) return;	
//
//				#region Estado Documento
//				if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Borrar From Compra Where idCompra = {0}", (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value), true))
//				{
//					MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.Informacion();
//					return;
//				}
//				int iEstadoDoc = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From Compra Where idCompra = {0}", 
//					(int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value));
//				if (!miAcceso.BVerSerialesProcesadas && iEstadoDoc == 3)
//				{
//					MessageBox.Show("El documento ya fue Procesado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.Informacion();
//					return;
//				}
//				if (!miAcceso.BVerSerialesProcesadas && iEstadoDoc == 9)
//				{
//					MessageBox.Show("El documento ya fue Facturado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.Informacion();
//					return;
//				}
//				#endregion Estado Documento			
//				#endregion Transacciones
//
//				#region Articulos
//				if (this.uGridArticulos.ActiveRow == null)
//				{
//					MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					return;
//				}
//
//				if (this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value == System.DBNull.Value)
//				{
//					MessageBox.Show("Seleccione un item", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					return;
//				}
//
//				if ((int)this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value == 0)
//				{
//					MessageBox.Show("Seleccione un item", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					return;
//				}
//
//				if ((int)this.uGridArticulos.ActiveRow.Cells["Cantidad"].Value == 0)
//				{
//					MessageBox.Show("Cantidad en 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					return;
//				}
//				#region Estado Articulos
//				if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit = 0 If Exists(Select idCompra From DetCompra Where idDetCompra = {0}) Set @Res = 1 Select @Res", 
//					(int)this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value), true))
//				{
//					MessageBox.Show("El Articulo Fue Eliminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.Informacion();
//					return;
//				}
//				if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit = 0 If ((Select idArticulo From DetCompra Where idDetCompra = {0}) = {1}) Set @Res = 1 Select @Res", 
//					(int)this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value, (int)this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value), true))
//				{
//					MessageBox.Show("Se ha cambiado de producto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.Informacion();
//					return;
//				}
//				if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit = 0 If ((Select Cantidad From DetCompra Where idDetCompra = {0}) = {1}) Set @Res = 1 Select @Res", 
//					(int)this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value, (int)this.uGridArticulos.ActiveRow.Cells["Cantidad"].Value), true))
//				{
//					MessageBox.Show(string.Format("Se ha modificado la cantidad a procesar del articulo '{0}'", this.uGridArticulos.ActiveRow.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.Informacion();
//					return;
//				}								
//				bool bEstadoSerial = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select EstadoSerial From DetCompra Where idDetCompra = {0}", 
//					(int)this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value), true);
//				if (!miAcceso.BVerSerialesProcesadas && bEstadoSerial)
//				{
//					MessageBox.Show("El Articulo ya esta PROCESADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.Informacion();
//					return;
//				}
//				#endregion Estado Articulos
//				#endregion Articulos
//
//				#region Valida Usuario Activo
//				string sSQLValRB = string.Format("Select COUNT(*) From RegistroBloqueaERMercaderia Where Bloqueo = 1 And idDetCompra = {0}", (int)this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value);
//				if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValRB, true) > 0)
//				{
//					string sSQLCB = string.Format("Select Top 1 Usuario From RegistroBloqueaERMercaderia Where Bloqueo = 1 And idDetCompra = {0} Order By Fecha Desc", (int)this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value);
//					string sUsuario = Funcion.sEscalarSQL(cdsSeteoF, sSQLCB);
//												
//					MessageBox.Show(string.Format("No puede ingresar a Registro de Seriales de este Articulo Hasta que el usuario '{0}' cierre esta Pantalla", sUsuario), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				
//					return;				
//				}										
//				#endregion Valida Usuario Activo
//				#endregion Validacion
//
//				#region Variables
//				int IdDetCompra = (int) this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value;
//				int IdCompra = (int) this.uGridArticulos.ActiveRow.Cells["idCompra"].Value;				
//				int Cantidad = (int) this.uGridArticulos.ActiveRow.Cells["Cantidad"].Value;
//				string Codigo = (string) this.uGridArticulos.ActiveRow.Cells["Codigo"].Value;
//				string Articulo = (string) this.uGridArticulos.ActiveRow.Cells["Articulo"].Value;			
//				int IdArticulo = (int) this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value;
//				int IdTipoFactura = (int)this.cmbTipoFactura.Value;
//				int iBodega = (int) this.cmbBodega.Value;				
//				bool ManejaSerial = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", IdArticulo), true);
//				
//				string sNum = "";				
//				int iDisponible = 0;
//				bool bUsado = true;				
//				#endregion Variables			
//
//				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", IdCompra);
//				idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);
//
//				string sSQLB = string.Format("Exec BloqueaERMercaderia 1, {0}, {1}", IdDetCompra, idRegistroBloqueaERMercaderia);
//				idRegistroBloqueaERMercaderia = Funcion.iEscalarSQL(cdsSeteoF, sSQLB, true);
//
//				#region Registro de Mercaderia
//				#region Registro Productos con Seriales
//				if (ManejaSerial)
//				{	
//					using (frmIngresoSeriales Serial = new frmIngresoSeriales(IdCompra, IdTipoFactura, sNum, IdDetCompra, Cantidad, IdArticulo, Codigo, Articulo, iBodega, Estado))
//					{							
//						if (DialogResult.OK == Serial.ShowDialog())
//						{
//							#region Seriales de importación mixta
//							if (IdTipoFactura == 2 && (int)this.uGridTransacciones.ActiveRow.Cells["Estado"].Value == 12)
//							{
//								string sSQLGenSer =  string.Format("Exec SerialesImpMixOCCompra {0}", IdCompra);
//								Funcion.EjecutaSQL(cdsSeteoF, sSQLGenSer, true);
//							}
//							#endregion Seriales de importación mixta
//																					
//							string sSQLEstado = string.Format("Select Estado From Compra Where idCompra = {0}", IdCompra);
//							int iEstado = Funcion.iEscalarSQL(cdsSeteoF, sSQLEstado);
//							
//							if (IdTipoFactura == 36)
//								if (iEstado == 3)
//									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridArticulos.Rows.All)
//										Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetCompra Set Principal = 1 Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));
//
//							if (iEstado == 3 || iEstado == 13) IdCompra = 0;
//
//							this.ListaDetalle();
//						}
//						else this.ListaDetalle();
//					}					
//				}
//				#endregion Registro Productos con Seriales
//
//				#region Registro Productos sin Seriales
//				if (!ManejaSerial)
//				{					
//					using (ArticulosSinSerial Serial = new ArticulosSinSerial(Cantidad, IdArticulo, Codigo, Articulo, IdCompra, sNum, iDisponible, IdTipoFactura, IdDetCompra, bUsado, iBodega, false, this.uGridArticulos.ActiveRow.Cells["Notas"].Value.ToString()))
//					{												
//						if (DialogResult.OK == Serial.ShowDialog())
//						{							
//							string sSQLEstado = string.Format("Select Estado From Compra Where idCompra = {0}", IdCompra);
//							int Est = Funcion.iEscalarSQL(cdsSeteoF, sSQLEstado);
//							if (Est == 3) IdCompra = 0;
//
//							this.ListaDetalle();
//						}
//						else this.ListaDetalle();
//					}					
//				}
//				#endregion Registro Productos sin Seriales
//				
//				string sSQLDesbloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, IdCompra);
//				Funcion.EjecutaSQL(cdsSeteoF, sSQLDesbloqueo, true);
//
//				string sSQLDB = string.Format("Exec BloqueaERMercaderia 0, {0}, {1}", IdDetCompra, idRegistroBloqueaERMercaderia);
//				Funcion.EjecutaSQL(cdsSeteoF, sSQLDB, true);
//
//				idRegistroBloqueaERMercaderia = 0;
//				idBloqueaTransacciones = 0;
//				#endregion Registro de Mercaderia

         #endregion Codigo Anterior

				#region Validacion
				int idRegistroBloqueaERMercaderia = 0;
				int idBloqueaTransacciones = 0;

				#region Transacciones
				if (this.uGridTransacciones.Rows.Count == 0)
				{
					MessageBox.Show("No existen documentos con estos parámetros de búsqueda.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (this.uGridTransacciones.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (Funcion.ValidaBloqueo((int)this.uGridArticulos.ActiveRow.Cells["idCompra"].Value, 1, cdsSeteoF)) return;	

				#region Estado Documento
				if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Borrar From Compra Where idCompra = {0}", (int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value), true))
				{
					MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.Informacion();
					return;
				}
				int iEstadoDoc = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From Compra Where idCompra = {0}", 
					(int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value));
				if (!miAcceso.BVerSerialesProcesadas && iEstadoDoc == 3)
				{
					MessageBox.Show("El documento ya fue Procesado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.Informacion();
					return;
				}
				if (!miAcceso.BVerSerialesProcesadas && iEstadoDoc == 9)
				{
					MessageBox.Show("El documento ya fue Facturado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.Informacion();
					return;
				}
				#endregion Estado Documento			
				#endregion Transacciones

				#region Articulos
				if (this.uGridArticulos.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un item", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if ((int)this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value == 0)
				{
					MessageBox.Show("Seleccione un item", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if ((int)this.uGridArticulos.ActiveRow.Cells["Cantidad"].Value == 0)
				{
					MessageBox.Show("Cantidad en 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#region Estado Articulos
				if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit = 0 If Exists(Select idCompra From DetCompra Where idDetCompra = {0}) Set @Res = 1 Select @Res", 
					(int)this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value), true))
				{
					MessageBox.Show("El Articulo Fue Eliminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.Informacion();
					return;
				}
				if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit = 0 If ((Select idArticulo From DetCompra Where idDetCompra = {0}) = {1}) Set @Res = 1 Select @Res", 
					(int)this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value, (int)this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value), true))
				{
					MessageBox.Show("Se ha cambiado de producto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.Informacion();
					return;
				}
				if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit = 0 If ((Select Cantidad From DetCompra Where idDetCompra = {0}) = {1}) Set @Res = 1 Select @Res", 
					(int)this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value, (int)this.uGridArticulos.ActiveRow.Cells["Cantidad"].Value), true))
				{
					MessageBox.Show(string.Format("Se ha modificado la cantidad a procesar del articulo '{0}'", this.uGridArticulos.ActiveRow.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.Informacion();
					return;
				}								
				bool bEstadoSerial = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select EstadoSerial From DetCompra Where idDetCompra = {0}", 
					(int)this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value), true);
				if (!miAcceso.BVerSerialesProcesadas && bEstadoSerial)
				{
					MessageBox.Show("El Articulo ya esta PROCESADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.Informacion();
					return;
				}
				#endregion Estado Articulos
				#endregion Articulos

				#region Valida Usuario Activo
				string sSQLValRB = string.Format("Select COUNT(*) From RegistroBloqueaERMercaderia Where Bloqueo = 1 And idDetCompra = {0}", (int)this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value);
				if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValRB, true) > 0)
				{
					string sSQLCB = string.Format("Select Top 1 Usuario From RegistroBloqueaERMercaderia Where Bloqueo = 1 And idDetCompra = {0} Order By Fecha Desc", (int)this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value);
					string sUsuario = Funcion.sEscalarSQL(cdsSeteoF, sSQLCB);
												
					MessageBox.Show(string.Format("No puede ingresar a Registro de Seriales de este Articulo Hasta que el usuario '{0}' cierre esta Pantalla", sUsuario), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				
					return;				
				}										
				#endregion Valida Usuario Activo
				#endregion Validacion

				#region Variables
				int IdDetCompra = (int) this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value;
				int IdCompra = (int) this.uGridArticulos.ActiveRow.Cells["idCompra"].Value;				
				int Cantidad = (int) this.uGridArticulos.ActiveRow.Cells["Cantidad"].Value;
				string Codigo = (string) this.uGridArticulos.ActiveRow.Cells["Codigo"].Value;
				string Articulo = (string) this.uGridArticulos.ActiveRow.Cells["Articulo"].Value;			
				int IdArticulo = (int) this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value;
				int IdTipoFactura = (int)this.cmbTipoFactura.Value;
				int iBodega = (int) this.cmbBodega.Value;				
				bool ManejaSerial = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", IdArticulo), true);
				int idGrupoMoto =0;
				idGrupoMoto = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" EXEC [ValidaColor] {0}", (int) this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value));	
		
				string sNum = "";				
				int iDisponible = 0;
				bool bUsado = true;				
				#endregion Variables			

				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", IdCompra);
				idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

				string sSQLB = string.Format("Exec BloqueaERMercaderia 1, {0}, {1}", IdDetCompra, idRegistroBloqueaERMercaderia);
				idRegistroBloqueaERMercaderia = Funcion.iEscalarSQL(cdsSeteoF, sSQLB, true);

				#region Registro de Mercaderia
				#region Registro Productos con Seriales
				if (ManejaSerial)
				{	
					using (frmIngresoSeriales Serial = new frmIngresoSeriales(IdCompra, IdTipoFactura, sNum, IdDetCompra, Cantidad, IdArticulo, Codigo, Articulo, iBodega, Estado, idGrupoMoto))
					{							
						if (DialogResult.OK == Serial.ShowDialog())
						{
							#region Seriales de importación mixta
							if (IdTipoFactura == 2 && (int)this.uGridTransacciones.ActiveRow.Cells["Estado"].Value == 12)
							{
								string sSQLGenSer =  string.Format("Exec SerialesImpMixOCCompra {0}", IdCompra);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLGenSer, true);
							}
							#endregion Seriales de importación mixta
																					
							string sSQLEstado = string.Format("Select Estado From Compra Where idCompra = {0}", IdCompra);
							int iEstado = Funcion.iEscalarSQL(cdsSeteoF, sSQLEstado);
							
							if (IdTipoFactura == 36)
								if (iEstado == 3)
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridArticulos.Rows.All)
										Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetCompra Set Principal = 1 Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value));

							if (iEstado == 3 || iEstado == 13) IdCompra = 0;

							this.ListaDetalle();
						}
						else this.ListaDetalle();
					}					
				}
				#endregion Registro Productos con Seriales

				#region Registro Productos sin Seriales
				if (!ManejaSerial)
				{					
					using (ArticulosSinSerial Serial = new ArticulosSinSerial(Cantidad, IdArticulo, Codigo, Articulo, IdCompra, sNum, iDisponible, IdTipoFactura, IdDetCompra, bUsado, iBodega, false, this.uGridArticulos.ActiveRow.Cells["Notas"].Value.ToString()))
					{												
						if (DialogResult.OK == Serial.ShowDialog())
						{							
							string sSQLEstado = string.Format("Select Estado From Compra Where idCompra = {0}", IdCompra);
							int Est = Funcion.iEscalarSQL(cdsSeteoF, sSQLEstado);
							if (Est == 3) IdCompra = 0;

							this.ListaDetalle();
						}
						else this.ListaDetalle();
					}					
				}
				#endregion Registro Productos sin Seriales
				
				string sSQLDesbloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, IdCompra);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDesbloqueo, true);

				string sSQLDB = string.Format("Exec BloqueaERMercaderia 0, {0}, {1}", IdDetCompra, idRegistroBloqueaERMercaderia);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDB, true);

				idRegistroBloqueaERMercaderia = 0;
				idBloqueaTransacciones = 0;
				#endregion Registro de Mercaderia
			}
			catch(Exception oExc)
			{
				MessageBox.Show(oExc.ToString() + "Error del Sistema Comuniquese Con Sistemas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
			}
		}

		private void uGridArticulos_DoubleClick(object sender, System.EventArgs e)
		{
			if ((int)this.cmbTipoFactura.Value == 7) return;

			this.btnIngresoSeriales_Click(sender, e);
		}

		private void dtDesde_Validated(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
				this.dtHasta.CalendarInfo.MinDate = Convert.ToDateTime(this.dtDesde.Value);
				this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
				this.dtHasta.Value = DateTime.Today;
			}
		}

		private void dtHasta_Validated(object sender, System.EventArgs e)
		{
			if (this.dtHasta.Value == System.DBNull.Value)
			{
				this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
				this.dtHasta.CalendarInfo.MinDate = Convert.ToDateTime(this.dtDesde.Value);
				this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
				this.dtHasta.Value = DateTime.Today;
			}
		}

		private void dtDesde_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) if (this.cmbBodega.Enabled) this.cmbBodega.Focus(); else this.cmbTipoFactura.Focus();
		}

		private void dtHasta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void uGridArticulos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) if ((int)this.cmbTipoFactura.Value == 7) return; else this.btnIngresoSeriales_Click(sender, e);
			Funcion.CaracteresEspeciales(sender, e);
		}

		private void uGridArticulos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			if ((int)this.cmbTipoFactura.Value == 51)
			{
				if (this.uGridTransacciones.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un Acta de Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Acta de Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				this.Cursor = Cursors.WaitCursor;

				string stFiltro = "{Compra.idCompra} = " + this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value.ToString();

				Reporte miRep = new Reporte("ActaDeEntrega.rpt", stFiltro);
				miRep.MdiParent = this.MdiParent;
				miRep.Show();

				this.Cursor = Cursors.Default;
			}
			else this.mnuImprimir.Show(this.btnImprimir, new Point(1, 25));
		}

		private void btnGuiaDeRemision_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.uGridTransacciones.Rows.Count == 0)
			{
				MessageBox.Show("No hay transferencias procesadas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBuscar.Focus();
				return;
			}

			if (this.uGridTransacciones.ActiveRow == null || this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridTransacciones.Focus();
				return;
			}

			if ((int)this.cmbTipoFactura.Value != 8)
			{
				MessageBox.Show("Solo para Transferencias.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipoFactura.Focus();
				return;
			}

			if (this.cmbEstados.Text.ToString().Trim() == "PENDIENTE")
			{
				MessageBox.Show("Solo para transferencias procesadas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbEstados.Focus();
				return;
			}
			#endregion Validacion

			bool bAutomatico = Funcion.bEscalarSQL(cdsSeteoF, string.Format("SELECT AutGuiaRemision FROM Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);

			using (frmGuiaDeRemision GuiaDeRemision = new frmGuiaDeRemision((int)this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value, (int)this.cmbTipoFactura.Value, (int)this.cmbBodega.Value, bAutomatico))
			{
				if (DialogResult.OK == GuiaDeRemision.ShowDialog())
				{	
					this.Cursor = Cursors.WaitCursor;

					DateTime dtInicio = (DateTime)GuiaDeRemision.dtIniTraslado.Value;
					DateTime dtFin = (DateTime)GuiaDeRemision.dtFinTraslado.Value;
					
					if (bAutomatico) 
						if (GuiaDeRemision.bNuevo) 
							if ((int)this.cmbBodega.Value == 95) this.cmbBodega.Value = 156;
							GuiaDeRemision.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 30, {0}, 0", (int)this.cmbBodega.Value));

					string sSQL = string.Format("Exec GGuiaRemision {0}, {1}, '{2}', '{3}', {4}, '{5}', {6}, {7}, {8}, '{9}', '{10}', {11}", 
						GuiaDeRemision.txtIdGuiaRemision.Value, GuiaDeRemision.txtIdCompra.Value, 
						dtInicio.ToString("yyyyMMdd"), dtFin.ToString("yyyyMMdd"), 
						(int)this.cmbBodega.Value, 
						GuiaDeRemision.txtNumero.Text.ToString(), 
						(int)this.cmbTipoFactura.Value,
						(int)GuiaDeRemision.cmbTraslado.Value, 
						(int)GuiaDeRemision.cmbResponsable.Value, 
						GuiaDeRemision.txtPlaca.Text.ToString(), 
						GuiaDeRemision.txtNotas.Text.ToString().Trim(),
						GuiaDeRemision.iBodegaDestino);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);				

					//GuiaDeRemision.btnImprimir_Click(sender, e);

					this.Cursor = Cursors.Default;
				}
			}
		}

		private void mnuImprimirSerialesGeneradas_Click(object sender, System.EventArgs e)
		{
			//			if ((int)this.cmbTipoFactura.Value != 2 || (int)this.cmbTipoFactura.Value != 7)
			//			{	
			//				MessageBox.Show("Solo para Ordenes de Compra y Produccion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				this.txtBuscar.Focus();
			//				return;
			//			}

			if (this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una Orden de Compra", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBuscar.Focus();
				return;
			}

			int idCompra = (int)uGridArticulos.ActiveRow.Cells["idCompra"].Value;

			if ((int)this.cmbTipoFactura.Value == 2)
			{			
				if (this.uGridArticulos.ActiveRow.Cells["idDetCompra"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Producto de la Orden de Compra Seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtBuscar.Focus();
					return;
				}

				bool SerialPropia = false;
				SerialPropia = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select SeriePropiaAlterna From Articulo Where idArticulo = {0}", (int)this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value), true);
				if (!SerialPropia)
				{
					MessageBox.Show("El Artículo Seleccionado No Genera Serial de POINT", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtBuscar.Focus();
					return;
				}				

				idCompra = (int)uGridArticulos.ActiveRow.Cells["idDetCompra"].Value;
			}
						
			string sSQL = string.Format("Exec GuardaTempoSeriales {0}, {1}", idCompra, (int)this.cmbTipoFactura.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			string sDir = Funcion.sEscalarSQL(cdsSeteoF, "Select ADLabelJoy From SeteoF");

			Process.Start(sDir, "");
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTipoFactura.Focus();
		}

		private void cmbTipoFactura_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbEstados.Focus();
		}

		private void cmbEstados_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.cmbBodegaDestino.Enabled) this.cmbEstados.Focus(); else this.txtBuscar.Focus();
		}

		private void cmbBodegaDestino_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void mnuImprimir_Popup(object sender, System.EventArgs e)
		{
		
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			if ((int)this.cmbTipoFactura.Value == 8 && MenuLatinium.stIdDB == 1) // Solo salida, Solo POINT
			{
				string codigoTransf = this.uGridTransacciones.ActiveRow.Cells["Numero"].Value.ToString();
    
				// Llamada simple al servicio (ya no ensuciamos el front con lógica de logs ni API)
				TransferenciaService service = new TransferenciaService();
				service.EjecutarProcesamiento(codigoTransf);
			}
		}

		private void btnRecepcionBodega_Click(object sender, System.EventArgs e)
		{
			try
			{
				#region VALIDACIONES

				if (this.uGridTransacciones.ActiveRow == null)
				{
					MessageBox.Show(
						"Seleccione una orden de compra.",
						"Sistema",
						MessageBoxButtons.OK,
						MessageBoxIcon.Warning);

					return;
				}

				if (this.cmbTipoFactura.Value == null)
				{
					MessageBox.Show(
						"Seleccione un tipo de factura.",
						"Sistema",
						MessageBoxButtons.OK,
						MessageBoxIcon.Warning);

					return;
				}

				if ((int)this.cmbTipoFactura.Value != 2)
				{
					MessageBox.Show(
						"Solo las Ordenes de compra pueden enviarse a RECEPCIÓN BODEGA.",
						"Sistema",
						MessageBoxButtons.OK,
						MessageBoxIcon.Warning);

					return;
				}

				int idCompra =
					Convert.ToInt32(
					this.uGridTransacciones.ActiveRow.Cells["idCompra"].Value);

				DialogResult result =
					MessageBox.Show(
					"¿Desea enviar la orden a RECEPCIÓN BODEGA?",
					"Confirmación",
					MessageBoxButtons.YesNo,
					MessageBoxIcon.Question);

				if (result != DialogResult.Yes)
					return;

				#endregion

				#region PROCESO

				Latinium.Services.OrdenDeCompra.Bodega.OrdenDeCompraService service =
					new Latinium.Services.OrdenDeCompra.Bodega.OrdenDeCompraService(
					this.cdsSeteoF);

				service.EnviarRecibidoBodega(
					idCompra,
					MenuLatinium.stUsuario);

				#endregion

				#region MENSAJE

				MessageBox.Show(
					"Orden enviada correctamente a RECEPCIÓN BODEGA.",
					"Sistema",
					MessageBoxButtons.OK,
					MessageBoxIcon.Information);

				#endregion

				#region RECARGAR INFORMACIÓN

				this.Informacion();

				#endregion
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					ex.Message,
					"Error",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error);
			}
		}

	}
}
