using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using System.IO;
using System.Data;
using System.Text;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmDTSeguimiento.
	/// </summary>
	public class frmDTSeguimiento : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugIngresosDT;
		private System.Windows.Forms.GroupBox gbProducto;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label30;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDaño;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugInformesDT;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label8;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.ContextMenu mnuImprimir;
		private System.Windows.Forms.MenuItem mnuImprimirNE;
		private System.Windows.Forms.MenuItem mnuImprimirEgreso;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFactura;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaCompra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProveedor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIngreso;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInforme;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIngreso;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaRetorno;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstados;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label72;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSerial;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdArticulo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerial;
		private System.ComponentModel.IContainer components;

		public frmDTSeguimiento()
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDTSeguimiento));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroNE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tecnico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Daño");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idarticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idserial");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PointCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroNE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tecnico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProveedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proveedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Daño");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idarticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idserial");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Informe");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Informe");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			this.label15 = new System.Windows.Forms.Label();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.ugIngresosDT = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.gbProducto = new System.Windows.Forms.GroupBox();
			this.txtFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtFechaCompra = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label30 = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtProveedor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDaño = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIngreso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.dtFechaIngreso = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaRetorno = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtInforme = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnGuardar = new System.Windows.Forms.Button();
			this.ugInformesDT = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label8 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btnImprimir = new System.Windows.Forms.Button();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.mnuImprimir = new System.Windows.Forms.ContextMenu();
			this.mnuImprimirNE = new System.Windows.Forms.MenuItem();
			this.mnuImprimirEgreso = new System.Windows.Forms.MenuItem();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label72 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtIdSerial = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.ugIngresosDT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.gbProducto.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDaño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRetorno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInforme)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugInformesDT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).BeginInit();
			this.SuspendLayout();
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.Location = new System.Drawing.Point(8, 11);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(37, 17);
			this.label15.TabIndex = 371;
			this.label15.Text = "Buscar";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Location = new System.Drawing.Point(280, 8);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(77, 23);
			this.btnBuscar.TabIndex = 1;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// ugIngresosDT
			// 
			this.ugIngresosDT.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugIngresosDT.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugIngresosDT.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugIngresosDT.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 100;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 20;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Nota de Entrega";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Width = 70;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Width = 180;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 140;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 18;
			ultraGridColumn10.Width = 180;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Width = 180;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Width = 220;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Factura";
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Width = 100;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.Caption = "Comprado";
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn16.Width = 70;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 14;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 15;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 16;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 17;
			ultraGridColumn20.Width = 220;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 19;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn24.Header.VisiblePosition = 23;
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
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.ugIngresosDT.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugIngresosDT.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ugIngresosDT.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugIngresosDT.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugIngresosDT.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ugIngresosDT.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugIngresosDT.DisplayLayout.Override.GroupByRowAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8.5F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugIngresosDT.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ugIngresosDT.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugIngresosDT.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugIngresosDT.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugIngresosDT.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ugIngresosDT.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugIngresosDT.Location = new System.Drawing.Point(16, 48);
			this.ugIngresosDT.Name = "ugIngresosDT";
			this.ugIngresosDT.Size = new System.Drawing.Size(1048, 160);
			this.ugIngresosDT.TabIndex = 661;
			this.ugIngresosDT.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ugIngresosDT_BeforeSelectChange);
			this.ugIngresosDT.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugIngresosDT_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn7.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn12.DefaultValue = 0;
			ultraDataColumn16.DataType = typeof(System.DateTime);
			ultraDataColumn16.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn19.DataType = typeof(System.DateTime);
			ultraDataColumn19.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
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
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// gbProducto
			// 
			this.gbProducto.Controls.Add(this.txtFactura);
			this.gbProducto.Controls.Add(this.label19);
			this.gbProducto.Controls.Add(this.label18);
			this.gbProducto.Controls.Add(this.dtFechaCompra);
			this.gbProducto.Controls.Add(this.label30);
			this.gbProducto.Controls.Add(this.txtRuc);
			this.gbProducto.Controls.Add(this.label7);
			this.gbProducto.Controls.Add(this.txtProveedor);
			this.gbProducto.Location = new System.Drawing.Point(8, 216);
			this.gbProducto.Name = "gbProducto";
			this.gbProducto.Size = new System.Drawing.Size(480, 80);
			this.gbProducto.TabIndex = 663;
			this.gbProducto.TabStop = false;
			this.gbProducto.Text = "Proveedor Servicio Tecnico Autorizado";
			// 
			// txtFactura
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFactura.Appearance = appearance9;
			this.txtFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFactura.Enabled = false;
			this.txtFactura.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtFactura.Location = new System.Drawing.Point(224, 24);
			this.txtFactura.Name = "txtFactura";
			this.txtFactura.Size = new System.Drawing.Size(96, 22);
			this.txtFactura.TabIndex = 293;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label19.Location = new System.Drawing.Point(168, 27);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(41, 17);
			this.label19.TabIndex = 292;
			this.label19.Text = "Factura";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(328, 27);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(34, 17);
			this.label18.TabIndex = 291;
			this.label18.Text = "Fecha";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaCompra
			// 
			this.dtFechaCompra.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaCompra.DateButtons.Add(dateButton1);
			this.dtFechaCompra.Enabled = false;
			this.dtFechaCompra.Format = "dd/MM/yyyy";
			this.dtFechaCompra.Location = new System.Drawing.Point(368, 25);
			this.dtFechaCompra.Name = "dtFechaCompra";
			this.dtFechaCompra.NonAutoSizeHeight = 23;
			this.dtFechaCompra.Size = new System.Drawing.Size(104, 21);
			this.dtFechaCompra.SpinButtonsVisible = true;
			this.dtFechaCompra.TabIndex = 290;
			this.dtFechaCompra.Value = ((object)(resources.GetObject("dtFechaCompra.Value")));
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(8, 27);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(36, 17);
			this.label30.TabIndex = 278;
			this.label30.Text = "R.U.C.";
			this.label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance10;
			this.txtRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(56, 24);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(104, 22);
			this.txtRuc.TabIndex = 277;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(8, 51);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(44, 17);
			this.label7.TabIndex = 270;
			this.label7.Text = "Nombre";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtProveedor
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProveedor.Appearance = appearance11;
			this.txtProveedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtProveedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProveedor.Enabled = false;
			this.txtProveedor.Location = new System.Drawing.Point(56, 48);
			this.txtProveedor.Name = "txtProveedor";
			this.txtProveedor.Size = new System.Drawing.Size(416, 22);
			this.txtProveedor.TabIndex = 269;
			this.txtProveedor.ValueChanged += new System.EventHandler(this.txtProveedor_ValueChanged);
			// 
			// txtDaño
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDaño.Appearance = appearance12;
			this.txtDaño.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDaño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDaño.Enabled = false;
			this.txtDaño.Location = new System.Drawing.Point(8, 26);
			this.txtDaño.Multiline = true;
			this.txtDaño.Name = "txtDaño";
			this.txtDaño.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtDaño.Size = new System.Drawing.Size(552, 44);
			this.txtDaño.TabIndex = 664;
			// 
			// txtIngreso
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngreso.Appearance = appearance13;
			this.txtIngreso.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngreso.Enabled = false;
			this.txtIngreso.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtIngreso.Location = new System.Drawing.Point(72, 304);
			this.txtIngreso.Name = "txtIngreso";
			this.txtIngreso.Size = new System.Drawing.Size(104, 22);
			this.txtIngreso.TabIndex = 295;
			this.txtIngreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtIngreso_KeyDown);
			this.txtIngreso.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtIngreso_KeyPress);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 307);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 17);
			this.label1.TabIndex = 294;
			this.label1.Text = "Numero";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 330);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(42, 17);
			this.label2.TabIndex = 295;
			this.label2.Text = "Ingreso";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaIngreso
			// 
			this.dtFechaIngreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaIngreso.DateButtons.Add(dateButton2);
			this.dtFechaIngreso.Enabled = false;
			this.dtFechaIngreso.Format = "dd/MM/yyyy";
			this.dtFechaIngreso.Location = new System.Drawing.Point(72, 328);
			this.dtFechaIngreso.Name = "dtFechaIngreso";
			this.dtFechaIngreso.NonAutoSizeHeight = 23;
			this.dtFechaIngreso.Size = new System.Drawing.Size(104, 21);
			this.dtFechaIngreso.SpinButtonsVisible = true;
			this.dtFechaIngreso.TabIndex = 294;
			this.dtFechaIngreso.Value = ((object)(resources.GetObject("dtFechaIngreso.Value")));
			this.dtFechaIngreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaIngreso_KeyDown);
			// 
			// dtFechaRetorno
			// 
			this.dtFechaRetorno.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaRetorno.DateButtons.Add(dateButton3);
			this.dtFechaRetorno.Enabled = false;
			this.dtFechaRetorno.Format = "dd/MM/yyyy";
			this.dtFechaRetorno.Location = new System.Drawing.Point(72, 376);
			this.dtFechaRetorno.Name = "dtFechaRetorno";
			this.dtFechaRetorno.NonAutoSizeHeight = 23;
			this.dtFechaRetorno.Size = new System.Drawing.Size(104, 21);
			this.dtFechaRetorno.SpinButtonsVisible = true;
			this.dtFechaRetorno.TabIndex = 666;
			this.dtFechaRetorno.Value = ((object)(resources.GetObject("dtFechaRetorno.Value")));
			this.dtFechaRetorno.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFeRetorno_KeyDown);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 378);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 17);
			this.label3.TabIndex = 667;
			this.label3.Text = "Retorno";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(376, 307);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(44, 17);
			this.label6.TabIndex = 673;
			this.label6.Text = "Informe";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtInforme
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInforme.Appearance = appearance14;
			this.txtInforme.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtInforme.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInforme.Enabled = false;
			this.txtInforme.Location = new System.Drawing.Point(376, 328);
			this.txtInforme.Multiline = true;
			this.txtInforme.Name = "txtInforme";
			this.txtInforme.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtInforme.Size = new System.Drawing.Size(680, 44);
			this.txtInforme.TabIndex = 672;
			this.txtInforme.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtInforme_KeyDown);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.ImageIndex = 8;
			this.btnNuevo.ImageList = this.imageList1;
			this.btnNuevo.Location = new System.Drawing.Point(8, 513);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(77, 23);
			this.btnNuevo.TabIndex = 677;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnNuevo, "Nuevo informe..");
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.CausesValidation = false;
			this.btnGuardar.Enabled = false;
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.ImageIndex = 9;
			this.btnGuardar.ImageList = this.imageList1;
			this.btnGuardar.Location = new System.Drawing.Point(88, 513);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(77, 23);
			this.btnGuardar.TabIndex = 675;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnGuardar, "Grabar informe..");
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// ugInformesDT
			// 
			this.ugInformesDT.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugInformesDT.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugInformesDT.DataSource = this.ultraDataSource2;
			appearance15.BackColor = System.Drawing.Color.White;
			this.ugInformesDT.DisplayLayout.Appearance = appearance15;
			this.ugInformesDT.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Width = 100;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 559;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.ugInformesDT.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugInformesDT.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.ugInformesDT.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugInformesDT.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugInformesDT.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.ugInformesDT.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugInformesDT.DisplayLayout.Override.GroupByRowAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 8.5F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugInformesDT.DisplayLayout.Override.HeaderAppearance = appearance19;
			this.ugInformesDT.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance20.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance20.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugInformesDT.DisplayLayout.Override.RowAlternateAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugInformesDT.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugInformesDT.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.ugInformesDT.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugInformesDT.Location = new System.Drawing.Point(376, 376);
			this.ugInformesDT.Name = "ugInformesDT";
			this.ugInformesDT.Size = new System.Drawing.Size(680, 161);
			this.ugInformesDT.TabIndex = 676;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn25.DataType = typeof(System.DateTime);
			ultraDataColumn25.DefaultValue = new System.DateTime(((long)(0)));
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtDaño);
			this.groupBox1.Location = new System.Drawing.Point(488, 216);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(568, 80);
			this.groupBox1.TabIndex = 678;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Daño";
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(8, 403);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(39, 16);
			this.label8.TabIndex = 695;
			this.label8.Text = "Estado";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.ImageIndex = 6;
			this.btnImprimir.ImageList = this.imageList1;
			this.btnImprimir.Location = new System.Drawing.Point(168, 513);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 698;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnImprimir, "Reeimprimir N/E Proveedores...");
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// cmbEstados
			// 
			this.cmbEstados.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DataSource = this.ultraDataSource3;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 300;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEstados.DisplayMember = "Nombre";
			this.cmbEstados.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstados.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstados.Enabled = false;
			this.cmbEstados.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstados.Location = new System.Drawing.Point(72, 352);
			this.cmbEstados.MaxDropDownItems = 30;
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(250, 21);
			this.cmbEstados.TabIndex = 697;
			this.cmbEstados.ValueMember = "idEstado";
			this.cmbEstados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstados_KeyDown);
			this.cmbEstados.ValueChanged += new System.EventHandler(this.cmbEstados_ValueChanged);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn27.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn27,
																																 ultraDataColumn28});
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(392, 11);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 705;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 36);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1064, 8);
			this.groupBox2.TabIndex = 706;
			this.groupBox2.TabStop = false;
			// 
			// txtBuscar
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance23;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtBuscar.Location = new System.Drawing.Point(64, 8);
			this.txtBuscar.MaxLength = 100;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(200, 22);
			this.txtBuscar.TabIndex = 0;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.ImageIndex = 10;
			this.btnCancelar.ImageList = this.imageList1;
			this.btnCancelar.Location = new System.Drawing.Point(248, 513);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 708;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// mnuImprimir
			// 
			this.mnuImprimir.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuImprimirNE,
																																								this.mnuImprimirEgreso});
			// 
			// mnuImprimirNE
			// 
			this.mnuImprimirNE.Index = 0;
			this.mnuImprimirNE.Text = "Nota de Entrega";
			this.mnuImprimirNE.Click += new System.EventHandler(this.mnuImprimirNE_Click);
			// 
			// mnuImprimirEgreso
			// 
			this.mnuImprimirEgreso.Index = 1;
			this.mnuImprimirEgreso.Text = "Autorizacion de Nota de Credito";
			this.mnuImprimirEgreso.Visible = false;
			this.mnuImprimirEgreso.Click += new System.EventHandler(this.mnuImprimirEgreso_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(8, 403);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(83, 17);
			this.label4.TabIndex = 710;
			this.label4.Text = "Bodega Destino";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 180;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(104, 400);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(256, 22);
			this.cmbBodega.TabIndex = 709;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// label72
			// 
			this.label72.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label72.AutoSize = true;
			this.label72.BackColor = System.Drawing.Color.Transparent;
			this.label72.Location = new System.Drawing.Point(8, 440);
			this.label72.Name = "label72";
			this.label72.Size = new System.Drawing.Size(89, 17);
			this.label72.TabIndex = 712;
			this.label72.Text = "Ingrese la Serial:";
			this.label72.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodigo
			// 
			this.txtCodigo.AcceptsTab = true;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance24;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(104, 464);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(256, 22);
			this.txtCodigo.TabIndex = 713;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 472);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(42, 17);
			this.label5.TabIndex = 714;
			this.label5.Text = "Articulo";
			// 
			// txtIdSerial
			// 
			this.txtIdSerial.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSerial.Enabled = false;
			this.txtIdSerial.FormatString = "#,##0.00";
			this.txtIdSerial.Location = new System.Drawing.Point(296, 312);
			this.txtIdSerial.Name = "txtIdSerial";
			this.txtIdSerial.PromptChar = ' ';
			this.txtIdSerial.Size = new System.Drawing.Size(16, 22);
			this.txtIdSerial.TabIndex = 716;
			this.txtIdSerial.Visible = false;
			// 
			// txtIdArticulo
			// 
			this.txtIdArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdArticulo.Enabled = false;
			this.txtIdArticulo.FormatString = "#,##0.00";
			this.txtIdArticulo.Location = new System.Drawing.Point(312, 312);
			this.txtIdArticulo.Name = "txtIdArticulo";
			this.txtIdArticulo.PromptChar = ' ';
			this.txtIdArticulo.Size = new System.Drawing.Size(16, 22);
			this.txtIdArticulo.TabIndex = 715;
			this.txtIdArticulo.Visible = false;
			// 
			// txtSerial
			// 
			this.txtSerial.AcceptsTab = true;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerial.Appearance = appearance25;
			this.txtSerial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerial.Enabled = false;
			this.txtSerial.Location = new System.Drawing.Point(104, 432);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(256, 22);
			this.txtSerial.TabIndex = 717;
			this.txtSerial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSerial_KeyDown);
			this.txtSerial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerial_KeyPress);
			// 
			// frmDTSeguimiento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1064, 543);
			this.Controls.Add(this.txtSerial);
			this.Controls.Add(this.txtIdSerial);
			this.Controls.Add(this.txtIdArticulo);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label72);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.ugInformesDT);
			this.Controls.Add(this.txtInforme);
			this.Controls.Add(this.dtFechaRetorno);
			this.Controls.Add(this.gbProducto);
			this.Controls.Add(this.ugIngresosDT);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.txtIngreso);
			this.Controls.Add(this.dtFechaIngreso);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmDTSeguimiento";
			this.Text = "Seguimiento Articulos en RMA";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDTSeguimiento_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmDTSeguimiento_Closing);
			this.Load += new System.EventHandler(this.frmDTSeguimiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.ugIngresosDT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.gbProducto.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDaño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRetorno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInforme)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugInformesDT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		bool bEdicion = false;
		int idIngreso = 0;
		int idPersonal = 0;
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void Consulta ()
		{
			string sSQL = string.Format("Exec DT_BuscaNEProveedores '{0}'", this.txtBuscar.Text.ToString().Trim());
			FuncionesProcedimientos.dsGeneral(sSQL, this.ugIngresosDT);
			
			this.lblContador.Text = this.ugIngresosDT.Rows.Count + " REGISTROS ENCONTRADOS";
		}


		



		private void ListaInformes (int idIngreso)
		{		
			string sSQL = string.Format("Exec DT_ListaInformesRMA '{0}'", idIngreso);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ugInformesDT);
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			this.Consulta();
		}

		private void frmDTSeguimiento_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "101502");

			if (!Funcion.Permiso("749",cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar a Impresión de Precios", "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));
			}
			
			this.txtIdSerial.Value = 0;
			this.txtIdArticulo.Value = 0;

			this.Consulta();
			
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTListaBodegasSalida 0");

			idPersonal = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario Where idUsuario = {0})) Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario Where idUsuario = {0}) Else Select 0", MenuLatinium.IdUsuario));

			this.dtFechaIngreso.CalendarInfo.MaxDate = DateTime.Today;
			this.dtFechaRetorno.CalendarInfo.MaxDate = DateTime.Today;

			
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnBuscar_Click(sender, e);
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{	
			this.txtBuscar.Enabled = false;
			this.btnBuscar.Enabled = false;
			this.ugIngresosDT.Enabled = false;
			
			if (this.txtIngreso.Text.ToString().Length == 0) 
			{
				this.txtIngreso.Enabled = true;
				this.dtFechaIngreso.Enabled = true;

				this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTListaEstados 9");

				this.cmbEstados.Value = 16;
			}
			else 
			{
				if ((int)this.ugIngresosDT.ActiveRow.Cells["PointCliente"].Value == 1) this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTListaEstados 10");
				else this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTListaEstados 12");

				this.cmbEstados.Enabled = true;
			}

			this.txtInforme.Enabled = true;
			
			this.btnNuevo.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;			

			bEdicion = true;

			if (this.txtIngreso.Text.ToString().Length == 0) this.txtIngreso.Focus(); else this.cmbEstados.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			
			if (!Validacion.vbTexto(this.txtIngreso, 5, 20, "Numero de Ingreso de Proveedor")) return;
			if ((int)this.cmbEstados.Value == 16 && !Validacion.vbFechaEnDocumentos(this.dtFechaIngreso, "Seleccione la Fecha de Ingreso", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbEstados, "Seleccione el estado")) return;
			if (!Validacion.vbTexto(this.txtInforme, 20, 1000, "Informe")) return;
			
			DateTime dtRetorno = DateTime.Parse("1900/01/01");
			if (this.dtFechaRetorno.Value != System.DBNull.Value) dtRetorno = (DateTime)this.dtFechaRetorno.Value;

			string sSQL = string.Format("Exec GuardaDT_Procesos {0}, 20, {1}, {2}, '{3}', 0, '', 0, 0, 0, 0, '{4}', '{5}', 0,'{6}','{7}'", 
				idIngreso, (int)this.cmbEstados.Value, idPersonal, this.txtInforme.Text.ToString().Trim(), this.txtIngreso.Text.ToString().Trim(), dtRetorno.ToString("yyyyMMdd"),this.txtSerial.Text.ToString(),this.txtCodigo.Text.ToString());
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			if ((int)this.cmbEstados.Value == 34 )
			{
				if (!Validacion.vbTexto(this.txtSerial, 1, 1000, "Serie")) return;
				#region Serial Si Existe
				/*
				if (Convert.ToInt32(this.txtIdSerial.Value) > 0)
					
				{
					int iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ISMValidaSerialArticulo {0}, {1}", Convert.ToInt32(this.txtIdSerial.Value), Convert.ToInt32(this.txtIdArticulo.Value)));
				
					if (iCont == 0 && DialogResult.No == MessageBox.Show("CAMBIO DE ARTICULO,\n\nSi registra esta serial con este articulo se modificara el historial de la serial a este Articulo,\n\n¿Desea Continuar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

					if (DialogResult.Yes == MessageBox.Show("Esta seguro de registar un Movimiento  de esta Serial", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
					{					
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleSeriales Set Disponible = 0 Where idSerial = {0}", Convert.ToInt32(this.txtIdSerial.Value)));
					
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Seriales Set Estado = 1 Where idSerial = {0}", Convert.ToInt32(this.txtIdSerial.Value)));

						if (iCont == 0) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Seriales Set idArticulo = {0} Where idSerial = {1}", (int)this.txtIdArticulo.Value, Convert.ToInt32(this.txtIdSerial.Value)));

						string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, 0, 0, 20, '', 1, 'CAMBIO DE PRODUCTO REGISTRADO ', 100", Convert.ToInt32(this.txtIdSerial.Value));
						Funcion.EjecutaSQL(cdsSeteoF, sSQLGDetalleSeriales);

						MessageBox.Show("Movimiento Registrado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
					}
				
				}
				else
				{
					string sSQLGSeriales = string.Format("Exec GuardaSeriales {0}, {1}, '{2}', 1, 20",Convert.ToInt32(this.txtIdSerial.Value),Convert.ToInt32(this.txtIdArticulo.Value),this.txtSerial.Text.ToString());
					Funcion.EjecutaSQL(cdsSeteoF, sSQLGSeriales);

					string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, 0, 0, 20, '', 1, 'CAMBIO DE PRODUCTO REGISTRADO ', 100", Convert.ToInt32(this.txtIdSerial.Value));
					Funcion.EjecutaSQL(cdsSeteoF, sSQLGDetalleSeriales);

				}*/

				#endregion Serial Si Existe

				int iSerie= 0;
				iSerie = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Exec [ValidaSerieRMA] '{0}' ", this.txtSerial.Text));
				if(iSerie > 0)
				{
					MessageBox.Show("La Serie ingresada ya Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerial.Text ="";
				
				}
				else
				{
					string sSQLGSeriales = string.Format("Exec GuardaSeriales {0}, {1}, '{2}', 1, 20",0,Convert.ToInt32(this.txtIdArticulo.Value),this.txtSerial.Text.ToString());
					this.txtIdSerial.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLGSeriales);

					string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, 0, 0, 20, '', 1, 'CAMBIO DE PRODUCTO REGISTRADO ', 19", Convert.ToInt32(this.txtIdSerial.Value));
					Funcion.EjecutaSQL(cdsSeteoF, sSQLGDetalleSeriales);
					MessageBox.Show("Movimiento Registrado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			if (this.cmbBodega.ActiveRow != null)
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DTIngreso Set BodegaDestino = {0} Where idIngreso = {1}", (int)this.cmbBodega.Value, idIngreso));

			this.btnCancelar_Click(sender, e);
			this.txtSerial.Text="";
			this.txtSerial.Enabled=false;
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.mnuImprimir.Show(this.btnImprimir, new Point(1, 25));
		}

		private void txtFeRetorno_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtInforme.Focus();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			#region Controles
			this.txtBuscar.Enabled = true;
			this.btnBuscar.Enabled = true;
			this.ugIngresosDT.Enabled = true;
			
			bEdicion = false;
			idIngreso = 0;

			this.txtRuc.Text = "";
			this.txtFactura.Text = "";
			this.dtFechaCompra.Value = System.DBNull.Value;
			this.txtProveedor.Text = "";
			this.txtDaño.Text = "";

			this.txtIngreso.Text = "";
			this.dtFechaIngreso.Value = System.DBNull.Value;			
			this.cmbEstados.Value = System.DBNull.Value;
			this.dtFechaRetorno.Value = System.DBNull.Value;
			this.cmbBodega.Value = System.DBNull.Value;	
			this.txtInforme.Text = "";
			this.txtSerial.Text = "";
			
			this.txtIngreso.Enabled = false;
			this.dtFechaIngreso.Enabled = false;
			this.cmbEstados.Enabled = false;
			this.dtFechaRetorno.Enabled = false;			
			this.cmbBodega.Enabled = false;
			this.txtInforme.Enabled = false;
			this.txtSerial.Enabled=false;

			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;		

			this.Consulta();
			this.ListaInformes(0);
			#endregion Controles
		}

		private void mnuImprimirNE_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue idNEDt = new ParameterDiscreteValue ();
			pfRet.ParameterFieldName = "@idIngreso";
			idNEDt.Value = (int)this.ugIngresosDT.ActiveRow.Cells["idIngreso"].Value;
			pfRet.CurrentValues.Add(idNEDt);
			paramFields.Add (pfRet);
				
			Reporte Reporte = new Reporte("NotaEntregaProveedor.rpt", "");
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Ingreso DT");
			Reporte.crVista.ParameterFieldInfo = paramFields;			
			Reporte.Show();
		}

		private void mnuImprimirEgreso_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue idIngreso = new ParameterDiscreteValue ();
			pfRet.ParameterFieldName = "@BuskIngrs";
			idIngreso.Value = (int)this.ugIngresosDT.ActiveRow.Cells["idIngreso"].Value;
			pfRet.CurrentValues.Add (idIngreso);
			paramFields.Add (pfRet);
				
			Reporte Reporte = new Reporte("AutorizacionNC.rpt", "");
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Ingreso DT");
			Reporte.crVista.ParameterFieldInfo = paramFields;			
			Reporte.Show();
		}

		private void ugIngresosDT_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{	
			idIngreso = 0;		
			this.txtIngreso.Text = "";
			this.dtFechaIngreso.Value = System.DBNull.Value;
			this.dtFechaRetorno.Value = System.DBNull.Value;
			this.cmbEstados.Value = System.DBNull.Value;			
			this.txtInforme.Text = "";
			
			idIngreso = (int)this.ugIngresosDT.ActiveRow.Cells["idIngreso"].Value;
			this.txtRuc.Text = this.ugIngresosDT.ActiveRow.Cells["Ruc"].Value.ToString().Trim();
			this.txtProveedor.Text = this.ugIngresosDT.ActiveRow.Cells["Proveedor"].Value.ToString().Trim();
			this.txtFactura.Value = this.ugIngresosDT.ActiveRow.Cells["NumeroCompra"].Value.ToString().Trim();
			this.dtFechaCompra.Value = (DateTime)this.ugIngresosDT.ActiveRow.Cells["FechaCompra"].Value;
			this.txtDaño.Value = this.ugIngresosDT.ActiveRow.Cells["Daño"].Value;
			this.txtCodigo.Text = this.ugIngresosDT.ActiveRow.Cells["Articulo"].Value.ToString().Trim();
			this.txtIdArticulo.Value = this.ugIngresosDT.ActiveRow.Cells["idArticulo"].Value;
			this.txtIdSerial.Value = this.ugIngresosDT.ActiveRow.Cells["idserial"].Value;



			this.txtIngreso.Value = this.ugIngresosDT.ActiveRow.Cells["Numero"].Value.ToString().Trim();			
			if (this.ugIngresosDT.ActiveRow.Cells["FechaIngreso"].Value != System.DBNull.Value) this.dtFechaIngreso.Value = Convert.ToDateTime(this.ugIngresosDT.ActiveRow.Cells["FechaIngreso"].Value);
			//if (this.ugIngresosDT.ActiveRow.Cells["FechaRespuesta"].Value != System.DBNull.Value) this.dtFechaRetorno.Value = (DateTime)this.ugIngresosDT.ActiveRow.Cells["FechaRespuesta"].Value;			
									
			ListaInformes(idIngreso);	
	
			this.btnImprimir.Enabled = true;
			this.btnNuevo.Enabled = true;
		}

		private void txtIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFechaIngreso.Focus();
		}

		private void dtFechaIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.cmbEstados.Enabled) this.cmbEstados.Focus(); else this.txtInforme.Focus();
		}

		private void cmbEstados_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.dtFechaRetorno.Enabled) this.dtFechaRetorno.Focus(); else this.txtInforme.Focus();
		}

		private void frmDTSeguimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);				
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

		private void frmDTSeguimiento_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una informe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void txtIngreso_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtInforme_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void cmbEstados_ValueChanged(object sender, System.EventArgs e)
		{

			
			if (!bEdicion) return;

			if (this.cmbEstados.ActiveRow == null) return;

			this.dtFechaRetorno.Enabled = false;
			this.cmbBodega.Enabled = false;

			if ((int)this.cmbEstados.Value == 19 || (int)this.cmbEstados.Value == 20 || (int)this.cmbEstados.Value == 21)
				{
			
			this.txtSerial.Enabled=false;
			}
		
			if ((int)this.cmbEstados.Value == 19 || (int)this.cmbEstados.Value == 20)
			{
				this.cmbBodega.Enabled = true; 	
				this.txtSerial.Enabled=false;
	
				this.cmbBodega.Value = (int)this.ugIngresosDT.ActiveRow.Cells["Bodega"].Value;
			}
			if ((int)this.cmbEstados.Value == 34 )
			{
				if (miAcceso.BManejaSerial) 
					
				{
					this.txtSerial.Enabled = true;
				}

				else
				{

					this.txtSerial.Enabled=false;
					this.txtCodigo.Enabled=false;
				}
			}
		}

		private void ugIngresosDT_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

	

	
		private void txtSerial_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtProveedor_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			#region  Serial no existe
			if (Convert.ToInt32(this.txtIdSerial.Value) == 0)
			{
				
				if (Convert.ToInt32(this.txtIdArticulo.Value) == 0)
				{
					MessageBox.Show("El Articulo Ingresado No Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCodigo.Focus();
					return;
				}
				
			}
			#endregion  Serial no existe

			#region Serial Si Existe
			if (Convert.ToInt32(this.txtIdSerial.Value) > 0)
			{
				if (Convert.ToInt32(this.txtIdArticulo.Value) == 0)
				{
					MessageBox.Show("Borro el Articulo, click en el Boton Buscar para Actualizar la informacion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					this.btnBuscar_Click(sender, e);
					return;				
				}
				/*
				int iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ISMValidaSerialArticulo {0}, {1}", Convert.ToInt32(this.txtIdSerial.Value), Convert.ToInt32(this.txtIdArticulo.Value)));
				
				if (iCont == 0 && DialogResult.No == MessageBox.Show("CAMBIO DE ARTICULO,\n\nSi registra esta serial con este articulo se modificara el historial de la serial a este Articulo,\n\n¿Desea Continuar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
				*/	
			}
			#endregion Serial Si Existe
		}

		private void txtSerial_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
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
		}

	
	}
}
