using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmConstanciaFisica.
	/// </summary>
	public class frmConstanciaFisica : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblVendedor;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaCierre;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdPersonal;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.GroupBox gbBotones;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbInventariador;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdSeries;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidConstanciaFisica;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbConstancia;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstado;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Label lblContador2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Button btnVer;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConstanciaFisica()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConstanciaFisica));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonalConstanciaFisica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idConstanciaFisica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Personal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonalConstanciaFisica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idConstanciaFisica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Personal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleConstanciaFisica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Num");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Ultimo_Movimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Constancia", -1, "cmbConstancia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado", -1, "cmbEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idConstanciaFisica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SConstancia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEstado");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleConstanciaFisica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Num");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Modelo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Color");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Ultimo_Movimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Constancia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idConstanciaFisica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SConstancia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEstado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Constancia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sConstancia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Constancia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sConstancia");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sEstado");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			this.label6 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbInventariador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFechaCierre = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.grdPersonal = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.grdSeries = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtidConstanciaFisica = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbConstancia = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExportar = new System.Windows.Forms.Button();
			this.lblContador2 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbInventariador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaCierre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdSeries)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidConstanciaFisica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbConstancia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 24);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 171;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 56);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(48, 16);
			this.lblBodega.TabIndex = 170;
			this.lblBodega.Text = "Almacen";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFecha.Location = new System.Drawing.Point(80, 24);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(160, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 168;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// cmbBodega
			// 
			appearance2.FontData.Name = "Tahoma";
			this.cmbBodega.Appearance = appearance2;
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(80, 56);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 22);
			this.cmbBodega.TabIndex = 169;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(8, 88);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(70, 16);
			this.lblVendedor.TabIndex = 173;
			this.lblVendedor.Text = "Inventariador";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbInventariador
			// 
			appearance3.FontData.Name = "Tahoma";
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbInventariador.Appearance = appearance3;
			this.cmbInventariador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbInventariador.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbInventariador.DataSource = this.ultraDataSource2;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.UseEditorMaskSettings = true;
			ultraGridColumn4.Width = 198;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 198;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbInventariador.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbInventariador.DisplayMember = "Nombre";
			this.cmbInventariador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbInventariador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbInventariador.Enabled = false;
			this.cmbInventariador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbInventariador.Location = new System.Drawing.Point(80, 88);
			this.cmbInventariador.Name = "cmbInventariador";
			this.cmbInventariador.Size = new System.Drawing.Size(200, 22);
			this.cmbInventariador.TabIndex = 172;
			this.cmbInventariador.ValueMember = "idPersonal";
			this.cmbInventariador.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbInventariador_KeyDown);
			this.cmbInventariador.ValueChanged += new System.EventHandler(this.cmbInventariador_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 121);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(70, 16);
			this.label1.TabIndex = 175;
			this.label1.Text = "Fecha Cierre";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaCierre
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaCierre.Appearance = appearance4;
			this.dtFechaCierre.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaCierre.DateButtons.Add(dateButton2);
			this.dtFechaCierre.Enabled = false;
			this.dtFechaCierre.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFechaCierre.Location = new System.Drawing.Point(80, 119);
			this.dtFechaCierre.Name = "dtFechaCierre";
			this.dtFechaCierre.NonAutoSizeHeight = 23;
			this.dtFechaCierre.Size = new System.Drawing.Size(160, 21);
			this.dtFechaCierre.SpinButtonsVisible = true;
			this.dtFechaCierre.TabIndex = 174;
			this.dtFechaCierre.Value = ((object)(resources.GetObject("dtFechaCierre.Value")));
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.BackColor = System.Drawing.SystemColors.Control;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(296, 88);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 16);
			this.lblEstado.TabIndex = 274;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grdPersonal
			// 
			this.grdPersonal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.grdPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdPersonal.DataSource = this.ultraDataSource1;
			appearance5.BackColor = System.Drawing.Color.White;
			this.grdPersonal.DisplayLayout.Appearance = appearance5;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 250;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Width = 120;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.grdPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdPersonal.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.grdPersonal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.grdPersonal.DisplayLayout.Override.CardAreaAppearance = appearance7;
			this.grdPersonal.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 9F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdPersonal.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdPersonal.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdPersonal.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdPersonal.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.grdPersonal.Enabled = false;
			this.grdPersonal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdPersonal.Location = new System.Drawing.Point(544, 16);
			this.grdPersonal.Name = "grdPersonal";
			this.grdPersonal.Size = new System.Drawing.Size(416, 136);
			this.grdPersonal.TabIndex = 275;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(-88, 180);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1200, 8);
			this.gbBotones.TabIndex = 276;
			this.gbBotones.TabStop = false;
			// 
			// grdSeries
			// 
			this.grdSeries.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdSeries.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdSeries.DataSource = this.ultraDataSource3;
			appearance12.BackColor = System.Drawing.Color.White;
			this.grdSeries.DisplayLayout.Appearance = appearance12;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "#";
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 35;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Width = 110;
			ultraGridColumn17.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 5;
			ultraGridColumn17.Width = 250;
			ultraGridColumn18.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 6;
			ultraGridColumn18.Width = 110;
			ultraGridColumn19.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 7;
			ultraGridColumn19.Width = 110;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 8;
			ultraGridColumn20.Width = 40;
			ultraGridColumn21.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 9;
			ultraGridColumn22.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 10;
			ultraGridColumn22.Width = 110;
			ultraGridColumn23.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.VisiblePosition = 11;
			ultraGridColumn23.Width = 150;
			ultraGridColumn24.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 12;
			ultraGridColumn24.Width = 151;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.Caption = "Fecha Ultimo Movimiento";
			ultraGridColumn25.Header.VisiblePosition = 13;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance13;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 14;
			ultraGridColumn27.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn27.Header.VisiblePosition = 15;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn27.Width = 100;
			ultraGridColumn28.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn28.Header.VisiblePosition = 16;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn28.Width = 100;
			ultraGridColumn29.Header.VisiblePosition = 17;
			ultraGridColumn29.Width = 300;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 18;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 19;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 20;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.VisiblePosition = 21;
			ultraGridColumn33.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33});
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grdSeries.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdSeries.DisplayLayout.Override.ActiveRowAppearance = appearance14;
			this.grdSeries.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grdSeries.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.grdSeries.DisplayLayout.Override.CardAreaAppearance = appearance15;
			this.grdSeries.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 9F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdSeries.DisplayLayout.Override.HeaderAppearance = appearance16;
			this.grdSeries.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSeries.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSeries.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSeries.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.grdSeries.Enabled = false;
			this.grdSeries.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdSeries.Location = new System.Drawing.Point(8, 200);
			this.grdSeries.Name = "grdSeries";
			this.grdSeries.Size = new System.Drawing.Size(952, 224);
			this.grdSeries.TabIndex = 277;
			this.grdSeries.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdSeries_AfterCellUpdate);
			this.grdSeries.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdSeries_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.DateTime);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30});
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(600, 440);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 278;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Visible = false;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(8, 440);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 279;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(88, 440);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 282;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(168, 440);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 281;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(248, 440);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 280;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(328, 440);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 283;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(408, 440);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 284;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-116, 424);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1200, 8);
			this.groupBox1.TabIndex = 285;
			this.groupBox1.TabStop = false;
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
			// txtidConstanciaFisica
			// 
			this.txtidConstanciaFisica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidConstanciaFisica.Enabled = false;
			this.txtidConstanciaFisica.Location = new System.Drawing.Point(48, 0);
			this.txtidConstanciaFisica.Name = "txtidConstanciaFisica";
			this.txtidConstanciaFisica.PromptChar = ' ';
			this.txtidConstanciaFisica.Size = new System.Drawing.Size(40, 21);
			this.txtidConstanciaFisica.TabIndex = 851;
			this.txtidConstanciaFisica.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(88, 0);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(40, 21);
			this.txtEstado.TabIndex = 852;
			this.txtEstado.Visible = false;
			// 
			// cmbConstancia
			// 
			this.cmbConstancia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbConstancia.DataSource = this.ultraDataSource4;
			ultraGridColumn34.Header.VisiblePosition = 0;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 1;
			ultraGridColumn35.Width = 250;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn34,
																										 ultraGridColumn35});
			this.cmbConstancia.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbConstancia.DisplayMember = "sConstancia";
			this.cmbConstancia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbConstancia.Location = new System.Drawing.Point(456, 248);
			this.cmbConstancia.Name = "cmbConstancia";
			this.cmbConstancia.Size = new System.Drawing.Size(272, 88);
			this.cmbConstancia.TabIndex = 981;
			this.cmbConstancia.ValueMember = "Constancia";
			this.cmbConstancia.Visible = false;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn31.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// cmbEstado
			// 
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.ultraDataSource5;
			ultraGridColumn36.Header.VisiblePosition = 0;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 1;
			ultraGridColumn37.Width = 250;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbEstado.DisplayMember = "sEstado";
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(168, 248);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(272, 88);
			this.cmbEstado.TabIndex = 982;
			this.cmbEstado.ValueMember = "Estado";
			this.cmbEstado.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn33.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn33,
																																 ultraDataColumn34});
			// 
			// btnExportar
			// 
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(448, 152);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 23);
			this.btnExportar.TabIndex = 983;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// lblContador2
			// 
			this.lblContador2.AutoSize = true;
			this.lblContador2.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador2.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador2.Location = new System.Drawing.Point(256, 120);
			this.lblContador2.Name = "lblContador2";
			this.lblContador2.Size = new System.Drawing.Size(0, 19);
			this.lblContador2.TabIndex = 984;
			this.lblContador2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance20;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(80, 152);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(264, 21);
			this.txtNombre.TabIndex = 985;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombre_KeyDown);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Enabled = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(360, 152);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(78, 23);
			this.btnVer.TabIndex = 986;
			this.btnVer.Text = "Consultar";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// frmConstanciaFisica
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(968, 470);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.lblContador2);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.cmbConstancia);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtidConstanciaFisica);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.grdSeries);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.grdPersonal);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.dtFechaCierre);
			this.Controls.Add(this.cmbInventariador);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbBodega);
			this.KeyPreview = true;
			this.Name = "frmConstanciaFisica";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Constancia Fisica";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConstanciaFisica_KeyDown);
			this.Load += new System.EventHandler(this.frmConstanciaFisica_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbInventariador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaCierre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdSeries)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidConstanciaFisica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbConstancia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();			
		}
		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmConstanciaFisica_Load(object sender, System.EventArgs e)
		{

			miAcceso = new Acceso(cdsSeteoF, "22");

			if (!Funcion.Permiso("1324", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a INVENTARIO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;


			RetornaFechaS(this.dtFecha, cdsSeteoF, true);
			this.cmbInventariador.DataSource  = Funcion.dvProcedimiento(cdsSeteoF, "Select n.idPersonal, (n.ApellidoPaterno + ' ' + n.ApellidoMaterno+' '+ n.PrimerNombre + ' ' + n.SegundoNombre) As Nombre From Nomina n Where  ISNULL((Select Top 1 FechaSalida From NominaEntrada Where idNomina = n.idNomina Order By FechaEntrada Desc), '20301231') > GETDATE () And ISNULL((Select Top 1 FechaEntrada From NominaEntrada Where idNomina = n.idNomina Order By FechaEntrada Desc), '20301231') <= GETDATE () Order By  Nombre");
			this.cmbConstancia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format(" Exec Listas {0},0", 70));
			this.cmbEstado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format(" Exec Listas {0},0", 71));
		}

		static public void RetornaFechaS(Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha, C1.Data.C1DataSet ds, bool bHora)
		{
			DateTime dtFechaServ = Funcion.dtEscalarSQL(ds, "Exec RetornaFechaServidor");

			if (!bHora) dtFecha.Value = dtFechaServ.Date;
			if (bHora) dtFecha.Value = dtFechaServ;

			//			dtFecha.CalendarInfo.MaxDate = dtFechaServ.Date;
		} 

		

		private void frmConstanciaFisica_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidConstanciaFisica.Value > 0)
				{
					//					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.F5) this.Consulta((int)this.txtidConstanciaFisica.Value);					
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

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;
      
			this.txtidConstanciaFisica.Value = 0;
			this.dtFecha.Value = System.DBNull.Value;
			this.cmbBodega.Value  = System.DBNull.Value;
			this.cmbInventariador.Value = System.DBNull.Value;
			this.dtFechaCierre.Value = System.DBNull.Value;
			this.lblEstado.Text = "";
			this.txtEstado.Value = 0;


			FuncionesProcedimientos.EstadoGrids(false,this.grdPersonal);
			FuncionesProcedimientos.EstadoGrids(false,this.grdSeries);
			ListaVendedores(this.grdPersonal,0,DateTime.Today, 0, cdsSeteoF);
			string sSQLPa = string.Format(" EXEC ConsultaGeneralPersonalConstanciaFisica   {0} ", 0);
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.grdPersonal);

			string sSQLPas = string.Format(" EXEC ConsultaGeneralDetalleConstanciaFisicaS   {0} ", 0);
			FuncionesProcedimientos.dsGeneral(sSQLPas, this.grdSeries);


			this.cmbBodega.Enabled = false;
			this.cmbInventariador.Enabled = false;
			this.dtFechaCierre.Enabled = false;
			this.btnVer.Enabled = false;
			this.txtNombre.Enabled = false;



			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnImprimir.Enabled = false;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{

			this.btnCancelar_Click(sender, e);
			this.txtidConstanciaFisica.Value = 0;
			RetornaFechaS(this.dtFecha, cdsSeteoF, true);
			this.lblEstado.Text = "PENDIENTE";
			this.txtEstado.Value = 1;
			this.cmbInventariador.Enabled = true;
			FuncionesProcedimientos.EstadoGrids(true,this.grdPersonal);
			FuncionesProcedimientos.EstadoGrids(true,this.grdSeries);
			this.cmbInventariador.DataSource  = Funcion.dvProcedimiento(cdsSeteoF, "Select n.idPersonal, (n.ApellidoPaterno + ' ' + n.ApellidoMaterno+' '+ n.PrimerNombre + ' ' + n.SegundoNombre) As Nombre From Nomina n Where  ISNULL((Select Top 1 FechaSalida From NominaEntrada Where idNomina = n.idNomina Order By FechaEntrada Desc), '20301231') > GETDATE () And ISNULL((Select Top 1 FechaEntrada From NominaEntrada Where idNomina = n.idNomina Order By FechaEntrada Desc), '20301231') <= GETDATE () Order By  Nombre");
			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			int iBodegaPredef  = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			if (iBodegaPredef == 0) 
			{
				this.cmbBodega.Enabled = true;
				this.cmbBodega.Value = 1; 				
			}
			else this.cmbBodega.Value = iBodegaPredef;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;					
			bNuevo = true;
			bEdicion = true;
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
			{
				
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 42, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
		
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null)
			{
				if (this.dtFecha.Value != System.DBNull.Value)
				{
					bNuevo = true ;
						
					//					int iEstado =  Funcion.iEscalarSQL(cdsSeteoF, string.Format(" exec ListaPersonalConstanciaFisica {0}, '{1}', {2}", (int)this.cmbBodega.Value,"20010101", 1));
					//					string sBodega =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Select Nombre from Bodega where bodega = {0} ", (int)this.cmbBodega.Value));
					//					if (iEstado == 0 )
					//					{
					ListaVendedores(this.grdPersonal,(int)this.cmbBodega.Value,(DateTime)this.dtFecha.Value, 0, cdsSeteoF);
					//					}
					//					if (iEstado == 1 )	 
					//					{
					//						MessageBox.Show(string.Format(" Ya existe un inventario para la Bodega '{0}', en un estado PENDIENTE ", sBodega), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
					//	           return;
					//						ListaVendedores(this.grdPersonal,(int)this.cmbBodega.Value,(DateTime)this.dtFecha.Value, 0, cdsSeteoF);
					//					}
					//					if (iEstado == 2 )
					//					{
					//						MessageBox.Show(string.Format(" Ya existe un inventario para la Bodega '{0}', en un estado TERMINADO ", sBodega), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
					//						return;
					//					}
     
				}
			}
		}

		public void ListaVendedores (Infragistics.Win.UltraWinGrid.UltraGrid uCombo,int Bodega,   DateTime dtDesde, int idPersonalConstanciaFisica, C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec ListaPersonalConstanciaFisica {0}, '{1}',{2}", Bodega,  dtDesde.ToString("yyyyMMdd"),0));
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la Fecha De Salida", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un Almacen")) return;
			if (!Validacion.vbComboVacio(this.cmbInventariador, "Seleccione un Inventariador")) return;
			DateTime dtfechacierre = DateTime.Parse("01/01/2000");
			if (this.dtFechaCierre.Value != System.DBNull.Value)dtfechacierre = (DateTime)this.dtFechaCierre.Value;
			#endregion Validacion

			

			if (DialogResult.Yes == MessageBox.Show("¿Desea terminar el Acta Fisica ?, Si presiona SI ya no podra modificar ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
					
			{
				#region Valida Grid
				if ((int) this.grdPersonal.Rows.Count == 0)
				{
					MessageBox.Show("Ingrese el Personal del Almacen", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.grdPersonal.Focus ();
					return;
				}

				if ((int) this.grdSeries.Rows.Count == 0)
				{
					MessageBox.Show("Ingrese el Inventario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.grdSeries.Focus ();
					return;
				}

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSeries.Rows.All)
				{	

//					if (dr.Cells["Constancia"].Value == System.DBNull.Value || Convert.ToInt32( dr.Cells["Constancia"].Value) == 0 )
//					{
//						MessageBox.Show("Seleccione el estado de la Constancia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						this.grdSeries.ActiveRow = dr;
//						this.grdSeries.ActiveRow.Selected = true;
//						return;
//					}	
				
//					if ((int)dr.Cells["Constancia"].Value ==1 || (int)dr.Cells["Constancia"].Value ==3)
//					{
//						if (dr.Cells["Estado"].Value == System.DBNull.Value || Convert.ToInt32( dr.Cells["Estado"].Value) == 0 )
//						{
//							MessageBox.Show("Seleccione el Estado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							this.grdSeries.ActiveRow = dr;
//							this.grdSeries.ActiveRow.Selected = true;
//							return;
//						}	
//					
//					}	
					//					if (dr.Cells["Observaciones"].Value != System.DBNull.Value)
					//					{
					//						if (dr.Cells["Observaciones"].Value.ToString().Length > 0)
					//						{
					//							if (dr.Cells["Observaciones"].Value.ToString().Trim().Length < 20 )
					//							{
					//								MessageBox.Show("El Detalle debe tener minimo 20 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//								this.grdSeries.ActiveRow = dr;
					//								this.grdSeries.ActiveRow.Selected = true;
					//								return;
					//							}
					//							if (dr.Cells["Observaciones"].Value.ToString().Trim().Length > 200 )
					//							{
					//								MessageBox.Show("El Detalle debe tener maximo 200 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//								this.grdSeries.ActiveRow = dr;
					//								this.grdSeries.ActiveRow.Selected = true;
					//								return;
					//							}		
					//						}	
					//					}
				}
				#endregion Valida Grid

				this.txtEstado.Value = 2; 
				RetornaFechaS(this.dtFechaCierre, cdsSeteoF, true);
				dtfechacierre = (DateTime) this.dtFechaCierre.Value;
			}
			#region Guardar      
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

					#region Graba Maestro
					string  sSQL1 = string.Format("Exec [GrabaConstanciaFisica] {0}, '{1}', {2}, {3}, '{4}', {5}" ,
						(int)this.txtidConstanciaFisica.Value, 
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm:ss"), 
						(int)this.cmbInventariador.Value,
						(int)this.cmbBodega.Value,
						Convert.ToDateTime(dtfechacierre).ToString("yyyyMMdd HH:mm:ss"),
						(int) this.txtEstado.Value);
					oCmdActualiza.CommandText = sSQL1;
					this.txtidConstanciaFisica.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Graba Maestro
        
					#region Graba DetalleMaestro
					if((int) this.grdPersonal.Rows.Count > 0)
					{						
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdPersonal.Rows.All)
						{
							string  sSQLDet = string.Format("Exec GrabaPersonalConstanciaFisica {0}, {1}, {2}, {3}" ,
								Convert.ToInt32 ( dr.Cells["idPersonalConstanciaFisica"].Value), 
								(int) this.txtidConstanciaFisica.Value, 
								Convert.ToInt32 ( dr.Cells["idPersonal"].Value),
								Convert.ToInt32 ( dr.Cells["idCargo"].Value)
								);
							oCmdActualiza.CommandText = sSQLDet;
							oCmdActualiza.ExecuteNonQuery();	
						}						
					}
					#endregion Graba DetalleMaestro

					#region Graba DetalleMaestroSeriales
					if((int) this.grdSeries.Rows.Count > 0)
					{						
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drs in this.grdSeries.Rows.All)
						{
							string  sSQLSer = string.Format("Exec GuardaDetalleConstanciaFisica {0}, {1}, {2}, {3}, {4}, '{5}', '{6}', '{7}', '{8}', {9}, '{10}', '{11}', '{12}', '{13}', '{14}', {15}, {16}, {17}, '{18}', {19}" ,
								Convert.ToInt32 ( drs.Cells["idDetalleConstanciaFisica"].Value), /*0*/
								(int) this.txtidConstanciaFisica.Value, /*1*/
								Convert.ToInt32 ( drs.Cells["Num"].Value), /*2*/
								Convert.ToInt32 ( drs.Cells["idSerial"].Value), /*3*/
								Convert.ToInt32 ( drs.Cells["idDetalleSerial"].Value), /*4*/
								Convert.ToString  ( drs.Cells["Codigo"].Value), /*5*/
								Convert.ToString  ( drs.Cells["Articulo"].Value),/*6*/
								Convert.ToString  ( drs.Cells["Marca"].Value),/*7*/
								Convert.ToString  ( drs.Cells["Modelo"].Value),/*8*/
								Convert.ToInt32 ( drs.Cells["Saldo"].Value), /*9*/
								Convert.ToString  ( drs.Cells["Color"].Value),/*10*/
								Convert.ToString  ( drs.Cells["Serial"].Value),/*11*/
								Convert.ToString  ( drs.Cells["Factura"].Value),/*12*/
								Convert.ToString  ( drs.Cells["Numero"].Value),/*13*/
								Convert.ToDateTime(drs.Cells["Fecha_Ultimo_Movimiento"].Value).ToString("yyyyMMdd"),/*14*/
								Convert.ToDecimal ( drs.Cells["Costo"].Value),/*15*/
								Convert.ToInt32 ( drs.Cells["Constancia"].Value), /*16*/
								Convert.ToInt32 ( drs.Cells["Estado"].Value), /*17*/
								Convert.ToString  ( drs.Cells["Observaciones"].Value),/*18*/
								Convert.ToInt32 ( drs.Cells["Bodega"].Value)/*19*/
								);
							oCmdActualiza.CommandText = sSQLSer;
							oCmdActualiza.ExecuteNonQuery();	
						}						
					}
					#endregion Graba DetalleMaestroSeriales

				   

					oTransaction.Commit();

					Cursor = Cursors.Default;	

					if( (int) this.txtEstado.Value == 1)
					{

						MessageBox.Show("Acta de Inventario se Guardo en un estado TEMPORAL", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					else
					{
						MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
					}
				
				}
				catch(Exception ex)
				{
					try
					{ 
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
				
			#endregion Guardar  
  
			FuncionesProcedimientos.EstadoGrids(false,this.grdPersonal);
			FuncionesProcedimientos.EstadoGrids(false,this.grdSeries);
			this.cmbInventariador.Enabled = false;

			this.Consulta((int)this.txtidConstanciaFisica.Value);
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaGeneralConstanciaFisica Busqueda = new frmConsultaGeneralConstanciaFisica())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridCreditos.ActiveRow.Cells["idConstanciaFisica"].Value);
				}
			}
		}

		private void Consulta (int Id)
		{			
			int iFecha = 0;
			string sSQL = string.Format("Exec ConsultaIndividualConstanciaFisica {0}", Id);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidConstanciaFisica.Value = Id;
				if (dr.GetValue(1) != System.DBNull.Value) 
				{
					this.dtFecha.Value = dr.GetDateTime(1);
				}
				else
				{
					iFecha = 1;
				}
				if (dr.GetValue(2) != System.DBNull.Value) this.cmbInventariador.Value = dr.GetInt32(2);
				if (dr.GetValue(3) != System.DBNull.Value) this.cmbBodega.Value = dr.GetInt32(3);
				if (dr.GetValue(4) != System.DBNull.Value) this.dtFechaCierre.Value = dr.GetDateTime(4);
				if (dr.GetValue(5) != System.DBNull.Value) this.txtEstado.Value = dr.GetInt32(5);

				
			}
			dr.Close();

			if ( iFecha > 0) RetornaFechaS(this.dtFecha, cdsSeteoF, true);
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnImprimir.Enabled = true;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar)
			{
				if ((int)this.txtEstado.Value == 1 || (int)this.txtEstado.Value ==0 ) 
				{
					this.btnEditar.Enabled = true;
					this.btnVer.Enabled = true;
					this.txtNombre.Enabled = true;
				}
				else
				{
					this.btnEditar.Enabled = false;
					
				}
			}
			if ((int)this.txtEstado.Value == 2  ) 
			{
				this.btnEditar.Enabled = false;

				#region imprime
				ParameterFields paramFields = new ParameterFields();

				ParameterField paramFieldCedula = new ParameterField ();
				ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
				paramFieldCedula.ParameterFieldName = "@idConstanciaFisica";
				ValCedula.Value = (int)this.txtidConstanciaFisica.Value;
				paramFieldCedula.CurrentValues.Add (ValCedula);
				paramFields.Add (paramFieldCedula);

//				ParameterField paramFieldCedulan = new ParameterField ();
//				ParameterDiscreteValue ValCedulan = new ParameterDiscreteValue ();
//				paramFieldCedulan.ParameterFieldName = "@Serial";
//				ValCedulan.Value = this.txtNombre.Text.ToString();
//				paramFieldCedulan.CurrentValues.Add (ValCedulan);
//				paramFields.Add (paramFieldCedulan);

				Reporte miRep = new Reporte("ReporteGeneralConstanciaFisica.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Planificacion De Carpa");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
				#endregion imprime
			}
			string sSQLPa = string.Format(" EXEC ConsultaGeneralPersonalConstanciaFisica   {0} ", (int)this.txtidConstanciaFisica.Value);
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.grdPersonal);

			string sSQLPas = string.Format(" EXEC ConsultaGeneralDetalleConstanciaFisicaS   {0} ", (int)this.txtidConstanciaFisica.Value);
			FuncionesProcedimientos.dsGeneral(sSQLPas, this.grdSeries);
			

			if ((int)this.txtEstado.Value == 1 ) this.lblEstado.Text = "PENDIENTE";
			if ((int)this.txtEstado.Value == 2 ) this.lblEstado.Text = "TERMINADO";

			this.lblContador2.Text = this.grdSeries.Rows.Count + " REGISTROS ENCONTRADOS";	


		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			
			ParameterFields paramFields = new ParameterFields();

			ParameterField paramFieldCedula = new ParameterField ();
			ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
			paramFieldCedula.ParameterFieldName = "@idConstanciaFisica";
			ValCedula.Value = (int)this.txtidConstanciaFisica.Value;
			paramFieldCedula.CurrentValues.Add (ValCedula);
			paramFields.Add (paramFieldCedula);

//			ParameterField paramFieldCedulan = new ParameterField ();
//			ParameterDiscreteValue ValCedulan = new ParameterDiscreteValue ();
//			paramFieldCedulan.ParameterFieldName = "@Serial";
//			ValCedulan.Value = (string)this.txtNombre.Text;
//			paramFieldCedulan.CurrentValues.Add (ValCedulan);
//			paramFields.Add (paramFieldCedulan);


			Reporte miRep = new Reporte("ReporteGeneralConstanciaFisica.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Planificacion De Carpa");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();

		}

		private void cmbInventariador_ValueChanged(object sender, System.EventArgs e)
		{
			
			if (this.cmbBodega.ActiveRow != null)
			{
				if (this.dtFecha.Value != System.DBNull.Value)
				{

					bNuevo = true ;
					ListaVendedores(this.grdPersonal,(int)this.cmbBodega.Value,(DateTime)this.dtFecha.Value, 0, cdsSeteoF);
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.cmbInventariador.Enabled = true;
			FuncionesProcedimientos.EstadoGrids(true,this.grdSeries);

			//			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;	

			if ((int)this.txtEstado.Value == 1)
			{
				this.btnVer.Enabled = true;
				this.txtNombre.Enabled = true;
			}
			else
			{
				this.btnVer.Enabled = false;
				this.txtNombre.Enabled = false;
			}

			bEdicion=true;
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.grdSeries);
		}

		bool bActualiza = false;

		private void grdSeries_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "Constancia")
			{	
				bActualiza = true;

				if ((int) grdSeries.ActiveRow.Cells["Constancia"].Value > 0)
				{					
										
					string sSQL = string.Format("UPDATE DetalleConstanciaFisica SET Constancia = {0} WHERE idDetalleConstanciaFisica = {1} ", 
						(int)	this.grdSeries.ActiveRow.Cells["Constancia"].Value, (int)	this.grdSeries.ActiveRow.Cells["idDetalleConstanciaFisica"].Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
									
				}
				bActualiza = false;
			}	
			if (e.Cell.Column.ToString() == "Estado")
			{	
				bActualiza = true;

				if ((int) grdSeries.ActiveRow.Cells["Estado"].Value > 0)
				{						
					string sSQL = string.Format("UPDATE DetalleConstanciaFisica SET Estado = {0} WHERE idDetalleConstanciaFisica = {1} ", 
						(int)	this.grdSeries.ActiveRow.Cells["Estado"].Value, (int)	this.grdSeries.ActiveRow.Cells["idDetalleConstanciaFisica"].Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
									
				}
				bActualiza = false;
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			string sSQLPa = string.Format(" EXEC ConsultaGeneralDetalleConstanciaFisicaS   {0},'{1}' ", (int)this.txtidConstanciaFisica.Value, this.txtNombre.Text.ToString ());
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.grdSeries);
		}

		private void cmbInventariador_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNombre.Focus();
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			this.btnVer_Click (sender, e);
		}

		private void grdSeries_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
