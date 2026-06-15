using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmMercaderiaRetiroPorLocales.
	/// </summary>
	public class frmMercaderiaRetiroPorLocales : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroFactura;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdMercaderiaRetiroLocales;
		private System.Windows.Forms.Label lblNota;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNota;
		private System.Windows.Forms.Label lblfecha;
		private System.Windows.Forms.Label lblNumerodeFactura;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label lblRecibidopor;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtLocal;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtEmision;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Button btnProcesar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.GroupBox groupBox5;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDTIngreso;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidcompra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCedula;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRecibidopor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstados;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRetiradoPor;
		private System.Windows.Forms.Button btnRecibido;
		private System.Windows.Forms.Label txtNumero;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnRecibidoM;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmMercaderiaRetiroPorLocales()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmMercaderiaRetiroPorLocales));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDTEstado", -1, "cmbEstados");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleMercaderiaLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio11");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio13");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio14");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio15");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio16");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio17");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio18");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio19");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio20");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("A", 0);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDTEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleMercaderiaLocales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio13");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio14");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio16");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio17");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio18");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio19");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio20");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDTEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			this.lblfecha = new System.Windows.Forms.Label();
			this.lblNumerodeFactura = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNumeroFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new System.Windows.Forms.Label();
			this.txtIdMercaderiaRetiroLocales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblBodega = new System.Windows.Forms.Label();
			this.lblNota = new System.Windows.Forms.Label();
			this.txtNota = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.lblRecibidopor = new System.Windows.Forms.Label();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.label11 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtLocal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtEmision = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnProcesar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnRecibido = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.uGridDTIngreso = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtidcompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCedula = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbRecibidopor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbRetiradoPor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnRecibidoM = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdMercaderiaRetiroLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtEmision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDTIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidcompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRecibidopor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetiradoPor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// lblfecha
			// 
			this.lblfecha.AutoSize = true;
			this.lblfecha.Location = new System.Drawing.Point(8, 10);
			this.lblfecha.Name = "lblfecha";
			this.lblfecha.Size = new System.Drawing.Size(36, 16);
			this.lblfecha.TabIndex = 476;
			this.lblfecha.Text = "Fecha";
			this.lblfecha.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumerodeFactura
			// 
			this.lblNumerodeFactura.AutoSize = true;
			this.lblNumerodeFactura.BackColor = System.Drawing.Color.Transparent;
			this.lblNumerodeFactura.Location = new System.Drawing.Point(200, 114);
			this.lblNumerodeFactura.Name = "lblNumerodeFactura";
			this.lblNumerodeFactura.Size = new System.Drawing.Size(101, 16);
			this.lblNumerodeFactura.TabIndex = 473;
			this.lblNumerodeFactura.Text = "Numero de Factura";
			this.lblNumerodeFactura.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblNumero.AutoSize = true;
			this.lblNumero.Location = new System.Drawing.Point(816, 10);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 471;
			this.lblNumero.Text = "Numero";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 477;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtNumeroFactura
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumeroFactura.Appearance = appearance2;
			this.txtNumeroFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumeroFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroFactura.Enabled = false;
			this.txtNumeroFactura.Location = new System.Drawing.Point(320, 112);
			this.txtNumeroFactura.MaxLength = 13;
			this.txtNumeroFactura.Name = "txtNumeroFactura";
			this.txtNumeroFactura.Size = new System.Drawing.Size(112, 21);
			this.txtNumeroFactura.TabIndex = 472;
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumero.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.txtNumero.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.txtNumero.Location = new System.Drawing.Point(880, 7);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(128, 22);
			this.txtNumero.TabIndex = 470;
			this.txtNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIdMercaderiaRetiroLocales
			// 
			this.txtIdMercaderiaRetiroLocales.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdMercaderiaRetiroLocales.Appearance = appearance3;
			this.txtIdMercaderiaRetiroLocales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdMercaderiaRetiroLocales.Enabled = false;
			this.txtIdMercaderiaRetiroLocales.FormatString = "###0";
			this.txtIdMercaderiaRetiroLocales.Location = new System.Drawing.Point(976, 104);
			this.txtIdMercaderiaRetiroLocales.MaxValue = 100000000;
			this.txtIdMercaderiaRetiroLocales.MinValue = 0;
			this.txtIdMercaderiaRetiroLocales.Name = "txtIdMercaderiaRetiroLocales";
			this.txtIdMercaderiaRetiroLocales.PromptChar = ' ';
			this.txtIdMercaderiaRetiroLocales.Size = new System.Drawing.Size(24, 21);
			this.txtIdMercaderiaRetiroLocales.TabIndex = 698;
			this.txtIdMercaderiaRetiroLocales.Visible = false;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(200, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(43, 16);
			this.lblBodega.TabIndex = 699;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNota
			// 
			this.lblNota.AutoSize = true;
			this.lblNota.BackColor = System.Drawing.Color.Transparent;
			this.lblNota.Location = new System.Drawing.Point(360, 53);
			this.lblNota.Name = "lblNota";
			this.lblNota.Size = new System.Drawing.Size(34, 16);
			this.lblNota.TabIndex = 702;
			this.lblNota.Text = "Notas";
			this.lblNota.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNota
			// 
			this.txtNota.AcceptsReturn = true;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNota.Appearance = appearance4;
			this.txtNota.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNota.Enabled = false;
			this.txtNota.Location = new System.Drawing.Point(416, 40);
			this.txtNota.MaxLength = 300;
			this.txtNota.Multiline = true;
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(440, 42);
			this.txtNota.TabIndex = 701;
			this.txtNota.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNota_KeyDown);
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Location = new System.Drawing.Point(8, 53);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(68, 16);
			this.label20.TabIndex = 705;
			this.label20.Text = "Retirado Por";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblRecibidopor
			// 
			this.lblRecibidopor.AutoSize = true;
			this.lblRecibidopor.Location = new System.Drawing.Point(472, 10);
			this.lblRecibidopor.Name = "lblRecibidopor";
			this.lblRecibidopor.Size = new System.Drawing.Size(72, 16);
			this.lblRecibidopor.TabIndex = 717;
			this.lblRecibidopor.Text = "Recibido Por:";
			this.lblRecibidopor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(0, 96);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1016, 8);
			this.groupBox3.TabIndex = 727;
			this.groupBox3.TabStop = false;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(456, 114);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(31, 16);
			this.label11.TabIndex = 737;
			this.label11.Text = "Local";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 114);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(45, 16);
			this.label7.TabIndex = 732;
			this.label7.Text = "Emisión";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(8, 146);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(40, 16);
			this.label8.TabIndex = 731;
			this.label8.Text = "Cédula";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(200, 146);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(44, 16);
			this.label9.TabIndex = 729;
			this.label9.Text = "Nombre";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtLocal
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLocal.Appearance = appearance5;
			this.txtLocal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLocal.Enabled = false;
			this.txtLocal.Location = new System.Drawing.Point(504, 112);
			this.txtLocal.MaxLength = 13;
			this.txtLocal.Name = "txtLocal";
			this.txtLocal.Size = new System.Drawing.Size(160, 21);
			this.txtLocal.TabIndex = 736;
			// 
			// dtEmision
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtEmision.Appearance = appearance6;
			this.dtEmision.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtEmision.DateButtons.Add(dateButton2);
			this.dtEmision.Enabled = false;
			this.dtEmision.Format = "dd/MM/yyyy";
			this.dtEmision.Location = new System.Drawing.Point(64, 112);
			this.dtEmision.Name = "dtEmision";
			this.dtEmision.NonAutoSizeHeight = 23;
			this.dtEmision.Size = new System.Drawing.Size(112, 21);
			this.dtEmision.SpinButtonsVisible = true;
			this.dtEmision.TabIndex = 733;
			this.dtEmision.Value = ((object)(resources.GetObject("dtEmision.Value")));
			// 
			// txtNombre
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance7;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(320, 144);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(344, 21);
			this.txtNombre.TabIndex = 728;
			// 
			// btnProcesar
			// 
			this.btnProcesar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnProcesar.Enabled = false;
			this.btnProcesar.Image = ((System.Drawing.Image)(resources.GetObject("btnProcesar.Image")));
			this.btnProcesar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnProcesar.Location = new System.Drawing.Point(417, 376);
			this.btnProcesar.Name = "btnProcesar";
			this.btnProcesar.Size = new System.Drawing.Size(104, 23);
			this.btnProcesar.TabIndex = 746;
			this.btnProcesar.Text = "&Enviar a Matriz";
			this.btnProcesar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 376);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 738;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(86, 376);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 739;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(170, 376);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 742;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(252, 376);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 741;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnRecibido
			// 
			this.btnRecibido.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnRecibido.Enabled = false;
			this.btnRecibido.Image = ((System.Drawing.Image)(resources.GetObject("btnRecibido.Image")));
			this.btnRecibido.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRecibido.Location = new System.Drawing.Point(335, 376);
			this.btnRecibido.Name = "btnRecibido";
			this.btnRecibido.Size = new System.Drawing.Size(76, 23);
			this.btnRecibido.TabIndex = 740;
			this.btnRecibido.Text = "&Guardar";
			this.btnRecibido.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRecibido.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(735, 376);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 744;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(653, 376);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 743;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(813, 376);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 745;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox4
			// 
			this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox4.Location = new System.Drawing.Point(0, 360);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1022, 8);
			this.groupBox4.TabIndex = 747;
			this.groupBox4.TabStop = false;
			// 
			// groupBox5
			// 
			this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox5.Location = new System.Drawing.Point(0, 168);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(1016, 8);
			this.groupBox5.TabIndex = 748;
			this.groupBox5.TabStop = false;
			// 
			// uGridDTIngreso
			// 
			this.uGridDTIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDTIngreso.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDTIngreso.DataSource = this.ultraDataSource1;
			appearance8.BackColor = System.Drawing.Color.White;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDTIngreso.DisplayLayout.Appearance = appearance8;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 120;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 160;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 220;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance9;
			ultraGridColumn7.Header.Caption = "Cant";
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Width = 50;
			ultraGridColumn8.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn8.Header.Caption = "Estado";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 100;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 300;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance10.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn31.CellAppearance = appearance10;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			appearance11.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance11.ImageVAlign = Infragistics.Win.VAlign.Middle;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance11.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn31.CellButtonAppearance = appearance11;
			ultraGridColumn31.Header.Caption = "...";
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn31.Width = 30;
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
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			this.uGridDTIngreso.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDTIngreso.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.uGridDTIngreso.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDTIngreso.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDTIngreso.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.uGridDTIngreso.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 7.75F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDTIngreso.DisplayLayout.Override.HeaderAppearance = appearance14;
			appearance15.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance15.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDTIngreso.DisplayLayout.Override.RowAlternateAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDTIngreso.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDTIngreso.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.uGridDTIngreso.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDTIngreso.Location = new System.Drawing.Point(7, 184);
			this.uGridDTIngreso.Name = "uGridDTIngreso";
			this.uGridDTIngreso.Size = new System.Drawing.Size(1002, 176);
			this.uGridDTIngreso.TabIndex = 749;
			this.uGridDTIngreso.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDTIngreso_AfterRowInsert);
			this.uGridDTIngreso.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDTIngreso_BeforeRowsDeleted);
			this.uGridDTIngreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDTIngreso_KeyDown);
			this.uGridDTIngreso.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDTIngreso_ClickCellButton);
			this.uGridDTIngreso.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDTIngreso_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(int);
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
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30});
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
			// txtidcompra
			// 
			this.txtidcompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidcompra.Appearance = appearance18;
			this.txtidcompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidcompra.Enabled = false;
			this.txtidcompra.FormatString = "###0";
			this.txtidcompra.Location = new System.Drawing.Point(976, 152);
			this.txtidcompra.MaxValue = 10000000;
			this.txtidcompra.MinValue = 0;
			this.txtidcompra.Name = "txtidcompra";
			this.txtidcompra.PromptChar = ' ';
			this.txtidcompra.Size = new System.Drawing.Size(24, 21);
			this.txtidcompra.TabIndex = 751;
			this.txtidcompra.Visible = false;
			// 
			// txtCedula
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCedula.Appearance = appearance19;
			this.txtCedula.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCedula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCedula.Enabled = false;
			this.txtCedula.Location = new System.Drawing.Point(64, 144);
			this.txtCedula.MaxLength = 10;
			this.txtCedula.Name = "txtCedula";
			this.txtCedula.Size = new System.Drawing.Size(112, 21);
			this.txtCedula.TabIndex = 753;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(264, 7);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 754;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// cmbRecibidopor
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRecibidopor.Appearance = appearance20;
			this.cmbRecibidopor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRecibidopor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn33.Header.VisiblePosition = 1;
			ultraGridColumn33.Width = 335;
			ultraGridColumn34.Header.VisiblePosition = 2;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			this.cmbRecibidopor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbRecibidopor.DisplayMember = "Nombre";
			this.cmbRecibidopor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRecibidopor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRecibidopor.Enabled = false;
			this.cmbRecibidopor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRecibidopor.Location = new System.Drawing.Point(560, 7);
			this.cmbRecibidopor.Name = "cmbRecibidopor";
			this.cmbRecibidopor.Size = new System.Drawing.Size(240, 22);
			this.cmbRecibidopor.TabIndex = 755;
			this.cmbRecibidopor.ValueMember = "idPersonal";
			this.cmbRecibidopor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbRecibidopor_KeyDown);
			// 
			// cmbEstados
			// 
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 78;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Width = 98;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn35,
																										 ultraGridColumn36});
			ultraGridBand3.SummaryFooterCaption = "Grand Summaries";
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEstados.DisplayMember = "Descripcion";
			this.cmbEstados.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbEstados.Location = new System.Drawing.Point(624, 224);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(100, 104);
			this.cmbEstados.TabIndex = 756;
			this.cmbEstados.ValueMember = "idDTEstado";
			this.cmbEstados.Visible = false;
			// 
			// cmbRetiradoPor
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRetiradoPor.Appearance = appearance21;
			this.cmbRetiradoPor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRetiradoPor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn37.Header.VisiblePosition = 0;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 1;
			ultraGridColumn38.Width = 335;
			ultraGridColumn39.Header.VisiblePosition = 2;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39});
			this.cmbRetiradoPor.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbRetiradoPor.DisplayMember = "Nombre";
			this.cmbRetiradoPor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRetiradoPor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRetiradoPor.Enabled = false;
			this.cmbRetiradoPor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRetiradoPor.Location = new System.Drawing.Point(96, 50);
			this.cmbRetiradoPor.Name = "cmbRetiradoPor";
			this.cmbRetiradoPor.Size = new System.Drawing.Size(240, 22);
			this.cmbRetiradoPor.TabIndex = 757;
			this.cmbRetiradoPor.ValueMember = "idPersonal";
			this.cmbRetiradoPor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbRetiradoPor_KeyDown);
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstado.Appearance = appearance22;
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.FormatString = "###0";
			this.txtEstado.Location = new System.Drawing.Point(976, 128);
			this.txtEstado.MaxValue = 10000000;
			this.txtEstado.MinValue = 0;
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(24, 21);
			this.txtEstado.TabIndex = 758;
			this.txtEstado.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 11F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblEstado.Location = new System.Drawing.Point(744, 128);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 21);
			this.lblEstado.TabIndex = 759;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnRecibidoM
			// 
			this.btnRecibidoM.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnRecibidoM.Enabled = false;
			this.btnRecibidoM.Image = ((System.Drawing.Image)(resources.GetObject("btnRecibidoM.Image")));
			this.btnRecibidoM.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRecibidoM.Location = new System.Drawing.Point(527, 376);
			this.btnRecibidoM.Name = "btnRecibidoM";
			this.btnRecibidoM.Size = new System.Drawing.Size(120, 23);
			this.btnRecibidoM.TabIndex = 760;
			this.btnRecibidoM.Text = "&Recibido en Matriz";
			this.btnRecibidoM.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRecibidoM.Click += new System.EventHandler(this.btnRecibidoM_Click);
			// 
			// frmMercaderiaRetiroPorLocales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1016, 406);
			this.Controls.Add(this.btnRecibidoM);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.cmbRetiradoPor);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.cmbRecibidopor);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtCedula);
			this.Controls.Add(this.txtidcompra);
			this.Controls.Add(this.uGridDTIngreso);
			this.Controls.Add(this.groupBox5);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.btnProcesar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnRecibido);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.lblRecibidopor);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.lblNota);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.lblfecha);
			this.Controls.Add(this.lblNumerodeFactura);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.txtLocal);
			this.Controls.Add(this.dtEmision);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.txtNota);
			this.Controls.Add(this.txtIdMercaderiaRetiroLocales);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtNumeroFactura);
			this.Controls.Add(this.txtNumero);
			this.KeyPreview = true;
			this.Name = "frmMercaderiaRetiroPorLocales";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Recepción Mercaderia de Retiro";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmMercaderiaRetiroPorLocales_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmMercaderiaRetiroPorLocales_Closing);
			this.Load += new System.EventHandler(this.frmMercaderiaRetiroPorLocales_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdMercaderiaRetiroLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtEmision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDTIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidcompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRecibidopor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetiradoPor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		bool bNuevo = false;
		bool bEdicion = false;
		int idBodegaPredef =0;
		int idBloqueaTransacciones =0;

		public frmMercaderiaRetiroPorLocales(int iEstado)
		{		
			InitializeComponent();
		
			this.txtEstado.Value = iEstado;
		}
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmMercaderiaRetiroPorLocales_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0833");

			if (!Funcion.Permiso("964", cdsSeteoF))
			{		
		
				MessageBox.Show("No tiene Acceso Retiro de Mercaderia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;

			}


			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT idDTEstado, Descripcion FROM DTEstados");
			FuncionesProcedimientos.EstadoGrids(false,this.uGridDTIngreso);
			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			idBodegaPredef = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click (sender,e);

			using (frmDTBuscaFacturaRetiro Busqueda = new frmDTBuscaFacturaRetiro())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.txtNumeroFactura.Text = Busqueda.uGridFacturas.ActiveRow.Cells["Numero"].Value.ToString();
					this.txtidcompra.Value =(int)Busqueda.uGridFacturas.ActiveRow.Cells["idCompra"].Value;
					this.txtCedula.Text = Busqueda.uGridFacturas.ActiveRow.Cells["RUC"].Value.ToString();
					this.txtNombre.Text = Busqueda.uGridFacturas.ActiveRow.Cells["Nombre"].Value.ToString();
					this.txtLocal.Text = Busqueda.uGridFacturas.ActiveRow.Cells["Local"].Value.ToString();
					this.dtEmision.Value = (DateTime)Busqueda.uGridFacturas.ActiveRow.Cells["Fecha"].Value;

					this.uGridDTIngreso.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DTConsultaDetalleRetiroPoLocales {0}", (int)this.txtidcompra.Value));

					this.dtFecha.Value = DateTime.Today;
					this.txtEstado.Value=1;
					
					this.cmbRetiradoPor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Codigo, Nombre From Personal Where idCargo In (2, 36, 37, 38) And FechaSalida IS NULL Order By Nombre");
			
					#region habilita controles 
					this.cmbRecibidopor.Enabled=true;
					this.cmbRetiradoPor.Enabled=true;
					this.txtNota.Enabled= true;
					
					FuncionesProcedimientos.EstadoGrids(true, this.uGridDTIngreso);

					this.btnNuevo.Enabled = false;
					this.btnConsultar.Enabled = false;
					this.btnEditar.Enabled = false;
					this.btnImprimir.Enabled = false;
					this.btnRecibido.Enabled = true;
					this.btnProcesar.Enabled  = false;
					this.btnRecibidoM.Enabled = false;
					this.btnAnular.Enabled = false;
					this.btnCancelar.Enabled = true;	

					bNuevo = true;
					bEdicion = true;

					if (idBodegaPredef == 0) 
					{
						this.cmbBodega.Value = 1; 
						this.cmbBodega.Enabled = true;
					}
					else this.cmbBodega.Value = idBodegaPredef;

					if(this.cmbBodega.Enabled) this.cmbBodega.Focus(); else this.cmbRecibidopor.Focus();
					#endregion habilita controles					
				}
			}			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)	
			{				
				if (this.cmbRecibidopor.ActiveRow != null) this.cmbRecibidopor.Value = System.DBNull.Value;
				if (this.cmbBodega.ActiveRow != null) 
					this.cmbRecibidopor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 43, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if (this.cmbBodega.ActiveRow != null)
				{
					if (this.cmbRecibidopor.ActiveRow != null) this.cmbRecibidopor.Value = System.DBNull.Value;
					this.cmbRecibidopor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
				}
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			int iEstadoGrid = 0;
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una Bodega")) return;
			if (!Validacion.vbComboVacio(this.cmbRecibidopor, "Seleccione la persona quien recibio")) return;
			if (!Validacion.vbComboVacio(this.cmbRetiradoPor, "Seleccione la persona quien retiro")) return;
			if (!Validacion.vbTexto(this.txtNota, 20, 300, "Notas")) return;

			#region Seriales
			bool bSeriales = false;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDTIngreso.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
					{
						if (dr.Cells["Serial"].Value.ToString().Length == 0)
						{
							this.uGridDTIngreso.ActiveRow = dr;
							this.uGridDTIngreso.ActiveRow.Selected = true;
							bSeriales = true;
							continue;
						}
					}
				}
			}

			if (bSeriales)
			{
				MessageBox.Show("Ingrese la serial de los articulos resaltados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Seriales
			
			#region Valida Grid 
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDTi in this.uGridDTIngreso.Rows.All)
			{
				
				if((int)drDTi.Cells["idDTEstado"].Value== 0)
				{
					iEstadoGrid =  iEstadoGrid+ 1;
					this.uGridDTIngreso.ActiveRow = drDTi;
					this.uGridDTIngreso.ActiveRow.Selected = true;
				}
				
				if(drDTi.Cells["Notas"].Value.ToString().Length < 10)
				{
						
					MessageBox.Show("El campo notas debe tener minimo 10 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDTIngreso.ActiveRow = drDTi;
					this.uGridDTIngreso.ActiveRow.Selected = true;
					return;
				}
			}
			if(iEstadoGrid > 0) {MessageBox.Show("Seleccione el Estado del producto " , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); return;}
			#endregion Valida Grid 

			#region Numeracion
			string sSQL = string.Format("Exec NumeracionLocales 63, {0}, 0", (int)this.cmbBodega.Value);
			this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, sSQL);			
			#endregion Numeracion
      
			#region Guardar Maestro
			string sSQLl = string.Format("Exec GuardaMercaderiaRetiroLocales {0},{1},{2},{3},'{4}',{5},{6},'{7}','{8}'",
				(int) this.txtIdMercaderiaRetiroLocales.Value,
				(int) this.txtidcompra.Value,
				(int) this.cmbRetiradoPor.Value,
				(int) this.cmbRecibidopor.Value,
				this.txtNumero.Text.ToString(), 
				(int) this.cmbBodega.Value,
				this.txtEstado.Value,
				this.txtNota.Text.ToString(), 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			this.txtIdMercaderiaRetiroLocales.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQLl);
			#endregion Guardar Maestro

			#region Guarda detalle Maestro
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.uGridDTIngreso.Rows.All)
			{
				int iEstado = (int)this.txtEstado.Value;
				string sSQLDetalle = string.Format("Exec GuardaDetalleMercaderiaLocales  {0},{1},{2},{3},'{4}','{5}','{6}',{7},{8},'{9}',{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26},{27},{28},{29}", 
					(int)drr.Cells["idDetalleMercaderiaLocales"].Value,
					(int)this.txtIdMercaderiaRetiroLocales.Value,
					(int)drr.Cells["idArticulo"].Value,
					(int)drr.Cells["idSerial"].Value,
					drr.Cells["Serial"].Value.ToString(),
					drr.Cells["Codigo"].Value.ToString(),
					drr.Cells["Articulo"].Value.ToString(),
					(int)drr.Cells["idDTEstado"].Value, 
					(int)drr.Cells["Cantidad"].Value, 
					drr.Cells["Notas"].Value.ToString(),(int)drr.Cells["idAccesorio1"].Value, (int)drr.Cells["idAccesorio2"].Value, (int)drr.Cells["idAccesorio3"].Value, (int)drr.Cells["idAccesorio4"].Value,
					(int)drr.Cells["idAccesorio5"].Value, (int)drr.Cells["idAccesorio6"].Value, (int)drr.Cells["idAccesorio7"].Value, (int)drr.Cells["idAccesorio8"].Value,
					(int)drr.Cells["idAccesorio9"].Value, (int)drr.Cells["idAccesorio10"].Value, (int)drr.Cells["idAccesorio11"].Value, (int)drr.Cells["idAccesorio12"].Value,
					(int)drr.Cells["idAccesorio13"].Value, (int)drr.Cells["idAccesorio14"].Value, (int)drr.Cells["idAccesorio15"].Value, (int)drr.Cells["idAccesorio16"].Value,
					(int)drr.Cells["idAccesorio17"].Value, (int)drr.Cells["idAccesorio18"].Value, (int)drr.Cells["idAccesorio19"].Value, (int)drr.Cells["idAccesorio20"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);

				
			}	
			#endregion Guarda detalle Maestro

			#region bloquea controles 
			//			FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, false, (int)this.txtIdMercaderiaRetiroLocales.Value, 11, cdsSeteoF);

			this.lblEstado.Text= "RECIBIDO";
			this.dtFecha.Enabled = false;
			
			
			this.cmbBodega.Enabled = false;
			this.cmbRecibidopor.Enabled = false;
			
			this.cmbRetiradoPor.Enabled = false;
			this.txtNota.Enabled = false;
			
			FuncionesProcedimientos.EstadoGrids(false, this.uGridDTIngreso);

			
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnImprimir.Enabled = true;
			this.btnRecibido.Enabled = false;
			this.btnProcesar.Enabled  = true;
			this.btnRecibidoM.Enabled = false;
			this.btnAnular.Enabled = true;
			this.btnCancelar.Enabled = true;

			
			bNuevo = false;
			bEdicion = false;
			#endregion bloquea controles
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (!bNuevo && bEdicion) 
				FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, false, (int)this.txtIdMercaderiaRetiroLocales.Value, 11, cdsSeteoF);

			bNuevo = false;
			bEdicion = false;			

			this.dtFecha.Value=System.DBNull.Value;
			this.txtEstado.Value=0;
			this.lblEstado.Text="";
			this.txtIdMercaderiaRetiroLocales.Value = 0;
			this.txtidcompra.Value=0;
			this.cmbBodega.Value=System.DBNull.Value;
			this.cmbRecibidopor.Value=System.DBNull.Value;
			this.txtNumero.Text="";
			this.cmbRetiradoPor.Value=System.DBNull.Value;
			this.txtNota.Text="";
			this.dtEmision.Value=System.DBNull.Value;;
			this.txtNumeroFactura.Text="";
			this.txtLocal.Text="";
			this.txtCedula.Text="";
			this.txtNombre.Text="";

			this.uGridDTIngreso.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DTConsultaDetalleRetiroPoLocales 0"));
			
			#region bloquea controles 
			this.dtFecha.Enabled = false;
			
			this.cmbBodega.Enabled = false;
			this.cmbRecibidopor.Enabled = false;
			
			this.cmbRetiradoPor.Enabled = false;
			this.txtNota.Enabled = false;
			
			FuncionesProcedimientos.EstadoGrids(false, this.uGridDTIngreso);

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnRecibido.Enabled = false;
			this.btnProcesar.Enabled  = false;
			this.btnRecibidoM.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			idBloqueaTransacciones =0;
			#endregion bloquea controles
		}

		private void btnProcesar_Click(object sender, System.EventArgs e)
		{
			ValidacioncambioEstado ();

			using (frmNotasMercaderiaRetiroLocales NotasMercaderias = new frmNotasMercaderiaRetiroLocales())
			{
				NotasMercaderias.txtEstado.Value = this.txtEstado.Value;
				NotasMercaderias.txtBodegas.Value = this.cmbBodega.Value;

				if (DialogResult.OK == NotasMercaderias .ShowDialog())
				{
					int IdMercaderiaRetiroLocales =(int) this.txtIdMercaderiaRetiroLocales.Value;
					string sSQL = string.Format("Update MercaderiaRetiroLocaless Set NotaEnviadoMatriz = '{0}',Estado = 2,FechaEnviadoMatriz = GETDATE(), IdTransportista ={1} Where idMercaderiaRetiroLocales = {2}",
						NotasMercaderias.txtNota.Text.ToString(),NotasMercaderias.cmbTransportista.Value, IdMercaderiaRetiroLocales);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);

					Consultar ((int) this.txtIdMercaderiaRetiroLocales.Value);
					
					Imprimir ();
				}
			}			
		}

		private void uGridDTIngreso_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDetalleMercaderiaLocales"].Value = 0;
				e.Row.Cells["Cantidad"].Value = 1;
				e.Row.Cells["idDTEstado"].Value = 0;
				e.Row.Cells["idArticulo"].Value = 0; 
				e.Row.Cells["idSerial"].Value = 0; 	
				e.Row.Cells["idAccesorio1"].Value = 0; 
				e.Row.Cells["idAccesorio2"].Value = 0; 
				e.Row.Cells["idAccesorio3"].Value = 0; 
				e.Row.Cells["idAccesorio4"].Value = 0;
				e.Row.Cells["idAccesorio5"].Value = 0; 
				e.Row.Cells["idAccesorio6"].Value = 0; 
				e.Row.Cells["idAccesorio7"].Value = 0; 
				e.Row.Cells["idAccesorio8"].Value = 0;
				e.Row.Cells["idAccesorio9"].Value = 0; 
				e.Row.Cells["idAccesorio10"].Value = 0; 
				e.Row.Cells["idAccesorio11"].Value = 0; 
				e.Row.Cells["idAccesorio12"].Value = 0;
				e.Row.Cells["idAccesorio13"].Value = 0; 
				e.Row.Cells["idAccesorio14"].Value = 0; 
				e.Row.Cells["idAccesorio15"].Value = 0; 
				e.Row.Cells["idAccesorio16"].Value = 0;
				e.Row.Cells["idAccesorio17"].Value = 0; 
				e.Row.Cells["idAccesorio18"].Value = 0; 
				e.Row.Cells["idAccesorio19"].Value = 0; 
				e.Row.Cells["idAccesorio20"].Value = 0;
			}
		}
		private void Consultar (int idMercaderiaRetiroLocales )
		{
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnRecibido.Enabled = false;
			this.btnProcesar.Enabled  = false;
			this.btnRecibidoM.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;

			#region Consulta Individual
			string ssql= string.Format ("Exec ConsultaIndividualMercaderiaRetiroLocales {0}",idMercaderiaRetiroLocales); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
			this.txtIdMercaderiaRetiroLocales.Value = dr.GetInt32 (0);
			if(dr.GetValue(1)!= System.DBNull.Value)this.txtidcompra.Value = dr.GetInt32 (1);
			if(dr.GetValue(2)!= System.DBNull.Value)this.dtFecha.Value = dr.GetDateTime (2);
			if(dr.GetValue(3)!= System.DBNull.Value)this.cmbBodega.Value =dr.GetInt32 (3);
			if(dr.GetValue(4)!= System.DBNull.Value)this.txtNumero.Text = dr.GetString (4);
			if(dr.GetValue(5)!= System.DBNull.Value)this.txtNota.Text = dr.GetString (5);
			if(dr.GetValue(6)!= System.DBNull.Value)this.cmbRetiradoPor.Value = dr.GetInt32  (6);
			if(dr.GetValue(7)!= System.DBNull.Value)this.cmbRecibidopor.Value = dr.GetInt32 (7);
			if(dr.GetValue(8)!= System.DBNull.Value)this.txtEstado.Value = dr.GetInt32 (8);
			if(dr.GetValue(9)!= System.DBNull.Value)this.dtEmision.Value = dr.GetDateTime(9);
			if(dr.GetValue(10)!= System.DBNull.Value)this.txtNumeroFactura.Text = dr.GetString (10);
			if(dr.GetValue(11)!= System.DBNull.Value)this.txtLocal.Text = dr.GetString (11);
			if(dr.GetValue(12)!= System.DBNull.Value)this.txtCedula.Text = dr.GetString (12);
			if(dr.GetValue(13)!= System.DBNull.Value)this.txtNombre.Text= dr.GetString (13);
			dr.Close();
			this.uGridDTIngreso.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIDML {0}", (int)this.txtIdMercaderiaRetiroLocales.Value));
			#endregion Consulta Individual

			this.cmbRecibidopor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
			this.cmbRetiradoPor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Codigo, Nombre From Personal Where idCargo In (2, 36, 37, 38) And FechaSalida IS NULL Order By Nombre");
						
			if( (int) this.txtEstado.Value == 1)
			{
				this.lblEstado.Text ="RECIBIDO";
			}
			if( (int) this.txtEstado.Value == 2)
			{
				this.lblEstado.Text ="ENVIADO A MATRIZ";
				
			}
			if( (int) this.txtEstado.Value == 3)
			{
				this.lblEstado.Text ="RECIBIDO EN MATRIZ ";
			}
			if( (int) this.txtEstado.Value == 4)
			{
				this.lblEstado.Text ="ANULADO";
			}

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnRecibido.Enabled = false;

			if( (int) this.txtEstado.Value != 4)
			{
				
				if((int)this.txtEstado.Value == 1 ) this.btnEditar.Enabled = true;
				this.btnImprimir.Enabled = true;
				
				if((int)this.txtEstado.Value == 1 )this.btnProcesar.Enabled  = true;
				if((int)this.txtEstado.Value == 2 )this.btnRecibidoM.Enabled = true;
				if((int)this.txtEstado.Value == 1 )this.btnAnular.Enabled = true;				
			}
			this.btnCancelar.Enabled = true;
		}
		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click (sender,e);

			using (frmReporteGeneralMRL consultab = new frmReporteGeneralMRL())
			{
				if (DialogResult.OK == consultab.ShowDialog())
				{
					this.Consultar((int)consultab.uGridReporte.ActiveRow.Cells["idMercaderiaRetiroLocales"].Value);					
				}	
			}
		}

		private void ValidacioncambioEstado ()
		{
			Consultar ((int) this.txtIdMercaderiaRetiroLocales.Value  );

			if( (int) this.txtEstado.Value == 4)
			{
				MessageBox.Show("REGISTRO ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;		
			}
			
			if( (int) this.txtEstado.Value == 2)
			{
				MessageBox.Show("MERCADERIA ENVIADA A MATRIZ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if( (int) this.txtEstado.Value == 3)
			{
				MessageBox.Show("MERCADERIA RECIBIDA EN MATRIZ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;		
			}

			if (Funcion.ValidaBloqueo((int)this.txtIdMercaderiaRetiroLocales.Value, 11, cdsSeteoF)) return;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			ValidacioncambioEstado ();

			FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, true, (int)this.txtIdMercaderiaRetiroLocales.Value, 11, cdsSeteoF);

			this.txtNota.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.uGridDTIngreso);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnRecibido.Enabled = true;
			this.btnProcesar.Enabled  = false;
			this.btnRecibidoM.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
		
			bEdicion=true;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();			
		}

		private void btnRecibidoM_Click(object sender, System.EventArgs e)
		{
			Consultar ((int) this.txtIdMercaderiaRetiroLocales.Value  );

			if( (int) this.txtEstado.Value == 4)
			{
				MessageBox.Show("REGISTRO ANULADA.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;		
			}
			
			if( (int) this.txtEstado.Value == 3)
			{
				MessageBox.Show("MERCADERIA RECIBIDA EN MATRIZ.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;		
			}

			using (frmNotasMercaderiaRetiroLocales NotasMercaderias = new frmNotasMercaderiaRetiroLocales())
			{
				NotasMercaderias.txtEstado.Value = this.txtEstado.Value;
				NotasMercaderias.txtBodegas.Value = this.cmbBodega.Value;

				if (DialogResult.OK == NotasMercaderias .ShowDialog())
				{
					int idPlanificacionCompras =(int) this.txtIdMercaderiaRetiroLocales.Value;
					string sSQL = string.Format("Update MercaderiaRetiroLocaless Set NotaRecibidoEnMatriz= '{0}',Estado = 3,FechaRecibidoEnMatriz = GETDATE(), idPersonal ={1} Where idMercaderiaRetiroLocales = {2}",
						NotasMercaderias.txtNota.Text.ToString(),NotasMercaderias.cmbTransportista.Value, idPlanificacionCompras);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);

					Consultar ((int) this.txtIdMercaderiaRetiroLocales.Value  );

					Imprimir ();
				}
			}			
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			ValidacioncambioEstado();			

			using (frmNotasMercaderiaRetiroLocales NotasMercaderias = new frmNotasMercaderiaRetiroLocales())
			{
				NotasMercaderias.txtEstado.Value = (int) this.txtEstado.Value + 3;
				NotasMercaderias.txtBodegas.Value = this.cmbBodega.Value;

				if (DialogResult.OK == NotasMercaderias .ShowDialog())
				{
					NotasMercaderias.cmbTransportista.Visible = false;
					int idPlanificacionCompras =(int) this.txtIdMercaderiaRetiroLocales.Value;
					string sSQL = string.Format("Update MercaderiaRetiroLocaless Set NotaAnulado = '{0}',Estado = 4,FechaAnulado= GETDATE() Where idMercaderiaRetiroLocales = {1}",
						NotasMercaderias.txtNota.Text.ToString(), idPlanificacionCompras);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
					
					Consultar ((int) this.txtIdMercaderiaRetiroLocales.Value  );
				}
			}			
		}

		private void uGridDTIngreso_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bEdicion) 
			{
				string sCol = "";
				string sFiltro = "";
				if (e.Cell.Column.ToString() == "A")
				{					
					for (int j = 9; j < 29; j++)
					{
						sCol = this.uGridDTIngreso.DisplayLayout.Bands[0].Columns[j].Header.Caption;

						if (sFiltro.Length == 0) sFiltro = e.Cell.Row.Cells[sCol].Value.ToString();
						if (sFiltro.Length > 0) sFiltro = sFiltro + "," + e.Cell.Row.Cells[sCol].Value.ToString();
					}
				}

				using (frmAccesorios Accesorios = new frmAccesorios (sFiltro))
				{
					if (DialogResult.OK == Accesorios.ShowDialog())
					{
						int iPos = 1;						
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Accesorios.ultraGrid1.Rows.All)
						{
							string sColAsigna = "idAccesorio";
							if ((bool)dr.Cells["Sel"].Value && (int)dr.Cells["idAccesorio"].Value > 0)
							{
								sColAsigna = sColAsigna + iPos.ToString();
								e.Cell.Row.Cells[sColAsigna].Value = (int)dr.Cells["idAccesorio"].Value;
								iPos++;				
							}
						}

						int i = 8;
						i = i + iPos;
						for (int k = i ; k <= 28; k++)
						{
							sCol = this.uGridDTIngreso.DisplayLayout.Bands[0].Columns[k].Header.Caption;
							e.Cell.Row.Cells[sCol].Value = 0;
						}
					}
				}
			}
		}

		private void uGridDTIngreso_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Serial")
			{
				string sSQLSerial = string.Format("Declare @idSerial Int = 0 If Exists(Select idSerial From Seriales Where Serial = '{0}') Select @idSerial = idSerial From Seriales Where Serial = '{0}' Else Set @idSerial = 0 Select @idSerial", 
					e.Cell.Row.Cells["Serial"].Value.ToString());
				int idSerial = Funcion.iEscalarSQL(cdsSeteoF, sSQLSerial);
				e.Cell.Row.Cells["idSerial"].Value = idSerial;
			}
		}

		private void uGridDTIngreso_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta Seguro de Borrar los registros seleccionados?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					if (!bNuevo)
					{						
						for (int i=0; i<e.Rows.Length; i++)
						{							
							e.DisplayPromptMsg = false;

							string sSQLElimina = string.Format("Delete From DetalleMercaderiaLocales Where idDetalleMercaderiaLocales = {0}", 
								(int)e.Rows[i].Cells["idDetalleMercaderiaLocales"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLElimina, true);						
						}
					}
					else e.DisplayPromptMsg = false;
				}
			}
			else e.Cancel = true;
		}

		private void frmMercaderiaRetiroPorLocales_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un Registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void frmMercaderiaRetiroPorLocales_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E)  if ((int)this.txtEstado.Value == 1) this.btnEditar_Click(sender, e); 																											
				if (e.KeyCode == Keys.I)  if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
				if (e.KeyCode == Keys.A)  if ((int)this.txtEstado.Value == 1) this.btnAnular_Click(sender, e); 																																	 
				if (e.KeyCode == Keys.F5)
				{
					this.Consultar((int)this.txtIdMercaderiaRetiroLocales.Value);					
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

		private void Imprimir ()
		{
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldId = new ParameterField ();
			ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
			paramFieldId.ParameterFieldName = "@idMercaderiaRetiroLocales";
			ValIdId.Value = int.Parse(this.txtIdMercaderiaRetiroLocales.Value.ToString());
			paramFieldId.CurrentValues.Add (ValIdId);
			paramFields.Add (paramFieldId);

			ParameterField paramFieldEstado = new ParameterField ();
			ParameterDiscreteValue ValEstadoEstado = new ParameterDiscreteValue ();
			paramFieldEstado.ParameterFieldName = "@Estado";
			ValEstadoEstado.Value = int.Parse(this.txtEstado.Value.ToString());
			paramFieldEstado.CurrentValues.Add (ValEstadoEstado);
			paramFields.Add (paramFieldEstado);

			Reporte miRep = new Reporte("RecepcionMercaderiaRetiros.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Recepción Mercaderia de Retiro");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Consultar ((int) this.txtIdMercaderiaRetiroLocales.Value  );

			if( (int) this.txtEstado.Value == 4)
			{
				MessageBox.Show("REGISTRO ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;		
			}

			Imprimir ();
		}

		private void uGridDTIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,uGridDTIngreso);
		}
		
		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbRecibidopor.Focus();
		}

		private void cmbRetiradoPor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbRetiradoPor.Focus();
		}

		private void cmbRecibidopor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNota.Focus();
		}

		private void txtNota_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.uGridDTIngreso.Focus();
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
