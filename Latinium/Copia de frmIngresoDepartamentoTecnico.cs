using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmIngresoDepartamentoTecnico.
	/// </summary>
	public class frmIngresoDepartamentoTecnico : DevExpress.XtraEditors.XtraForm
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDTIngreso;
		private System.Windows.Forms.Label label36;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label39;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdIngreso;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.ComponentModel.IContainer components;

		public frmIngresoDepartamentoTecnico()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}


		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngresoMultiple");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Garantia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoGarantia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingresos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Daño");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Accesorios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empaque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio11");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio13");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio14");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio15");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio16");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio17");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio18");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio19");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio20");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnboundColumn1", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmIngresoDepartamentoTecnico));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngresoMultiple");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Garantia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TiempoGarantia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Daño");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Accesorios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empaque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio13");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio14");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio16");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio17");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio18");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio19");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio20");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.uGridDTIngreso = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label36 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label32 = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label39 = new System.Windows.Forms.Label();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.label8 = new System.Windows.Forms.Label();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtIdIngreso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblEstado = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.uGridDTIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridDTIngreso
			// 
			this.uGridDTIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDTIngreso.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDTIngreso.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDTIngreso.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 120;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 180;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Factura";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 75;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 55;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 50;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "T G";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 40;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 50;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 250;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 204;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 250;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn36.Header.VisiblePosition = 35;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn37.Header.VisiblePosition = 36;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			appearance2.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance2.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn38.CellButtonAppearance = appearance2;
			ultraGridColumn38.Header.Caption = "A";
			ultraGridColumn38.Header.VisiblePosition = 37;
			ultraGridColumn38.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn38.Width = 30;
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
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38});
			this.uGridDTIngreso.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDTIngreso.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.uGridDTIngreso.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDTIngreso.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDTIngreso.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridDTIngreso.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 7.75F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDTIngreso.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDTIngreso.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDTIngreso.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDTIngreso.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridDTIngreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDTIngreso.Location = new System.Drawing.Point(8, 240);
			this.uGridDTIngreso.Name = "uGridDTIngreso";
			this.uGridDTIngreso.Size = new System.Drawing.Size(1148, 312);
			this.uGridDTIngreso.TabIndex = 3;
			this.uGridDTIngreso.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDTIngreso_AfterRowInsert);
			this.uGridDTIngreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDTIngreso_KeyDown);
			this.uGridDTIngreso.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDTIngreso_ClickCellButton);
			this.uGridDTIngreso.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDTIngreso_AfterCellUpdate);
			this.uGridDTIngreso.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDTIngreso_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn10.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn11.DataType = typeof(bool);
			ultraDataColumn11.DefaultValue = false;
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(bool);
			ultraDataColumn17.DefaultValue = false;
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
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(int);
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
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37});
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.Location = new System.Drawing.Point(8, 11);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(30, 16);
			this.label36.TabIndex = 357;
			this.label36.Text = "Local";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn39.Header.VisiblePosition = 2;
			ultraGridColumn40.Header.VisiblePosition = 1;
			ultraGridColumn40.Width = 180;
			ultraGridColumn41.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(72, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 21);
			this.cmbBodega.TabIndex = 358;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodega_Validating);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// label32
			// 
			this.label32.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label32.AutoSize = true;
			this.label32.Location = new System.Drawing.Point(956, 11);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(44, 16);
			this.label32.TabIndex = 355;
			this.label32.Text = "Número";
			this.label32.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblNumero.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblNumero.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblNumero.Location = new System.Drawing.Point(1028, 8);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(128, 22);
			this.lblNumero.TabIndex = 353;
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 43);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 352;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.BackColor = System.Drawing.Color.Transparent;
			this.label39.Location = new System.Drawing.Point(264, 195);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(38, 16);
			this.label39.TabIndex = 369;
			this.label39.Text = "Celular";
			// 
			// txtCelular
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance9;
			this.txtCelular.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Enabled = false;
			this.txtCelular.Location = new System.Drawing.Point(328, 192);
			this.txtCelular.MaxLength = 10;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(112, 22);
			this.txtCelular.TabIndex = 368;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 195);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 367;
			this.label1.Text = "Teléfono";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(16, 163);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(50, 16);
			this.label2.TabIndex = 365;
			this.label2.Text = "Dirección";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(16, 99);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(38, 16);
			this.label6.TabIndex = 363;
			this.label6.Text = "Codigo";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(16, 131);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(44, 16);
			this.label5.TabIndex = 361;
			this.label5.Text = "Nombre";
			// 
			// txtTelefono
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefono.Appearance = appearance10;
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Location = new System.Drawing.Point(88, 192);
			this.txtTelefono.MaxLength = 9;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(112, 22);
			this.txtTelefono.TabIndex = 366;
			// 
			// txtDireccion
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccion.Appearance = appearance11;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Enabled = false;
			this.txtDireccion.Location = new System.Drawing.Point(88, 160);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(350, 22);
			this.txtDireccion.TabIndex = 364;
			// 
			// txtCodigo
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance12;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(88, 96);
			this.txtCodigo.MaxLength = 13;
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(112, 22);
			this.txtCodigo.TabIndex = 362;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCodigo_KeyPress);
			this.txtCodigo.Validating += new System.ComponentModel.CancelEventHandler(this.txtCodigo_Validating);
			this.txtCodigo.Validated += new System.EventHandler(this.txtCodigo_Validated);
			this.txtCodigo.ValueChanged += new System.EventHandler(this.txtCodigo_ValueChanged);
			// 
			// txtNombre
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance13;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(88, 128);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(352, 22);
			this.txtNombre.TabIndex = 360;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 576);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 371;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(84, 576);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 372;
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
			this.btnEditar.Location = new System.Drawing.Point(166, 576);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 375;
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
			this.btnImprimir.Location = new System.Drawing.Point(246, 576);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 374;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(327, 576);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 373;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(487, 576);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 377;
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
			this.btnAnular.Location = new System.Drawing.Point(407, 576);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 376;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 0;
			this.btnSalir.ImageList = this.imageList1;
			this.btnSalir.Location = new System.Drawing.Point(563, 576);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 378;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// groupBox5
			// 
			this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox5.Location = new System.Drawing.Point(0, 560);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(1164, 8);
			this.groupBox5.TabIndex = 370;
			this.groupBox5.TabStop = false;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(728, 96);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(76, 16);
			this.label8.TabIndex = 384;
			this.label8.Text = "Observaciones";
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance14;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Enabled = false;
			this.txtObservaciones.Location = new System.Drawing.Point(728, 128);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(424, 96);
			this.txtObservaciones.TabIndex = 383;
			this.txtObservaciones.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtObservaciones_KeyPress);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1164, 8);
			this.groupBox1.TabIndex = 385;
			this.groupBox1.TabStop = false;
			// 
			// txtIdIngreso
			// 
			this.txtIdIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance15.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtIdIngreso.Appearance = appearance15;
			this.txtIdIngreso.Enabled = false;
			this.txtIdIngreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdIngreso.FormatString = "";
			this.txtIdIngreso.Location = new System.Drawing.Point(1076, 576);
			this.txtIdIngreso.Name = "txtIdIngreso";
			this.txtIdIngreso.PromptChar = ' ';
			this.txtIdIngreso.Size = new System.Drawing.Size(16, 21);
			this.txtIdIngreso.TabIndex = 386;
			// 
			// txtidCliente
			// 
			this.txtidCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance16.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtidCliente.Appearance = appearance16;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtidCliente.FormatString = "";
			this.txtidCliente.Location = new System.Drawing.Point(1092, 576);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(16, 21);
			this.txtidCliente.TabIndex = 387;
			// 
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance17.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtBodegaPredef.Appearance = appearance17;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtBodegaPredef.FormatString = "";
			this.txtBodegaPredef.Location = new System.Drawing.Point(1108, 576);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(16, 21);
			this.txtBodegaPredef.TabIndex = 388;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance18.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtEstado.Appearance = appearance18;
			this.txtEstado.Enabled = false;
			this.txtEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEstado.FormatString = "";
			this.txtEstado.Location = new System.Drawing.Point(1124, 576);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 389;
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
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(956, 40);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 391;
			// 
			// dtFecha
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance19;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 40);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 392;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			// 
			// frmIngresoDepartamentoTecnico
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1164, 606);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.txtidCliente);
			this.Controls.Add(this.txtIdIngreso);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label39);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label36);
			this.Controls.Add(this.label32);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox5);
			this.Controls.Add(this.txtCelular);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.txtDireccion);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.uGridDTIngreso);
			this.Name = "frmIngresoDepartamentoTecnico";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Ingreso Multiple Departamento Técnico";
			this.Load += new System.EventHandler(this.frmIngresoDepartamentoTecnico_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDTIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void ModificaGrilla(bool bModifica)
		{
			this.uGridDTIngreso.Enabled = true;

			if (bModifica)
			{
				this.uGridDTIngreso.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridDTIngreso.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				this.uGridDTIngreso.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridDTIngreso.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void frmIngresoDepartamentoTecnico_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "1003");

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			ModificaGrilla(false);

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}", MenuLatinium.IdUsuario));

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			this.txtBodegaPredef.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			if ((int)this.txtBodegaPredef.Value == 0) 
			{
				this.cmbBodega.Value = 1;
				this.cmbBodega.Enabled = true;
			}
			else this.cmbBodega.Value = (int)this.txtBodegaPredef.Value;

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;

			this.txtCodigo.Enabled = true;
			this.txtObservaciones.Enabled = true;

			ModificaGrilla(true);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;			

			this.dtFecha.Value = DateTime.Today;			
			this.txtEstado.Value = 0;

			if ((int)this.txtBodegaPredef.Value == 0) this.cmbBodega.Focus();
			else if (miAcceso.CambiarFecha) this.dtFecha.Focus();
			else this.txtCodigo.Focus();

			bNuevo = true;
			bEdicion = true;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			this.txtIdIngreso.Value = 0;
			this.txtidCliente.Value = 0;
			this.txtEstado.Value = 0;
			this.cmbBodega.Value = System.DBNull.Value;
			this.dtFecha.Value = System.DBNull.Value;
			this.lblNumero.Text = "";
			this.lblEstado.Text = "";
			this.txtCodigo.Text = "";
			this.txtNombre.Text = "";
			this.txtDireccion.Text = "";
			this.txtTelefono.Text = "";
			this.txtCelular.Text = "";
			this.txtObservaciones.Text = "";
			
			this.uGridDTIngreso.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTConsultaDetalleIM 0");

			this.cmbBodega.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.txtObservaciones.Enabled = false;

			ModificaGrilla(false);

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Bodega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			  this.cmbBodega.Focus();
			}
			else if (this.txtCodigo.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Numero de Identificación del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();			
			}
			else
			{
				string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', 4)", this.txtCodigo.Text.ToString());
				string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
				if (VlCed.StartsWith("Error"))
				{
					MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCodigo.Focus();
				}				
				else
				{
					string sSQL = string.Format("Exec DTIMValidacliente '{0}'", this.txtCodigo.Text.ToString());
					int iRes = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

					if (iRes == 1)
					{
						MessageBox.Show("El Cliente No Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtCodigo.Focus();
					}
					else if (iRes == 2)
					{
						MessageBox.Show("El Cliente No es Corporativo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtCodigo.Focus();
					}
					else
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDTIngreso.Rows.All)
						{
							if (dr.Cells["Daño"].Value.ToString().Length == 0)
							{
								MessageBox.Show(string.Format("Ingrese el Daño del Articulo '{0}'.", dr.Cells["Articulo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
								this.uGridDTIngreso.ActiveRow = dr;
								this.uGridDTIngreso.ActiveRow.Selected = true;
								return;
							}
						}

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

								#region Numeracion
								if (bNuevo)
								{
									string sSQLNumero = string.Format("Exec NumeracionLocales 53, {0}, 0", (int)this.cmbBodega.Value);
									oCmdActualiza.CommandText = sSQLNumero;
									this.lblNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
								}	
								#endregion Numeracion
			
								#region Graba Registro
								#region Maestro
								DateTime dtFecha = (DateTime)this.dtFecha.Value;																

								string sSQLGuarda = string.Format("Exec DTGuardaIM {0}, '{1}', '{2}', {3}, {4}, '{5}', {6}", 
									(int) this.txtIdIngreso.Value, 
									this.lblNumero.Text.ToString(),
									dtFecha.ToString("yyyyMMdd"),
									(int)this.cmbBodega.Value,
									(int)this.txtidCliente.Value,
									this.txtObservaciones.Text.ToString(),
									(int)this.txtEstado.Value);
								oCmdActualiza.CommandText = sSQLGuarda;
								this.txtIdIngreso.Value = (int)oCmdActualiza.ExecuteScalar();
								#endregion Maestro

								#region Detalle
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDTIngreso.Rows.All)
								{										
									string sSQLDetalle = string.Format("Exec DTGuardaDetalleIM {0}, {1}, {2}, {3}, '{4}', {5}, {6}, {7}, '{8}', '{9}', '{10}', {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}",
										(int)dr.Cells["idDetalleIngreso"].Value, (int)this.txtIdIngreso.Value,
										(int)dr.Cells["idSerial"].Value, (int)dr.Cells["idArticulo"].Value,
										(int)dr.Cells["idDetCompra"].Value, (bool)dr.Cells["Garantia"].Value, 
										(int)dr.Cells["TiempoGarantia"].Value, (int)dr.Cells["Ingresos"].Value,
										dr.Cells["Daño"].Value.ToString(), dr.Cells["Accesorios"].Value.ToString(), 
										dr.Cells["Observaciones"].Value.ToString(), (bool)dr.Cells["Empaque"].Value, (int)dr.Cells["idCompra"].Value,
										(int)dr.Cells["idAccesorio1"].Value, (int)dr.Cells["idAccesorio2"].Value, (int)dr.Cells["idAccesorio3"].Value, (int)dr.Cells["idAccesorio4"].Value,
										(int)dr.Cells["idAccesorio5"].Value, (int)dr.Cells["idAccesorio6"].Value, (int)dr.Cells["idAccesorio7"].Value, (int)dr.Cells["idAccesorio8"].Value,
										(int)dr.Cells["idAccesorio9"].Value, (int)dr.Cells["idAccesorio10"].Value, (int)dr.Cells["idAccesorio11"].Value, (int)dr.Cells["idAccesorio12"].Value,
										(int)dr.Cells["idAccesorio13"].Value, (int)dr.Cells["idAccesorio14"].Value, (int)dr.Cells["idAccesorio15"].Value, (int)dr.Cells["idAccesorio16"].Value,
										(int)dr.Cells["idAccesorio17"].Value, (int)dr.Cells["idAccesorio18"].Value, (int)dr.Cells["idAccesorio19"].Value, (int)dr.Cells["idAccesorio20"].Value);
									oCmdActualiza.CommandText = sSQLDetalle;
									oCmdActualiza.ExecuteNonQuery();
								}
								#endregion Detalle	

								string sSQLActDetalle = string.Format("Exec DTConsultaDetalleIM {0}", (int)this.txtIdIngreso.Value);
								oCmdActualiza.CommandText = sSQLActDetalle;
								SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
								DataTable oDTDetalle = new DataTable();
								oDADetalle.Fill(oDTDetalle);
								this.uGridDTIngreso.DataSource = oDTDetalle;							
								#endregion Graba Factura
					
								oTransaction.Commit();	
								
								#region Controles
								this.cmbBodega.Enabled = false;
								this.dtFecha.Enabled = false;
								this.txtCodigo.Enabled = false;								
								this.txtObservaciones.Enabled = false;

								if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
								if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
								if (miAcceso.Editar) this.btnEditar.Enabled = true;
								if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;								
								if (miAcceso.Anular) this.btnAnular.Enabled = true;																
								this.btnGuardar.Enabled = false;
								this.btnCancelar.Enabled = true;
								
								ModificaGrilla(false);			

								bNuevo = false;
								bEdicion = false;
								#endregion Controles
							}
							catch (Exception ex)
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
					}
				}
			}
		}

		private void txtCodigo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes Busqueda = new BuscaClientes(this.txtCodigo.Text.ToString(), 50, 1, 4, (int) this.cmbBodega.Value, false))
				{
					if (DialogResult.OK == Busqueda.ShowDialog())
					{
						this.txtCodigo.Text = Busqueda.Ruc;

						this.txtObservaciones.Focus();
					}
				}
			}

			if (e.KeyCode == Keys.Enter) this.txtObservaciones.Focus();
		}

		private void txtCodigo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtCodigo_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				string sSQL = string.Format("Exec DTIMBuscaCliente '{0}'", this.txtCodigo.Text.ToString());
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);

				dr.Read();
				if (dr.HasRows)
				{
					this.txtidCliente.Value = dr.GetInt32(0);
					this.txtCodigo.Value = dr.GetString(1);
					this.txtNombre.Value = dr.GetString(2);
					this.txtDireccion.Value = dr.GetString(3);
					this.txtTelefono.Value = dr.GetString(4);
					this.txtCelular.Value = dr.GetString(5);							
				}
				dr.Close();
			}
		}

		private void uGridDTIngreso_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtBodegaPredef.Value == 0) this.cmbBodega.Enabled = true;

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;

			this.txtCodigo.Enabled = true;
			this.txtObservaciones.Enabled = true;

			ModificaGrilla(true);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;			

			if ((int)this.txtBodegaPredef.Value == 0) this.cmbBodega.Focus();
			else if (miAcceso.CambiarFecha) this.dtFecha.Focus();
			else this.txtCodigo.Focus();

			bNuevo = false;
			bEdicion = true;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			btnCancelar_Click(sender, e);

			using (frmDTBuscaIM miBusqueda = new frmDTBuscaIM())
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					try
					{	
						#region Maestro
						this.txtIdIngreso.Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["idIngresoMultiple"].Value;
						this.cmbBodega.Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["Bodega"].Value;
						this.dtFecha.Value = (DateTime)miBusqueda.ultraGrid1.ActiveRow.Cells["Fecha"].Value;
						this.lblNumero.Text = miBusqueda.ultraGrid1.ActiveRow.Cells["Numero"].Value.ToString();
            this.txtidCliente.Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["idCliente"].Value;
						this.txtCodigo.Text = miBusqueda.ultraGrid1.ActiveRow.Cells["Ruc"].Value.ToString();
						this.txtNombre.Text = miBusqueda.ultraGrid1.ActiveRow.Cells["Nombre"].Value.ToString();				
						this.txtDireccion.Text = miBusqueda.ultraGrid1.ActiveRow.Cells["Direccion"].Value.ToString();
						this.txtTelefono.Text = miBusqueda.ultraGrid1.ActiveRow.Cells["Telefono"].Value.ToString();				
						this.txtCelular.Text = miBusqueda.ultraGrid1.ActiveRow.Cells["Celular"].Value.ToString();
						this.txtObservaciones.Text = miBusqueda.ultraGrid1.ActiveRow.Cells["Observaciones"].Value.ToString();
						this.txtEstado.Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["Estado"].Value;						
						#endregion Maestro
						
						this.uGridDTIngreso.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DTConsultaDetalleIM {0}", (int)this.txtIdIngreso.Value));
	
						if (miAcceso.Editar) this.btnEditar.Enabled = true;
						if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;						
						if (miAcceso.Anular) this.btnAnular.Enabled = true;						
						
						this.btnCancelar.Enabled = true;
					}
					catch(Exception Exc)
					{
						MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}		

		}

		private void txtObservaciones_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.uGridDTIngreso.DisplayLayout.Bands[0].AddNew();
				this.uGridDTIngreso.ActiveCell = this.uGridDTIngreso.Rows[0].Cells["Serial"];
			}
		}

		private void txtCodigo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{				
				if (this.txtCodigo.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Numero de Identificación del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCodigo.Focus();
					e.Cancel = true;
				}
				else
				{
					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', 4)", this.txtCodigo.Text.ToString());
					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
					if (VlCed.StartsWith("Error"))
					{
						MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtCodigo.Focus();
						e.Cancel = true;			
					}				
					else
					{
						string sSQL = string.Format("Exec DTIMValidacliente '{0}'", this.txtCodigo.Text.ToString());
						int iRes = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

						if (iRes == 1)
						{
							MessageBox.Show("El Cliente No Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtCodigo.Focus();
							e.Cancel = true;
						}
						else if (iRes == 2)
						{
							MessageBox.Show("El Cliente No es Corporativo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtCodigo.Focus();
							e.Cancel = true;
						}
					}
				}
			}
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (miAcceso.CambiarFecha) this.dtFecha.Focus(); else this.txtCodigo.Focus();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCodigo.Focus();
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value) this.dtFecha.Value = DateTime.Today;
		}

		private void cmbBodega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbBodega.ActiveRow == null)
				{
					MessageBox.Show("Seleccione una Bodega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbBodega.Focus();
					e.Cancel = true;
				}
			}
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridDTIngreso_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Serial")
			{
				if (e.Cell.Row.Cells["Serial"].Value.ToString().Length > 0)
				{
					int iOrigen = 0;
					string sMensaje = "";

					string sSQLValida = string.Format("Exec DTValidaSerialIM '{0}', {1}, '{2}'", e.Cell.Row.Cells["Serial"].Value.ToString(), (int)this.txtidCliente.Value, this.txtCodigo.Text.ToString());
					SqlDataReader drVal = Funcion.rEscalarSQL(cdsSeteoF, sSQLValida, true);
					drVal.Read();
					if (drVal.HasRows)
					{
						iOrigen = drVal.GetInt32(0);
						sMensaje = drVal.GetString(1);
					}
					drVal.Close();

					if (iOrigen == 0)
					{					
						MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
						e.Cell.Row.Cells["Serial"].Value = "";
					}
					else 
					{
						string sSQL = "";
						if (iOrigen == 1) sSQL = string.Format("Exec DTBuscaArticuloSerialIM '{0}', {1}", e.Cell.Row.Cells["Serial"].Value.ToString(), (int)this.txtidCliente.Value);
						if (iOrigen == 2) sSQL = string.Format("Exec DTBuscaArticuloSerialEMIM '{0}', {1}", e.Cell.Row.Cells["Serial"].Value.ToString(), this.txtCodigo.Value);
						
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
						dr.Read();
						if (dr.HasRows)
						{
							e.Cell.Row.Cells["idSerial"].Value = dr.GetInt32(0);
							e.Cell.Row.Cells["idArticulo"].Value = dr.GetInt32(2);
							e.Cell.Row.Cells["Codigo"].Value = dr.GetString(3);
							e.Cell.Row.Cells["idDetCompra"].Value = dr.GetInt32(4);
							e.Cell.Row.Cells["idCompra"].Value = dr.GetInt32(5);
							e.Cell.Row.Cells["Numero"].Value = dr.GetString(6);
							e.Cell.Row.Cells["Fecha"].Value = (DateTime)dr.GetDateTime(7);
							e.Cell.Row.Cells["Garantia"].Value = (bool)dr.GetValue(8);
							e.Cell.Row.Cells["TiempoGarantia"].Value = dr.GetInt32(9);
							e.Cell.Row.Cells["Ingresos"].Value = dr.GetInt32(10);
						}
						dr.Close();
					}
				}
			}
		}

		private void uGridDTIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (uGridDTIngreso.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{							
							#region Flechas
						case (int) Keys.Up:
							if (uGridDTIngreso.ActiveCell.DroppedDown) return;

							uGridDTIngreso.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridDTIngreso.PerformAction(UltraGridAction.AboveCell, false, false);
							e.Handled = true;
							uGridDTIngreso.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Left:
							uGridDTIngreso.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridDTIngreso.PerformAction(UltraGridAction.PrevCellByTab, false, false);
							e.Handled = true;
							uGridDTIngreso.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Down:
							if (uGridDTIngreso.ActiveCell.DroppedDown) return;
			
							uGridDTIngreso.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridDTIngreso.PerformAction(UltraGridAction.BelowCell, false, false);
							e.Handled = true;
							uGridDTIngreso.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
							#endregion Flechas
							#region Enter
						case (int) Keys.Enter:
							uGridDTIngreso.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridDTIngreso.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							uGridDTIngreso.PerformAction(UltraGridAction.EnterEditMode, false, false);
							if(uGridDTIngreso.ActiveCell.DroppedDown == false)
								uGridDTIngreso.ActiveCell.DroppedDown = true;
							//						uGridDTIngreso.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
							break;
							#endregion Enter
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error del sistema.");
				}
			}
		}

		private void uGridDTIngreso_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDetalleIngreso"].Value = 0;
				e.Row.Cells["idIngresoMultiple"].Value = 0;				
				e.Row.Cells["Empaque"].Value = false;	
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

		private void uGridDTIngreso_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bEdicion) 
			{
				string sCol = "";
				string sFiltro = "";
				if (e.Cell.Column.ToString() == "UnboundColumn1")
				{
					for (int j = 17; j <= 36; j++)
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

						int i = 16;
						i = i + iPos;
						for (int k = i ; k <= 36; k++)
						{
							sCol = this.uGridDTIngreso.DisplayLayout.Bands[0].Columns[k].Header.Caption;
							e.Cell.Row.Cells[sCol].Value = 0;							
						}
					}
				}
			}
		}
	}
}

