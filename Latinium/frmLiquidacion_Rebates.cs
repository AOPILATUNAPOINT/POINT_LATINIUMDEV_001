using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmLiquidacion_Rebates.
	/// </summary>
	public class frmLiquidacion_Rebates : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblEstado;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidRebates_Liquidacion;
		private System.Windows.Forms.Label lablEstado1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label label26;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBancos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmLiquidacion_Rebates()
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
		private Acceso miAcceso;
		bool bNuevo = false;
		bool bEdicion = false;	
		private void UnloadMe()
		{
			this.Close();
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRebates_Liquidacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRebates_Liquidacion_Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seleccionar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rebate");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proveedor");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmLiquidacion_Rebates));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRebates_Liquidacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRebates_Liquidacion_Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rebate");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.label8 = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label13 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lablEstado1 = new System.Windows.Forms.Label();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.txtidRebates_Liquidacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label26 = new System.Windows.Forms.Label();
			this.cmbBancos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidRebates_Liquidacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBancos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			this.SuspendLayout();
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(256, 56);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(82, 16);
			this.label8.TabIndex = 384;
			this.label8.Text = "Forma de Pago";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn6.DefaultValue = false;
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn7.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn8.DefaultValue = "";
			ultraDataColumn9.DefaultValue = "";
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn12.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn13.DefaultValue = "";
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
																																 ultraDataColumn13});
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(16, 96);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(37, 16);
			this.label13.TabIndex = 387;
			this.label13.Text = "Desde";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(256, 96);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(34, 16);
			this.label11.TabIndex = 383;
			this.label11.Text = "Hasta";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(776, 16);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 382;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 56);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 381;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(256, 16);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(56, 16);
			this.lblCliente.TabIndex = 380;
			this.lblCliente.Text = "Proveedor";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(1024, 419);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(29, 16);
			this.label1.TabIndex = 379;
			this.label1.Text = "Total";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance1;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(104, 96);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(120, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 386;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// dtHasta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance2;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(360, 96);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 369;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.Location = new System.Drawing.Point(428, 416);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(96, 23);
			this.btnAsiento.TabIndex = 375;
			this.btnAsiento.Text = "&Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(16, 416);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(96, 23);
			this.btnNuevo.TabIndex = 365;
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
			this.btnConsultar.Location = new System.Drawing.Point(121, 416);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(92, 23);
			this.btnConsultar.TabIndex = 377;
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
			this.btnEditar.Location = new System.Drawing.Point(222, 416);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(96, 23);
			this.btnEditar.TabIndex = 376;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(327, 416);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(92, 23);
			this.btnGuardar.TabIndex = 372;
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
			this.btnAnular.Location = new System.Drawing.Point(533, 416);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(92, 23);
			this.btnAnular.TabIndex = 385;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(634, 416);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(96, 23);
			this.btnCancelar.TabIndex = 373;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(739, 416);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(92, 23);
			this.btnSalir.TabIndex = 374;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance3;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "_DetCompra";
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 89;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 109;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.Caption = "Sel";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 34;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 87;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 97;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 192;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance6;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 192;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 244;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance7;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Width = 101;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.VisiblePosition = 11;
			ultraGridColumn13.Width = 296;
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
																										 ultraGridColumn13});
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.White;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 9F;
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
			this.ultraGrid1.Enabled = false;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 176);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1264, 216);
			this.ultraGrid1.TabIndex = 371;
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// cmbFormaPago
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance15;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(360, 53);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(200, 22);
			this.cmbFormaPago.TabIndex = 370;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.ValueChanged += new System.EventHandler(this.cmbFormaPago_ValueChanged);
			// 
			// txtNumero
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance16;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(864, 14);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 21);
			this.txtNumero.TabIndex = 367;
			// 
			// dtFecha
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance17;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton3);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(104, 54);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(120, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 368;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// txtNombre
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance18;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(360, 14);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(352, 21);
			this.txtNombre.TabIndex = 366;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance19;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(1104, 415);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(168, 24);
			this.txtTotal.TabIndex = 378;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(70, 16);
			this.label2.TabIndex = 389;
			this.label2.Text = "Identificación";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance20;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(104, 14);
			this.txtRuc.MaxLength = 13;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(120, 21);
			this.txtRuc.TabIndex = 388;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 128);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(56, 16);
			this.label3.TabIndex = 391;
			this.label3.Text = "Proveedor";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBusqueda
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance21;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Enabled = false;
			this.txtBusqueda.Location = new System.Drawing.Point(104, 126);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(368, 21);
			this.txtBusqueda.TabIndex = 390;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(776, -456);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 392;
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
			// lablEstado1
			// 
			this.lablEstado1.AutoSize = true;
			this.lablEstado1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lablEstado1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lablEstado1.ForeColor = System.Drawing.Color.Firebrick;
			this.lablEstado1.Location = new System.Drawing.Point(992, 11);
			this.lablEstado1.Name = "lablEstado1";
			this.lablEstado1.Size = new System.Drawing.Size(0, 26);
			this.lablEstado1.TabIndex = 408;
			this.lablEstado1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(840, 416);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(16, 21);
			this.txtIdCliente.TabIndex = 409;
			this.txtIdCliente.Visible = false;
			// 
			// btnVer
			// 
			this.btnVer.Enabled = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(496, 125);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(88, 23);
			this.btnVer.TabIndex = 1007;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// txtidRebates_Liquidacion
			// 
			this.txtidRebates_Liquidacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidRebates_Liquidacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidRebates_Liquidacion.Enabled = false;
			this.txtidRebates_Liquidacion.Location = new System.Drawing.Point(856, 416);
			this.txtidRebates_Liquidacion.Name = "txtidRebates_Liquidacion";
			this.txtidRebates_Liquidacion.PromptChar = ' ';
			this.txtidRebates_Liquidacion.Size = new System.Drawing.Size(16, 21);
			this.txtidRebates_Liquidacion.TabIndex = 1008;
			this.txtidRebates_Liquidacion.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(872, 416);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 1009;
			this.txtEstado.Visible = false;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Location = new System.Drawing.Point(584, 56);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(36, 16);
			this.label26.TabIndex = 1011;
			this.label26.Text = "Banco";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBancos
			// 
			this.cmbBancos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBancos.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbBancos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBancos.DisplayMember = "Descripcion";
			this.cmbBancos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBancos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBancos.Enabled = false;
			this.cmbBancos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBancos.Location = new System.Drawing.Point(640, 53);
			this.cmbBancos.Name = "cmbBancos";
			this.cmbBancos.Size = new System.Drawing.Size(216, 22);
			this.cmbBancos.TabIndex = 1010;
			this.cmbBancos.ValueMember = "idCuenta";
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(888, 416);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 21);
			this.txtIdAsiento.TabIndex = 1012;
			this.txtIdAsiento.Visible = false;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(904, 416);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 21);
			this.txtIdCompra.TabIndex = 1013;
			this.txtIdCompra.Visible = false;
			// 
			// frmLiquidacion_Rebates
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1280, 446);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.cmbBancos);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtidRebates_Liquidacion);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.lablEstado1);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.label8);
			this.Name = "frmLiquidacion_Rebates";
			this.Text = "Liquidacion Rebates";
			this.Load += new System.EventHandler(this.frmLiquidacion_Rebates_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidRebates_Liquidacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBancos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(true, 2);
				if (sIdentificacion.Length > 0)
				{						
					this.txtRuc.Text = sIdentificacion;
					
					CargaCliente();					
				}				
			}		
		}
		private void CargaCliente()
		{
			string[] sParametros = new string[11];
			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(true, 2, this.txtRuc.Text.ToString(), cdsSeteoF);
			
			if (sParametros[0] != null) this.txtIdCliente.Value = int.Parse(sParametros[0]);
			if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];
		}

		private void frmLiquidacion_Rebates_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0638");
			
			if (!Funcion.Permiso("2182", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar a LIQUIDACION REBATES ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;

			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF,
				string.Format("SELECT idFormaPago,FormaPago FROM CompraFormaPago where idFormaPago in (4,7,34)"));

			this.cmbBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaCuentasDepositos "));
		}
		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.txtRuc.Enabled = true;

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			DateTime hoy   = DateTime.Today;
			DateTime desde = new DateTime(hoy.Year, hoy.Month, 1);
			DateTime hasta = new DateTime(hoy.Year, hoy.Month, DateTime.DaysInMonth(hoy.Year, hoy.Month));

			this.dtDesde.Value = desde; 
			this.dtHasta.Value = hasta;
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;			
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;		
			this.cmbFormaPago.Enabled = true; 
			this.dtDesde.Enabled = true;
			this.dtHasta.Enabled = true;
			this.btnVer.Enabled = true;
			this.txtBusqueda.Enabled = true;
			this.cmbBancos.Enabled = true; 
			this.txtRuc.Focus();

			bNuevo = true;
			bEdicion = true;
		
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";		
			this.lablEstado1.Text = "";
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFecha.CalendarInfo.MinDate = Funcion.dtEscalarSQL(cdsSeteoF, "Select CONVERT(Date, FechaInicio) From SeteoF");			
			this.txtNumero.Text = "";
			this.txtTotal.Value = 0.00m;		
			this.cmbFormaPago.Value = System.DBNull.Value;  
			this.dtDesde.Value = null;
			this.dtDesde.Text  = string.Empty;
			this.dtHasta.Value = null;
			this.dtHasta.Text  = string.Empty;
			this.txtidRebates_Liquidacion.Value = 0;
			this.cmbBancos.Value = System.DBNull.Value;  
			this.txtIdCompra.Value = 0;
			this.txtIdAsiento.Value = 0;

			this.txtRuc.Enabled = false;
			this.txtNombre.Enabled = false;			
			this.dtFecha.Enabled = false;		
			this.txtNumero.Enabled = false;	
			this.dtFecha.Enabled = false;
			this.cmbFormaPago.Enabled = false; 
			this.txtBusqueda.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			this.cmbBancos.Enabled = false; 

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;			
			this.btnEditar.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnGuardar.Enabled = false;						
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Text = "Anular";
			this.btnVer.Enabled = false;
			bNuevo = false;
			bEdicion = false;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
				
		}

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Seleccionar")
			{
				this.ultraGrid1.UpdateData();
				
				string sSQL = string.Format("Exec InsertaRebatesLiquidacion {0}, {1}, {2},{3},{4}", 
					(int)e.Cell.Row.Cells["idRebates_Liquidacion_Detalle"].Value,(int)e.Cell.Row.Cells["idRebates_Liquidacion"].Value, 
					(int)e.Cell.Row.Cells["idDetCompra"].Value,(decimal)e.Cell.Row.Cells["Rebate"].Value,
					((e.Cell.Row.Cells["Seleccionar"].Value is bool) && (bool)e.Cell.Row.Cells["Seleccionar"].Value) ? 1 : 0);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);									

				Suma();				
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (bNuevo)
			{
				string sSQLNumero = string.Format("Exec NumeracionLocales 101, 0, 0");
				this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, sSQLNumero, true);
			}

			int estado = 1;

			System.Windows.Forms.DialogResult dr = System.Windows.Forms.MessageBox.Show(
				"¿Desea liquidar ahora?","Point Technology", System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question );
			if (dr == System.Windows.Forms.DialogResult.Yes) estado = 2;

			DateTime dtFecha = (DateTime)this.dtFecha.Value;
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			string sSQL = string.Format("Exec GuardaMaestroRebatesLiquidacion {0}, '{1}','{2}','{3}','{4}',{5},{6},{7},'{8}',{9},{10}", 
				(int)this.txtidRebates_Liquidacion.Value,	this.txtNumero.Text.ToString(),dtFecha.ToString("yyyyMMdd"),
				dtDesde.ToString("yyyyMMdd"),dtHasta.ToString("yyyyMMdd"),(int)this.txtIdCliente.Value,(int)cmbFormaPago.Value,
				(decimal)txtTotal.Value,"",1,(int)this.cmbBancos.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			string ruc = (this.txtRuc.Text == null) ? "" : this.txtRuc.Text.Trim();

			if (estado == 2)
			{
				int formaPago = 0;
				if (this.cmbFormaPago.Value != null && this.cmbFormaPago.Value != System.DBNull.Value)
					try { formaPago = Convert.ToInt32(this.cmbFormaPago.Value); } 
					catch { formaPago = 0; }

				if (formaPago == 4)
				{
					using (frmDevolucionCompra f = new frmDevolucionCompra(ruc,(string)this.txtNombre.Value,(decimal)this.txtTotal.Value,true,(int)this.txtidRebates_Liquidacion.Value))
					{
						f.ShowDialog(this);  
					}
				}
				else if (formaPago == 34)
				{					
					if (ruc != "")
					{
						string sSQLCli = string.Format( "SELECT COUNT(1) FROM Cliente WHERE Ruc = '{0}' AND Proveedor = 0",ruc.Replace("'", "''"));

						int  sCount = Funcion.iEscalarSQL(cdsSeteoF, sSQLCli, true);
						int n = 0;
						try { n = sCount; } 
						catch { n = 0; }

						if (n == 0)
						{		
							FuncionesProcedimientos.AbreFacturacionClientes(4, ruc, false, true);
						}
					}
					using (frmVenta f = new frmVenta(ruc,true,(decimal)this.txtTotal.Value,(int)this.txtIdCliente.Value,(string)this.txtNombre.Value,(int)this.txtidRebates_Liquidacion.Value))
					{
						f.ShowDialog(this);
					}
				}
				else if (formaPago == 7)
				{
					string sSQLAsiento = string.Format("Exec AsientoLiquidacionRebates '{0}',{1},{2},{3}", 
						(string)this.txtNumero.Value,0,(int)this.cmbBancos.Value,(decimal)this.txtTotal.Value );			
					this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento, true);

					string sSQLUpd = string.Format("UPDATE Rebates_Liquidacion SET Estado = 2,idAsiento ={0} WHERE idRebates_Liquidacion = {1}",
						(int)this.txtIdAsiento.Value,(int)this.txtidRebates_Liquidacion.Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQLUpd, true);
				}
			}

			this.dtFecha.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;		
			this.txtBusqueda.Enabled = false;
			this.btnVer.Enabled = false;
			this.cmbFormaPago.Enabled = false;
			this.cmbBancos.Enabled = false;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;			
			this.btnAsiento.Enabled = true;	
		
			Consulta((int)txtidRebates_Liquidacion.Value);

			bEdicion = false;
			bNuevo = false;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format( "EXEC BuscaRebatesLiquidacion {0},'{1}','{2}'", (int)this.txtIdCliente.Value, Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}
		private void Suma()
		{
			decimal total = 0m;
			int i;

			for (i = 0; i < this.ultraGrid1.Rows.Count; i++)
			{
				Infragistics.Win.UltraWinGrid.UltraGridRow r = this.ultraGrid1.Rows[i];

				bool sel = false;
				if (r.Cells.Exists("Seleccionar"))
				{
					if (r.Cells["Seleccionar"].Value != null && r.Cells["Seleccionar"].Value != System.DBNull.Value)
					{
						if (r.Cells["Seleccionar"].Value is bool)
						{
							sel = (bool)r.Cells["Seleccionar"].Value;
						}
					}
				}
				if (sel)
				{
					if (r.Cells.Exists("Rebate"))
					{
						if (r.Cells["Rebate"].Value != null && r.Cells["Rebate"].Value != System.DBNull.Value)
						{
							try { total += System.Convert.ToDecimal(r.Cells["Rebate"].Value); } 
							catch {}
						}
					}
				}
			}

			this.txtTotal.Text = total.ToString("N2");
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaLiquidacionRebates Busqueda = new frmBuscaLiquidacionRebates())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.grdFacturas.ActiveRow.Cells["idRebates_Liquidacion"].Value);
				}
			}
		}
		private void Consulta (int idRebates_Liquidacion)
		{
			string sSQL = string.Format("Exec BuscaMaestroLiquidacionRebates {0}", idRebates_Liquidacion);
			SqlDataReader drDeposito = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);

			if (drDeposito != null && drDeposito.Read())
			{                
				this.txtidRebates_Liquidacion.Value = drDeposito.GetInt32(0);
				this.txtRuc.Value = drDeposito.GetString(1);
				this.txtNumero.Text = drDeposito.GetString(2);
				this.dtFecha.Value = drDeposito.GetDateTime(3);
				this.cmbFormaPago.Value = drDeposito.GetInt32(4);
				this.dtDesde.Value = drDeposito.GetDateTime(5);
				this.dtHasta.Value = drDeposito.GetDateTime(6);
				this.txtIdCliente.Value = drDeposito.GetInt32(7);
				this.txtNombre.Text = drDeposito.GetString(8);
				this.txtTotal.Value = drDeposito.GetDecimal(9);
				this.txtEstado.Value = drDeposito.GetInt32(10);
				this.cmbBancos.Value = drDeposito.GetInt32(11);		
				this.txtIdCompra.Value = drDeposito.GetInt32(12);
				this.txtIdAsiento.Value = drDeposito.GetInt32(13);
			}
			drDeposito.Close();
						
			if ((int)this.txtEstado.Value == 1) 
			{					
				this.lablEstado1.Text = "PENDIENTE";				
				this.btnEditar.Enabled = true;				
				this.btnAnular.Enabled = true;				
			}

			if ((int)this.txtEstado.Value == 2) 
			{
				this.lablEstado1.Text = "LIQUIDADO";

				this.dtFecha.Enabled = false;
				this.dtDesde.Enabled = false;
				this.dtHasta.Enabled = false;		
				this.txtBusqueda.Enabled = false;
				this.btnVer.Enabled = false;
				this.cmbFormaPago.Enabled = false;
				this.cmbBancos.Enabled = false;

				this.btnNuevo.Enabled = true;
				this.btnConsultar.Enabled = true;
				this.btnEditar.Enabled = true;
				this.btnGuardar.Enabled = false;			
				if ((int)this.txtIdAsiento.Value > 0)
					this.btnAsiento.Enabled = true;
				else
					this.btnAsiento.Enabled = false;
				this.btnAnular.Enabled = true;	
			}
			if ((int)this.txtEstado.Value == 3) this.lablEstado1.Text = "ANULADO";

			this.btnCancelar.Enabled = true;				
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{		
			this.dtFecha.Enabled = true;
			this.dtDesde.Enabled = true;
			this.dtHasta.Enabled = true;
			this.btnVer.Enabled = true;
			this.txtBusqueda.Enabled = true;
			this.cmbFormaPago.Enabled = true;
			this.cmbBancos.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;			
			this.btnAsiento.Enabled = false;			
			this.btnCancelar.Enabled = true;		
	
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);			

			bNuevo = false;
			bEdicion = true;

			this.dtFecha.Focus();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (!Validacion.vbClienteExiste(this.txtRuc, true, 2, 6, cdsSeteoF)) 
				{
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaGrupoProveedor '{0}'", this.txtRuc.Text.ToString())) == 0)
				{
					MessageBox.Show("En la Ficha Proveedores Seleccione el Grupo al cual Pertenece el Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
			}
		}

		private void cmbFormaPago_ValueChanged(object sender, System.EventArgs e)
		{
			int id = 0;
			if (this.cmbFormaPago.Value != null && this.cmbFormaPago.Value != System.DBNull.Value)
				id = Convert.ToInt32(this.cmbFormaPago.Value);

			bool requiereBanco = (id == 7);

			this.cmbBancos.Visible = requiereBanco;
			if (this.label26 != null) this.label26.Visible = requiereBanco;

			if (requiereBanco)
			{
				if (this.cmbBancos.Value == null || this.cmbBancos.Value == System.DBNull.Value)
					this.cmbBancos.Focus(); 
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{		
			if (this.txtIdCompra == null || this.txtIdCompra.Value == null || this.txtIdCompra.Value == System.DBNull.Value)
			{
				MessageBox.Show("No hay documento para anular.", "Point Technology",
					MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (this.cmbFormaPago == null || this.cmbFormaPago.Value == null)
			{
				MessageBox.Show("No se pudo determinar la forma de pago.", "Point Technology",
					MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			int idCompra  = System.Convert.ToInt32(this.txtIdCompra.Value);
			int idFP = System.Convert.ToInt32(this.cmbFormaPago.Value);

			if (DialogResult.Yes != MessageBox.Show("¿Está seguro de anular este documento \n Esto no puede ser revertido.?",
				"Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				return;

			if (idFP == 4)
			{
				if (Facturacion.AnulacionCompraVenta(idCompra, 6, cdsSeteoF))
				{
					string sSQLLiquidacion = string.Format("UPDATE Rebates_Liquidacion SET Estado = 3 WHERE idRebates_Liquidacion = {0}", (int)this.txtidRebates_Liquidacion.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLLiquidacion);

					string sSQLDelet = string.Format("DELETE FROM Rebates_Liquidacion_Detalle WHERE idRebates_Liquidacion = {0}", (int)this.txtidRebates_Liquidacion.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLDelet);

					MessageBox.Show("Documento Anulado Correctamente", "Point Technology",
						MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				else
				{
					MessageBox.Show("No fue posible anular el documento.", "Point Technology",
						MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				return;
			}

			if (idFP == 34)
			{
				string sSQL = string.Format("SELECT idFE_EstadoComprobante FROM Compra WHERE idCompra = {0}", idCompra);
				int feEstado = Funcion.iEscalarSQL(cdsSeteoF, sSQL, true);

				if (feEstado != 2)
				{
					if (Facturacion.AnulacionCompraVenta(idCompra, 1, cdsSeteoF))
					{
						string sSQLLiquidacion = string.Format("UPDATE Rebates_Liquidacion SET Estado = 3 WHERE idRebates_Liquidacion = {0}",(int)this.txtidRebates_Liquidacion.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLLiquidacion);
						
						string sSQLDelet = string.Format("DELETE FROM Rebates_Liquidacion_Detalle WHERE idRebates_Liquidacion = {0}", (int)this.txtidRebates_Liquidacion.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLDelet);

						MessageBox.Show("Documento Anulado Correctamente", "Point Technology",
							MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					else
					{
						MessageBox.Show("No fue posible anular el documento.", "Point Technology",
							MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
				else
				{
					MessageBox.Show(
						"No se puede anular: el comprobante AUTORIZADO POR EL SRI.",
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
				return;
			}

			if (idFP == 7)
			{
				this.txtIdCompra.Value = 0;

				int idAsiento = (this.txtIdAsiento != null && this.txtIdAsiento.Value != null && this.txtIdAsiento.Value != System.DBNull.Value)
					? System.Convert.ToInt32(this.txtIdAsiento.Value) : 0;

				if (idAsiento == 0)
				{
					MessageBox.Show("No existe asiento para marcar como anulado.", "Point Technology",
						MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				string notaExtra = "ANULACION DESDE LIQUIDACION REBATES - " + MenuLatinium.stUsuario;
				string sSQLAnula = string.Format(
					"UPDATE Asiento SET Borrar = 1, Estado = 3, Notas = ISNULL(Notas, '') + '//' + '{1}' WHERE idAsiento = {0}",idAsiento, notaExtra.Replace("'", "''"));
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAnula);

				string sSQLLiquidacion = string.Format(
					"UPDATE Rebates_Liquidacion set Estado = 3 WHERE idRebates_Liquidacion = {0}", (int)this.txtidRebates_Liquidacion.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLLiquidacion);
			
				string sSQLDelet = string.Format("DELETE FROM Rebates_Liquidacion_Detalle WHERE idRebates_Liquidacion = {0}", (int)this.txtidRebates_Liquidacion.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDelet);

				MessageBox.Show("Liquidacion Anulada.", "Point Technology",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			MessageBox.Show("Forma de pago no soportada para anulación específica.", "Point Technology",
				MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtIdAsiento.Value == 0)
			{
				MessageBox.Show("No existe asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}			
			frmAsientos Asientos = new frmAsientos(((int)this.txtIdAsiento.Value));
			Asientos.MdiParent = MdiParent;
			Asientos.Show();	
		}
	}
}
