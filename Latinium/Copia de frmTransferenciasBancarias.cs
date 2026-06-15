using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmTransferenciasBancarias.
	/// </summary>
	public class frmTransferenciasBancarias : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lblIce;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaOrigen;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDetalleEgreso;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroEgreso;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEgreso;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaDestino;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroIngreso;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIngreso;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDetalleIngreso;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTransferencia;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstadoEgreso;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdIngreso;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdEgreso;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstadoIngreso;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroDeCosto;
		private System.Windows.Forms.Label lblEstadoEgreso;
		private System.Windows.Forms.Label lblEstadoIngreso;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ContextMenu mnuVer;
		private System.Windows.Forms.MenuItem mnuVerEgreso;
		private System.Windows.Forms.MenuItem mnuVerIngreso;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.MenuItem mnuImprimirEgreso;
		private System.Windows.Forms.MenuItem mnuImprimirDeposito;
		private System.Windows.Forms.ContextMenu mnuImprimir;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTransferenciasBancarias()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTransferenciasBancarias));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			this.cmbCuentaOrigen = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label5 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.lblIce = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.cmbCentroDeCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtDetalleEgreso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumeroEgreso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFechaEgreso = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbCuentaDestino = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNumeroIngreso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.dtFechaIngreso = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label8 = new System.Windows.Forms.Label();
			this.txtDetalleIngreso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtEstadoIngreso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTransferencia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstadoEgreso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdIngreso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdEgreso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstadoIngreso = new System.Windows.Forms.Label();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.lblEstadoEgreso = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.mnuVer = new System.Windows.Forms.ContextMenu();
			this.mnuVerEgreso = new System.Windows.Forms.MenuItem();
			this.mnuVerIngreso = new System.Windows.Forms.MenuItem();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.mnuImprimir = new System.Windows.Forms.ContextMenu();
			this.mnuImprimirEgreso = new System.Windows.Forms.MenuItem();
			this.mnuImprimirDeposito = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalleEgreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroEgreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEgreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalleIngreso)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTransferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoEgreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdEgreso)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbCuentaOrigen
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaOrigen.Appearance = appearance1;
			this.cmbCuentaOrigen.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaOrigen.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaOrigen.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 608;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCuentaOrigen.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuentaOrigen.DisplayMember = "Descripcion";
			this.cmbCuentaOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaOrigen.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaOrigen.Enabled = false;
			this.cmbCuentaOrigen.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaOrigen.Location = new System.Drawing.Point(120, 72);
			this.cmbCuentaOrigen.Name = "cmbCuentaOrigen";
			this.cmbCuentaOrigen.Size = new System.Drawing.Size(608, 22);
			this.cmbCuentaOrigen.TabIndex = 314;
			this.cmbCuentaOrigen.ValueMember = "idCuenta";
			this.cmbCuentaOrigen.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCuentaOrigen_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(24, 74);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(37, 16);
			this.label5.TabIndex = 313;
			this.label5.Text = "Origen";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(8, 448);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(728, 8);
			this.groupBox1.TabIndex = 312;
			this.groupBox1.TabStop = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(12, 464);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 304;
			this.btnNuevo.Text = "Nuevo";
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
			this.btnConsultar.Location = new System.Drawing.Point(89, 464);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 305;
			this.btnConsultar.Text = "Consultar";
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
			this.btnEditar.Location = new System.Drawing.Point(172, 464);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 308;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(333, 464);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 306;
			this.btnGuardar.Text = "Guardar";
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
			this.btnAnular.Location = new System.Drawing.Point(499, 464);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 310;
			this.btnAnular.Text = "Anular";
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
			this.btnCancelar.Location = new System.Drawing.Point(580, 464);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 311;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(24, 170);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 16);
			this.label4.TabIndex = 302;
			this.label4.Text = "Detalle";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIce
			// 
			this.lblIce.AutoSize = true;
			this.lblIce.Location = new System.Drawing.Point(24, 235);
			this.lblIce.Name = "lblIce";
			this.lblIce.Size = new System.Drawing.Size(43, 16);
			this.lblIce.TabIndex = 301;
			this.lblIce.Text = "Valor $.";
			this.lblIce.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(544, 42);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(53, 16);
			this.label1.TabIndex = 300;
			this.label1.Text = "Egreso N.";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(248, 42);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(85, 16);
			this.lblBodega.TabIndex = 299;
			this.lblBodega.Text = "Centro de Costo";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(24, 42);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 298;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCentroDeCosto
			// 
			this.cmbCentroDeCosto.CausesValidation = false;
			this.cmbCentroDeCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroDeCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroDeCosto.DataSource = this.ultraDataSource1;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbCentroDeCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCentroDeCosto.DisplayMember = "Nombre";
			this.cmbCentroDeCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroDeCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentroDeCosto.Enabled = false;
			this.cmbCentroDeCosto.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroDeCosto.Location = new System.Drawing.Point(352, 40);
			this.cmbCentroDeCosto.MaxDropDownItems = 30;
			this.cmbCentroDeCosto.Name = "cmbCentroDeCosto";
			this.cmbCentroDeCosto.Size = new System.Drawing.Size(176, 22);
			this.cmbCentroDeCosto.TabIndex = 303;
			this.cmbCentroDeCosto.ValueMember = "idProyecto";
			this.cmbCentroDeCosto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCentroDeCosto_KeyDown);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// txtDetalleEgreso
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDetalleEgreso.Appearance = appearance2;
			this.txtDetalleEgreso.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDetalleEgreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDetalleEgreso.Enabled = false;
			this.txtDetalleEgreso.Location = new System.Drawing.Point(120, 136);
			this.txtDetalleEgreso.Multiline = true;
			this.txtDetalleEgreso.Name = "txtDetalleEgreso";
			this.txtDetalleEgreso.Size = new System.Drawing.Size(608, 88);
			this.txtDetalleEgreso.TabIndex = 296;
			this.txtDetalleEgreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDetalleEgreso_KeyDown);
			// 
			// txtValor
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance3;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(120, 232);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(112, 22);
			this.txtValor.TabIndex = 297;
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			// 
			// txtNumeroEgreso
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumeroEgreso.Appearance = appearance4;
			this.txtNumeroEgreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroEgreso.Enabled = false;
			this.txtNumeroEgreso.Location = new System.Drawing.Point(616, 39);
			this.txtNumeroEgreso.Name = "txtNumeroEgreso";
			this.txtNumeroEgreso.Size = new System.Drawing.Size(112, 22);
			this.txtNumeroEgreso.TabIndex = 295;
			// 
			// dtFechaEgreso
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEgreso.Appearance = appearance5;
			this.dtFechaEgreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaEgreso.CausesValidation = false;
			dateButton1.Caption = "Today";
			this.dtFechaEgreso.DateButtons.Add(dateButton1);
			this.dtFechaEgreso.Enabled = false;
			this.dtFechaEgreso.Format = "dd/MM/yyyy";
			this.dtFechaEgreso.Location = new System.Drawing.Point(120, 40);
			this.dtFechaEgreso.Name = "dtFechaEgreso";
			this.dtFechaEgreso.NonAutoSizeHeight = 23;
			this.dtFechaEgreso.Size = new System.Drawing.Size(112, 21);
			this.dtFechaEgreso.SpinButtonsVisible = true;
			this.dtFechaEgreso.TabIndex = 294;
			this.dtFechaEgreso.Value = ((object)(resources.GetObject("dtFechaEgreso.Value")));
			this.dtFechaEgreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaEgreso_KeyDown);
			// 
			// cmbCuentaDestino
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaDestino.Appearance = appearance6;
			this.cmbCuentaDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaDestino.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaDestino.DataSource = this.ultraDataSource3;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 608;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbCuentaDestino.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCuentaDestino.DisplayMember = "Descripcion";
			this.cmbCuentaDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaDestino.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaDestino.Enabled = false;
			this.cmbCuentaDestino.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaDestino.Location = new System.Drawing.Point(120, 104);
			this.cmbCuentaDestino.Name = "cmbCuentaDestino";
			this.cmbCuentaDestino.Size = new System.Drawing.Size(608, 22);
			this.cmbCuentaDestino.TabIndex = 316;
			this.cmbCuentaDestino.ValueMember = "idCuenta";
			this.cmbCuentaDestino.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCuentaDestino_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(24, 106);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(42, 16);
			this.label2.TabIndex = 315;
			this.label2.Text = "Destino";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(24, 307);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(62, 16);
			this.label3.TabIndex = 318;
			this.label3.Text = "Deposito N.";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumeroIngreso
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumeroIngreso.Appearance = appearance7;
			this.txtNumeroIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroIngreso.Enabled = false;
			this.txtNumeroIngreso.Location = new System.Drawing.Point(104, 304);
			this.txtNumeroIngreso.Name = "txtNumeroIngreso";
			this.txtNumeroIngreso.Size = new System.Drawing.Size(152, 22);
			this.txtNumeroIngreso.TabIndex = 317;
			this.txtNumeroIngreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumeroIngreso_KeyDown);
			this.txtNumeroIngreso.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumeroIngreso_KeyPress);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(288, 307);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(34, 16);
			this.label7.TabIndex = 320;
			this.label7.Text = "Fecha";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaIngreso
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaIngreso.Appearance = appearance8;
			this.dtFechaIngreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaIngreso.CausesValidation = false;
			dateButton2.Caption = "Today";
			this.dtFechaIngreso.DateButtons.Add(dateButton2);
			this.dtFechaIngreso.Enabled = false;
			this.dtFechaIngreso.Format = "dd/MM/yyyy";
			this.dtFechaIngreso.Location = new System.Drawing.Point(344, 305);
			this.dtFechaIngreso.Name = "dtFechaIngreso";
			this.dtFechaIngreso.NonAutoSizeHeight = 23;
			this.dtFechaIngreso.Size = new System.Drawing.Size(112, 21);
			this.dtFechaIngreso.SpinButtonsVisible = true;
			this.dtFechaIngreso.TabIndex = 319;
			this.dtFechaIngreso.Value = ((object)(resources.GetObject("dtFechaIngreso.Value")));
			this.dtFechaIngreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaIngreso_KeyDown);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(24, 96);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(39, 16);
			this.label8.TabIndex = 322;
			this.label8.Text = "Detalle";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDetalleIngreso
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDetalleIngreso.Appearance = appearance9;
			this.txtDetalleIngreso.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDetalleIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDetalleIngreso.Enabled = false;
			this.txtDetalleIngreso.Location = new System.Drawing.Point(96, 64);
			this.txtDetalleIngreso.Multiline = true;
			this.txtDetalleIngreso.Name = "txtDetalleIngreso";
			this.txtDetalleIngreso.Size = new System.Drawing.Size(624, 88);
			this.txtDetalleIngreso.TabIndex = 321;
			this.txtDetalleIngreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDetalleIngreso_KeyDown);
			// 
			// btnVer
			// 
			this.btnVer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Enabled = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(414, 464);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(80, 23);
			this.btnVer.TabIndex = 323;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(657, 464);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 324;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.txtEstadoIngreso);
			this.groupBox2.Controls.Add(this.txtIdTransferencia);
			this.groupBox2.Controls.Add(this.txtEstadoEgreso);
			this.groupBox2.Controls.Add(this.txtIdIngreso);
			this.groupBox2.Controls.Add(this.txtIdEgreso);
			this.groupBox2.Controls.Add(this.txtDetalleIngreso);
			this.groupBox2.Controls.Add(this.lblEstadoIngreso);
			this.groupBox2.Controls.Add(this.label8);
			this.groupBox2.Location = new System.Drawing.Point(8, 272);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(728, 168);
			this.groupBox2.TabIndex = 325;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Ingreso";
			// 
			// txtEstadoIngreso
			// 
			this.txtEstadoIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstadoIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstadoIngreso.Enabled = false;
			this.txtEstadoIngreso.Location = new System.Drawing.Point(72, 128);
			this.txtEstadoIngreso.Name = "txtEstadoIngreso";
			this.txtEstadoIngreso.PromptChar = ' ';
			this.txtEstadoIngreso.Size = new System.Drawing.Size(16, 22);
			this.txtEstadoIngreso.TabIndex = 271;
			this.txtEstadoIngreso.Visible = false;
			// 
			// txtIdTransferencia
			// 
			this.txtIdTransferencia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTransferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTransferencia.Enabled = false;
			this.txtIdTransferencia.Location = new System.Drawing.Point(8, 128);
			this.txtIdTransferencia.Name = "txtIdTransferencia";
			this.txtIdTransferencia.PromptChar = ' ';
			this.txtIdTransferencia.Size = new System.Drawing.Size(16, 22);
			this.txtIdTransferencia.TabIndex = 270;
			this.txtIdTransferencia.Visible = false;
			// 
			// txtEstadoEgreso
			// 
			this.txtEstadoEgreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstadoEgreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstadoEgreso.Enabled = false;
			this.txtEstadoEgreso.Location = new System.Drawing.Point(56, 128);
			this.txtEstadoEgreso.Name = "txtEstadoEgreso";
			this.txtEstadoEgreso.PromptChar = ' ';
			this.txtEstadoEgreso.Size = new System.Drawing.Size(16, 22);
			this.txtEstadoEgreso.TabIndex = 269;
			this.txtEstadoEgreso.Visible = false;
			// 
			// txtIdIngreso
			// 
			this.txtIdIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdIngreso.Enabled = false;
			this.txtIdIngreso.Location = new System.Drawing.Point(40, 128);
			this.txtIdIngreso.Name = "txtIdIngreso";
			this.txtIdIngreso.PromptChar = ' ';
			this.txtIdIngreso.Size = new System.Drawing.Size(16, 22);
			this.txtIdIngreso.TabIndex = 268;
			this.txtIdIngreso.Visible = false;
			// 
			// txtIdEgreso
			// 
			this.txtIdEgreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdEgreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdEgreso.Enabled = false;
			this.txtIdEgreso.Location = new System.Drawing.Point(24, 128);
			this.txtIdEgreso.Name = "txtIdEgreso";
			this.txtIdEgreso.PromptChar = ' ';
			this.txtIdEgreso.Size = new System.Drawing.Size(16, 22);
			this.txtIdEgreso.TabIndex = 267;
			this.txtIdEgreso.Visible = false;
			// 
			// lblEstadoIngreso
			// 
			this.lblEstadoIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstadoIngreso.AutoSize = true;
			this.lblEstadoIngreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstadoIngreso.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstadoIngreso.Location = new System.Drawing.Point(464, 33);
			this.lblEstadoIngreso.Name = "lblEstadoIngreso";
			this.lblEstadoIngreso.Size = new System.Drawing.Size(0, 26);
			this.lblEstadoIngreso.TabIndex = 242;
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.lblEstadoEgreso);
			this.groupBox3.Location = new System.Drawing.Point(8, 8);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(728, 264);
			this.groupBox3.TabIndex = 326;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Egreso";
			// 
			// lblEstadoEgreso
			// 
			this.lblEstadoEgreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstadoEgreso.AutoSize = true;
			this.lblEstadoEgreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstadoEgreso.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstadoEgreso.Location = new System.Drawing.Point(240, 226);
			this.lblEstadoEgreso.Name = "lblEstadoEgreso";
			this.lblEstadoEgreso.Size = new System.Drawing.Size(0, 26);
			this.lblEstadoEgreso.TabIndex = 241;
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
			// mnuVer
			// 
			this.mnuVer.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																					 this.mnuVerEgreso,
																																					 this.mnuVerIngreso});
			// 
			// mnuVerEgreso
			// 
			this.mnuVerEgreso.Enabled = false;
			this.mnuVerEgreso.Index = 0;
			this.mnuVerEgreso.Text = "Egreso";
			this.mnuVerEgreso.Click += new System.EventHandler(this.mnuVerEgreso_Click);
			// 
			// mnuVerIngreso
			// 
			this.mnuVerIngreso.Enabled = false;
			this.mnuVerIngreso.Index = 1;
			this.mnuVerIngreso.Text = "Depósito";
			this.mnuVerIngreso.Click += new System.EventHandler(this.mnuVerIngreso_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(253, 464);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.TabIndex = 327;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// mnuImprimir
			// 
			this.mnuImprimir.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuImprimirEgreso,
																																								this.mnuImprimirDeposito});
			// 
			// mnuImprimirEgreso
			// 
			this.mnuImprimirEgreso.Enabled = false;
			this.mnuImprimirEgreso.Index = 0;
			this.mnuImprimirEgreso.Text = "Egreso";
			this.mnuImprimirEgreso.Click += new System.EventHandler(this.mnuImprimirEgreso_Click);
			// 
			// mnuImprimirDeposito
			// 
			this.mnuImprimirDeposito.Enabled = false;
			this.mnuImprimirDeposito.Index = 1;
			this.mnuImprimirDeposito.Text = "Depósito";
			this.mnuImprimirDeposito.Click += new System.EventHandler(this.mnuImprimirDeposito_Click);
			// 
			// frmTransferenciasBancarias
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(744, 494);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.lblIce);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtFechaIngreso);
			this.Controls.Add(this.txtNumeroIngreso);
			this.Controls.Add(this.cmbCuentaDestino);
			this.Controls.Add(this.cmbCuentaOrigen);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.cmbCentroDeCosto);
			this.Controls.Add(this.txtDetalleEgreso);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.txtNumeroEgreso);
			this.Controls.Add(this.dtFechaEgreso);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox3);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmTransferenciasBancarias";
			this.Text = "Transferencias Entre Bancos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmTransferenciasBancarias_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmTransferenciasBancarias_Closing);
			this.Load += new System.EventHandler(this.frmTransferenciasBancarias_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalleEgreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroEgreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEgreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalleIngreso)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTransferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoEgreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdEgreso)).EndInit();
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;
		int idBloqueaTransacciones = 0;

		private void UnloadMe()
		{
			this.Close();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void Listas()
		{
			this.cmbCuentaOrigen.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select c.idCuenta, c.Descripcion From Cuenta c Inner Join CuentaCorriente cc On cc.idCuenta = c.idCuenta Order By c.Descripcion");
			this.cmbCuentaDestino.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select c.idCuenta, c.Descripcion From Cuenta c Inner Join CuentaCorriente cc On cc.idCuenta = c.idCuenta Order By c.Descripcion");

			this.cmbCentroDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order By Nombre");
		}

		private void frmTransferenciasBancarias_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "040805");
				
			if (!Funcion.Permiso("467", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Transferencias Entre Bancos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				Listas();

				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (!bNuevo && bEdicion) 
				FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, false, (int)this.txtIdTransferencia.Value, 9, cdsSeteoF);

			this.txtIdTransferencia.Value = 0;
			this.txtIdEgreso.Value = 0;
			this.txtIdIngreso.Value = 0;
			this.txtEstadoEgreso.Value = 0;
			this.txtEstadoIngreso.Value = 0;
			
			this.dtFechaEgreso.Value = System.DBNull.Value;
			this.cmbCentroDeCosto.Value = System.DBNull.Value;
			this.txtNumeroEgreso.Text = "";
			this.cmbCuentaOrigen.Value = System.DBNull.Value;
			this.cmbCuentaDestino.Value = System.DBNull.Value;
			this.txtDetalleEgreso.Text = "";
			this.txtValor.Value = 0;
						
			this.txtNumeroIngreso.Text = "";
			this.dtFechaIngreso.Value = System.DBNull.Value;
			this.txtDetalleIngreso.Text = "";

			this.lblEstadoEgreso.Text = "";
			this.lblEstadoIngreso.Text = "";

			this.dtFechaEgreso.Enabled = false;
			this.cmbCentroDeCosto.Enabled = false;
			this.txtNumeroEgreso.Enabled = false;
			this.cmbCuentaOrigen.Enabled = false;
			this.cmbCuentaDestino.Enabled = false;
			this.txtDetalleEgreso.Enabled = false;
			this.txtValor.Enabled = false;
						
			this.txtNumeroIngreso.Enabled = false;
			this.dtFechaIngreso.Enabled = false;
			this.txtDetalleIngreso.Enabled = false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;			
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnVer.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;

			idBloqueaTransacciones = 0;
			bEdicion = false;
			bNuevo = false;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFechaEgreso, cdsSeteoF, false);
			//this.dtFechaEgreso.Value = DateTime.Today;			

			this.cmbCentroDeCosto.Enabled = true;			
			this.cmbCuentaOrigen.Enabled = true;
			this.cmbCuentaDestino.Enabled = true;
			this.txtDetalleEgreso.Enabled = true;
			this.txtValor.Enabled = true;
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;	
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnVer.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			
			this.lblEstadoEgreso.Text = "PENDIENTE";			

			bNuevo = true;
			bEdicion = true;
			
			if (miAcceso.CambiarFecha) 
			{
				this.dtFechaEgreso.Enabled = true;
				this.dtFechaEgreso.Focus();
			}
			else this.cmbCentroDeCosto.Focus();
		}

		private void dtFechaEgreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCentroDeCosto.Focus();
		}

		private void cmbCentroDeCosto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCuentaOrigen.Focus();
		}

		private void cmbCuentaOrigen_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCuentaDestino.Focus();
		}

		private void cmbCuentaDestino_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDetalleEgreso.Focus();
		}

		private void txtDetalleEgreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtValor.Focus();
		}

		private void txtValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if ((int)this.txtEstadoEgreso.Value == 2) this.txtNumeroIngreso.Focus();
				else this.btnGuardar.Focus();
			}			
		}

		private void txtNumeroIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFechaIngreso.Focus();
		}

		private void dtFechaIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDetalleIngreso.Focus();
		}

		private void txtDetalleIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdTransferencia.Value);

			#region Validacion
			if ((int)this.txtEstadoEgreso.Value == 4)
			{
				MessageBox.Show("La Transferencia esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstadoEgreso.Value == 1)
			{
				MessageBox.Show("La Transferencia esta autorizada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}						
			if ((int)this.txtEstadoEgreso.Value == 3)
			{
				MessageBox.Show("El valor de la Transferencia ya esta depositado en la cuenta destino", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.ValidaBloqueo((int)this.txtIdTransferencia.Value, 9, cdsSeteoF)) return;
			#endregion Validacion

			if ((int)this.txtEstadoEgreso.Value == 0)
			{
				this.dtFechaEgreso.Enabled = true;
				this.cmbCentroDeCosto.Enabled = true;				
				this.cmbCuentaOrigen.Enabled = true;
				this.cmbCuentaDestino.Enabled = true;
				this.txtDetalleEgreso.Enabled = true;
				this.txtValor.Enabled = true;				

				this.dtFechaEgreso.Focus();
			}
			else if ((int)this.txtEstadoEgreso.Value == 2)
			{				
				this.dtFechaIngreso.Enabled = true;
				this.txtNumeroIngreso.Enabled = true;
				this.txtDetalleIngreso.Enabled = true;			
				
				this.dtFechaIngreso.Focus();
			}			

			FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, true, (int)this.txtIdTransferencia.Value, 9, cdsSeteoF);
															
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnVer.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;

			bNuevo = false;
			bEdicion = true;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion Datos Egreso
			if ((int)this.txtEstadoEgreso.Value == 0)
			{
				if (!Validacion.vbFechaEnDocumentos(this.dtFechaEgreso, "Seleccione la Fecha del Egreso", true, "Bancos", cdsSeteoF)) return;
			
				if (!Validacion.vbComboVacio(this.cmbCentroDeCosto, "Seleccione el Centro de Costo")) return;

				if (!Validacion.vbComboVacio(this.cmbCuentaOrigen, "Seleccione la Cuenta de Origen")) return;

				if (!Validacion.vbComboVacio(this.cmbCuentaDestino, "Seleccione la Cuenta Destino")) return;		

				if ((int)this.cmbCuentaOrigen.Value == (int)this.cmbCuentaDestino.Value)
				{
					MessageBox.Show("La Cuenta Destino no puede ser igual a la Cuenta Origen", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbCuentaDestino.Focus();
					return;
				}				
				
				if (!Validacion.vbTexto(this.txtDetalleEgreso, 20, 100, "Detalle Egreso")) return;
			}
			#endregion Validacion Datos Egreso

			#region Validacion Datos Ingreso
			if ((int)this.txtEstadoEgreso.Value == 2)
			{
				if (!Validacion.vbFechaEnDocumentos(this.dtFechaEgreso, "Seleccione la Fecha del Egreso", true, "Bancos", cdsSeteoF)) return;

				if (!Validacion.vbTexto(this.txtDetalleEgreso, 20, 100, "Detalle Egreso")) return;

				if (!FuncionesContabilidad.vbValidaDocumentoBancario((int)this.cmbCuentaDestino.Value, this.txtNumeroIngreso, 2, (int)this.txtIdIngreso.Value, cdsSeteoF)) return;

				if (!FuncionesContabilidad.vBancosValidaEstadoConciliacion((int)this.cmbCuentaDestino.Value, (DateTime)this.dtFechaIngreso.Value, cdsSeteoF)) return;

				if (!Validacion.vbTexto(this.txtDetalleEgreso, 20, 100, "Detalle Egreso")) return;
			}
			#endregion Validacion Datos Ingreso

			DateTime FechaEgreso = (DateTime)this.dtFechaEgreso.Value;
			DateTime FechaIngreso = DateTime.Today;
			bool bFechaIng = false;

			if (this.dtFechaIngreso.Value != System.DBNull.Value)
			{
				bFechaIng = true;
				FechaIngreso = (DateTime)this.dtFechaIngreso.Value;
			}
			
			string sSQL = string.Format("Exec GuardaTransferenciasEntreBancos {0}, {1}, {2}, {3}, '{4}', {5}, {6}, {7}, '{8}', {9}, '{10}', '{11}', '{12}', {13}",
				(int)this.txtIdTransferencia.Value, (int)this.txtIdIngreso.Value, (int)this.txtEstadoEgreso.Value, (int)this.txtEstadoIngreso.Value, FechaEgreso.ToString("yyyyMMdd"), (int)this.cmbCentroDeCosto.Value,
				(int)this.cmbCuentaOrigen.Value, (int)this.cmbCuentaDestino.Value, this.txtDetalleEgreso.Text.ToString(), Convert.ToDecimal(this.txtValor.Value), this.txtNumeroIngreso.Text.ToString(), FechaIngreso.ToString("yyyyMMdd"), 
				this.txtDetalleIngreso.Text.ToString(), bFechaIng);
			this.txtIdTransferencia.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			#region Autorización de Pagos
			if ((int)this.txtEstadoEgreso.Value == 0)
			{												
				string sSQLAutorizacion = string.Format("Exec CreaAutorizacionesDePago 12, {0}, '{1}', '{2}', {3}, 2, '1 DE 1', '{4}', {5}", 
					(int)this.txtIdTransferencia.Value, FechaEgreso.ToString("yyyyMMdd"), FechaEgreso.ToString("yyyyMMdd"), Convert.ToDecimal(this.txtValor.Value), this.txtDetalleEgreso.Text.ToString(), (int)this.cmbCuentaOrigen.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAutorizacion, true);
			}
			#endregion Autorización de Pagos

			#region Asiento Ingreso
			if ((int)this.txtEstadoEgreso.Value == 2)
			{
				if (this.txtNumeroIngreso.Text.ToString().Length > 0 && this.dtFechaIngreso.Value != System.DBNull.Value && this.txtNumeroIngreso.Text.ToString().Length > 0)
				{					
					string sSQLAsiento = string.Format("Exec AsientoIngresoTEB {0}", (int)this.txtIdTransferencia.Value); 
					this.txtIdIngreso.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento);

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasEntreBancos Set EstadoEgreso = 3 Where idTransferencia = {0}", (int)this.txtIdTransferencia.Value));
				}
			}
			#endregion Asiento Ingreso

			FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, false, (int)this.txtIdTransferencia.Value, 9, cdsSeteoF);

			this.dtFechaEgreso.Enabled = false;
			this.cmbCentroDeCosto.Enabled = false;
			this.txtNumeroEgreso.Enabled = false;
			this.cmbCuentaOrigen.Enabled = false;
			this.cmbCuentaDestino.Enabled = false;
			this.txtDetalleEgreso.Enabled = false;
			this.txtValor.Enabled = false;
			
			this.txtNumeroIngreso.Enabled = false;
			this.dtFechaIngreso.Enabled = false;
			this.txtDetalleIngreso.Enabled = false;

			this.Consulta((int)this.txtIdTransferencia.Value);

			bNuevo = false;
			bEdicion = false;
		}

		private void Consulta(int idTransferencia)
		{
			try
			{	
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;	
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnVer.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;

				this.mnuImprimirEgreso.Enabled = false;
				this.mnuImprimirDeposito.Enabled = false;

				this.mnuVerEgreso.Enabled = false;
				this.mnuVerIngreso.Enabled = false;

				string sSQL = string.Format("Exec ConsultaIndividualTransferenciasBancarias {0}", idTransferencia);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdTransferencia.Value = idTransferencia;
					this.cmbCentroDeCosto.Value = dr.GetInt32(1);
					this.txtIdEgreso.Value = dr.GetInt32(2);
					this.txtEstadoEgreso.Value = dr.GetInt32(3);
					this.dtFechaEgreso.Value = dr.GetDateTime(4);
					this.txtNumeroEgreso.Text = dr.GetString(5);
					this.cmbCuentaOrigen.Value = dr.GetInt32(6);
					this.txtDetalleEgreso.Text = dr.GetString(7);

					this.txtIdIngreso.Value = dr.GetInt32(8);
					this.txtEstadoIngreso.Value = dr.GetInt32(9);
					if (dr.GetValue(10) != System.DBNull.Value) this.dtFechaIngreso.Value = dr.GetDateTime(10);
					this.txtNumeroIngreso.Text = dr.GetString(11);
					this.cmbCuentaDestino.Value = dr.GetInt32(12);
					this.txtDetalleIngreso.Text = dr.GetString(13);
					this.txtValor.Value = dr.GetDecimal(14);
					this.lblEstadoEgreso.Text = dr.GetString(15);
				}
				dr.Close();

				if ((int)this.txtEstadoEgreso.Value == 0)
				{
					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					if (miAcceso.Anular) this.btnAnular.Enabled = true;
//					if (miAcceso.Imprimir){
////				this.btnImprimir.Enabled = true;
//					this.mnuImprimirEgreso.Enabled = true;
//				} 
				}

				if ((int)this.txtEstadoEgreso.Value == 2)
				{
					if (miAcceso.Editar) this.btnEditar.Enabled = true;

					if (miAcceso.Imprimir)
					{
						this.btnImprimir.Enabled = true;
						this.mnuImprimirEgreso.Enabled = true;
					}					

					if (miAcceso.CrearAsiento) 
					{
						this.btnVer.Enabled = true;
						this.mnuVerEgreso.Enabled = true;
					}
				}

				if ((int)this.txtEstadoEgreso.Value == 3) 
				{
					if (miAcceso.Imprimir)
					{
						this.btnImprimir.Enabled = true;
						this.mnuImprimirEgreso.Enabled = true;
						this.mnuImprimirDeposito.Enabled = true;
					}		
			
					if (miAcceso.CrearAsiento) 
					{
						this.btnVer.Enabled = true;
						this.mnuVerEgreso.Enabled = true;
						this.mnuVerIngreso.Enabled = true;
					}
				}
				
				this.btnCancelar.Enabled = true;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaTransferenciasEntreBancos Busqueda = new frmBuscaTransferenciasEntreBancos())
			{
				if (Busqueda.ShowDialog() == DialogResult.OK)
				{
					this.Consulta((int)Busqueda.ultraGrid1.ActiveRow.Cells["idTransferencia"].Value);						
				}
			}
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtValor_Enter(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void mnuVerEgreso_Click(object sender, System.EventArgs e)
		{	
			this.Consulta((int)this.txtIdTransferencia.Value);

			#region Validacion
			if ((int)this.txtEstadoEgreso.Value == 4)
			{
				MessageBox.Show("La Transferencia esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}			
			if ((int)this.txtEstadoEgreso.Value == 1)
			{
				MessageBox.Show("No se ha generado el egreso de esta transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene acceso a ver el egreso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion

			Cursor = Cursors.WaitCursor;
			frmAsientos Asiento = new frmAsientos(((int)this.txtIdEgreso.Value));
			Asiento.MdiParent = MdiParent;
			Asiento.Show();
			Cursor = Cursors.Default;			
		}

		private void mnuVerIngreso_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdTransferencia.Value);

			#region Validacion
			if ((int)this.txtEstadoEgreso.Value == 4)
			{
				MessageBox.Show("La Transferencia esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}	
			if ((int)this.txtEstadoEgreso.Value != 3)
			{
				MessageBox.Show("No se ha generado el deposito de esta transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene acceso a ver el ingreso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion

			Cursor = Cursors.WaitCursor;
			frmAsientos Asiento = new frmAsientos(((int)this.txtIdIngreso.Value));
			Asiento.MdiParent = MdiParent;
			Asiento.Show();
			Cursor = Cursors.Default;			
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if ((int)this.txtEstadoEgreso.Value == 4)
			{
				MessageBox.Show("La Transferencia esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (!miAcceso.Anular)
			{
				MessageBox.Show("No tiene acceso a Anular", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstadoEgreso.Value == 1)
			{
				MessageBox.Show("La Transferencia esta autorizada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}						
			if ((int)this.txtEstadoEgreso.Value == 2)
			{
				MessageBox.Show("El egreso de esta Transferencia esta generado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstadoEgreso.Value == 3)
			{
				MessageBox.Show("El valor de la Transferencia ya esta depositado en la cuenta destino", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.ValidaBloqueo((int)this.txtIdTransferencia.Value, 9, cdsSeteoF)) return;
			#endregion Validacion

			if (DialogResult.No == MessageBox.Show("¿Esta seguro de Anular esta Transferencia?", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2)) return;

			using (frmJustAnulaVenta Nota = new frmJustAnulaVenta())
			{
				if (DialogResult.OK == Nota.ShowDialog())
				{
					try
					{
						string sSQL = string.Format("Exec AnulaTransferenciasBancarias {0}, {1}, {2}, '{3}'", 
							(int)this.txtIdTransferencia.Value, (int)this.txtIdEgreso.Value, (int)this.txtIdIngreso.Value, Nota.txtNotas.Text.ToString());
						Funcion.EjecutaSQL(cdsSeteoF, sSQL);

						MessageBox.Show("Registro Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

						this.Consulta((int)this.txtIdTransferencia.Value);
					}
					catch(Exception Exc)
					{
						MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
		}

		private void txtNumeroIngreso_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdTransferencia.Value);

			#region Validacion
			if ((int)this.txtEstadoEgreso.Value == 4)
			{
				MessageBox.Show("La Transferencia esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene acceso a ver los movimientos de esta transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstadoEgreso.Value == 1)
			{
				MessageBox.Show("No se ha generado el egreso de esta transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion

			this.mnuVer.Show(this.btnVer, new Point(1, 23));
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdTransferencia.Value);

			#region Validacion
			if ((int)this.txtEstadoEgreso.Value == 4)
			{
				MessageBox.Show("La Transferencia esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstadoEgreso.Value == 1)
			{
				MessageBox.Show("No se ha generado el egreso de esta transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion

			this.mnuImprimir.Show(this.btnImprimir, new Point(1, 23));
		}

		private void mnuImprimirEgreso_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdTransferencia.Value);

			#region Validacion
			if ((int)this.txtEstadoEgreso.Value == 4)
			{
				MessageBox.Show("La Transferencia esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}		
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}		
			if ((int)this.txtEstadoEgreso.Value == 1)
			{
				MessageBox.Show("No se ha generado el egreso de esta transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion

			Cursor = Cursors.WaitCursor;

			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfidAsiento = new ParameterField ();
			ParameterDiscreteValue ValIdAsiento = new ParameterDiscreteValue ();
			pfidAsiento.ParameterFieldName = "@idAsiento";
			ValIdAsiento.Value =  (int)this.txtIdEgreso.Value;
			pfidAsiento.CurrentValues.Add (ValIdAsiento);
			paramFields.Add (pfidAsiento);
				
			ParameterField pfidTipo = new ParameterField ();
			ParameterDiscreteValue ValidTipo = new ParameterDiscreteValue ();
			pfidTipo.ParameterFieldName = "@idTipo";
			ValidTipo.Value = 12;
			pfidTipo.CurrentValues.Add (ValidTipo);
			paramFields.Add (pfidTipo);
				
			ParameterField pfidAsientoDet = new ParameterField ();
			ParameterDiscreteValue ValIdAsientoDet = new ParameterDiscreteValue ();
			pfidAsientoDet.ParameterFieldName = "@idAsientoDet";
			ValIdAsientoDet.Value =  (int)this.txtIdEgreso.Value;
			pfidAsientoDet.CurrentValues.Add (ValIdAsientoDet);
			paramFields.Add (pfidAsientoDet);
				
			Reporte miRepor = new Reporte("ComprobanteDeEgreso.rpt", "");
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("Egreso");
			miRepor.crVista.ParameterFieldInfo = paramFields;
			miRepor.Show();	

			Cursor = Cursors.Default;
		}

		private void mnuImprimirDeposito_Click(object sender, System.EventArgs e)
		{			
			this.Consulta((int)this.txtIdTransferencia.Value);

			#region Validacion
			if ((int)this.txtEstadoEgreso.Value == 4)
			{
				MessageBox.Show("La Transferencia esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}			
			if ((int)this.txtEstadoEgreso.Value != 3)
			{
				MessageBox.Show("No se ha generado el deposito de esta transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion

			string strFiltro = "{Asiento.idAsiento} = " + this.txtIdIngreso.Value.ToString();
			Reporte miReporte = new Reporte("Asientos.rpt", strFiltro);			
			miReporte.MdiParent = MdiParent;
			miReporte.Show();
		}

		private void frmTransferenciasBancarias_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);				
				if (e.KeyCode == Keys.E) if ((int)this.txtEstadoEgreso.Value == 0) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
				if (e.KeyCode == Keys.I) if ((int)this.txtEstadoEgreso.Value == 2 || (int)this.txtEstadoEgreso.Value == 3) if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
				if (e.KeyCode == Keys.A) if ((int)this.txtEstadoEgreso.Value == 0) if (miAcceso.Anular) if (this.btnAnular.Enabled) this.btnAnular_Click(sender, e);
				if (e.KeyCode == Keys.F5)
				{
					miAcceso = new Acceso(cdsSeteoF, "040805");
					Listas();
					this.Consulta((int)this.txtIdTransferencia.Value);
				}	

				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
				if (e.KeyCode == Keys.F5) Listas();
			}
		}

		private void frmTransferenciasBancarias_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una Transferencia Bancaria", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}
	}
}

