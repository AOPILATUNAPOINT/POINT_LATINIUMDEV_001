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
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmControlMantenimiento.
	/// </summary>
	public class frmControlMantenimiento : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkInventario;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnFacrutaCompra;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTrabajosExteriores;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEquipos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCuartoMaquinas;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCocina;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObserbaciones;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProveedor;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTrabajosInteriores;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipoMantenimiento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbHabitaciones;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempo;
		private System.Windows.Forms.Button btnAnular;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private System.Windows.Forms.Button btnanticiposPendientes;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnOrdenCocina;
		private System.Windows.Forms.Button btnClaveDescuento;
		private System.Windows.Forms.Button btnExtras;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVideo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCajero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredefinida;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidControlMantenimiento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcionTrabajo;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTiposdeTrabajo;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.GroupBox groupBox6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorCompra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumFactura;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCompra;
	
		public enum Trabajos 
		{
			Ninguno = 0,
			TrabajosExteriores = 1 << 0,
			TrabajosInteriores = 1 << 1,
			Cocina = 1 << 2,
			Equipos = 1 << 3,
			CuartosMaquinas = 1 << 4
		}
		
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmControlMantenimiento()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmControlMantenimiento));
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("nombreTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nombreTrabajo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			this.label19 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.optTipoMantenimiento = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnFacrutaCompra = new System.Windows.Forms.Button();
			this.txtObserbaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.chkInventario = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.label9 = new System.Windows.Forms.Label();
			this.chkTrabajosExteriores = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.chkTrabajosInteriores = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.chkEquipos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.chkCuartoMaquinas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCocina = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtProveedor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label4 = new System.Windows.Forms.Label();
			this.txtDescripcionTrabajo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbHabitaciones = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label14 = new System.Windows.Forms.Label();
			this.txtidControlMantenimiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbTiposdeTrabajo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label8 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.groupBox6 = new System.Windows.Forms.GroupBox();
			this.txtValorCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.optTipoMantenimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObserbaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcionTrabajo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHabitaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidControlMantenimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiposdeTrabajo)).BeginInit();
			this.groupBox2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.groupBox5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).BeginInit();
			this.SuspendLayout();
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(184, 26);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(43, 16);
			this.label19.TabIndex = 1054;
			this.label19.Text = "Bodega";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label19.Click += new System.EventHandler(this.label19_Click);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 60);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(89, 16);
			this.label6.TabIndex = 1074;
			this.label6.Text = "Tipos de Trabajo";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optTipoMantenimiento
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipoMantenimiento.Appearance = appearance1;
			this.optTipoMantenimiento.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTipoMantenimiento.CheckedIndex = 0;
			this.optTipoMantenimiento.Enabled = false;
			this.optTipoMantenimiento.ItemAppearance = appearance2;
			this.optTipoMantenimiento.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Preventivo";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Correctivo";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Predictivo";
			this.optTipoMantenimiento.Items.Add(valueListItem1);
			this.optTipoMantenimiento.Items.Add(valueListItem2);
			this.optTipoMantenimiento.Items.Add(valueListItem3);
			this.optTipoMantenimiento.ItemSpacingVertical = 10;
			this.optTipoMantenimiento.Location = new System.Drawing.Point(528, 56);
			this.optTipoMantenimiento.Name = "optTipoMantenimiento";
			this.optTipoMantenimiento.Size = new System.Drawing.Size(280, 24);
			this.optTipoMantenimiento.TabIndex = 1076;
			this.optTipoMantenimiento.Text = "Preventivo";
			// 
			// btnFacrutaCompra
			// 
			this.btnFacrutaCompra.Enabled = false;
			this.btnFacrutaCompra.Location = new System.Drawing.Point(16, 247);
			this.btnFacrutaCompra.Name = "btnFacrutaCompra";
			this.btnFacrutaCompra.Size = new System.Drawing.Size(96, 23);
			this.btnFacrutaCompra.TabIndex = 1077;
			this.btnFacrutaCompra.Text = "Factura Compra";
			this.btnFacrutaCompra.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// txtObserbaciones
			// 
			this.txtObserbaciones.AcceptsReturn = true;
			this.txtObserbaciones.AcceptsTab = true;
			this.txtObserbaciones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObserbaciones.Appearance = appearance3;
			this.txtObserbaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObserbaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObserbaciones.Enabled = false;
			this.txtObserbaciones.Location = new System.Drawing.Point(112, 88);
			this.txtObserbaciones.Multiline = true;
			this.txtObserbaciones.Name = "txtObserbaciones";
			this.txtObserbaciones.Size = new System.Drawing.Size(280, 72);
			this.txtObserbaciones.TabIndex = 1078;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 112);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(84, 16);
			this.label7.TabIndex = 1079;
			this.label7.Text = "Obserbaciones:";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(400, 136);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(95, 16);
			this.label3.TabIndex = 1080;
			this.label3.Text = "Ingresa Inventario";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkInventario
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkInventario.Appearance = appearance4;
			this.chkInventario.BackColor = System.Drawing.Color.Transparent;
			this.chkInventario.Enabled = false;
			this.chkInventario.Location = new System.Drawing.Point(504, 134);
			this.chkInventario.Name = "chkInventario";
			this.chkInventario.Size = new System.Drawing.Size(16, 21);
			this.chkInventario.TabIndex = 1081;
			// 
			// dtFecha
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance5;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 24);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1083;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(200, 296);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(72, 23);
			this.btnEditar.TabIndex = 1087;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(288, 296);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(72, 23);
			this.btnGuardar.TabIndex = 1086;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(376, 296);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 1088;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(464, 296);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 1089;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(24, 296);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 1084;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(112, 296);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(72, 23);
			this.btnConsultar.TabIndex = 1085;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(456, 24);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(58, 16);
			this.label9.TabIndex = 1090;
			this.label9.Text = "Habitación";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkTrabajosExteriores
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTrabajosExteriores.Appearance = appearance6;
			this.chkTrabajosExteriores.BackColor = System.Drawing.Color.Transparent;
			this.chkTrabajosExteriores.Enabled = false;
			this.chkTrabajosExteriores.Location = new System.Drawing.Point(128, 192);
			this.chkTrabajosExteriores.Name = "chkTrabajosExteriores";
			this.chkTrabajosExteriores.Size = new System.Drawing.Size(16, 21);
			this.chkTrabajosExteriores.TabIndex = 1094;
			this.chkTrabajosExteriores.VisibleChanged += new System.EventHandler(this.chkTrabajosExteriores_VisibleChanged);
			this.chkTrabajosExteriores.CheckedChanged += new System.EventHandler(this.chkTrabajosExteriores_CheckedChanged);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(16, 192);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(102, 16);
			this.label10.TabIndex = 1095;
			this.label10.Text = "Trabajos Exteriores";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(152, 192);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(99, 16);
			this.label11.TabIndex = 1096;
			this.label11.Text = "Trabajos Interiores";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkTrabajosInteriores
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTrabajosInteriores.Appearance = appearance7;
			this.chkTrabajosInteriores.BackColor = System.Drawing.Color.Transparent;
			this.chkTrabajosInteriores.Enabled = false;
			this.chkTrabajosInteriores.Location = new System.Drawing.Point(264, 192);
			this.chkTrabajosInteriores.Name = "chkTrabajosInteriores";
			this.chkTrabajosInteriores.Size = new System.Drawing.Size(16, 21);
			this.chkTrabajosInteriores.TabIndex = 1097;
			this.chkTrabajosInteriores.Text = "Trabajos Interiores";
			this.chkTrabajosInteriores.CheckedChanged += new System.EventHandler(this.chkTrabajosInteriores_CheckedChanged);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(520, 192);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(45, 16);
			this.label12.TabIndex = 1098;
			this.label12.Text = "Equipos";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkEquipos
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEquipos.Appearance = appearance8;
			this.chkEquipos.BackColor = System.Drawing.Color.Transparent;
			this.chkEquipos.Enabled = false;
			this.chkEquipos.Location = new System.Drawing.Point(488, 192);
			this.chkEquipos.Name = "chkEquipos";
			this.chkEquipos.Size = new System.Drawing.Size(16, 21);
			this.chkEquipos.TabIndex = 1099;
			this.chkEquipos.CheckedChanged += new System.EventHandler(this.chkEquipos_CheckedChanged);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(368, 192);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(106, 16);
			this.label13.TabIndex = 1100;
			this.label13.Text = "Cuarto de Maquinas";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkCuartoMaquinas
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCuartoMaquinas.Appearance = appearance9;
			this.chkCuartoMaquinas.BackColor = System.Drawing.Color.Transparent;
			this.chkCuartoMaquinas.Enabled = false;
			this.chkCuartoMaquinas.Location = new System.Drawing.Point(576, 192);
			this.chkCuartoMaquinas.Name = "chkCuartoMaquinas";
			this.chkCuartoMaquinas.Size = new System.Drawing.Size(16, 21);
			this.chkCuartoMaquinas.TabIndex = 1101;
			this.chkCuartoMaquinas.CheckedChanged += new System.EventHandler(this.chkCuartoMaquinas_CheckedChanged);
			// 
			// chkCocina
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCocina.Appearance = appearance10;
			this.chkCocina.BackColor = System.Drawing.Color.Transparent;
			this.chkCocina.Enabled = false;
			this.chkCocina.Location = new System.Drawing.Point(344, 192);
			this.chkCocina.Name = "chkCocina";
			this.chkCocina.Size = new System.Drawing.Size(16, 21);
			this.chkCocina.TabIndex = 1103;
			this.chkCocina.CheckedChanged += new System.EventHandler(this.chkCocina_CheckedChanged);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(392, 60);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(122, 16);
			this.label15.TabIndex = 1104;
			this.label15.Text = "Tipo de Mantenimiento:";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(120, 250);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(43, 16);
			this.lblNumero.TabIndex = 1106;
			this.lblNumero.Text = "Factura";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtProveedor
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProveedor.Appearance = appearance11;
			this.txtProveedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProveedor.Enabled = false;
			this.txtProveedor.Location = new System.Drawing.Point(344, 248);
			this.txtProveedor.Name = "txtProveedor";
			this.txtProveedor.Size = new System.Drawing.Size(112, 21);
			this.txtProveedor.TabIndex = 1107;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(280, 250);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(56, 16);
			this.label1.TabIndex = 1108;
			this.label1.Text = "Proveedor";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(472, 250);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(29, 16);
			this.label2.TabIndex = 1109;
			this.label2.Text = "Total";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(400, 104);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(121, 16);
			this.label4.TabIndex = 1112;
			this.label4.Text = "Descripción de Trabajo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcionTrabajo
			// 
			this.txtDescripcionTrabajo.AcceptsReturn = true;
			this.txtDescripcionTrabajo.AcceptsTab = true;
			this.txtDescripcionTrabajo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcionTrabajo.Appearance = appearance12;
			this.txtDescripcionTrabajo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcionTrabajo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcionTrabajo.Enabled = false;
			this.txtDescripcionTrabajo.Location = new System.Drawing.Point(528, 88);
			this.txtDescripcionTrabajo.Multiline = true;
			this.txtDescripcionTrabajo.Name = "txtDescripcionTrabajo";
			this.txtDescripcionTrabajo.Size = new System.Drawing.Size(280, 72);
			this.txtDescripcionTrabajo.TabIndex = 1113;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// cmbBodega
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance13;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(240, 24);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(208, 21);
			this.cmbBodega.TabIndex = 1115;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// cmbHabitaciones
			// 
			appearance14.FontData.Name = "Tahoma";
			this.cmbHabitaciones.Appearance = appearance14;
			this.cmbHabitaciones.CausesValidation = false;
			this.cmbHabitaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbHabitaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbHabitaciones.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 250;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbHabitaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbHabitaciones.DisplayMember = "Nombre";
			this.cmbHabitaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbHabitaciones.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbHabitaciones.Enabled = false;
			this.cmbHabitaciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbHabitaciones.Location = new System.Drawing.Point(528, 24);
			this.cmbHabitaciones.MaxDropDownItems = 30;
			this.cmbHabitaciones.Name = "cmbHabitaciones";
			this.cmbHabitaciones.Size = new System.Drawing.Size(208, 22);
			this.cmbHabitaciones.TabIndex = 1116;
			this.cmbHabitaciones.ValueMember = "idHabitacion";
			this.cmbHabitaciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbHabitaciones_InitializeLayout);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(296, 192);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(39, 16);
			this.label14.TabIndex = 1102;
			this.label14.Text = "Cocina";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtidControlMantenimiento
			// 
			this.txtidControlMantenimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidControlMantenimiento.Enabled = false;
			this.txtidControlMantenimiento.Location = new System.Drawing.Point(760, 296);
			this.txtidControlMantenimiento.Name = "txtidControlMantenimiento";
			this.txtidControlMantenimiento.PromptChar = ' ';
			this.txtidControlMantenimiento.Size = new System.Drawing.Size(56, 21);
			this.txtidControlMantenimiento.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidControlMantenimiento.TabIndex = 1117;
			this.txtidControlMantenimiento.Visible = false;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 26);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(36, 16);
			this.label5.TabIndex = 1082;
			this.label5.Text = "Fecha";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbTiposdeTrabajo
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTiposdeTrabajo.Appearance = appearance15;
			this.cmbTiposdeTrabajo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTiposdeTrabajo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTiposdeTrabajo.DataSource = this.ultraDataSource2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbTiposdeTrabajo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTiposdeTrabajo.DisplayMember = "nombreTrabajo";
			this.cmbTiposdeTrabajo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTiposdeTrabajo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTiposdeTrabajo.Enabled = false;
			this.cmbTiposdeTrabajo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTiposdeTrabajo.Location = new System.Drawing.Point(112, 58);
			this.cmbTiposdeTrabajo.Name = "cmbTiposdeTrabajo";
			this.cmbTiposdeTrabajo.Size = new System.Drawing.Size(208, 21);
			this.cmbTiposdeTrabajo.TabIndex = 1114;
			this.cmbTiposdeTrabajo.ValueMember = "idTipoTrabajo";
			this.cmbTiposdeTrabajo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTiposdeTrabajo_InitializeLayout);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(16, 168);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(89, 16);
			this.label8.TabIndex = 1121;
			this.label8.Text = "Tipo De Trabajo";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.groupBox3);
			this.groupBox2.Controls.Add(this.groupBox1);
			this.groupBox2.Location = new System.Drawing.Point(-248, 280);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1323, 8);
			this.groupBox2.TabIndex = 1122;
			this.groupBox2.TabStop = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1323, 8);
			this.groupBox1.TabIndex = 1123;
			this.groupBox1.TabStop = false;
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Controls.Add(this.groupBox4);
			this.groupBox3.Location = new System.Drawing.Point(0, 0);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1323, 8);
			this.groupBox3.TabIndex = 1124;
			this.groupBox3.TabStop = false;
			// 
			// groupBox4
			// 
			this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox4.Location = new System.Drawing.Point(0, 0);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1323, 8);
			this.groupBox4.TabIndex = 1123;
			this.groupBox4.TabStop = false;
			// 
			// groupBox5
			// 
			this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox5.Controls.Add(this.groupBox6);
			this.groupBox5.Location = new System.Drawing.Point(-248, 224);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(1323, 8);
			this.groupBox5.TabIndex = 1123;
			this.groupBox5.TabStop = false;
			// 
			// groupBox6
			// 
			this.groupBox6.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox6.Location = new System.Drawing.Point(0, 0);
			this.groupBox6.Name = "groupBox6";
			this.groupBox6.Size = new System.Drawing.Size(1323, 8);
			this.groupBox6.TabIndex = 1123;
			this.groupBox6.TabStop = false;
			// 
			// txtValorCompra
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorCompra.Appearance = appearance16;
			this.txtValorCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorCompra.Enabled = false;
			this.txtValorCompra.FormatString = "#,##0.00";
			this.txtValorCompra.Location = new System.Drawing.Point(512, 248);
			this.txtValorCompra.MinValue = 0;
			this.txtValorCompra.Name = "txtValorCompra";
			this.txtValorCompra.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorCompra.PromptChar = ' ';
			this.txtValorCompra.Size = new System.Drawing.Size(112, 21);
			this.txtValorCompra.TabIndex = 1124;
			// 
			// txtNumFactura
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumFactura.Appearance = appearance17;
			this.txtNumFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumFactura.Enabled = false;
			this.txtNumFactura.Location = new System.Drawing.Point(168, 248);
			this.txtNumFactura.Name = "txtNumFactura";
			this.txtNumFactura.Size = new System.Drawing.Size(112, 21);
			this.txtNumFactura.TabIndex = 1125;
			// 
			// txtidCompra
			// 
			this.txtidCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCompra.Enabled = false;
			this.txtidCompra.Location = new System.Drawing.Point(760, 240);
			this.txtidCompra.Name = "txtidCompra";
			this.txtidCompra.PromptChar = ' ';
			this.txtidCompra.Size = new System.Drawing.Size(56, 21);
			this.txtidCompra.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidCompra.TabIndex = 1126;
			this.txtidCompra.Visible = false;
			// 
			// frmControlMantenimiento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(824, 326);
			this.Controls.Add(this.txtidCompra);
			this.Controls.Add(this.txtNumFactura);
			this.Controls.Add(this.txtValorCompra);
			this.Controls.Add(this.groupBox5);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.txtidControlMantenimiento);
			this.Controls.Add(this.cmbHabitaciones);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbTiposdeTrabajo);
			this.Controls.Add(this.txtDescripcionTrabajo);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtProveedor);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.chkCocina);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.chkCuartoMaquinas);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.chkEquipos);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.chkTrabajosInteriores);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.chkTrabajosExteriores);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.chkInventario);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtObserbaciones);
			this.Controls.Add(this.btnFacrutaCompra);
			this.Controls.Add(this.optTipoMantenimiento);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label5);
			this.KeyPreview = true;
			this.Name = "frmControlMantenimiento";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Control de Mantenimiento";
			this.Load += new System.EventHandler(this.frmControlMantenimiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.optTipoMantenimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObserbaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcionTrabajo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHabitaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidControlMantenimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiposdeTrabajo)).EndInit();
			this.groupBox2.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			this.groupBox5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtValorCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void label19_Click(object sender, System.EventArgs e)
		{
		
		}

		int idHabitacion = 0;
		int iBodega = 0;
		int iTarifaExpress = 0;
		int idCompraV = 0;
		bool bMenuCuadricula = false;
		bool bServicioPagado = false;

		bool bNuevo = false;
		bool bEdicion = false;

		int contadorTrabajos = 0;
		int idCompra = 0;
		int retornarId = 0;

		private Acceso miAcceso;

		private void Consultar(int ControlMantenimientoId)
		{
			try
			{
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnCancelar.Enabled = false;


				string sSQLBusqueda = string.Format("Exec [sp_BuscaGeneralControlMantenimiento] {0}", ControlMantenimientoId);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLBusqueda, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtidControlMantenimiento.Value = dr.GetInt32(0);
					if (dr.GetValue(1) != System.DBNull.Value) this.dtFecha.Value = dr.GetDateTime(1);
					if (dr.GetValue(2) != System.DBNull.Value) this.cmbHabitaciones.Value =  dr.GetInt32(2);
					if (dr.GetValue(3) != System.DBNull.Value) this.cmbBodega.Value =  dr.GetInt32(3);
					if (dr.GetValue(4) != System.DBNull.Value) this.cmbTiposdeTrabajo.Value = dr.GetInt32(4);
					if (dr.GetValue(5) != System.DBNull.Value) this.optTipoMantenimiento.Value = dr.GetInt32(5);
					if (dr.GetValue(6) != System.DBNull.Value) this.txtDescripcionTrabajo.Text = dr.GetString(6); 
					if (dr.GetValue(7) != System.DBNull.Value) this.txtidCompra.Value =  dr.GetInt32(7);
					if (dr.GetValue(8) != System.DBNull.Value) this.chkInventario.Checked =  dr.GetBoolean(8);
					if (dr.GetValue(9) != System.DBNull.Value) this.txtObserbaciones.Text = dr.GetString(9); 
					if (dr.GetValue(10) != System.DBNull.Value) this.chkTrabajosExteriores.Checked =   dr.GetBoolean(10);
					if (dr.GetValue(11) != System.DBNull.Value) this.chkTrabajosInteriores.Checked =  dr.GetBoolean(11);
					if (dr.GetValue(12) != System.DBNull.Value) this.chkEquipos.Checked =  dr.GetBoolean(12);
					if (dr.GetValue(13) != System.DBNull.Value) this.chkCuartoMaquinas.Checked =  dr.GetBoolean(13);
					if (dr.GetValue(14) != System.DBNull.Value) this.chkCocina.Checked =  dr.GetBoolean(14);
					if (dr.GetValue(15) != System.DBNull.Value) this.txtNumFactura.Text = dr.GetString(15); 
           if (dr.GetValue(16) != System.DBNull.Value) this.txtProveedor.Text = dr.GetString(16); 
          if (dr.GetValue(17) != System.DBNull.Value) this.txtValorCompra.Value = dr.GetDecimal(17);

//					if (dr.GetValue(14) != System.DBNull.Value) this.optTipoMantenimiento.Value = dr.GetInt32(14);
				}
				dr.Close();
			
				this.btnNuevo.Enabled = true;
				this.btnConsultar.Enabled = true;
				this.btnEditar.Enabled = true;
				this.btnCancelar.Enabled = true;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}			
		}



		private void ValidaTiempo(int idCompraC)
		{
			int TiempoEstadia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Jornadas from compra where idCompra = {0}", idCompraC));
			
			if (TiempoEstadia > 8)
				this.txtTiempo.MaxValue = 12;
		}

		// Consulta
		
	

		private void frmControlMantenimiento_Load(object sender, System.EventArgs e)
		{
			/*traer la fecha del servidor*/
			//FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));
			this.cmbTiposdeTrabajo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaTipoTrabajo"));
		}

		private void ultraCheckEditor3_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbRegion_ValueChanged(object sender, System.EventArgs e)
		{
	
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
		/*
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local"))
			{
				this.cmbBodega.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbHabitaciones, "Seleccione una habitación")) 
			{
				this.cmbHabitaciones.Focus();
				return;
			}
*/
			using (frmBuscaTransacciones Busqueda = new frmBuscaTransacciones(4, 0))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{	
					this.txtidCompra.Value = (int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value;
					this.txtProveedor.Value = Busqueda.grdFacturas.ActiveRow.Cells["Nombre"].Value.ToString();
					this.txtNumFactura.Text = Busqueda.grdFacturas.ActiveRow.Cells["Numero"].Value.ToString();
					this.txtValorCompra.Value = Convert.ToDecimal(Busqueda.grdFacturas.ActiveRow.Cells["Total"].Value);
				}
			}

			/*
			if (!Validacion.vbComboVacio (this.cmbTiposdeTrabajo, "Seleccione tipos de trabajo")) 
			{
				this.cmbTiposdeTrabajo.Focus();
				return;
			}
			
			if (!this.chkTrabajosExteriores.Checked &&
				!this.chkTrabajosInteriores.Checked &&
				!this.chkCocina.Checked &&
				!this.chkEquipos.Checked && 
				!this.chkCuartoMaquinas.Checked) 
			{
				MessageBox.Show("Selecciona un lugar de trabajo ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (this.txtDescripcionTrabajo.Text == "") 
			{
				MessageBox.Show("Por favor ingresa el tipo de trabajo ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (!this.chkInventario.Checked)
			{
				MessageBox.Show("Ingresa inventario?", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (this.optTipoMantenimiento.Value == null )
			{
				MessageBox.Show("Ingresa tipo de mantenimiento?", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			//
			*/

		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cmbRegion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			// validacion
		}

		private void cmbRegion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void ultraCombo2_ValueChanged(object sender, System.EventArgs e)
		{
			if (bNuevo && bEdicion)
			{
				if (this.cmbHabitaciones.ActiveRow != null)
				{
					bool bActivoFact = false;

					bActivoFact = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select top 1 Activo From TiempoModificacionEstadia Where Bodega = {0} Order by FechaSistema desc", 
						(int)this.cmbBodega.Value), true);
				}
			}
		}

		private void ultraCombo2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);

			#region bloqueo datos
			this.cmbBodega.Enabled = true;
			this.cmbHabitaciones.Enabled = true;
			this.cmbTiposdeTrabajo.Enabled = true;
			this.optTipoMantenimiento.Enabled = true;
			this.txtObserbaciones.Enabled = true;
			this.txtDescripcionTrabajo.Enabled = true;
			this.chkInventario.Enabled = true;
			this.chkTrabajosExteriores.Enabled = true;
			this.chkTrabajosInteriores.Enabled = true;
			this.chkCocina.Enabled = true;
			this.chkEquipos.Enabled = true;
			this.chkCuartoMaquinas.Enabled = true;
			this.btnFacrutaCompra.Enabled = true;
			#endregion bloqueo datos

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			bNuevo = true;
			bEdicion = true;
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null)			
				this.cmbHabitaciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idHabitacion, Nombre From Habitaciones Where Bodega = {0} Order by NumeroHabitacion", (int)this.cmbBodega.Value));
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region validaciones
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;
			if (!Validacion.vbComboVacio(this.cmbHabitaciones, "Seleccione una habitación")) return;
			if (!Validacion.vbComboVacio(this.cmbTiposdeTrabajo, "Seleccione tipos de trabajo")) return;

			if (!this.chkTrabajosExteriores.Checked &&
				!this.chkTrabajosInteriores.Checked &&
				!this.chkCocina.Checked &&
				!this.chkEquipos.Checked && 
				!this.chkCuartoMaquinas.Checked) 
			{
				MessageBox.Show("Selecciona un lugar de trabajo ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!Validacion.vbTexto(this.txtDescripcionTrabajo, 10, 254, "Detalle")) return;

			if (!this.chkInventario.Checked)
			{
				MessageBox.Show("Ingresa inventario?", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			# endregion validaciones

			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int) this.cmbBodega.Value;

			int iHabiatcion = 0;
			if (this.cmbHabitaciones.ActiveRow != null) iHabiatcion = (int) this.cmbHabitaciones.Value;
      
			if((int) this.txtidCompra.Value == 0 )
			{
				MessageBox.Show("Seleccione una factura de Compra.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			int iTipodeTrabajo = 0;
			if (this.cmbTiposdeTrabajo.ActiveRow != null) iTipodeTrabajo = (int) this.cmbTiposdeTrabajo.Value;
			int intControlMantenimientoId = (int)this.txtidControlMantenimiento.Value;
//			MessageBox.Show(intControlMantenimientoId.ToString(), "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
			if (idCompra == 0) idCompra = Convert.ToInt32(this.txtidControlMantenimiento.Value);
//			MessageBox.Show(idCompra.ToString(), "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
			int chkTrabajosExteriores = this.chkTrabajosExteriores.Checked ? 1 : 0;
			int chkTrabajosInteriores = this.chkTrabajosInteriores.Checked ? 1 : 0;
			int chkCocina = this.chkCocina.Checked ? 1 : 0;
			int chkEquipos = this.chkEquipos.Checked ? 1 : 0;
			int chkCuartoMaquinas = this.chkCuartoMaquinas.Checked ? 1 : 0;

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
					DateTime dtFecha = (DateTime) this.dtFecha.Value;
					string sSQL1 = string.Format("Exec [sp_InsertarControlMantenimiento] {0}, {1}, '{2}', {3}, {4}, '{5}', {6}, {7}, '{8}', {9}, {10}, {11}, {12}, {13}, {14}",
						intControlMantenimientoId,
						iBodega,
						dtFecha.ToString("yyyyMMdd"), 
						iHabiatcion,
						iTipodeTrabajo,
						this.txtDescripcionTrabajo.Text,
						(int)this.optTipoMantenimiento.Value,
						(bool)this.chkInventario.Checked ? 1 : 0, 
						this.txtObserbaciones.Text, 
						chkTrabajosExteriores,
						chkTrabajosInteriores,
						chkEquipos,
						chkCuartoMaquinas,
						chkCocina,
						(int) this.txtidCompra.Value);
					oCmdActualiza.CommandText = sSQL1;

					this.txtidControlMantenimiento.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Graba Maestro
					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

					#region Controles

					#region bloqueo datos
					this.cmbBodega.Enabled = false;
					this.cmbHabitaciones.Enabled = false;
					this.cmbTiposdeTrabajo.Enabled = false;
					this.optTipoMantenimiento.Enabled = false;
					this.txtObserbaciones.Enabled = false;
					this.txtDescripcionTrabajo.Enabled = false;
					this.chkInventario.Enabled = false;
					this.chkTrabajosExteriores.Enabled = false;
					this.chkTrabajosInteriores.Enabled = false;
					this.chkCocina.Enabled = false;
					this.chkEquipos.Enabled = false;
					this.chkCuartoMaquinas.Enabled = false;
					this.txtNumFactura.Enabled = false;
					this.txtProveedor.Enabled = false;;
					this.txtValorCompra.Enabled = false;
					this.btnFacrutaCompra.Enabled = false;
					#endregion bloqueo datos			

					this.btnNuevo.Enabled = true;
					this.btnConsultar.Enabled = true;
					this.btnEditar.Enabled = false;
					this.btnGuardar.Enabled = false;
					this.btnCancelar.Enabled = true;

					this.Consultar((int) this.txtidControlMantenimiento.Value);

					bNuevo = false;
					bEdicion = false;
					#endregion Controles

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

		}

		private void cmbTiposdeTrabajo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void chkTrabajosExteriores_VisibleChanged(object sender, System.EventArgs e)
		{
			
		}

		private void chkTrabajosExteriores_CheckedChanged(object sender, System.EventArgs e)
		{
			
		}

		private void chkTrabajosInteriores_CheckedChanged(object sender, System.EventArgs e)
		{
			
		}

		private void chkCocina_CheckedChanged(object sender, System.EventArgs e)
		{
			
		}

		private void chkEquipos_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkCuartoMaquinas_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbHabitaciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmControlMantenimientoConsulta Busqueda = new frmControlMantenimientoConsulta())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consultar((int)Busqueda.ultraGrid1.ActiveRow.Cells["idControlMantenimiento"].Value);
					/*idCompra = (int)Busqueda.ultraGrid1.ActiveRow.Cells["idCompra"].Value;
					retornarId = 1;*/
				}
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region vaciar datos
			this.dtFecha.Value = System.DBNull.Value;
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbHabitaciones.Value = System.DBNull.Value;
			this.cmbTiposdeTrabajo.Value = System.DBNull.Value;
			this.optTipoMantenimiento.Value = 1;
			this.txtObserbaciones.Text ="";
			this.txtDescripcionTrabajo.Text ="";
			this.chkInventario.Checked = false;
			this.chkTrabajosExteriores.Checked = false;
			this.chkTrabajosInteriores.Checked = false;
			this.chkCocina.Checked = false;
			this.chkEquipos.Checked = false;
			this.chkCuartoMaquinas.Checked = false;
			this.txtNumFactura.Text ="";
			this.txtProveedor.Text ="";
			this.txtidControlMantenimiento.Value = 0;
			this.txtValorCompra.Value = 0;
			this.txtidCompra.Value = 0;
			#endregion vaciar datos

			#region bloqueo datos
			this.cmbBodega.Enabled = false;
			this.cmbHabitaciones.Enabled = false;
			this.cmbTiposdeTrabajo.Enabled = false;
			this.optTipoMantenimiento.Enabled = false;
			this.txtObserbaciones.Enabled = false;
			this.txtDescripcionTrabajo.Enabled = false;
			this.chkInventario.Enabled = false;
			this.chkTrabajosExteriores.Enabled = false;
			this.chkTrabajosInteriores.Enabled = false;
			this.chkCocina.Enabled = false;
			this.chkEquipos.Enabled = false;
			this.chkCuartoMaquinas.Enabled = false;
			this.txtNumFactura.Enabled = false;
			this.txtProveedor.Enabled = false;;
			this.txtValorCompra.Enabled = false;
			this.btnFacrutaCompra.Enabled = false;
			#endregion bloqueo datos

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;

		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.cmbBodega.Enabled = true;
			this.cmbHabitaciones.Enabled = true;
			this.cmbTiposdeTrabajo.Enabled = true;
			this.optTipoMantenimiento.Enabled = true;
			this.optTipoMantenimiento.CheckedIndex = 0;
			this.txtObserbaciones.Enabled = true;
			this.txtDescripcionTrabajo.Enabled = true;
			this.chkInventario.Enabled = true;
			this.chkTrabajosExteriores.Enabled = true;
			this.chkTrabajosInteriores.Enabled = true;
			this.chkCocina.Enabled = true;
			this.chkEquipos.Enabled = true;
			this.chkCuartoMaquinas.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;			
			this.btnGuardar.Enabled = true;			
			this.btnCancelar.Enabled = true;
		}

		private void txtOrdenSpf_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
