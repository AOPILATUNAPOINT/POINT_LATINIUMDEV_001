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
	/// Descripción breve de frmUsuario_maquina.
	/// </summary>
	public class frmUsuario_maquina : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtUsuario;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEstacion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkActivo;
		private System.Windows.Forms.GroupBox gbBotones;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidUsuario_Maquina;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnConsultar;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtbusqueda;
		private System.Windows.Forms.Label lblBusqueda;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultgconsulta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmUsuario_maquina()
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmUsuario_maquina));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuario_maquina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha Ingreso");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario_maquina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha Ingreso");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			this.label5 = new System.Windows.Forms.Label();
			this.txtUsuario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEstacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.chkActivo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtidUsuario_Maquina = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.txtbusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblBusqueda = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultgconsulta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidUsuario_Maquina)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtbusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgconsulta)).BeginInit();
			this.SuspendLayout();
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(16, 18);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(43, 16);
			this.label5.TabIndex = 866;
			this.label5.Text = "Usuario";
			// 
			// txtUsuario
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUsuario.Appearance = appearance1;
			this.txtUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUsuario.Enabled = false;
			this.txtUsuario.Location = new System.Drawing.Point(112, 16);
			this.txtUsuario.MaxLength = 200;
			this.txtUsuario.Name = "txtUsuario";
			this.txtUsuario.Size = new System.Drawing.Size(184, 21);
			this.txtUsuario.TabIndex = 1036;
			// 
			// txtEstacion
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstacion.Appearance = appearance2;
			this.txtEstacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstacion.Enabled = false;
			this.txtEstacion.Location = new System.Drawing.Point(112, 54);
			this.txtEstacion.MaxLength = 200;
			this.txtEstacion.Name = "txtEstacion";
			this.txtEstacion.Size = new System.Drawing.Size(184, 21);
			this.txtEstacion.TabIndex = 1038;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(89, 16);
			this.label1.TabIndex = 1037;
			this.label1.Text = "Estación Trabajo";
			// 
			// chkActivo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkActivo.Appearance = appearance3;
			this.chkActivo.BackColor = System.Drawing.Color.Transparent;
			this.chkActivo.Checked = true;
			this.chkActivo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkActivo.Enabled = false;
			this.chkActivo.Location = new System.Drawing.Point(316, 56);
			this.chkActivo.Name = "chkActivo";
			this.chkActivo.Size = new System.Drawing.Size(96, 22);
			this.chkActivo.TabIndex = 1039;
			this.chkActivo.Text = "Activo";
			// 
			// btnGuardar
			// 
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(176, 112);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 1044;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(96, 112);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(72, 23);
			this.btnEditar.TabIndex = 1041;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(264, 112);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 1042;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(344, 112);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 1043;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(16, 112);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 1040;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// gbBotones
			// 
			this.gbBotones.Location = new System.Drawing.Point(0, 88);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(616, 8);
			this.gbBotones.TabIndex = 1045;
			this.gbBotones.TabStop = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 152);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(616, 8);
			this.groupBox1.TabIndex = 1046;
			this.groupBox1.TabStop = false;
			// 
			// txtidUsuario_Maquina
			// 
			this.txtidUsuario_Maquina.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidUsuario_Maquina.Enabled = false;
			this.txtidUsuario_Maquina.Location = new System.Drawing.Point(448, 56);
			this.txtidUsuario_Maquina.Name = "txtidUsuario_Maquina";
			this.txtidUsuario_Maquina.PromptChar = ' ';
			this.txtidUsuario_Maquina.Size = new System.Drawing.Size(48, 21);
			this.txtidUsuario_Maquina.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidUsuario_Maquina.TabIndex = 1047;
			this.txtidUsuario_Maquina.Visible = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(416, 176);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 1051;
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(320, 174);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(72, 23);
			this.btnConsultar.TabIndex = 1050;
			this.btnConsultar.Text = "Buscar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// txtbusqueda
			// 
			appearance4.BackColorDisabled = System.Drawing.Color.Silver;
			appearance4.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance4.BorderColor = System.Drawing.Color.Black;
			appearance4.BorderColor3DBase = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtbusqueda.Appearance = appearance4;
			this.txtbusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtbusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtbusqueda.Location = new System.Drawing.Point(104, 175);
			this.txtbusqueda.Name = "txtbusqueda";
			this.txtbusqueda.Size = new System.Drawing.Size(208, 21);
			this.txtbusqueda.TabIndex = 1049;
			this.txtbusqueda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtbusqueda_KeyDown);
			// 
			// lblBusqueda
			// 
			this.lblBusqueda.AutoSize = true;
			this.lblBusqueda.Location = new System.Drawing.Point(8, 177);
			this.lblBusqueda.Name = "lblBusqueda";
			this.lblBusqueda.Size = new System.Drawing.Size(90, 16);
			this.lblBusqueda.TabIndex = 1048;
			this.lblBusqueda.Text = "Usuario/Estación";
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
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(bool);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// ultgconsulta
			// 
			this.ultgconsulta.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultgconsulta.DataSource = this.ultraDataSource2;
			appearance5.BackColor = System.Drawing.Color.White;
			appearance5.ForegroundAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultgconsulta.DisplayLayout.Appearance = appearance5;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 150;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 150;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 130;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 110;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.ultgconsulta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultgconsulta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultgconsulta.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultgconsulta.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8.5F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultgconsulta.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgconsulta.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgconsulta.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultgconsulta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultgconsulta.Location = new System.Drawing.Point(8, 216);
			this.ultgconsulta.Name = "ultgconsulta";
			this.ultgconsulta.Size = new System.Drawing.Size(608, 176);
			this.ultgconsulta.TabIndex = 1052;
			this.ultgconsulta.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultgconsulta_BeforeSelectChange);
			// 
			// frmUsuario_maquina
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(624, 406);
			this.Controls.Add(this.ultgconsulta);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.txtbusqueda);
			this.Controls.Add(this.lblBusqueda);
			this.Controls.Add(this.txtidUsuario_Maquina);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.chkActivo);
			this.Controls.Add(this.txtEstacion);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtUsuario);
			this.Controls.Add(this.label5);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmUsuario_maquina";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Configuración Estación de Trabajo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUsuario_maquina_KeyDown);
			this.Load += new System.EventHandler(this.frmUsuario_maquina_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidUsuario_Maquina)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtbusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgconsulta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		
		bool bNuevo = false;
		bool bEdicion = false;

		private void frmUsuario_maquina_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "2906");

			if (!Funcion.Permiso("2178", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a ESTACIÓN DE TRABAJO ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			string sSQL = string.Format("Exec ConultaGeneralValidaUsuario_maquina  '{0}'", this.txtbusqueda.Text.ToString());
			this.ultgconsulta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region Bloqueo y vaciado de controles 
			string sSQL = string.Format("Exec ConultaGeneralValidaUsuario_maquina  '{0}'", this.txtbusqueda.Text.ToString());
			this.ultgconsulta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.txtEstacion.Text ="";
			this.txtUsuario.Text ="";
			this.chkActivo.Checked = false;
			this.txtidUsuario_Maquina.Value = 0;

			this.txtEstacion.Enabled = false;
			this.txtUsuario.Enabled = false;
			this.chkActivo.Checked = false;
			this.txtbusqueda.Text ="";


			this.btnNuevo.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;

			this.btnConsultar.Enabled = true;
			this.txtbusqueda.Enabled= true;
	
			#endregion Bloqueo y vaciado de controles 
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			#region Habilitacion de controles 
			this.btnCancelar_Click(sender,e);

			this.txtEstacion.Enabled = true;
			this.txtUsuario.Enabled = true;
			this.chkActivo.Enabled = true;
			this.chkActivo.Checked = true;

			this.txtbusqueda.Text="";
    //  FuncionesProcedimientos.EstadoGrids(true,this.ultgconsulta);
			this.btnNuevo.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnConsultar.Enabled = false;
			this.txtbusqueda.Enabled= false;

		

			bNuevo = true;
			bEdicion = true;

	
			#endregion Habilitacion de controles 
		}

		private void frmUsuario_maquina_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidUsuario_Maquina.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);													
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

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.txtEstacion.Enabled = true;
			this.txtUsuario.Enabled = true;
			this.chkActivo.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnConsultar.Enabled = false;
			this.txtbusqueda.Enabled= false;
		//	FuncionesProcedimientos.EstadoGrids(true,this.ultgconsulta);
		
			bNuevo = false;
			bEdicion = true;	
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
		  if (!Validacion.vbTexto(this.txtEstacion, 1, 50, "Estación")) return;
			if (!Validacion.vbTexto(this.txtUsuario, 1, 50, "Usuario")) return;
			if(bNuevo) this.chkActivo.Checked = true;

			#region Validacion
			#region Valida Usuario Existe
			string Mensaje ="";
			Mensaje = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec ValidaUsuario_maquina {0}, {1}",
				this.txtUsuario.Text.ToString(), this.txtEstacion.Text.ToString()));		

			if ( Mensaje.Length > 0)
			{
				MessageBox.Show(string.Format("Resultado:  '{0}'", Mensaje ), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}


      #endregion Valida Usuario Existe
			#endregion Validacion

			if (DialogResult.No == MessageBox.Show("¿Esta seguro de Guardar? ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				return;
			}
      
			#region Guardar

			//			DateTime dtInicia = DateTime.Today;
			//			DateTime dtRetorna = DateTime.Today;
			//				dtInicia = (DateTime)this.dtInicia.Value;
			//				dtRetorna = (DateTime)this.dtRetorna.Value;

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

					

					

					#region Graba Ingreso Gastos
					
							string  sSQLDeta = string.Format("Exec GrabaUsuario_maquina {0}, '{1}', '{2}', {3} " ,
								Convert.ToInt32 ( this.txtidUsuario_Maquina.Value), 
								Convert.ToString ( this.txtUsuario.Text), 
								Convert.ToString  ( this.txtEstacion.Text),
								Convert.ToInt32( this.chkActivo.Checked)
								);

							oCmdActualiza.CommandText = sSQLDeta;
							oCmdActualiza.ExecuteNonQuery();	
					
					#endregion Graba Ingreso Gastos
          
					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					#region Controles
					this.txtUsuario.Enabled = false;
					this.txtEstacion.Enabled = false;
					this.chkActivo.Enabled = false;

					this.btnNuevo.Enabled = true;
					this.btnConsultar.Enabled = true;
					this.btnEditar.Enabled = true;
					this.btnGuardar.Enabled = false;
					this.btnCancelar.Enabled = true;
					this.btnConsultar.Enabled = true;
					this.txtbusqueda.Enabled= true;

					string sSQL = string.Format("Exec ConultaGeneralValidaUsuario_maquina  '{0}'", this.txtbusqueda.Text.ToString());
					this.ultgconsulta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		
			   // FuncionesProcedimientos.EstadoGrids(false, this.ultgconsulta);
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

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();	
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ConultaGeneralValidaUsuario_maquina  '{0}'", this.txtbusqueda.Text.ToString());
			this.ultgconsulta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			
			if (this.ultgconsulta.Rows.Count > 0) this.ultgconsulta.Focus(); else this.txtbusqueda.Focus();			

			this.btnNuevo.Enabled = true;
			this.btnEditar.Enabled = true;			
			this.btnCancelar.Enabled = true;	
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}

		private void txtbusqueda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode==Keys.Enter)
				this.btnConsultar_Click(sender,e);
		}

		private void ultgconsulta_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			#region Capturar datos del datagrid a los cajas de texto
			if(!bEdicion)
			{
				if (this.ultgconsulta.ActiveRow.Cells["idUsuario_maquina"].Value == DBNull.Value) return;

				this.txtidUsuario_Maquina.Value = (int)this.ultgconsulta.ActiveRow.Cells["idUsuario_maquina"].Value;
				this.txtUsuario.Text = this.ultgconsulta.ActiveRow.Cells["Usuario"].Value.ToString();
				this.txtEstacion.Text= this.ultgconsulta.ActiveRow.Cells["Estacion"].Value.ToString();
        this.chkActivo.Checked =  (bool)this.ultgconsulta.ActiveRow.Cells["Activo"].Value;

				this.btnNuevo.Enabled = true;
			//	this.btnConsultar.Enabled = true;
				this.btnEditar.Enabled = true;
				this.btnGuardar.Enabled = false;
				this.btnCancelar.Enabled = true;
			}
			#endregion Capturar datos del datagrid a los cajas de texto
		}
	}
}
