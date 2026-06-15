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
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmGestorDeDocumentos.
	/// </summary>
	public class frmGestorDeDocumentos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidGestorDocumentos;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidPErsonal;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoContrato;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPersonal;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtUrl;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaVence;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasVencido;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtestado;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmGestorDeDocumentos()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
    

		int idContrato = 0 ;
		public frmGestorDeDocumentos(int IdContrato)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idContrato = IdContrato;

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
		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;
		bool bError = true;

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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmGestorDeDocumentos));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoContrato");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleGestorContratos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGestorContratos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Vencimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iDiasVencimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("URL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bArchivo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleGestorContratos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGestorContratos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iDiasVencimiento");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("URL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bArchivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCarga", 0);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BtnVer", 1);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BtnDescargar", 2);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.label5 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label14 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtidGestorDocumentos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidPErsonal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbTipoContrato = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtUrl = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.dtFechaVence = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtDiasVencido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtestado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidGestorDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPErsonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoContrato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUrl)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaVence)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasVencido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtestado)).BeginInit();
			this.SuspendLayout();
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(94, 16);
			this.label5.TabIndex = 730;
			this.label5.Text = "Administrado Por:";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label5.Click += new System.EventHandler(this.label5_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 16);
			this.label1.TabIndex = 731;
			this.label1.Text = "Contrato:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 124);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(97, 16);
			this.label2.TabIndex = 732;
			this.label2.Text = "Detalle /Concepto:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 74);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(71, 16);
			this.label3.TabIndex = 733;
			this.label3.Text = "Fecha Firma:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtConcepto
			// 
			this.txtConcepto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtConcepto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Enabled = false;
			this.txtConcepto.Location = new System.Drawing.Point(112, 104);
			this.txtConcepto.MaxLength = 1000;
			this.txtConcepto.Multiline = true;
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(480, 56);
			this.txtConcepto.TabIndex = 750;
			this.txtConcepto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtConcepto_KeyDown);
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(112, 72);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 751;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Enabled = false;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(536, 8);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(44, 17);
			this.label14.TabIndex = 985;
			this.label14.Text = "Número";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label14.Visible = false;
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColor = System.Drawing.Color.Brown;
			appearance2.ForeColorDisabled = System.Drawing.Color.Brown;
			this.txtNumero.Appearance = appearance2;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(592, 5);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 23);
			this.txtNumero.TabIndex = 984;
			this.txtNumero.Visible = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(-192, 168);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1112, 8);
			this.groupBox2.TabIndex = 986;
			this.groupBox2.TabStop = false;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(424, 345);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1017;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 344);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 24);
			this.btnNuevo.TabIndex = 1015;
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
			this.btnConsultar.Location = new System.Drawing.Point(88, 344);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 1014;
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
			this.btnEditar.Location = new System.Drawing.Point(176, 344);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 24);
			this.btnEditar.TabIndex = 1013;
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
			this.btnGuardar.Location = new System.Drawing.Point(264, 344);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 1012;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(344, 344);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 24);
			this.btnCancelar.TabIndex = 1011;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtidGestorDocumentos
			// 
			this.txtidGestorDocumentos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidGestorDocumentos.Enabled = false;
			this.txtidGestorDocumentos.Location = new System.Drawing.Point(416, 8);
			this.txtidGestorDocumentos.Name = "txtidGestorDocumentos";
			this.txtidGestorDocumentos.PromptChar = ' ';
			this.txtidGestorDocumentos.Size = new System.Drawing.Size(96, 21);
			this.txtidGestorDocumentos.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidGestorDocumentos.TabIndex = 1018;
			this.txtidGestorDocumentos.Visible = false;
			// 
			// txtidPErsonal
			// 
			this.txtidPErsonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidPErsonal.Enabled = false;
			this.txtidPErsonal.Location = new System.Drawing.Point(416, 40);
			this.txtidPErsonal.Name = "txtidPErsonal";
			this.txtidPErsonal.PromptChar = ' ';
			this.txtidPErsonal.Size = new System.Drawing.Size(96, 21);
			this.txtidPErsonal.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidPErsonal.TabIndex = 1019;
			this.txtidPErsonal.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(bool);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11});
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
			// cmbTipoContrato
			// 
			this.cmbTipoContrato.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoContrato.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoContrato.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 280;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 280;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipoContrato.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoContrato.DisplayMember = "Nombre";
			this.cmbTipoContrato.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoContrato.Enabled = false;
			this.cmbTipoContrato.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoContrato.Location = new System.Drawing.Point(112, 40);
			this.cmbTipoContrato.Name = "cmbTipoContrato";
			this.cmbTipoContrato.Size = new System.Drawing.Size(288, 22);
			this.cmbTipoContrato.TabIndex = 1021;
			this.cmbTipoContrato.ValueMember = "idTipoContrato";
			this.cmbTipoContrato.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoContrato_KeyDown);
			// 
			// cmbPersonal
			// 
			this.cmbPersonal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonal.DataSource = this.ultraDataSource3;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 280;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbPersonal.DisplayMember = "Nombre";
			this.cmbPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPersonal.Enabled = false;
			this.cmbPersonal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonal.Location = new System.Drawing.Point(112, 8);
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.Size = new System.Drawing.Size(288, 22);
			this.cmbPersonal.TabIndex = 1022;
			this.cmbPersonal.ValueMember = "idPersonal";
			this.cmbPersonal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbPersonal_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// txtUrl
			// 
			this.txtUrl.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtUrl.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtUrl.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUrl.Enabled = false;
			this.txtUrl.Location = new System.Drawing.Point(600, 136);
			this.txtUrl.MaxLength = 1000;
			this.txtUrl.Multiline = true;
			this.txtUrl.Name = "txtUrl";
			this.txtUrl.Size = new System.Drawing.Size(88, 24);
			this.txtUrl.TabIndex = 1023;
			this.txtUrl.Visible = false;
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource2;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance3;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 300;
			ultraGridColumn8.Header.Caption = "Fecha Vence";
			ultraGridColumn8.Header.VisiblePosition = 3;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance4;
			ultraGridColumn9.Header.Caption = "Dias Aviso";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn10.Header.VisiblePosition = 11;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 8;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn12.Header.VisiblePosition = 6;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn13.Header.Caption = "Archivo";
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn14.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			appearance5.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellButtonAppearance = appearance5;
			ultraGridColumn14.Header.Caption = "...";
			ultraGridColumn14.Header.VisiblePosition = 8;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn14.Width = 28;
			ultraGridColumn15.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance6.Image = ((object)(resources.GetObject("appearance6.Image")));
			appearance6.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellButtonAppearance = appearance6;
			ultraGridColumn15.Header.Caption = "..";
			ultraGridColumn15.Header.VisiblePosition = 9;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn15.Width = 28;
			ultraGridColumn16.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance7.Image = ((object)(resources.GetObject("appearance7.Image")));
			appearance7.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellButtonAppearance = appearance7;
			ultraGridColumn16.Header.Caption = ".";
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn16.Width = 30;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn16});
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand3.SummaryFooterCaption = "Totales";
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance9;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8.5F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance11.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.uGridDetalle.Enabled = false;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 184);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(696, 152);
			this.uGridDetalle.TabIndex = 987;
			this.uGridDetalle.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalle_AfterRowInsert);
			this.uGridDetalle.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalle_BeforeRowsDeleted);
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			this.uGridDetalle.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_ClickCellButton);
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
			// 
			// dtFechaVence
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaVence.Appearance = appearance14;
			this.dtFechaVence.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaVence.DateButtons.Add(dateButton2);
			this.dtFechaVence.Enabled = false;
			this.dtFechaVence.Format = "dd/MM/yyyy";
			this.dtFechaVence.Location = new System.Drawing.Point(304, 72);
			this.dtFechaVence.Name = "dtFechaVence";
			this.dtFechaVence.NonAutoSizeHeight = 23;
			this.dtFechaVence.Size = new System.Drawing.Size(104, 21);
			this.dtFechaVence.SpinButtonsVisible = true;
			this.dtFechaVence.TabIndex = 1025;
			this.dtFechaVence.Value = ((object)(resources.GetObject("dtFechaVence.Value")));
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(232, 74);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(58, 16);
			this.label4.TabIndex = 1024;
			this.label4.Text = "Fecha Fin:";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(416, 74);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(61, 16);
			this.label6.TabIndex = 1026;
			this.label6.Text = "Dias Aviso:";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(600, 72);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 17);
			this.lblEstado.TabIndex = 1028;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasVencido
			// 
			this.txtDiasVencido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasVencido.Enabled = false;
			this.txtDiasVencido.Location = new System.Drawing.Point(480, 72);
			this.txtDiasVencido.Name = "txtDiasVencido";
			this.txtDiasVencido.PromptChar = ' ';
			this.txtDiasVencido.Size = new System.Drawing.Size(64, 21);
			this.txtDiasVencido.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtDiasVencido.TabIndex = 1030;
			// 
			// txtestado
			// 
			this.txtestado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtestado.Enabled = false;
			this.txtestado.Location = new System.Drawing.Point(616, 64);
			this.txtestado.Name = "txtestado";
			this.txtestado.PromptChar = ' ';
			this.txtestado.Size = new System.Drawing.Size(48, 21);
			this.txtestado.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtestado.TabIndex = 1031;
			this.txtestado.Visible = false;
			// 
			// frmGestorDeDocumentos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(712, 382);
			this.Controls.Add(this.txtestado);
			this.Controls.Add(this.txtDiasVencido);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtFechaVence);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtUrl);
			this.Controls.Add(this.cmbPersonal);
			this.Controls.Add(this.cmbTipoContrato);
			this.Controls.Add(this.txtidPErsonal);
			this.Controls.Add(this.txtidGestorDocumentos);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.uGridDetalle);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.txtConcepto);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label5);
			this.KeyPreview = true;
			this.Name = "frmGestorDeDocumentos";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Gestor de Contratos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGestorDeDocumentos_KeyDown);
			this.Load += new System.EventHandler(this.frmGestorDeDocumentos_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidGestorDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPErsonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoContrato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUrl)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaVence)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasVencido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtestado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void label5_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();		
		}

		private void frmGestorDeDocumentos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar)this.btnConsultar_Click(sender, e);
				if ((int)this.txtidGestorDocumentos.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.F5) if (miAcceso.Buscar)this.Consulta((int)this.txtidGestorDocumentos.Value);					
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
		
		static public void RetornaFechaS(Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha, C1.Data.C1DataSet ds, bool bHora)
		{
			DateTime dtFechaServ = Funcion.dtEscalarSQL(ds, "Exec RetornaFechaServidor");

			if (!bHora) dtFecha.Value = dtFechaServ.Date;
			if (bHora) dtFecha.Value = dtFechaServ;

//				dtFecha.CalendarInfo.MinDate = dtFechaServ.Date;
		} 

		private void frmGestorDeDocumentos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0225");

			if (!Funcion.Permiso("1337", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Proformas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.cmbPersonal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format(" Exec [ConsultaAdministradorDocumentos] {0},{1}", 0, 0 ));
			this.cmbTipoContrato.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT idTipoContrato, Nombre  FROM TipoContrato ");
//			FuncionesProcedimientos.EstadoGrids(false,this.uGridDetalle);

			this.uGridDetalle.DisplayLayout.Bands[0].Columns["btnCarga"].Hidden = true;
			this.uGridDetalle.DisplayLayout.Bands[0].Columns["URL"].Hidden = true;
      if (idContrato > 0) this.Consulta (idContrato);

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;
			bError = true;
			carpeta ="";
			this.txtUrl.Text ="";

			#region Vacia Datos
			this.txtidGestorDocumentos.Value=0;
			this.txtidGestorDocumentos.Value=0;
			this.txtNumero.Value = 0;
			this.cmbPersonal.Value = System.DBNull.Value;
			this.cmbTipoContrato.Value = System.DBNull.Value;
			this.dtHasta.Value = System.DBNull.Value;
			this.txtConcepto.Text ="";
			this.txtUrl.Text ="";
			this.dtFechaVence.Value = System.DBNull.Value;
			this.txtestado.Value = 0;
			this.txtDiasVencido.Value = 0;
			this.lblEstado.Text ="";

      #endregion Vacia Datos

			#region Bloquea Datos
			this.uGridDetalle.DisplayLayout.Bands[0].Columns["btnCarga"].Hidden = true;
			this.uGridDetalle.DisplayLayout.Bands[0].Columns["URL"].Hidden = true;

			this.cmbPersonal.Enabled = false;
			this.cmbTipoContrato.Enabled = false;
			this.dtHasta.Enabled = false;
			this.txtConcepto.Enabled = false;
			this.dtFechaVence.Enabled = false;
			this.txtDiasVencido.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false,this.uGridDetalle);
			string sSQLPar = string.Format("Exec ConsultaIndividualDetalleGestorContratos {0}", 0);
			FuncionesProcedimientos.dsGeneral(sSQLPar, this.uGridDetalle);
				#endregion Bloquea Datos

			#region Controles
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar)this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			#endregion Controles
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click (sender, e);
			RetornaFechaS(this.dtHasta, cdsSeteoF, true);
			RetornaFechaS(this.dtFechaVence, cdsSeteoF, true);

		   int idPersonal = 0;

				string ssqlDo = string.Format ("Exec ConsultaAdministradorDocumentos {0}, 1", MenuLatinium.IdUsuario ); 
				SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
				drDO.Read();
				if(drDO.HasRows)
				{
					idPersonal  = drDO.GetInt32 (0);
				}
				drDO.Close();

			

			#region Habilita Datos
			if(miAcceso.BIrA) this.cmbPersonal.Enabled = true;
			this.cmbTipoContrato.Enabled = true;
			this.dtHasta.Enabled = true;
			this.txtConcepto.Enabled = true;
			FuncionesProcedimientos.EstadoGrids(true,this.uGridDetalle);
			this.uGridDetalle.DisplayLayout.Bands[0].Columns["btnCarga"].Hidden = false;
			this.uGridDetalle.DisplayLayout.Bands[0].Columns["URL"].Hidden = true;
			this.dtFechaVence.Enabled = true;
			this.txtDiasVencido.Enabled = true;

			#endregion Habilita Datos

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
		
			if (miAcceso.BIrA)
			{
				this.cmbPersonal.Value = idPersonal;
				this.cmbPersonal.Focus ();
			}
			else
			{
				this.cmbPersonal.Value = idPersonal;
				this.cmbPersonal.Enabled = false;
				this.cmbTipoContrato.Focus ();
			}
			
			//vaeriables de validacion inicualizadas
			bNuevo = true;
			bEdicion = true;
			
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion			
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la Fecha de la frima del Contrato ", false, "", cdsSeteoF)) return;

			if (!Validacion.vbFechaEnDocumentos(this.dtFechaVence, "Seleccione la Fecha de Finalización del contrato ", false, "", cdsSeteoF)) return;

			if (!Validacion.vbComboVacio(this.cmbPersonal, "Seleccione el Administrador del Documento")) return;

			if (!Validacion.vbComboVacio(this.cmbTipoContrato, "Seleccione el tipo de contrato")) return;

			if (!Validacion.vbTexto(this.txtConcepto, 10, 298, "Detalle/Concepto")) return;

		  int	iRangodiaE = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" SELECT DATEDIFF (DAY, '{0}','{1}')", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
				         Convert.ToDateTime(this.dtFechaVence.Value).ToString("yyyyMMdd")));	
	
			if ( (int) this.txtDiasVencido.Value > iRangodiaE )
			{
				MessageBox.Show("El dia de aviso no puede ser mayor a " + iRangodiaE + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
      
			}


			#endregion Validacion

			#region Valida Grid
			if ((int) this.uGridDetalle.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese El Detalle ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridDetalle.Focus ();
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
			{	
			 int iRangodia = 0;
				
			
				if (dr.Cells["Fecha_Vencimiento"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la Fecha de Vencimiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDetalle.ActiveRow = dr;
					this.uGridDetalle.ActiveRow.Selected = true;
					return;
				}
				if (Convert.ToDateTime(dr.Cells["Fecha_Vencimiento"].Value) < Convert.ToDateTime(this.dtHasta.Value))
				{
					MessageBox.Show("La Fecha de Vencimiento no puede ser menor a la fecha de la firma del Contrato", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDetalle.ActiveRow = dr;
					this.uGridDetalle.ActiveRow.Selected = true;
					return;
				}
				
				if ((int)dr.Cells["iEstado"].Value == 0)
				{
			
					if (dr.Cells["Concepto"].Value.ToString().Trim().Length < 10 )
					{
						MessageBox.Show("El Detalle debe tener minimo 10 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Concepto"].Value.ToString().Trim().Length > 300 )
					{
						MessageBox.Show("El Detalle debe tener maximo 300 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}

					if (dr.Cells["iDiasVencimiento"].Value == System.DBNull.Value)
					{
						MessageBox.Show("El día de aviso no puede ser vacio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}

					if ((int)dr.Cells["iDiasVencimiento"].Value == 0)
					{
						MessageBox.Show("Los dias de aviso no puede estar en 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}

					if ((int)dr.Cells["iEstado"].Value == 0)
					{
						iRangodia = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" SELECT DATEDIFF (DAY, '{0}','{1}')", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
							Convert.ToDateTime(dr.Cells["Fecha_Vencimiento"].Value).ToString("yyyyMMdd")));		
						if ( (int)dr.Cells["iDiasVencimiento"].Value > iRangodia )
						{
							MessageBox.Show("El dia de aviso no puede ser mayor a " + iRangodia + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
      
						}

					}
				}
								
			}
			#endregion Valida Grid

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
					
					string  sSQL1 = string.Format("Exec GuardaGestorContratos {0}, {1}, {2}, '{3}', '{4}', '{5}', {6}, {7}, {8}" ,
						(int)this.txtidGestorDocumentos.Value, 
						(int) this.cmbPersonal.Value,
						(int)this.cmbTipoContrato.Value,
						Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 
						this.txtConcepto.Text.ToString(),
						Convert.ToDateTime(this.dtFechaVence.Value).ToString("yyyyMMdd"),
						(int) this.txtDiasVencido.Value, (int)this.txtestado.Value , 1);
					oCmdActualiza.CommandText = sSQL1;
					this.txtidGestorDocumentos.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Graba Maestro

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					{
						if ((int)dr.Cells["iEstado"].Value == 0)
						{
							string  sSQLDet = string.Format("Exec GuardaDetalleGestorContratos {0}, {1}, '{2}', '{3}', '{4}', {5}" ,
								Convert.ToInt32 ( dr.Cells["idDetalleGestorContratos"].Value), 
								(int) this.txtidGestorDocumentos.Value, 
								dr.Cells["Concepto"].Value,
								Convert.ToDateTime(dr.Cells["Fecha_Vencimiento"].Value).ToString("yyyyMMdd"),
								dr.Cells["URL"].Value,
								Convert.ToInt32 ( dr.Cells["iDiasVencimiento"].Value)
							
								);
							oCmdActualiza.CommandText = sSQLDet;
							oCmdActualiza.ExecuteNonQuery();	
						}
					}	
          
					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

					#region Controles


					#region Bloquea Datos
					this.cmbPersonal.Enabled = false;
					this.cmbTipoContrato.Enabled = false;
					this.dtHasta.Enabled = false;
					this.dtFechaVence.Enabled = false;
						this.txtDiasVencido.Enabled = false;
					this.txtConcepto.Enabled = false;
						FuncionesProcedimientos.EstadoGrids(false,this.uGridDetalle);
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["btnCarga"].Hidden = true;
					#endregion Bloquea Datos



					if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			  	if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
					if (miAcceso.Editar)this.btnEditar.Enabled = false;

					this.btnGuardar.Enabled = false;
					this.btnCancelar.Enabled = true;

					this.Consulta((int)this.txtidGestorDocumentos.Value);
					bNuevo = false;
					bEdicion = false;
					#endregion Controles		
					this.txtDiasVencido.Visible = true;

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

		private void cmbPersonal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbTipoContrato.Focus();
		}

		private void cmbTipoContrato_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtConcepto.Focus();
		}

		private void txtConcepto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaGeneralGestorContratos Busqueda = new frmConsultaGeneralGestorContratos(miAcceso.BIrA))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridCreditos.ActiveRow.Cells["idGestorContratos"].Value);
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["btnCarga"].Hidden = true;
//					this.uGridDetalle.DisplayLayout.Bands[0].Columns["URL"].Hidden = false;
					FuncionesProcedimientos.EstadoGrids(false,this.uGridDetalle);

				}
			}
		}

		private void Consulta (int Id)
		{			
			string sSQL = string.Format("Exec ConsultaIndividualGestorContratos {0}", Id);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidGestorDocumentos.Value = Id;
				   if (dr.GetValue(1) != System.DBNull.Value) this.cmbTipoContrato.Value = dr.GetInt32(1);
           if (dr.GetValue(2) != System.DBNull.Value) this.cmbPersonal.Value = dr.GetInt32(2);
				   if (dr.GetValue(3) != System.DBNull.Value) this.dtHasta.Value = dr.GetDateTime(3);
				   if (dr.GetValue(4) != System.DBNull.Value) this.txtConcepto.Text = dr.GetString(4);  
				   if (dr.GetValue(5) != System.DBNull.Value) this.dtFechaVence.Value = dr.GetDateTime(5);  
				   if (dr.GetValue(6) != System.DBNull.Value) this.txtestado.Value = dr.GetInt32(6);  
				   if (dr.GetValue(7) != System.DBNull.Value) this.txtDiasVencido.Value = dr.GetInt32(7);  
			}
			dr.Close();

			string sSQLPar = string.Format("Exec ConsultaIndividualDetalleGestorContratos {0}", (int)this.txtidGestorDocumentos.Value);
			FuncionesProcedimientos.dsGeneral(sSQLPar, this.uGridDetalle);


			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar)this.btnEditar.Enabled = true;
		}

		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["iEstado"].Value = 0;
			e.Row.Cells["idDetalleGestorContratos"].Value = 0;
			e.Row.Cells["idGestorContratos"].Value = 0;
			e.Row.Cells["iDiasVencimiento"].Value = 0;
			
			
		}

		private void uGridDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.uGridDetalle);
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			#region Habilita Datos
			this.Consulta ( (int) this.txtidGestorDocumentos.Value);
			if(miAcceso.BIrA) this.cmbPersonal.Enabled = true;
			this.cmbTipoContrato.Enabled = true;
			this.dtHasta.Enabled = true;
			this.txtConcepto.Enabled = true;
			this.dtFechaVence.Enabled = true;
			this.txtDiasVencido.Enabled = true;
			FuncionesProcedimientos.EstadoGrids(true,this.uGridDetalle);
			this.uGridDetalle.DisplayLayout.Bands[0].Columns["btnCarga"].Hidden = false;
//			this.uGridDetalle.DisplayLayout.Bands[0].Columns["URL"].Hidden = false;

			#endregion Habilita Datos

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;			
			this.btnGuardar.Enabled = true;			
			this.btnCancelar.Enabled = true;

			bNuevo = false;
			bEdicion = true;

		}

		string carpeta ;
		private void Archivo()
		{
			string sAdministrador = Funcion.sEscalarSQL(cdsSeteoF, string.Format
				 (" Select  (n.ApellidoPaterno + ' '+ n.ApellidoMaterno + ' '+ n.PrimerNombre + ' '+ n.SegundoNombre ) as Administrador from nomina n where idpersonal = {0}"
				,(int)this.cmbPersonal.Value));		
//				  carpeta = @"C:\Documents and Settings\Infoelec\Escritorio\" + MenuLatinium.stDirFacturacion+ @"\"+ sAdministrador+ (int)this.txtidGestorDocumentos.Value ;
			 carpeta = @"\\192.168.1.10\Proyectos\Latinium\CONTRATOS\" + MenuLatinium.stDirFacturacion+ @"\"+ sAdministrador+ (int)this.txtidGestorDocumentos.Value ;
//					 + sRuc+sNumeroFactura;  /* aqui creamos el nombre de la carpeta con el #cedual el cliente*/
//			carpeta = @"\\192.168.1.187\Latinium\ed\"+ sRuc+sNumeroFactura; 
			try
			{
				if (Directory.Exists(carpeta))
				{
					Directorio (carpeta , true);
				}
				else
				{
					Directory.CreateDirectory(carpeta);
					Directorio (carpeta , true);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show ("Error: "+ ex.Message);

			}
		}

		
		private void Directorio (string sCarpeta, bool bArchivoGuarda)
		{ 
			this.txtUrl.Text="";
			string URL ="" ;
			string sArchivo="";
			string sNomCarpeta ="";

			if (bArchivoGuarda)
			{
				try 
				{
					OpenFileDialog buscar = new OpenFileDialog ();
					if (buscar.ShowDialog () == DialogResult.OK)
					{
						sNomCarpeta = buscar.FileName;
						URL=  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",sNomCarpeta,@"\"));
						//						this.txtURL.Text = sArchivo;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show ("Error: "+ ex.Message);
					return;

				}
			}
			
			if ( URL.Length > 0)
			{
				try 
				{
					sCarpeta  = (sCarpeta + @"\" + URL);
					System.IO.File.Copy (  sNomCarpeta ,sCarpeta, false);
//					MessageBox.Show(string.Format("El archivo '{0}' se a Guardado en la carpeta '{1}'  ", URL,sCarpeta), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
					sNomCarpeta= sCarpeta;
					this.txtUrl.Text = sNomCarpeta;
				}
				catch (Exception ex)
				{
					bError = false;
					MessageBox.Show ("Error: "+ ex.Message);
					return;
				}
			}
			/*AQUI PONER EL CANCELAT */
		}

		private void uGridDetalle_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
			try 
			{
				
				if (e.Cell.Column.ToString() == "btnCarga")
				{
					if ((int)e.Cell.Row.Cells["iEstado"].Value > 0)
					{
						MessageBox.Show(string.Format("El registro ya fue REVISADO,  "), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

						return;
					}
					if (e.Cell.Row.Cells["URL"].Value.ToString().Length > 0)
					{
						if (DialogResult.Yes == MessageBox.Show(" Ya existe un Archivo Guardado.¿Desea Eliminar el Actual?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
						{
							Archivo();
						}
						else
						{
							return;
						}
					}
					else
					{
						Archivo();
					}
					if (bError)
					{
						e.Cell.Row.Cells["URL"].Value = Convert.ToString( this.txtUrl.Text);
						this.txtUrl.Text="";
					}
					
				}
				

				if (e.Cell.Column.ToString() == "BtnVer")
				{
					if (e.Cell.Row.Cells["URL"].Value.ToString().Length < 5)
					{
						MessageBox.Show(string.Format("No  existe un Archivo Guardado"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

						return;
					}
					else
					{
						Process.Start ("explorer.exe",e.Cell.Row.Cells["URL"].Value.ToString() );
						//FileStream archivo = new FileStream ( e.Cell.Row.Cells["DireccionUrlWSP"].Value.ToString() , FileMode.Open ); 
					}
				}

			}
			catch (Exception ex)
			{
				MessageBox.Show ("Error: "+ ex.Message);
				return;
			}
     bError = true;
		}
   bool bActualiza = false;

		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "Concepto")
			{	
				bActualiza = true;

				if ((int) uGridDetalle.ActiveRow.Cells["iEstado"].Value >0)
				{			
		
					if (uGridDetalle.ActiveRow.Cells["Concepto"].Value != System.DBNull.Value || uGridDetalle.ActiveRow.Cells["Concepto"].Value == System.DBNull.Value )
					{					
						e.Cell.Row.Cells["Concepto"].Value = (string)e.Cell.OriginalValue;
						//return;
					}
				}	
			
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "Fecha_Vencimiento")
			{	
				bActualiza = true;

				if ((int) uGridDetalle.ActiveRow.Cells["iEstado"].Value >0)
				{					
					if (uGridDetalle.ActiveRow.Cells["Fecha_Vencimiento"].Value != System.DBNull.Value || uGridDetalle.ActiveRow.Cells["Fecha_Vencimiento"].Value == System.DBNull.Value)
					{						
						e.Cell.Row.Cells["Fecha_Vencimiento"].Value = (DateTime)e.Cell.OriginalValue;
						//return;
					}
				}	
			
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "iDiasVencimiento")
			{	
				bActualiza = true;

				if ((int) uGridDetalle.ActiveRow.Cells["iEstado"].Value >0)
				{					
					if (uGridDetalle.ActiveRow.Cells["iDiasVencimiento"].Value != System.DBNull.Value || uGridDetalle.ActiveRow.Cells["iDiasVencimiento"].Value == System.DBNull.Value)
					{						
						e.Cell.Row.Cells["iDiasVencimiento"].Value = (int)e.Cell.OriginalValue;
					}
					//return;
				}	
			
				bActualiza = false;
			}
		}

		private void uGridDetalle_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			#region Elimina
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro seleccionado?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{				
				for (int i=0; i < e.Rows.Length; i++)
				{	
					if ((int)e.Rows[i].Cells["idDetalleGestorContratos"].Value > 0)
					{
						if (!miAcceso.Eliminar) return;

						string sSQL = string.Format("Delete From DetalleGestorContratos where idDetalleGestorContratos = {0} ", (int)e.Rows[i].Cells["idDetalleGestorContratos"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						e.DisplayPromptMsg = false;

					}
					else e.DisplayPromptMsg = false;
				}					
			}
			else e.Cancel = true;
			#endregion Elimina
		}

	}
}
