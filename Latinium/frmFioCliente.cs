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
	/// Descripción breve de frmFioCliente.
	/// </summary>
	public class frmFioCliente : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidFioCredito;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidFormularioSolicitud;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidFormularioCuenta;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroSolicitud;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaSolicitud;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaAcreditado;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaPago;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMonto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtComision;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNroCedula;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreAcreditado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTipoCuenta;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumcuenta;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNuovo;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBanco;
		private System.Windows.Forms.Label label15;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtOrigen;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label label16;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroCuenta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFioCliente()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int iFio = 0;
		public frmFioCliente(int IFio )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			iFio = IFio;
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
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFioCliente));
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
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
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtidFioCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidFormularioSolicitud = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidFormularioCuenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.txtNumeroSolicitud = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFechaSolicitud = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFechaAcreditado = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.dtFechaPago = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtMonto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtComision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label16 = new System.Windows.Forms.Label();
			this.txtNumeroCuenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtBanco = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtNumcuenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.txtTipoCuenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtNombreAcreditado = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtNroCedula = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.txtNuovo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.txtOrigen = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidFioCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidFormularioSolicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidFormularioCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroSolicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSolicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaAcreditado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMonto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumcuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTipoCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreAcreditado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNroCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNuovo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
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
			// txtidFioCredito
			// 
			this.txtidFioCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidFioCredito.Enabled = false;
			this.txtidFioCredito.Location = new System.Drawing.Point(632, 408);
			this.txtidFioCredito.Name = "txtidFioCredito";
			this.txtidFioCredito.PromptChar = ' ';
			this.txtidFioCredito.Size = new System.Drawing.Size(56, 21);
			this.txtidFioCredito.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidFioCredito.TabIndex = 869;
			this.txtidFioCredito.Visible = false;
			// 
			// txtidFormularioSolicitud
			// 
			this.txtidFormularioSolicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidFormularioSolicitud.Enabled = false;
			this.txtidFormularioSolicitud.Location = new System.Drawing.Point(632, 376);
			this.txtidFormularioSolicitud.Name = "txtidFormularioSolicitud";
			this.txtidFormularioSolicitud.PromptChar = ' ';
			this.txtidFormularioSolicitud.Size = new System.Drawing.Size(56, 21);
			this.txtidFormularioSolicitud.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidFormularioSolicitud.TabIndex = 870;
			this.txtidFormularioSolicitud.Visible = false;
			// 
			// txtidFormularioCuenta
			// 
			this.txtidFormularioCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidFormularioCuenta.Enabled = false;
			this.txtidFormularioCuenta.Location = new System.Drawing.Point(632, 344);
			this.txtidFormularioCuenta.Name = "txtidFormularioCuenta";
			this.txtidFormularioCuenta.PromptChar = ' ';
			this.txtidFormularioCuenta.Size = new System.Drawing.Size(56, 21);
			this.txtidFormularioCuenta.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidFormularioCuenta.TabIndex = 871;
			this.txtidFormularioCuenta.Visible = false;
			// 
			// txtidCliente
			// 
			this.txtidCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.Location = new System.Drawing.Point(568, 344);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(56, 21);
			this.txtidCliente.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidCliente.TabIndex = 872;
			this.txtidCliente.Visible = false;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(16, 18);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(40, 16);
			this.lblIdentificacion.TabIndex = 876;
			this.lblIdentificacion.Text = "Cédula";
			this.lblIdentificacion.Click += new System.EventHandler(this.lblIdentificacion_Click);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(184, 18);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(44, 16);
			this.lblCliente.TabIndex = 875;
			this.lblCliente.Text = "Nombre";
			this.lblCliente.Click += new System.EventHandler(this.lblCliente_Click);
			// 
			// txtRuc
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance1;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(64, 16);
			this.txtRuc.MaxLength = 10;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 21);
			this.txtRuc.TabIndex = 873;
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// txtNombre
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance2;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(240, 16);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(264, 21);
			this.txtNombre.TabIndex = 874;
			this.txtNombre.ValueChanged += new System.EventHandler(this.txtNombre_ValueChanged);
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(544, 18);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(44, 17);
			this.label14.TabIndex = 985;
			this.label14.Text = "Número";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumeroSolicitud
			// 
			this.txtNumeroSolicitud.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumeroSolicitud.Appearance = appearance3;
			this.txtNumeroSolicitud.Enabled = false;
			this.txtNumeroSolicitud.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumeroSolicitud.Location = new System.Drawing.Point(600, 15);
			this.txtNumeroSolicitud.Name = "txtNumeroSolicitud";
			this.txtNumeroSolicitud.Size = new System.Drawing.Size(112, 23);
			this.txtNumeroSolicitud.TabIndex = 984;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(16, 48);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(82, 16);
			this.label6.TabIndex = 987;
			this.label6.Text = "Fecha Solicitud";
			// 
			// dtFechaSolicitud
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaSolicitud.Appearance = appearance4;
			this.dtFechaSolicitud.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaSolicitud.DateButtons.Add(dateButton1);
			this.dtFechaSolicitud.Enabled = false;
			this.dtFechaSolicitud.Format = "dd/MM/yyyy";
			this.dtFechaSolicitud.Location = new System.Drawing.Point(112, 46);
			this.dtFechaSolicitud.Name = "dtFechaSolicitud";
			this.dtFechaSolicitud.NonAutoSizeHeight = 23;
			this.dtFechaSolicitud.Size = new System.Drawing.Size(112, 21);
			this.dtFechaSolicitud.SpinButtonsVisible = true;
			this.dtFechaSolicitud.TabIndex = 986;
			this.dtFechaSolicitud.Value = ((object)(resources.GetObject("dtFechaSolicitud.Value")));
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(248, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(93, 16);
			this.label1.TabIndex = 989;
			this.label1.Text = "Fecha Acreditado";
			// 
			// dtFechaAcreditado
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaAcreditado.Appearance = appearance5;
			this.dtFechaAcreditado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaAcreditado.DateButtons.Add(dateButton2);
			this.dtFechaAcreditado.Enabled = false;
			this.dtFechaAcreditado.Format = "dd/MM/yyyy";
			this.dtFechaAcreditado.Location = new System.Drawing.Point(352, 46);
			this.dtFechaAcreditado.Name = "dtFechaAcreditado";
			this.dtFechaAcreditado.NonAutoSizeHeight = 23;
			this.dtFechaAcreditado.Size = new System.Drawing.Size(112, 21);
			this.dtFechaAcreditado.SpinButtonsVisible = true;
			this.dtFechaAcreditado.TabIndex = 988;
			this.dtFechaAcreditado.Value = ((object)(resources.GetObject("dtFechaAcreditado.Value")));
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(496, 48);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(65, 16);
			this.label2.TabIndex = 991;
			this.label2.Text = "Fecha Pago";
			// 
			// dtFechaPago
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaPago.Appearance = appearance6;
			this.dtFechaPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaPago.DateButtons.Add(dateButton3);
			this.dtFechaPago.Enabled = false;
			this.dtFechaPago.Format = "dd/MM/yyyy";
			this.dtFechaPago.Location = new System.Drawing.Point(576, 46);
			this.dtFechaPago.Name = "dtFechaPago";
			this.dtFechaPago.NonAutoSizeHeight = 23;
			this.dtFechaPago.Size = new System.Drawing.Size(112, 21);
			this.dtFechaPago.SpinButtonsVisible = true;
			this.dtFechaPago.TabIndex = 990;
			this.dtFechaPago.Value = ((object)(resources.GetObject("dtFechaPago.Value")));
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(224, 155);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(58, 17);
			this.label3.TabIndex = 1015;
			this.label3.Text = "Comisión";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(408, 155);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(33, 17);
			this.label12.TabIndex = 1012;
			this.label12.Text = "Total";
			// 
			// txtMonto
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMonto.Appearance = appearance7;
			this.txtMonto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMonto.Enabled = false;
			this.txtMonto.FormatString = "#,##0.00";
			this.txtMonto.Location = new System.Drawing.Point(120, 152);
			this.txtMonto.Name = "txtMonto";
			this.txtMonto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMonto.PromptChar = ' ';
			this.txtMonto.Size = new System.Drawing.Size(96, 23);
			this.txtMonto.TabIndex = 1014;
			// 
			// txtComision
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComision.Appearance = appearance8;
			this.txtComision.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComision.Enabled = false;
			this.txtComision.FormatString = "#,##0.00";
			this.txtComision.Location = new System.Drawing.Point(296, 152);
			this.txtComision.MinValue = 0;
			this.txtComision.Name = "txtComision";
			this.txtComision.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtComision.PromptChar = ' ';
			this.txtComision.Size = new System.Drawing.Size(96, 23);
			this.txtComision.TabIndex = 1013;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.label16);
			this.groupBox1.Controls.Add(this.txtNumeroCuenta);
			this.groupBox1.Controls.Add(this.label13);
			this.groupBox1.Controls.Add(this.txtBanco);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Controls.Add(this.txtTotal);
			this.groupBox1.Controls.Add(this.label9);
			this.groupBox1.Controls.Add(this.txtNumcuenta);
			this.groupBox1.Controls.Add(this.label8);
			this.groupBox1.Controls.Add(this.txtTipoCuenta);
			this.groupBox1.Controls.Add(this.label7);
			this.groupBox1.Controls.Add(this.txtNombreAcreditado);
			this.groupBox1.Controls.Add(this.label5);
			this.groupBox1.Controls.Add(this.txtNroCedula);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.txtMonto);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.txtComision);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(8, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(704, 192);
			this.groupBox1.TabIndex = 1018;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Datos acreditación";
			// 
			// label16
			// 
			this.label16.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.Location = new System.Drawing.Point(528, 43);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(54, 17);
			this.label16.TabIndex = 1023;
			this.label16.Text = "Solicitud";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumeroCuenta
			// 
			this.txtNumeroCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumeroCuenta.Appearance = appearance9;
			this.txtNumeroCuenta.Enabled = false;
			this.txtNumeroCuenta.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumeroCuenta.Location = new System.Drawing.Point(592, 40);
			this.txtNumeroCuenta.Name = "txtNumeroCuenta";
			this.txtNumeroCuenta.Size = new System.Drawing.Size(104, 23);
			this.txtNumeroCuenta.TabIndex = 1022;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(16, 80);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(41, 17);
			this.label13.TabIndex = 1021;
			this.label13.Text = "Banco";
			// 
			// txtBanco
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBanco.Appearance = appearance10;
			this.txtBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBanco.Enabled = false;
			this.txtBanco.Location = new System.Drawing.Point(104, 80);
			this.txtBanco.MaxLength = 10;
			this.txtBanco.Name = "txtBanco";
			this.txtBanco.Size = new System.Drawing.Size(168, 23);
			this.txtBanco.TabIndex = 1020;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 155);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(95, 17);
			this.label4.TabIndex = 1019;
			this.label4.Text = "Monto Acreditar";
			// 
			// txtTotal
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance11;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(456, 152);
			this.txtTotal.MinValue = 0;
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(96, 23);
			this.txtTotal.TabIndex = 1018;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(16, 119);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(70, 17);
			this.label9.TabIndex = 884;
			this.label9.Text = "Nro Cuenta";
			// 
			// txtNumcuenta
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumcuenta.Appearance = appearance12;
			this.txtNumcuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumcuenta.Enabled = false;
			this.txtNumcuenta.Location = new System.Drawing.Point(104, 116);
			this.txtNumcuenta.MaxLength = 10;
			this.txtNumcuenta.Name = "txtNumcuenta";
			this.txtNumcuenta.Size = new System.Drawing.Size(168, 23);
			this.txtNumcuenta.TabIndex = 883;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(288, 80);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(74, 17);
			this.label8.TabIndex = 882;
			this.label8.Text = "Tipo Cuenta";
			// 
			// txtTipoCuenta
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTipoCuenta.Appearance = appearance13;
			this.txtTipoCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTipoCuenta.Enabled = false;
			this.txtTipoCuenta.Location = new System.Drawing.Point(376, 80);
			this.txtTipoCuenta.MaxLength = 10;
			this.txtTipoCuenta.Name = "txtTipoCuenta";
			this.txtTipoCuenta.Size = new System.Drawing.Size(168, 23);
			this.txtTipoCuenta.TabIndex = 881;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(192, 43);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(50, 17);
			this.label7.TabIndex = 880;
			this.label7.Text = "Nombre";
			// 
			// txtNombreAcreditado
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreAcreditado.Appearance = appearance14;
			this.txtNombreAcreditado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreAcreditado.Enabled = false;
			this.txtNombreAcreditado.Location = new System.Drawing.Point(256, 40);
			this.txtNombreAcreditado.Name = "txtNombreAcreditado";
			this.txtNombreAcreditado.Size = new System.Drawing.Size(264, 23);
			this.txtNombreAcreditado.TabIndex = 879;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(16, 43);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(45, 17);
			this.label5.TabIndex = 878;
			this.label5.Text = "Cédula";
			// 
			// txtNroCedula
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNroCedula.Appearance = appearance15;
			this.txtNroCedula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNroCedula.Enabled = false;
			this.txtNroCedula.Location = new System.Drawing.Point(72, 40);
			this.txtNroCedula.MaxLength = 10;
			this.txtNroCedula.Name = "txtNroCedula";
			this.txtNroCedula.Size = new System.Drawing.Size(112, 23);
			this.txtNroCedula.TabIndex = 877;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(16, 80);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(40, 16);
			this.label10.TabIndex = 1020;
			this.label10.Text = "Celular";
			// 
			// txtCelular
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance16;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Enabled = false;
			this.txtCelular.Location = new System.Drawing.Point(64, 78);
			this.txtCelular.MaxLength = 10;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(160, 21);
			this.txtCelular.TabIndex = 1019;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(240, 80);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(76, 16);
			this.label11.TabIndex = 1022;
			this.label11.Text = "Codigo Nuovo";
			// 
			// txtNuovo
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNuovo.Appearance = appearance17;
			this.txtNuovo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNuovo.Enabled = false;
			this.txtNuovo.Location = new System.Drawing.Point(336, 80);
			this.txtNuovo.MaxLength = 10;
			this.txtNuovo.Name = "txtNuovo";
			this.txtNuovo.Size = new System.Drawing.Size(160, 21);
			this.txtNuovo.TabIndex = 1021;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(352, 120);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 1023;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(16, 344);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(34, 16);
			this.label15.TabIndex = 1024;
			this.label15.Text = "Notas";
			// 
			// txtOrigen
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOrigen.Appearance = appearance18;
			this.txtOrigen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOrigen.Enabled = false;
			this.txtOrigen.Location = new System.Drawing.Point(64, 328);
			this.txtOrigen.Multiline = true;
			this.txtOrigen.Name = "txtOrigen";
			this.txtOrigen.Size = new System.Drawing.Size(488, 56);
			this.txtOrigen.TabIndex = 1025;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(104, 400);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1032;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click_1);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(360, 400);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 24);
			this.btnNuevo.TabIndex = 1030;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Visible = false;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(16, 400);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 1029;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Visible = false;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click_1);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(440, 400);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 24);
			this.btnEditar.TabIndex = 1028;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Visible = false;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(192, 400);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 1027;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Visible = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(280, 400);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 24);
			this.btnCancelar.TabIndex = 1026;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Visible = false;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click_1);
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(568, 384);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(56, 21);
			this.txtEstado.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtEstado.TabIndex = 1033;
			this.txtEstado.Visible = false;
			// 
			// frmFioCliente
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(720, 438);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.txtOrigen);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.txtNuovo);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.txtCelular);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.dtFechaPago);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFechaAcreditado);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtFechaSolicitud);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtNumeroSolicitud);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtidCliente);
			this.Controls.Add(this.txtidFormularioCuenta);
			this.Controls.Add(this.txtidFormularioSolicitud);
			this.Controls.Add(this.txtidFioCredito);
			this.KeyPreview = true;
			this.Name = "frmFioCliente";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "frmFioCliente";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFioCliente_KeyDown);
			this.Load += new System.EventHandler(this.frmFioCliente_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidFioCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidFormularioSolicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidFormularioCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroSolicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSolicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaAcreditado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMonto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumcuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTipoCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreAcreditado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNroCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNuovo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

	
		
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void btnCliente_Click(object sender, System.EventArgs e)
		{
			

			if ( iFio > 0 )
			{

			}
		}

		
		private void Consulta(int idFioCredito)
		{
			try
			{
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = false;
				//this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;				

				string sSQLBusqueda = string.Format("Exec consultaIndivisualRegistroFio  {0}", idFioCredito);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLBusqueda, true);
				dr.Read();
				if (dr.HasRows)
				{
					if (dr.GetValue(0) != System.DBNull.Value) this.txtidFioCredito.Value = dr.GetInt32(0);		
					if (dr.GetValue(1) != System.DBNull.Value) this.txtidFormularioSolicitud.Value = dr.GetInt32(1);		
					if (dr.GetValue(2) != System.DBNull.Value) this.txtidFormularioCuenta.Value = dr.GetInt32(2);		
					if (dr.GetValue(3) != System.DBNull.Value) this.txtidCliente.Value = dr.GetInt32(3);		
					if (dr.GetValue(4) != System.DBNull.Value) this.txtRuc.Text = dr.GetValue(4).ToString();
					if (dr.GetValue(5) != System.DBNull.Value) this.txtNombre.Text = dr.GetValue(5).ToString();
					if (dr.GetValue(6) != System.DBNull.Value) this.txtNumeroSolicitud.Text = dr.GetValue(6).ToString();
					if (dr.GetValue(7) != System.DBNull.Value) this.dtFechaSolicitud.Value = dr.GetDateTime(7);
					if (dr.GetValue(8) != System.DBNull.Value) this.dtFechaAcreditado.Value = dr.GetDateTime(8);
					if (dr.GetValue(9) != System.DBNull.Value) this.dtFechaPago.Value = dr.GetDateTime(9);
					if (dr.GetValue(10) != System.DBNull.Value) this.txtCelular.Text = dr.GetValue(10).ToString();
					if (dr.GetValue(11) != System.DBNull.Value) this.txtNuovo.Text = dr.GetValue(11).ToString();
					if (dr.GetValue(12) != System.DBNull.Value) this.txtNroCedula.Text = dr.GetValue(12).ToString();
					if (dr.GetValue(13) != System.DBNull.Value) this.txtNombreAcreditado.Text = dr.GetValue(13).ToString();
					if (dr.GetValue(14) != System.DBNull.Value) this.txtNumeroCuenta.Text = dr.GetValue(14).ToString();
					if (dr.GetValue(15) != System.DBNull.Value) this.txtBanco.Text = dr.GetValue(15).ToString();
					if (dr.GetValue(16) != System.DBNull.Value) this.txtTipoCuenta.Text = dr.GetValue(16).ToString();
					if (dr.GetValue(17) != System.DBNull.Value) this.txtNumcuenta.Text = dr.GetValue(17).ToString();
					if (dr.GetValue(18) != System.DBNull.Value) this.txtMonto.Value = dr.GetDecimal(18);
					if (dr.GetValue(19) != System.DBNull.Value) this.txtComision.Value = dr.GetDecimal(19);
					if (dr.GetValue(20) != System.DBNull.Value) this.txtTotal.Value = dr.GetDecimal(20);
					if (dr.GetValue(21) != System.DBNull.Value) this.txtEstado.Value = dr.GetInt32(21);		
					if (dr.GetValue(22) != System.DBNull.Value) this.txtOrigen.Text = dr.GetValue(22).ToString();

					

				}
				dr.Close();
								
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
      this.btnCancelar.Enabled = true;							
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}			
		}



		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
		
		}

	
		private void frmFioCliente_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				
				//if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click_1(sender, e);
				//				if ((int)this.txtidFioCredito.Value > 0)
				//				{	
				//					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
				//					//					if (e.KeyCode == Keys.I) 
				//					//					{
				//					//						if (miAcceso.Imprimir && !bImpreso) this.btnImprimir_Click(sender, e);
				//					//						if (miAcceso.ReImprimir && bImpreso) this.btnImprimir_Click(sender, e);
				//					//					}
				//					if (e.KeyCode == Keys.F5)
				//					{
				//						this.Consulta((int)this.txtidFioCredito.Value);				
				//					}
				//				}
				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			//			else
			//			{
			//				if (e.KeyCode == Keys.Escape) this.btnCancelar_Click_1(sender, e);
			//				//if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			//			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
		

		}

		private void lblCliente_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void lblIdentificacion_Click(object sender, System.EventArgs e)
		{
		
		}
		bool bNuevo = false;
		bool bEdicion = false;
		private void btnConsultar_Click_1(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click_1(sender, e);

			using (frmBuscaFioCliente Busqueda = new frmBuscaFioCliente())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.ultraGrid1.ActiveRow.Cells["idFioCredito"].Value);
				}
			}
		}




		private void btnCancelar_Click_1(object sender, System.EventArgs e)
		{

			#region vaciar
			this.txtRuc.Text= "";
			this.txtNombre.Text= "";
			this.txtNumeroSolicitud.Text= "";
			this.dtFechaSolicitud.Value= System.DBNull.Value;
			this.dtFechaAcreditado.Value= System.DBNull.Value;
			this.dtFechaPago.Value= System.DBNull.Value;
			this.txtCelular.Text= "";
			this.txtNuovo.Text= "";

			this.txtNroCedula.Text= "";
			this.txtNombreAcreditado.Text= "";
			this.txtNumeroCuenta.Text= "";
			this.txtBanco.Text= "";
			this.txtTipoCuenta.Text= "";
			this.txtNumcuenta.Text= "";
			this.txtMonto.Value = 0;
			this.txtComision.Value = 0;
			this.txtTotal.Value = 0;
			this.txtOrigen.Text= "";

			this.txtidCliente.Value = 0;
			this.txtidFormularioCuenta.Value = 0;
			this.txtidFioCredito.Value = 0;
			this.txtidFormularioSolicitud.Value = 0;
			this.txtEstado.Value = 0;
			#endregion vaciar

			#region bloqueo
			//this.dtFecha.Enabled = false;
			//this.cmbTipoRuc.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtNombre.Enabled = false;
			this.txtNumeroSolicitud.Enabled = false;
			this.dtFechaSolicitud.Enabled = false;
			this.dtFechaAcreditado.Enabled = false;
			this.dtFechaPago.Enabled = false;
			this.txtCelular.Enabled = false;
			this.txtNuovo.Enabled = false;

			this.txtNroCedula.Enabled = false;
			this.txtNombreAcreditado.Enabled = false;
			this.txtNumeroCuenta.Enabled = false;
			this.txtBanco.Enabled = false;
			this.txtTipoCuenta.Enabled = false;
			this.txtNumcuenta.Enabled = false;
			this.txtMonto.Enabled = false;
			this.txtComision.Enabled = false;
			this.txtTotal.Enabled = false;
			this.txtOrigen.Enabled = false;

			#endregion bloqueo
		}

		private void frmFioCliente_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "201202");

			if (!Funcion.Permiso("1916", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Credito ALFIO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar) this.btnConsultar.Visible = true;
			//if (miAcceso.Nuevo) this.btnDeposito.Visible = true;


			if (iFio > 0)
			{
				 this.Consulta(iFio);
			}
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnSalir_Click_1(object sender, System.EventArgs e)
		{
					this.Close ();
		}
	}
}
