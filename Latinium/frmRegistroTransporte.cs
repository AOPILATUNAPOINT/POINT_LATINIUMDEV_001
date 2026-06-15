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
	/// Descripción breve de frmRegistroTransporte.
	/// </summary>
	public class frmRegistroTransporte : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidRegistroTransporte;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label67;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label lblEstado;
		public System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtChofer;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtPlaca;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHoraInicio;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHoraFin;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaInicio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtKilometrajeI;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtKilometrajeF;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtLugarI;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtLugarF;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNota;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.GroupBox groupBox2;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredefinida;
		public System.Windows.Forms.Button brnSubir;
		private System.Windows.Forms.LinkLabel linkLabel1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRegistroTransporte()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		bool bPantalla =  false;
		int id = 0;
		public frmRegistroTransporte(bool BPantalla, int Id )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			bPantalla =  BPantalla;
			id = Id;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRegistroTransporte));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidRegistroTransporte = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label14 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtChofer = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtPlaca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label67 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.dtHoraInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHoraFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label12 = new System.Windows.Forms.Label();
			this.txtKilometrajeI = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtKilometrajeF = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtLugarI = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtLugarF = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.txtNota = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtBodegaPredefinida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.brnSubir = new System.Windows.Forms.Button();
			this.linkLabel1 = new System.Windows.Forms.LinkLabel();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidRegistroTransporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtChofer)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHoraInicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHoraFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaInicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtKilometrajeI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtKilometrajeF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLugarI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLugarF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredefinida)).BeginInit();
			this.SuspendLayout();
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(752, 304);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(56, 21);
			this.txtEstado.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtEstado.TabIndex = 868;
			this.txtEstado.Visible = false;
			// 
			// txtidRegistroTransporte
			// 
			this.txtidRegistroTransporte.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidRegistroTransporte.Enabled = false;
			this.txtidRegistroTransporte.Location = new System.Drawing.Point(688, 304);
			this.txtidRegistroTransporte.Name = "txtidRegistroTransporte";
			this.txtidRegistroTransporte.PromptChar = ' ';
			this.txtidRegistroTransporte.Size = new System.Drawing.Size(56, 21);
			this.txtidRegistroTransporte.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidRegistroTransporte.TabIndex = 867;
			this.txtidRegistroTransporte.Visible = false;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(16, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 870;
			this.label6.Text = "Fecha";
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
			this.dtFecha.Location = new System.Drawing.Point(64, 9);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 869;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(16, 43);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 872;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			this.cmbBodega.Location = new System.Drawing.Point(64, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 871;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(600, 8);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(44, 17);
			this.label14.TabIndex = 985;
			this.label14.Text = "Número";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance3;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(656, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(168, 23);
			this.txtNumero.TabIndex = 984;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(16, 72);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(70, 16);
			this.lblIdentificacion.TabIndex = 989;
			this.lblIdentificacion.Text = "Identificación";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(240, 72);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(85, 16);
			this.lblCliente.TabIndex = 988;
			this.lblCliente.Text = "N. Identificación";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbTipoRuc
			// 
			appearance4.FontData.Name = "Tahoma";
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance4;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 233;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(96, 69);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(136, 22);
			this.cmbTipoRuc.TabIndex = 986;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoRuc_KeyDown);
			this.cmbTipoRuc.ValueChanged += new System.EventHandler(this.cmbTipoRuc_ValueChanged);
			// 
			// txtRuc
			// 
			this.txtRuc.AcceptsTab = true;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance5;
			this.txtRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(336, 70);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(128, 21);
			this.txtRuc.TabIndex = 987;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(472, 72);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(56, 16);
			this.label3.TabIndex = 990;
			this.label3.Text = "Conductor";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtChofer
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtChofer.Appearance = appearance6;
			this.txtChofer.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtChofer.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtChofer.Enabled = false;
			this.txtChofer.Location = new System.Drawing.Point(536, 70);
			this.txtChofer.Name = "txtChofer";
			this.txtChofer.Size = new System.Drawing.Size(264, 21);
			this.txtChofer.TabIndex = 991;
			this.txtChofer.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtChofer_KeyDown);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(16, 104);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(48, 16);
			this.label7.TabIndex = 993;
			this.label7.Text = "Teléfono";
			// 
			// txtTelefono
			// 
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtTelefono.Appearance = appearance7;
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Location = new System.Drawing.Point(96, 102);
			this.txtTelefono.MaxLength = 10;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(136, 21);
			this.txtTelefono.TabIndex = 992;
			this.txtTelefono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTelefono_KeyDown);
			this.txtTelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTelefono_KeyPress);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(240, 104);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 16);
			this.label1.TabIndex = 995;
			this.label1.Text = "Placa";
			// 
			// txtPlaca
			// 
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtPlaca.Appearance = appearance8;
			this.txtPlaca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPlaca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPlaca.Enabled = false;
			this.txtPlaca.Location = new System.Drawing.Point(288, 102);
			this.txtPlaca.MaxLength = 7;
			this.txtPlaca.Name = "txtPlaca";
			this.txtPlaca.Size = new System.Drawing.Size(136, 21);
			this.txtPlaca.TabIndex = 994;
			this.txtPlaca.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPlaca_KeyDown);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(384, 138);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(48, 16);
			this.label4.TabIndex = 1022;
			this.label4.Text = "Hora Fin";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label67
			// 
			this.label67.AutoSize = true;
			this.label67.BackColor = System.Drawing.Color.Transparent;
			this.label67.Location = new System.Drawing.Point(216, 138);
			this.label67.Name = "label67";
			this.label67.Size = new System.Drawing.Size(58, 16);
			this.label67.TabIndex = 1020;
			this.label67.Text = "Hora Inicio";
			this.label67.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(16, 138);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(65, 16);
			this.label5.TabIndex = 1017;
			this.label5.Text = "Fecha Inicia";
			// 
			// dtHoraInicio
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHoraInicio.Appearance = appearance9;
			this.dtHoraInicio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHoraInicio.DateButtons.Add(dateButton2);
			this.dtHoraInicio.Enabled = false;
			this.dtHoraInicio.Format = "HH:mm";
			this.dtHoraInicio.Location = new System.Drawing.Point(288, 136);
			this.dtHoraInicio.Name = "dtHoraInicio";
			this.dtHoraInicio.NonAutoSizeHeight = 23;
			this.dtHoraInicio.Size = new System.Drawing.Size(88, 21);
			this.dtHoraInicio.SpinButtonsVisible = true;
			this.dtHoraInicio.TabIndex = 1023;
			this.dtHoraInicio.Value = ((object)(resources.GetObject("dtHoraInicio.Value")));
			this.dtHoraInicio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHoraInicio_KeyDown);
			// 
			// dtHoraFin
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHoraFin.Appearance = appearance10;
			this.dtHoraFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtHoraFin.DateButtons.Add(dateButton3);
			this.dtHoraFin.Enabled = false;
			this.dtHoraFin.Format = "HH:mm";
			this.dtHoraFin.Location = new System.Drawing.Point(440, 136);
			this.dtHoraFin.Name = "dtHoraFin";
			this.dtHoraFin.NonAutoSizeHeight = 23;
			this.dtHoraFin.Size = new System.Drawing.Size(88, 21);
			this.dtHoraFin.SpinButtonsVisible = true;
			this.dtHoraFin.TabIndex = 1021;
			this.dtHoraFin.Value = ((object)(resources.GetObject("dtHoraFin.Value")));
			this.dtHoraFin.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHoraFin_KeyDown);
			// 
			// dtFechaInicio
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaInicio.Appearance = appearance11;
			this.dtFechaInicio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtFechaInicio.DateButtons.Add(dateButton4);
			this.dtFechaInicio.Enabled = false;
			this.dtFechaInicio.Format = "dd/MM/yyyy";
			this.dtFechaInicio.Location = new System.Drawing.Point(96, 136);
			this.dtFechaInicio.Name = "dtFechaInicio";
			this.dtFechaInicio.NonAutoSizeHeight = 23;
			this.dtFechaInicio.Size = new System.Drawing.Size(112, 21);
			this.dtFechaInicio.SpinButtonsVisible = true;
			this.dtFechaInicio.TabIndex = 1016;
			this.dtFechaInicio.Value = ((object)(resources.GetObject("dtFechaInicio.Value")));
			this.dtFechaInicio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaInicio_KeyDown);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(432, 104);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(91, 16);
			this.label12.TabIndex = 1024;
			this.label12.Text = "Kilometraje Inicio";
			// 
			// txtKilometrajeI
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtKilometrajeI.Appearance = appearance12;
			this.txtKilometrajeI.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtKilometrajeI.Enabled = false;
			this.txtKilometrajeI.FormatString = "#,##0.00";
			this.txtKilometrajeI.Location = new System.Drawing.Point(536, 102);
			this.txtKilometrajeI.Name = "txtKilometrajeI";
			this.txtKilometrajeI.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtKilometrajeI.PromptChar = ' ';
			this.txtKilometrajeI.Size = new System.Drawing.Size(96, 21);
			this.txtKilometrajeI.TabIndex = 1025;
			this.txtKilometrajeI.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtKilometrajeI_KeyDown);
			this.txtKilometrajeI.Click += new System.EventHandler(this.txtKilometrajeI_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(640, 104);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 16);
			this.label2.TabIndex = 1026;
			this.label2.Text = "Kilometraje Fin";
			// 
			// txtKilometrajeF
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtKilometrajeF.Appearance = appearance13;
			this.txtKilometrajeF.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtKilometrajeF.Enabled = false;
			this.txtKilometrajeF.FormatString = "#,##0.00";
			this.txtKilometrajeF.Location = new System.Drawing.Point(720, 102);
			this.txtKilometrajeF.Name = "txtKilometrajeF";
			this.txtKilometrajeF.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtKilometrajeF.PromptChar = ' ';
			this.txtKilometrajeF.Size = new System.Drawing.Size(96, 21);
			this.txtKilometrajeF.TabIndex = 1027;
			this.txtKilometrajeF.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtKilometrajeF_KeyDown);
			this.txtKilometrajeF.Click += new System.EventHandler(this.txtKilometrajeF_Click);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(408, 192);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(75, 16);
			this.label15.TabIndex = 1030;
			this.label15.Text = "Lugar Destino";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(16, 192);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(86, 16);
			this.label9.TabIndex = 1028;
			this.label9.Text = "Lugar de Origen";
			// 
			// txtLugarI
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLugarI.Appearance = appearance14;
			this.txtLugarI.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtLugarI.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLugarI.Enabled = false;
			this.txtLugarI.Location = new System.Drawing.Point(104, 169);
			this.txtLugarI.Multiline = true;
			this.txtLugarI.Name = "txtLugarI";
			this.txtLugarI.Size = new System.Drawing.Size(296, 63);
			this.txtLugarI.TabIndex = 1031;
			this.txtLugarI.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLugarI_KeyDown);
			// 
			// txtLugarF
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLugarF.Appearance = appearance15;
			this.txtLugarF.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtLugarF.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLugarF.Enabled = false;
			this.txtLugarF.Location = new System.Drawing.Point(488, 169);
			this.txtLugarF.Multiline = true;
			this.txtLugarF.Name = "txtLugarF";
			this.txtLugarF.Size = new System.Drawing.Size(296, 63);
			this.txtLugarF.TabIndex = 1029;
			this.txtLugarF.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLugarF_KeyDown);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(16, 256);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(68, 16);
			this.label8.TabIndex = 1033;
			this.label8.Text = "Observación";
			// 
			// txtNota
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNota.Appearance = appearance16;
			this.txtNota.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNota.Enabled = false;
			this.txtNota.Location = new System.Drawing.Point(104, 240);
			this.txtNota.Multiline = true;
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(296, 63);
			this.txtNota.TabIndex = 1032;
			this.txtNota.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNota_KeyDown);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(448, 264);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(31, 16);
			this.label10.TabIndex = 1035;
			this.label10.Text = "Valor";
			// 
			// txtValor
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance17;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(488, 262);
			this.txtValor.MinValue = 0;
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(96, 21);
			this.txtValor.TabIndex = 1034;
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(568, 136);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 22);
			this.lblEstado.TabIndex = 1036;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(656, 344);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 1044;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Visible = false;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(416, 344);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1043;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(16, 344);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 24);
			this.btnNuevo.TabIndex = 1041;
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
			this.btnConsultar.Location = new System.Drawing.Point(96, 344);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 1040;
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
			this.btnEditar.TabIndex = 1039;
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
			this.btnGuardar.Location = new System.Drawing.Point(256, 344);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 1038;
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
			this.btnCancelar.Location = new System.Drawing.Point(336, 344);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 24);
			this.btnCancelar.TabIndex = 1037;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(-40, 328);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(928, 8);
			this.groupBox2.TabIndex = 1045;
			this.groupBox2.TabStop = false;
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
			// txtBodegaPredefinida
			// 
			this.txtBodegaPredefinida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredefinida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredefinida.Enabled = false;
			this.txtBodegaPredefinida.Location = new System.Drawing.Point(696, 272);
			this.txtBodegaPredefinida.Name = "txtBodegaPredefinida";
			this.txtBodegaPredefinida.PromptChar = ' ';
			this.txtBodegaPredefinida.Size = new System.Drawing.Size(16, 21);
			this.txtBodegaPredefinida.TabIndex = 1046;
			this.txtBodegaPredefinida.Visible = false;
			// 
			// brnSubir
			// 
			this.brnSubir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.brnSubir.CausesValidation = false;
			this.brnSubir.Image = ((System.Drawing.Image)(resources.GetObject("brnSubir.Image")));
			this.brnSubir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.brnSubir.Location = new System.Drawing.Point(744, 344);
			this.brnSubir.Name = "brnSubir";
			this.brnSubir.Size = new System.Drawing.Size(77, 23);
			this.brnSubir.TabIndex = 1047;
			this.brnSubir.Text = "&Subir PDF";
			this.brnSubir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.brnSubir.Visible = false;
			this.brnSubir.Click += new System.EventHandler(this.brnSubir_Click);
			// 
			// linkLabel1
			// 
			this.linkLabel1.Location = new System.Drawing.Point(656, 40);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(72, 16);
			this.linkLabel1.TabIndex = 1048;
			this.linkLabel1.TabStop = true;
			this.linkLabel1.Text = "Voucher";
			this.linkLabel1.Visible = false;
			this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
			// 
			// frmRegistroTransporte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(832, 382);
			this.Controls.Add(this.linkLabel1);
			this.Controls.Add(this.brnSubir);
			this.Controls.Add(this.txtBodegaPredefinida);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.txtNota);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.txtLugarI);
			this.Controls.Add(this.txtLugarF);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtKilometrajeF);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtKilometrajeI);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label67);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.dtHoraInicio);
			this.Controls.Add(this.dtHoraFin);
			this.Controls.Add(this.dtFechaInicio);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtPlaca);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.txtChofer);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtidRegistroTransporte);
			this.KeyPreview = true;
			this.Name = "frmRegistroTransporte";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "frmreppr";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRegistroTransporte_KeyDown);
			this.Load += new System.EventHandler(this.frmRegistroTransporte_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidRegistroTransporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtChofer)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHoraInicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHoraFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaInicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtKilometrajeI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtKilometrajeF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLugarI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLugarF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredefinida)).EndInit();
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
		bool bAnulado = false;
		bool bEliminaAlValidar = false;
		private bool bActualiza = true;
		int idBloqueaTransacciones = 0;
		string sCarpeta = "";
		string sUrlNube = "";
		string txtDirectorioarchivo ="";
		private void frmRegistroTransporte_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "201102");

			if (!Funcion.Permiso("1817", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Voucher de Transporte", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}


			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			this.txtBodegaPredefinida.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;


			//this.dtHoraInicio.CalendarInfo.MaxDate = DateTime.Today;
			//this.dtHoraFin.CalendarInfo.MaxDate = DateTime.Today;
			//this.dtHoraInicio.CalendarInfo.MinDate = DateTime.Today;
			//this.dtHoraFin.CalendarInfo.MinDate = DateTime.Today;

			string carpeta_1 ="";
			carpeta_1 = @"C:\Latinium\Reportes\ReporteTransporte1.rpt";
			if (!File.Exists(carpeta_1))
			{
				Process.Start("cmd.exe", "/C" + "net use \\192.168.1.10  /user:latinium 10dgpv15.p");
				Directory.GetDirectories(@"\\192.168.1.10\Proyectos");
				System.IO.File.Copy  (  @"\\192.168.1.10\Proyectos\Latinium\Reportes\ReporteTransporte1.rpt" ,@"C:\Latinium\Reportes\ReporteTransporte1.rpt", false);
			}
			if( id > 0)
			{
				Consulta(id);
				this.btnNuevo.Visible = false;
				this.btnConsultar.Visible = false;
				this.btnEditar.Visible = false;
				this.btnGuardar.Visible = false;
				this.btnCancelar.Visible = false;	

			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region vaciar
			this.dtFecha.Value = System.DBNull.Value;
			this.txtNumero.Text ="";
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbTipoRuc.Value = System.DBNull.Value;
			this.txtRuc.Text ="";
			this.txtChofer.Text ="";
			this.txtTelefono.Text ="";
			this.txtPlaca.Text ="";
			this.txtKilometrajeI.Value = 0;
			this.txtKilometrajeF.Value = 0;
			this.dtFechaInicio.Value = System.DBNull.Value;
			this.dtHoraInicio.Value = System.DBNull.Value;
			this.dtHoraFin.Value = System.DBNull.Value;
			this.txtLugarI.Text ="";
			this.txtLugarF.Text ="";
			this.txtNota.Text ="";
			this.txtValor.Value = 0;
			this.txtidRegistroTransporte.Value = 0;
			this.txtEstado.Value = 0;
			this.lblEstado.Text ="";
			#endregion vaciar
			#region bloqueo
			//this.dtFecha.Enabled = false;
			//this.txtNumero.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbTipoRuc.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtChofer.Enabled = false;
			this.txtTelefono.Enabled = false;
			this.txtPlaca.Enabled = false;
			this.txtKilometrajeI.Enabled = false;
			this.txtKilometrajeF.Enabled = false;
			this.dtFechaInicio.Enabled = false;
			this.dtHoraInicio.Enabled = false;
			this.dtHoraFin.Enabled = false;
			this.txtLugarI.Enabled = false;
			this.txtLugarF.Enabled = false;
			this.txtNota.Enabled = false;
			this.txtValor.Enabled = false;
			this.brnSubir.Visible = false;
			this.linkLabel1.Visible = false;
			txtDirectorioarchivo ="";
			sCarpeta = "";
			sUrlNube = "";

			//this.txtidRegistroTransporte.Enabled = false;
			//this.txtEstado.Enabled = false;
			//this.lblEstado.Enabled = false;
			#endregion bloqueo

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnImprimir.Visible = false;


		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
			{
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 1, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
			}
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFechaInicio, cdsSeteoF, false);
			this.dtHoraInicio.Value = DateTime.Today ;
			this.dtHoraFin.Value = DateTime.Today ;
			//FuncionesProcedimientos.RetornaFechaServidor(this.dtHoraInicio, cdsSeteoF, true);
     // FuncionesProcedimientos.RetornaFechaServidor(this.dtHoraFin, cdsSeteoF, true);
			if ((int)this.txtBodegaPredefinida.Value == 0) this.cmbBodega.Value = 1;
			else this.cmbBodega.Value = (int)this.txtBodegaPredefinida.Value;

			if ((int)this.txtBodegaPredefinida.Value == 0) this.cmbBodega.Enabled = true;
			else this.cmbBodega.Enabled = false;

			#region bloqueo
			//this.dtFecha.Enabled = false;
			//this.txtNumero.Enabled = false;
			//this.cmbBodega.Enabled = false;
			this.cmbTipoRuc.Enabled = true;
			this.txtRuc.Enabled = true;
			this.txtChofer.Enabled = true;
			this.txtTelefono.Enabled = true;
			this.txtPlaca.Enabled = true;
			this.txtKilometrajeI.Enabled = true;
			this.txtKilometrajeF.Enabled = true;
			this.dtFechaInicio.Enabled = true;
			this.dtHoraInicio.Enabled = true;
			this.dtHoraFin.Enabled = true;
			this.txtLugarI.Enabled = true;
			this.txtLugarF.Enabled = true;
			this.txtNota.Enabled = true;
			this.txtValor.Enabled = true;
			//this.txtidRegistroTransporte.Enabled = false;
			//this.txtEstado.Enabled = false;
			//this.lblEstado.Enabled = false;
			#endregion bloqueo
      this.txtEstado.Value = 1;
			this.lblEstado.Text = "PENDIENTE";
			this.brnSubir.Visible = false;

			this.linkLabel1.Visible = false;

			#region controles
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;	
			#endregion controles
			
			bNuevo = true;
			bEdicion = true;

			this.cmbTipoRuc.Focus();		

		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null)
			{				
				this.cmbTipoRuc.DataSource = FuncionesProcedimientos.Lista_TiposDeIdentificacion((int)this.cmbBodega.Value, 1, cdsSeteoF);
			}
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.cmbTipoRuc.Focus();
		}

		private void cmbTipoRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtRuc.Focus();
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtChofer.Focus();
		}

		private void txtChofer_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtTelefono.Focus();
		}

		private void txtTelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtPlaca.Focus();
		}

		private void txtPlaca_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtKilometrajeI.Focus();
		}

		private void txtKilometrajeI_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtKilometrajeF.Focus();
		}

		private void txtKilometrajeF_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtFechaInicio.Focus();
		}

		private void dtFechaInicio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtHoraInicio.Focus();
		}

		private void dtHoraInicio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtHoraFin.Focus();
		}

		private void dtHoraFin_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtLugarI.Focus();
		}

		private void txtLugarI_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtLugarF.Focus();
		}

		private void txtLugarF_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNota.Focus();
		}

		private void txtNota_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtValor.Focus();
		}

		private void txtValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void txtKilometrajeI_Click(object sender, System.EventArgs e)
		{
				this.txtKilometrajeI.SelectAll();
		}

		private void txtKilometrajeF_Click(object sender, System.EventArgs e)
		{
			this.txtKilometrajeF.SelectAll();
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void frmRegistroTransporte_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e); 																																																										 
				if (e.KeyCode == Keys.F5)	this.Consulta((int)this.txtidRegistroTransporte.Value);
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

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaGeneralRegistroTransporte Busqueda = new frmConsultaGeneralRegistroTransporte())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridCreditos.ActiveRow.Cells["idRegistroTransporte"].Value);
				}
			}
		}

		private void Consulta (int Id)
		{			
			string sSQL = string.Format("Exec ConsultaIndividualRegistroTransporte {0}", Id);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidRegistroTransporte.Value = Id;
				if (dr.GetValue(1) != System.DBNull.Value) this.dtFecha.Value = dr.GetDateTime(1);
				if (dr.GetValue(2) != System.DBNull.Value) this.cmbBodega.Value = dr.GetInt32(2);
				if (dr.GetValue(3) != System.DBNull.Value) this.txtNumero.Text = dr.GetString(3); 
				if (dr.GetValue(4) != System.DBNull.Value) this.txtChofer.Text= dr.GetString(4);
				if (dr.GetValue(5) != System.DBNull.Value) this.cmbTipoRuc.Value = dr.GetInt32(5);
				if (dr.GetValue(6) != System.DBNull.Value) this.txtRuc.Text= dr.GetString(6);
				if (dr.GetValue(7) != System.DBNull.Value) this.txtTelefono.Text= dr.GetString(7);
				if (dr.GetValue(8) != System.DBNull.Value) this.txtPlaca.Text= dr.GetString(8);
				if (dr.GetValue(9) != System.DBNull.Value) this.dtFechaInicio.Value = dr.GetDateTime(9);
				if (dr.GetValue(10) != System.DBNull.Value) this.dtHoraInicio.Value = dr.GetDateTime(10);
				if (dr.GetValue(11) != System.DBNull.Value) this.dtHoraFin.Value = dr.GetDateTime(11);
				if (dr.GetValue(12) != System.DBNull.Value) this.txtKilometrajeI.Value= dr.GetDecimal(12);
				if (dr.GetValue(13) != System.DBNull.Value) this.txtKilometrajeF.Value= dr.GetDecimal(13);
				if (dr.GetValue(14) != System.DBNull.Value) this.txtLugarI.Text= dr.GetString(14);
				if (dr.GetValue(15) != System.DBNull.Value) this.txtLugarF.Text= dr.GetString(15);
				if (dr.GetValue(16) != System.DBNull.Value) this.txtValor.Value = dr.GetDecimal(16);
			
				if (dr.GetValue(18) != System.DBNull.Value) this.txtEstado.Value = dr.GetInt32(18);
				if (dr.GetValue(19) != System.DBNull.Value) this.lblEstado.Text= dr.GetString(19);	
				if (dr.GetValue(17) != System.DBNull.Value) this.txtNota.Text= dr.GetString(17);
				if (dr.GetValue(20) != System.DBNull.Value) sCarpeta = dr.GetString(20).ToString();
				if (dr.GetValue(21) != System.DBNull.Value) sUrlNube = dr.GetString(21).ToString();
				if (sUrlNube.Length > 10 )
				{
					this.linkLabel1.Visible = true;
					//this.btnSubir.Visible = true;
				}
			}
			dr.Close();

			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar)
				if ((int)this.txtEstado.Value == 1 )
				{
					this.btnEditar.Enabled = true; 
					this.brnSubir.Visible = true;
				}

			if ((int)this.txtEstado.Value == 1 ) // PENDIENTE - AUTORIZADO
			{
				this.brnSubir.Visible = true;
			}
			if ((int)this.txtEstado.Value > 1 ) // PENDIENTE - AUTORIZADO
			{
				this.linkLabel1.Visible = true;
			}
			if (miAcceso.Imprimir)this.btnImprimir.Visible = true;
		}

	


		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			Consulta ((int)this.txtidRegistroTransporte.Value);
			
			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 1)
			{
				#region bloqueo
				//this.dtFecha.Enabled = false;
				//this.txtNumero.Enabled = false;
				this.cmbBodega.Enabled = false;
				this.cmbTipoRuc.Enabled = true;
				this.txtRuc.Enabled = true;
				this.txtChofer.Enabled = true;
				this.txtTelefono.Enabled = true;
				this.txtPlaca.Enabled = true;
				this.txtKilometrajeI.Enabled = true;
				this.txtKilometrajeF.Enabled = true;
				this.dtFechaInicio.Enabled = true;
				this.dtHoraInicio.Enabled = true;
				this.dtHoraFin.Enabled = true;
				this.txtLugarI.Enabled = true;
				this.txtLugarF.Enabled = true;
				this.txtNota.Enabled = true;
				this.txtValor.Enabled = true;
				//this.txtidRegistroTransporte.Enabled = false;
				//this.txtEstado.Enabled = false;
				//this.lblEstado.Enabled = false;
				#endregion bloqueo
			}

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.brnSubir.Visible = false;
			this.btnEditar.Enabled = false;			
			this.btnGuardar.Enabled = true;			
			this.btnCancelar.Enabled = true;
			this.btnImprimir.Visible = false;

			this.linkLabel1.Visible = false;
		//	this.btnAnular.Enabled = false;


			bNuevo = false;
			bEdicion = true;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", false, "VENTA", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;
			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select count(1) from dbo.TransporteValores where Bodega =  {0}", (int) this.cmbBodega.Value), true);
			if( iBodega == 0)
			{
				MessageBox.Show("Local Seleccionado no esta configurado, comunicarse Ext.506", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();	
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione el tipo de Identificación")) return;
			if (!Validacion.vbTexto(this.txtRuc, 6, 13, "Identificación")) return;
			if (!Validacion.vbTexto(this.txtChofer, 3, 120, "Nombre")) return;
		
		//	if (!Validacion.vbTexto(this.txtDireccion, 3, 200, "Dirección")) return;

			if ((int)this.txtTelefono.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el numero de teléfono convencional o celular del cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTelefono.Focus();	
				return;
			}
			if (!Validacion.vbValidaPlaca(this.txtPlaca)) return;
			if (!Validacion.vbCampoDecimalVacio (this.txtKilometrajeI,"Ingrese el Kilometraje Inicial",1,1000000)) return;	
			if (!Validacion.vbCampoDecimalVacio (this.txtKilometrajeF,"Ingrese el Kilometraje Final",1,1000000)) return;	
			if (Convert.ToDecimal (this.txtKilometrajeI.Value) > Convert.ToDecimal (this.txtKilometrajeF.Value))
			{
				MessageBox.Show("El Kilometraje Inicial no puede ser superior al Final.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtKilometrajeI.Focus();	
				return;
			}
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaInicio, "Seleccione la Fecha De Retorno", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHoraInicio, "Seleccione la Hora De Inicio", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHoraFin, "Seleccione la Hora De Retorno", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtFechaInicio.Value), Convert.ToDateTime(this.dtFecha.Value), this.dtFechaInicio, "La fecha de Inicio no puede ser superior a la fecha", false)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtHoraInicio.Value), Convert.ToDateTime(this.dtHoraFin.Value), this.dtHoraInicio, "La Hora de Inicio no puede ser superior a la Hora Final", false)) return;
			
			if (!Validacion.vbTexto(this.txtLugarI, 10, 200, "Origen")) return;
			if (!Validacion.vbTexto(this.txtLugarF, 10, 200, "Destino")) return;
			if (!Validacion.vbTexto(this.txtNota, 10, 200, "Observación")) return;
			//decimal dValor = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select top 1 Valor from dbo.TransporteValores where Bodega =  {0}", (int) this.cmbBodega.Value), true);
			decimal dValor = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select top 1 Valor from dbo.TransporteValores where Bodega =  {0}", (int) this.cmbBodega.Value));
	
			if (!Validacion.vbCampoDecimalVacio (this.txtValor,"Ingrese el Valor",1,dValor)) return;	
		
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

					#region Numeracion
					if (bNuevo)
					{
						string ssSQL = string.Format("Exec NumeracionLocales 82, {0}, 0",(int)this.cmbBodega.Value);
						oCmdActualiza.CommandText = ssSQL;
						this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();						
					}
					#endregion Numeracion

					#region Graba Maestro
					int iBodegaa = 0;
					int iTipoRuc = 0;
					if (this.cmbBodega.ActiveRow != null) iBodegaa = (int) this.cmbBodega.Value;
					if(this.cmbTipoRuc.ActiveRow != null) iTipoRuc = (int) this.cmbTipoRuc.Value;

					string  sSQL1 = string.Format("Exec [GuardaRegistroTransporte] {0}, {1}, {2}, {3}, {4}, {5}, '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', '{17}', {18}" ,
               (int)this.txtidRegistroTransporte.Value, //0
						iBodegaa, //1
						iTipoRuc, //2
						Convert.ToDecimal( this.txtKilometrajeI.Value), //3
						Convert.ToDecimal( this.txtKilometrajeF.Value), //4
						(int) this.txtEstado.Value,//5
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), //6
						(string) this.txtNumero.Text, //7
						(string) this.txtChofer.Text,//8
						(string) this.txtRuc.Text, //9
						(string) this.txtTelefono.Text, //10
						(string) this.txtPlaca.Text,//11
						Convert.ToDateTime(this.dtFechaInicio.Value).ToString("yyyyMMdd"), //12
						Convert.ToDateTime(this.dtHoraInicio.Value).ToString("yyyyMMdd HH:mm:ss"),//13
						Convert.ToDateTime(this.dtHoraFin.Value).ToString("yyyyMMdd HH:mm:ss"),//14
            (string) this.txtLugarI.Text, //15
						(string) this.txtLugarF.Text, //16
						(string) this.txtNota.Text,//17
						Convert.ToDecimal( this.txtValor.Value) //18
						);
					oCmdActualiza.CommandText = sSQL1;
					this.txtidRegistroTransporte.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Graba Maestro

					#region nube 
					if(txtDirectorioarchivo.Length > 10 )
					{
						NuevoGoogle (txtDirectorioarchivo);
					}
					if (sUrlNube.Length > 10 )
					{
						string sSQLGoogle = string.Format("Exec [GrabaUrlNubeTransporteNube] {0}, '{1}', '{2}'", (int)this.txtidRegistroTransporte.Value, sCarpeta, sUrlNube );
						oCmdActualiza.CommandText = sSQLGoogle;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion nube 


          
					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

					#region Controles
					#region bloqueo
					//this.dtFecha.Enabled = false;
					//this.txtNumero.Enabled = false;
					this.cmbBodega.Enabled = false;
					this.cmbTipoRuc.Enabled = false;
					this.txtRuc.Enabled = false;
					this.txtChofer.Enabled = false;
					this.txtTelefono.Enabled = false;
					this.txtPlaca.Enabled = false;
					this.txtKilometrajeI.Enabled = false;
					this.txtKilometrajeF.Enabled = false;
					this.dtFechaInicio.Enabled = false;
					this.dtHoraInicio.Enabled = false;
					this.dtHoraFin.Enabled = false;
					this.txtLugarI.Enabled = false;
					this.txtLugarF.Enabled = false;
					this.txtNota.Enabled = false;
					this.txtValor.Enabled = false;
					//this.txtidRegistroTransporte.Enabled = false;
					//this.txtEstado.Enabled = false;
					//this.lblEstado.Enabled = false;
					#endregion bloqueo

					if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
					if (miAcceso.Editar)this.btnEditar.Enabled = true;
          if (miAcceso.Imprimir)this.btnImprimir.Visible = true;
					this.btnGuardar.Enabled = false;
					this.btnCancelar.Enabled = true;

					this.Consulta((int)this.txtidRegistroTransporte.Value);
					
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

		private void cmbTipoRuc_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{

				if (bNuevo )
				{
					this.txtRuc.Enabled = true;
					this.txtRuc.Value = "";
					this.txtChofer.Value = "";
				}

				if (this.cmbTipoRuc.ActiveRow != null)
				{
					if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 6) this.txtRuc.MaxLength = 13;
					if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;

				}
			}
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) e.Cancel = true;
		}

		private void txtTelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			
				ParameterFields paramFields = new ParameterFields();

				ParameterField paramFieldCedula = new ParameterField ();
				ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
				paramFieldCedula.ParameterFieldName = "@idRegistroTransporte";
				ValCedula.Value = (int)this.txtidRegistroTransporte.Value;
				paramFieldCedula.CurrentValues.Add (ValCedula);
				paramFields.Add (paramFieldCedula);

				Reporte miRep = new Reporte("ReporteTransporte1.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Voucher Transporte");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();

		
		}

		public void  NuevoGoogle (string sDirectorio)
		{
			string carpeta_1="";
			string carpeta_2="";
			string carpeta_3="";
			string carpeta ="";
			//string txtDirectorioarchivo ="";
			string txtURL ="";


			#region Editar google 
			carpeta_1 = @"\\192.168.1.10\Proyectos\Transporte";
			carpeta_2 = @"\\192.168.1.10\Proyectos\Transporte\"+MenuLatinium.stDirFacturacion;

			carpeta = @"\\192.168.1.10\Proyectos\Transporte\"+MenuLatinium.stDirFacturacion+@"\" + this.txtNumero.Text.Replace(" ", "_") +"_"+this.txtRuc.Text+"_"+ Convert.ToString (this.txtidRegistroTransporte.Value);
			carpeta_3 = @""+MenuLatinium.stDirFacturacion+@"\" + this.txtRuc.Text +"_"+this.txtNumero.Text+"_"+ Convert.ToString (this.txtidRegistroTransporte.Value);
			try
			{
				Directory.GetDirectories(@"\\192.168.1.10\Proyectos");
				Process.Start("cmd.exe", "/C" + "net use \\192.168.1.10  /user:latinium 10dgpv15.p");
				if (!Directory.Exists(carpeta_1))
				{
					Directory.CreateDirectory(carpeta);
					
				}
				if (!Directory.Exists(carpeta_2))
				{
					Directory.CreateDirectory(carpeta);
					
				}

				if (!Directory.Exists(carpeta))
				{
					Directory.CreateDirectory(carpeta);
					
				}
				

				//	Directory.CreateDirectory(carpeta);

			
				txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",sDirectorio,@"\"));
					
				//	txtURL.Replace(" ", "_");
           
				carpeta  = (carpeta + @"\" + txtURL.Replace(" ", "_"));
          
				System.IO.File.Copy  (  txtDirectorioarchivo ,carpeta, false);
				sCarpeta = carpeta;
				sUrlNube = carpeta_3+ @"\" + txtURL.Replace(" ", "_");
				MessageBox.Show(string.Format(" Archivo '{0}' subido . ", txtURL.Replace(" ", "_")), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
								
			}
			catch (Exception ex)
			{
				MessageBox.Show ("Error: "+ ex.Message);

			}
			#endregion Editar google 


		}

		private void brnSubir_Click(object sender, System.EventArgs e)
		{
			if((int) this.txtidRegistroTransporte.Value == 0)
			{
				return;
			}
			#region valida dia
			string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaDiaSemanaTransporte] {0}, {1}, '{2}'", 0, (int) this.cmbBodega.Value, Convert.ToDateTime(this.dtFechaInicio.Value).ToString("yyyyMMdd")));
				
			if ((int)sMensaje.Length > 0 )
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			#endregion valida dia

			#region valida semana
			string sMensajes = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaDiaSemanaTransporte] {0}, {1}, '{2}'", 1, (int) this.cmbBodega.Value, Convert.ToDateTime(this.dtFechaInicio.Value).ToString("yyyyMMdd")));
				
			if ((int)sMensajes.Length  > 0)
			{
				MessageBox.Show(sMensajes, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			#endregion valida semana

			#region valida mes
			string sMensajem = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaDiaSemanaTransporte] {0}, {1}, '{2}'", 2, (int) this.cmbBodega.Value, Convert.ToDateTime(this.dtFechaInicio.Value).ToString("yyyyMMdd")));
				
			if ((int)sMensajem.Length > 0)
			{
				MessageBox.Show(sMensajem, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			#endregion valida mes

			try
			{
				MessageBox.Show("Acontinuación seleccione el voucher de transporte en formato PDF.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

				OpenFileDialog buscar = new OpenFileDialog ();
				buscar.Filter = "Text files (*.pdf)|*.pdf";
				if (buscar.ShowDialog () == DialogResult.OK)
				{
					txtDirectorioarchivo = buscar.FileName;
					//txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",txtDirectorioarchivo,@"\"));
					this.txtEstado.Value = 2;
					this.brnSubir.Visible = false;
					this.linkLabel1.Visible = true;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show ("Error: "+ ex.Message);
			}
			this.btnGuardar_Click (sender,e);
		}

		private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			if (sUrlNube.Length > 10 )
			{
				System.Diagnostics.Process.Start("chrome.exe", sUrlNube);
			}
		}
	}
}
