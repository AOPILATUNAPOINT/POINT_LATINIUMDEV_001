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
	/// Descripción breve de frmEntregaRecibosManuales.
	/// </summary>
	public class frmEntregaRecibosManuales : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label9;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNota;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHasta;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDesde;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidEntregaRecibosManuales;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbLocal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEntrega;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCodigo;
		private System.Windows.Forms.Label label5;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmEntregaRecibosManuales()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmEntregaRecibosManuales));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCodigoRecibo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidEntregaRecibosManuales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label9 = new System.Windows.Forms.Label();
			this.txtNota = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.txtHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDesde = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbLocal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCobrador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbEntrega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbCodigo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label5 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEntregaRecibosManuales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodigo)).BeginInit();
			this.SuspendLayout();
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Enabled = false;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(368, 96);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 1057;
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(432, 8);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(56, 21);
			this.txtEstado.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtEstado.TabIndex = 1056;
			this.txtEstado.Visible = false;
			// 
			// txtidEntregaRecibosManuales
			// 
			this.txtidEntregaRecibosManuales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidEntregaRecibosManuales.Enabled = false;
			this.txtidEntregaRecibosManuales.Location = new System.Drawing.Point(360, 8);
			this.txtidEntregaRecibosManuales.Name = "txtidEntregaRecibosManuales";
			this.txtidEntregaRecibosManuales.PromptChar = ' ';
			this.txtidEntregaRecibosManuales.Size = new System.Drawing.Size(56, 21);
			this.txtidEntregaRecibosManuales.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidEntregaRecibosManuales.TabIndex = 1055;
			this.txtidEntregaRecibosManuales.Visible = false;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(488, 192);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1054;
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
			this.btnNuevo.Location = new System.Drawing.Point(8, 192);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 24);
			this.btnNuevo.TabIndex = 1052;
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
			this.btnConsultar.Location = new System.Drawing.Point(88, 192);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 1051;
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
			this.btnEditar.Location = new System.Drawing.Point(168, 192);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 24);
			this.btnEditar.TabIndex = 1050;
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
			this.btnGuardar.Location = new System.Drawing.Point(248, 192);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 1049;
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
			this.btnCancelar.Location = new System.Drawing.Point(408, 192);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 24);
			this.btnCancelar.TabIndex = 1048;
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
			this.btnAnular.Location = new System.Drawing.Point(328, 192);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(72, 23);
			this.btnAnular.TabIndex = 1053;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Enabled = false;
			this.groupBox2.Location = new System.Drawing.Point(-16, 176);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(664, 8);
			this.groupBox2.TabIndex = 1042;
			this.groupBox2.TabStop = false;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(8, 136);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(28, 16);
			this.label9.TabIndex = 1041;
			this.label9.Text = "Nota";
			// 
			// txtNota
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNota.Appearance = appearance1;
			this.txtNota.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNota.Enabled = false;
			this.txtNota.Location = new System.Drawing.Point(64, 136);
			this.txtNota.Multiline = true;
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(560, 32);
			this.txtNota.TabIndex = 1040;
			this.txtNota.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNota_KeyDown);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Enabled = false;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(192, 104);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 1037;
			this.label8.Text = "Hasta";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 104);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(37, 16);
			this.label7.TabIndex = 1036;
			this.label7.Text = "Desde";
			// 
			// txtHasta
			// 
			this.txtHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHasta.Enabled = false;
			this.txtHasta.Location = new System.Drawing.Point(240, 104);
			this.txtHasta.Name = "txtHasta";
			this.txtHasta.PromptChar = ' ';
			this.txtHasta.Size = new System.Drawing.Size(104, 21);
			this.txtHasta.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtHasta.TabIndex = 1035;
			this.txtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtHasta_KeyDown);
			this.txtHasta.Enter += new System.EventHandler(this.txtHasta_Enter);
			// 
			// txtDesde
			// 
			this.txtDesde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDesde.Enabled = false;
			this.txtDesde.Location = new System.Drawing.Point(64, 104);
			this.txtDesde.Name = "txtDesde";
			this.txtDesde.PromptChar = ' ';
			this.txtDesde.Size = new System.Drawing.Size(104, 21);
			this.txtDesde.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtDesde.TabIndex = 1034;
			this.txtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDesde_KeyDown);
			this.txtDesde.Enter += new System.EventHandler(this.txtDesde_Enter);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Enabled = false;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(8, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(31, 16);
			this.label3.TabIndex = 1030;
			this.label3.Text = "Local";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Enabled = false;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(8, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(51, 16);
			this.label2.TabIndex = 1029;
			this.label2.Text = "Cobrador";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(312, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 1028;
			this.label1.Text = "Entrega";
			// 
			// cmbLocal
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbLocal.Appearance = appearance2;
			this.cmbLocal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbLocal.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 275;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbLocal.DisplayMember = "Nombre";
			this.cmbLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbLocal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbLocal.Enabled = false;
			this.cmbLocal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLocal.Location = new System.Drawing.Point(64, 40);
			this.cmbLocal.Name = "cmbLocal";
			this.cmbLocal.Size = new System.Drawing.Size(240, 21);
			this.cmbLocal.TabIndex = 1026;
			this.cmbLocal.ValueMember = "Bodega";
			this.cmbLocal.ValueChanged += new System.EventHandler(this.cmbLocal_ValueChanged);
			// 
			// cmbCobrador
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCobrador.Appearance = appearance3;
			this.cmbCobrador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 180;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbCobrador.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCobrador.DisplayMember = "Nombre";
			this.cmbCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCobrador.Enabled = false;
			this.cmbCobrador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador.Location = new System.Drawing.Point(64, 72);
			this.cmbCobrador.Name = "cmbCobrador";
			this.cmbCobrador.Size = new System.Drawing.Size(240, 21);
			this.cmbCobrador.TabIndex = 1025;
			this.cmbCobrador.ValueMember = "idNomina";
			this.cmbCobrador.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCobrador_KeyDown);
			// 
			// cmbEntrega
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEntrega.Appearance = appearance4;
			this.cmbEntrega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEntrega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 180;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbEntrega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEntrega.DisplayMember = "Nombre";
			this.cmbEntrega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntrega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEntrega.Enabled = false;
			this.cmbEntrega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntrega.Location = new System.Drawing.Point(368, 40);
			this.cmbEntrega.Name = "cmbEntrega";
			this.cmbEntrega.Size = new System.Drawing.Size(256, 21);
			this.cmbEntrega.TabIndex = 1024;
			this.cmbEntrega.ValueMember = "idNomina";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(8, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 1023;
			this.label6.Text = "Fecha";
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
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(136, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1022;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
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
			// cmbCodigo
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCodigo.Appearance = appearance6;
			this.cmbCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCodigo.Cursor = System.Windows.Forms.Cursors.WaitCursor;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbCodigo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCodigo.DisplayMember = "Nombre";
			this.cmbCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCodigo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCodigo.Enabled = false;
			this.cmbCodigo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodigo.Location = new System.Drawing.Point(368, 72);
			this.cmbCodigo.Name = "cmbCodigo";
			this.cmbCodigo.Size = new System.Drawing.Size(128, 21);
			this.cmbCodigo.TabIndex = 1059;
			this.cmbCodigo.ValueMember = "idCodigoRecibo";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Enabled = false;
			this.label5.Location = new System.Drawing.Point(320, 72);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(40, 16);
			this.label5.TabIndex = 1058;
			this.label5.Text = "Codigo";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmEntregaRecibosManuales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(632, 222);
			this.Controls.Add(this.cmbCodigo);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtidEntregaRecibosManuales);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtNota);
			this.Controls.Add(this.txtHasta);
			this.Controls.Add(this.txtDesde);
			this.Controls.Add(this.cmbLocal);
			this.Controls.Add(this.cmbCobrador);
			this.Controls.Add(this.cmbEntrega);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmEntregaRecibosManuales";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Entrega Recibos Manuales";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmEntregaRecibosManuales_KeyDown);
			this.Load += new System.EventHandler(this.frmEntregaRecibosManuales_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEntregaRecibosManuales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodigo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;
			#region Vacia
			this.dtFecha.Value  = System.DBNull.Value;
			this.cmbLocal.Value  = System.DBNull.Value;
			this.cmbEntrega.Value  = System.DBNull.Value;
			this.cmbCobrador.Value  = System.DBNull.Value;
			this.cmbCodigo.Value  = System.DBNull.Value;
			this.txtDesde.Value = 0;
			this.txtHasta.Value = 0;
			this.txtNota.Text ="";
			this.txtEstado.Value = 0;
			this.txtidEntregaRecibosManuales.Value = 0;
			this.lblEstado.Text ="";
			#endregion Vacia
			iBodegaPredefinida = 0;

			#region Bloquea

			this.cmbLocal.Enabled = false;
			this.cmbCodigo.Enabled = false;
			this.cmbEntrega.Enabled = false;
			this.cmbCobrador.Enabled = false;
			this.txtDesde.Enabled = false;
			this.txtHasta.Enabled = false;
			this.txtNota.Enabled = false;

			if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Enabled = false;

			#endregion Bloquea

		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click (sender, e);
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
    
			if (iBodegaPredefinida== 0) this.cmbLocal.Enabled = true; 
			else 
			{
				this.cmbLocal.Enabled = false;
				this.cmbLocal.Value = iBodegaPredefinida;
			} 
			#region Bloquea
			this.cmbEntrega.Enabled = true;
  		this.cmbLocal.Enabled = true;
			this.cmbCobrador.Enabled = true;
			this.cmbCodigo.Enabled = true;

			this.txtDesde.Enabled = true;
			this.txtHasta.Enabled = true;
			this.txtNota.Enabled = true;
			this.txtEstado.Value =1;
			this.lblEstado.Text = "PENDIENTE";
			#endregion Bloquea
      
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;	

			bNuevo = true;
			bEdicion = true;

			this.cmbLocal.Focus();	

		}

		private void cmbCobrador_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDesde.Focus();
		}

		private void txtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtHasta.Focus();
		}

		private void txtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNota.Focus();
		}

		private void txtNota_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void txtDesde_Enter(object sender, System.EventArgs e)
		{
				this.txtDesde.SelectAll();
		}

		private void txtHasta_Enter(object sender, System.EventArgs e)
		{
				this.txtHasta.SelectAll();
		}

		private void frmEntregaRecibosManuales_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidEntregaRecibosManuales.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.F5) this.Consulta((int)this.txtidEntregaRecibosManuales.Value);					
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
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;
		bool bEliminaAlValidar = false;
		int iBodegaPredefinida = 0;


		private void frmEntregaRecibosManuales_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0840");

			if (!Funcion.Permiso("1448", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Entrega de Recibos Manuales ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}


			this.cmbEntrega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RetornaNombres {0}", MenuLatinium.IdUsuario));
//			this.cmbLocal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC [RetornaBodega]"));


			this.cmbLocal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec FacturacionListaBodegas {0}, 24, '{1}'", MenuLatinium.IdUsuario, Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd")));
			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			iBodegaPredefinida  = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			




			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using ( frmConsultaGeneralReciboManualesRecibidos Busqueda = new frmConsultaGeneralReciboManualesRecibidos())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
				this.Consulta((int)Busqueda.uGridResumen.ActiveRow.Cells["idEntregaRecibosManuales"].Value);
				}
			}
		}

		int ibodega = 0;
		private void Consulta (int Id)
		{			
			ibodega = 0;
			string sSQL = string.Format("Exec [ConsultaIndividualEntregaReciboManuales] {0}", Id);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidEntregaRecibosManuales.Value = Id;
				if (dr.GetValue(1) != System.DBNull.Value) this.dtFecha.Value = dr.GetDateTime(1);
				if (dr.GetValue(2) != System.DBNull.Value) ibodega = dr.GetInt32(2);
				if (dr.GetValue(3) != System.DBNull.Value) this.cmbEntrega.Value = dr.GetInt32(3);
				if (dr.GetValue(4) != System.DBNull.Value) this.cmbCobrador.Value = dr.GetInt32(4);
				if (dr.GetValue(5) != System.DBNull.Value) this.txtDesde.Value = dr.GetInt32(5);
				if (dr.GetValue(6) != System.DBNull.Value) this.txtHasta.Value = dr.GetInt32(6);
				if (dr.GetValue(7) != System.DBNull.Value) this.txtNota.Text = dr.GetString(7);
				if (dr.GetValue(8) != System.DBNull.Value) this.lblEstado.Text = dr.GetString(8);
				if (dr.GetValue(9) != System.DBNull.Value) this.cmbCodigo.Value = dr.GetInt32(9);
				if (dr.GetValue(10) != System.DBNull.Value) this.txtHasta.Value = dr.GetInt32(6);
			}
			dr.Close();
			this.cmbLocal.Value = ibodega;

			this.cmbEntrega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RetornaNombres {0}", MenuLatinium.IdUsuario));

			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar)
				if ((int)this.txtEstado.Value == 0 || (int)this.txtEstado.Value == 1  ) 
					this.btnEditar.Enabled = true;

			if ((int)this.txtEstado.Value == 0 || (int)this.txtEstado.Value == 1) // PENDIENTE - AUTORIZADO
			{
				if (miAcceso.Anular)this.btnAnular.Enabled = true;
			}

			if ((int)this.txtEstado.Value == 2) // PENDIENTE - AUTORIZADO
			{
				this.btnAnular.Enabled = false;
			}

			//			if ((int)this.txtEstado.Value == 2) // PENDIENTE - AUTORIZADO
			//			{
			//				if (miAcceso.Anular)this.btnAnular.Enabled = true;
			//			}

			if ((int)this.txtEstado.Value == 0) this.lblEstado.Text ="PENDIENTE";
			if ((int)this.txtEstado.Value == 1) this.lblEstado.Text ="ENTREGADO";

		
	
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
	
	
		}

		private void cmbLocal_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbLocal.ActiveRow != null)
			{
				if (this.cmbCobrador.ActiveRow != null) this.cmbCobrador.Value = System.DBNull.Value;
				if (this.cmbEntrega.ActiveRow != null) this.cmbEntrega.Value = System.DBNull.Value;
				//				this.cmbRecibidopor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);

				if ((int) this.cmbLocal.Value == 1)
				{
					
	     this.cmbEntrega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RetornaNombres {0}", MenuLatinium.IdUsuario));
				}
				else
				{
					this.cmbEntrega.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Select  n.idNomina, (n.ApellidoPaterno + ' ' + n.ApellidoMaterno + ' ' + n.PrimerNombre + ' ' + n.SegundoNombre) As Nombre, n.Codigo From Com_AsignacionDeVendedores a Inner Join Nomina n On n.idPersonal = a.idPersonal Where a.idCargo in ( 2 ,1)And a.Bodega = {0} And '{1}' Between CONVERT(Date, a.Desde) And CONVERT(Date, a.Hasta)",
						this.cmbLocal.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));	
				}
				int iRegion = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Exec Listas {0}, {1}", 90, (int) this.cmbLocal.Value));
				this.cmbCobrador.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec Listas {0}, {1}",
					89, iRegion ));	

			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la Fecha", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbLocal, "Seleccione el Local")) return;
			if (!Validacion.vbComboVacio(this.cmbEntrega, "Seleccione la Persona que entrega")) return;
			if (!Validacion.vbComboVacio(this.cmbCobrador, "Seleccione el Cobrador ")) return;
	    if (!Validacion.vbComboVacio(this.cmbCodigo, "Seleccione el Codigo ")) return;
			if (!Validacion.vbCampoEnteroVacio (this.txtDesde,"Ingrese la numeración Inicial ",1,10000)) return;	
			if (!Validacion.vbCampoEnteroVacio (this.txtHasta,"Ingrese la numeración Final ",1,10000)) return;
			if (!Validacion.vbTexto(this.txtNota, 10, 100, "Nota")) return;
			#region Guardar

			if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Guardar ?, ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
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
						int idLocal = 0;
						int idEntrega= 0;
						int idCobrador = 0;
						int idCodigo = 0;

						if (this.cmbLocal.ActiveRow != null) idLocal = (int) this.cmbLocal.Value;
						if(this.cmbEntrega.ActiveRow != null) idEntrega = (int) this.cmbEntrega.Value;
						if (this.cmbCobrador.ActiveRow != null) idCobrador = (int) this.cmbCobrador.Value;
						if (this.cmbCodigo.ActiveRow != null) idCodigo = (int) this.cmbCodigo.Value;

						string  sSQL1 = string.Format("Exec [GrabaEntregaReciboManuales] {0}, '{1}', {2}, {3}, {4}, {5}, {6}, '{7}', {8}, {9}" ,
							(int)this.txtidEntregaRecibosManuales.Value,
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
							idLocal ,
							idEntrega ,
							idCobrador ,
							(int) this.txtDesde.Value,
							(int) this.txtHasta.Value,
							this.txtNota.Text,
							(int) this.txtEstado.Value,
							idCodigo
							);
						oCmdActualiza.CommandText = sSQL1;
						this.txtidEntregaRecibosManuales.Value = (int)oCmdActualiza.ExecuteScalar();
						#endregion Graba Maestro					

						oTransaction.Commit();

						Cursor = Cursors.Default;	

						MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

						this.lblEstado.Text ="ENTREGADO";

						#region Controles
					
						this.cmbEntrega.Enabled = false;
						this.cmbLocal.Enabled = false;
						this.cmbCobrador.Enabled = false;

						this.txtDesde.Enabled = false;
						this.txtHasta.Enabled = false;
						this.txtNota.Enabled = false;


						if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
						if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
						if (miAcceso.Editar)this.btnEditar.Enabled = false;

						this.btnGuardar.Enabled = false;
						this.btnCancelar.Enabled = true;
						this.Consulta((int)this.txtidEntregaRecibosManuales.Value);
						//					this.Consulta((int)this.txtidGastosviajes.Value);
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
			}
				
			#endregion Guardar

		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
					this.Close();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			Consulta ((int)this.txtidEntregaRecibosManuales.Value);
			
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}


			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value >=0)
			{

				this.cmbCobrador.Enabled = true;
				this.cmbLocal.Enabled = true;
				this.cmbEntrega.Enabled = true;
				this.txtDesde.Enabled = true;
				this.txtHasta.Enabled = true;;
				this.txtNota.Enabled = true;	
			}

			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;			
			this.btnGuardar.Enabled = true;			
			this.btnCancelar.Enabled = true;
			this.btnAnular.Enabled= false;

			bNuevo = false;
			bEdicion = true;
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consulta ((int) this.txtidEntregaRecibosManuales.Value);
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if((int) this.txtEstado.Value < 2)
			{
				if (!miAcceso.Anular)
				{
					MessageBox.Show("No tiene acceso Anular", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			

			if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Anular ?, ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				using (frmNotas Nota = new frmNotas( ))
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{
						string notas = Nota.txtNotas.Text;

						string sSQLAsiEs = string.Format("update EntregaRecibosManuales set Estado = 2, Nota = Nota + '{0}'  where idEntregaRecibosManuales = {1}", notas, (int)this.txtidEntregaRecibosManuales.Value);			
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiEs);

						this.Consulta ((int) this.txtidEntregaRecibosManuales.Value);
					}
				}
			}
		}
	}
}
