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
	/// Descripción breve de frmDepositos.
	/// </summary>
	public class frmDepositos : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidDeposito;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private System.Windows.Forms.Label lblCuenta;
		private System.Windows.Forms.Label label53;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtnDeposito;
		private System.Windows.Forms.Label lblnDeposito;
		private System.Windows.Forms.Label lblValor;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblCedula;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		bool bNuevo = false;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnGuardar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label lblNombre;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		private System.Windows.Forms.Label lblEstado;
		bool bEdicion = false;
		int idBloqueaTransacciones = 0;
		private Acceso miAcceso;
		public frmDepositos()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDepositos));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			this.txtidDeposito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblCuenta = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.txtnDeposito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblnDeposito = new System.Windows.Forms.Label();
			this.lblValor = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblCedula = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblNombre = new System.Windows.Forms.Label();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtidDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtnDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// txtidDeposito
			// 
			this.txtidDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidDeposito.Enabled = false;
			this.txtidDeposito.Location = new System.Drawing.Point(480, 16);
			this.txtidDeposito.Name = "txtidDeposito";
			this.txtidDeposito.PromptChar = ' ';
			this.txtidDeposito.Size = new System.Drawing.Size(8, 22);
			this.txtidDeposito.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidDeposito.TabIndex = 794;
			this.txtidDeposito.Visible = false;
			// 
			// dtFecha
			// 
			appearance1.FontData.Name = "Tahoma";
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 104);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 797;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			// 
			// cmbCuenta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance2;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 352;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.Enabled = false;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(72, 72);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(200, 21);
			this.cmbCuenta.TabIndex = 798;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCuenta_KeyDown);
			// 
			// lblCuenta
			// 
			this.lblCuenta.AutoSize = true;
			this.lblCuenta.BackColor = System.Drawing.Color.Transparent;
			this.lblCuenta.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblCuenta.Location = new System.Drawing.Point(8, 74);
			this.lblCuenta.Name = "lblCuenta";
			this.lblCuenta.Size = new System.Drawing.Size(39, 17);
			this.lblCuenta.TabIndex = 799;
			this.lblCuenta.Text = "Cuenta";
			this.lblCuenta.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label53.Location = new System.Drawing.Point(8, 106);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 17);
			this.label53.TabIndex = 800;
			this.label53.Text = "Fecha";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtnDeposito
			// 
			this.txtnDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtnDeposito.Enabled = false;
			this.txtnDeposito.Location = new System.Drawing.Point(72, 136);
			this.txtnDeposito.MinValue = 0;
			this.txtnDeposito.Name = "txtnDeposito";
			this.txtnDeposito.PromptChar = ' ';
			this.txtnDeposito.Size = new System.Drawing.Size(104, 22);
			this.txtnDeposito.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtnDeposito.TabIndex = 801;
			this.txtnDeposito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtnDeposito_KeyDown);
			// 
			// lblnDeposito
			// 
			this.lblnDeposito.AutoSize = true;
			this.lblnDeposito.BackColor = System.Drawing.Color.Transparent;
			this.lblnDeposito.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblnDeposito.Location = new System.Drawing.Point(8, 138);
			this.lblnDeposito.Name = "lblnDeposito";
			this.lblnDeposito.Size = new System.Drawing.Size(44, 17);
			this.lblnDeposito.TabIndex = 802;
			this.lblnDeposito.Text = "Numero";
			this.lblnDeposito.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblValor
			// 
			this.lblValor.AutoSize = true;
			this.lblValor.BackColor = System.Drawing.Color.Transparent;
			this.lblValor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblValor.Location = new System.Drawing.Point(8, 170);
			this.lblValor.Name = "lblValor";
			this.lblValor.Size = new System.Drawing.Size(30, 17);
			this.lblValor.TabIndex = 803;
			this.lblValor.Text = "Valor";
			this.lblValor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGuardar
			// 
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(248, 216);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 810;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardarComoCliente_Click);
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
			this.btnEditar.Location = new System.Drawing.Point(168, 216);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(72, 23);
			this.btnEditar.TabIndex = 807;
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
			this.btnCancelar.Location = new System.Drawing.Point(336, 216);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 808;
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
			this.btnSalir.Location = new System.Drawing.Point(416, 216);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 809;
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
			this.btnNuevo.Location = new System.Drawing.Point(8, 216);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 805;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(88, 216);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(72, 23);
			this.btnConsultar.TabIndex = 806;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(464, 16);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(8, 22);
			this.txtEstado.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtEstado.TabIndex = 811;
			this.txtEstado.Visible = false;
			// 
			// lblCedula
			// 
			this.lblCedula.AutoSize = true;
			this.lblCedula.BackColor = System.Drawing.Color.Transparent;
			this.lblCedula.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCedula.Location = new System.Drawing.Point(8, 10);
			this.lblCedula.Name = "lblCedula";
			this.lblCedula.Size = new System.Drawing.Size(38, 17);
			this.lblCedula.TabIndex = 812;
			this.lblCedula.Text = "Cedula";
			this.lblCedula.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance3.BackColorDisabled = System.Drawing.Color.Silver;
			appearance3.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance3.BorderColor = System.Drawing.Color.Black;
			appearance3.BorderColor3DBase = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance3;
			this.txtRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(72, 8);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(104, 22);
			this.txtRuc.TabIndex = 813;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 200);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(504, 8);
			this.groupBox1.TabIndex = 814;
			this.groupBox1.TabStop = false;
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
			// txtValor
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance4;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(72, 168);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(104, 22);
			this.txtValor.TabIndex = 815;
			// 
			// lblNombre
			// 
			this.lblNombre.AutoSize = true;
			this.lblNombre.BackColor = System.Drawing.Color.Transparent;
			this.lblNombre.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblNombre.Location = new System.Drawing.Point(72, 40);
			this.lblNombre.Name = "lblNombre";
			this.lblNombre.Size = new System.Drawing.Size(0, 20);
			this.lblNombre.TabIndex = 816;
			this.lblNombre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtidCliente
			// 
			this.txtidCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.Location = new System.Drawing.Point(448, 16);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(8, 22);
			this.txtidCliente.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidCliente.TabIndex = 817;
			this.txtidCliente.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(232, 167);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 818;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmDepositos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(504, 246);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtidCliente);
			this.Controls.Add(this.lblNombre);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.lblCedula);
			this.Controls.Add(this.lblValor);
			this.Controls.Add(this.lblnDeposito);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.lblCuenta);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.txtnDeposito);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtidDeposito);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmDepositos";
			this.Text = "Depositos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDepositos_KeyDown);
			this.Load += new System.EventHandler(this.frmDepositos_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtidDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtnDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmDepositos_Load(object sender, System.EventArgs e)
		{
			
			miAcceso = new Acceso(cdsSeteoF, "0230");

			if (!Funcion.Permiso("1509",cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Depositos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			
			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCuentasCierreCaja");
			
		}

		private void txtValor_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)

		{
			if (idBloqueaTransacciones > 0)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 4", idBloqueaTransacciones, (int)this.txtidDeposito.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			bNuevo = false;
			bEdicion = false;

			#region vacia Campos
      this.txtnDeposito.Value=0;
			this.cmbCuenta.Value = System.DBNull.Value;
			this.txtRuc.Text="";
			this.dtFecha.Value= System.DBNull.Value;
			this.txtnDeposito.Value=0;
			this.txtValor.Value =0;
	    this.txtEstado.Value =0;
			this.txtidCliente.Value=0;
			this.lblEstado.Text ="";
			this.lblNombre.Text ="";
			this.txtidDeposito.Value=0;
			#endregion vacia Campos

			#region Bloquea Campos
			this.txtnDeposito.Enabled=false;
			this.cmbCuenta.Enabled=false;
			this.txtRuc.Enabled=false;
			this.dtFecha.Enabled=false;
			this.txtnDeposito.Enabled=false;
			this.txtValor.Enabled=false;
			this.txtEstado.Enabled=false;
			#endregion bloquea Campos

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			idBloqueaTransacciones = 0;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{			
			this.btnCancelar_Click(sender, e);

			this.dtFecha.Value = DateTime.Today;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;	
			this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-7);
			
			#region Desbloquea Campos
			this.txtnDeposito.Enabled = true;
			this.cmbCuenta.Enabled =  true;
			this.txtRuc.Enabled = true;
			this.dtFecha.Enabled = true;
			this.txtnDeposito.Enabled = true;
			this.txtValor.Enabled = true;
			this.txtEstado.Enabled = true;
			#endregion Desbloquea  Campos
      
			this.txtEstado.Value = 1;
			 this.Estado();
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;		
			
			//vaeriables de validacion inicualizadas
			bNuevo = true;
			bEdicion = true;
			//salto de objeto a

			this.txtRuc.Focus();
		}

		private void  Estado ()
		{
			if((int)this.txtEstado.Value == 1){
			
			this.lblEstado.Text ="PENDIENTE";
			}
			if((int)this.txtEstado.Value == 2)
			{
			
			this.lblEstado.Text ="AUTORIZADO";
			}
			if((int)this.txtEstado.Value == 3)
			{
			
			this.lblEstado.Text ="NO AUTORIZADO";
			}
			if((int)this.txtEstado.Value == 4)
			{
			this.lblEstado.Text ="ANULADO";
			
			}
		}
		
		
		
		private void Consulta (int idDeposito)
		{
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

			#region Consulta Individual
			string ssql= string.Format ("Exec  ConsultaDepositos {0}",idDeposito); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidDeposito.Value = dr.GetInt32(0);
				if(dr.GetValue(1)!= System.DBNull.Value)this.cmbCuenta.Value = dr.GetInt32(1);
				if(dr.GetValue(2)!= System.DBNull.Value)this.txtidCliente.Value = dr.GetInt32(2);
				if(dr.GetValue(3)!= System.DBNull.Value)this.txtRuc.Text = dr.GetString(3);
				if(dr.GetValue(4)!= System.DBNull.Value)this.dtFecha.Value = dr.GetDateTime(4);
				if(dr.GetValue(5)!= System.DBNull.Value)this.txtnDeposito.Value = dr.GetInt32(5);
				if(dr.GetValue(6)!= System.DBNull.Value)this.txtValor.Value = dr.GetDecimal(6);
				if(dr.GetValue(7)!= System.DBNull.Value)this.txtEstado.Value = dr.GetInt32(7);
				if(dr.GetValue(8)!= System.DBNull.Value)this.lblNombre.Text = dr.GetString(8);
			}
			dr.Close();
			#endregion Consulta Individual
			this.Estado();
			
		}

		
		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscarDeposito bd = new frmBuscarDeposito(true))
			{
				if (DialogResult.OK == bd.ShowDialog())
				{
				this.Consulta((int)bd.ultraGrid1.ActiveRow.Cells["idDeposito"].Value);
				}
			}
     this.Estado();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			if (Funcion.ValidaBloqueo((int)this.txtidDeposito.Value, 4, cdsSeteoF)) return;
			this.Consulta((int)this.txtidDeposito.Value);
			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 4", (int)this.txtidDeposito.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);
			#region Desbloquea Campos
			this.txtnDeposito.Enabled = true;
			this.cmbCuenta.Enabled =  true;
			this.txtRuc.Enabled = true;
			this.dtFecha.Enabled = true;
			this.txtnDeposito.Enabled = true;
			this.txtValor.Enabled = true;
			this.txtEstado.Enabled = true;
			#endregion Desbloquea  Campos
      
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			
			bEdicion=true;
		}

		

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();		
		}

		private void frmDepositos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidDeposito.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.F5) this.Consulta((int)this.txtidDeposito.Value);					
				}
				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardarComoCliente_Click(sender, e);				
			}
		}

		private void btnGuardarComoCliente_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (!Validacion.vbComboVacio(this.cmbCuenta, "Seleccione una cuenta")) return;
			if (!Validacion.vbTexto(this.txtRuc, 10, 13, "Por favor ingrese C.I/RUC ")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha del deposito", true, "", cdsSeteoF)) return;
    //if (!Validacion.vbTexto(this.txtnDeposito, 2, 20, "Por favor el # del Deposito")) return;
			
			if (this.txtnDeposito.Text.ToString().Length < 4)
			{
				MessageBox.Show("Ingrese el numero del Deposito ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtnDeposito.Focus();
				return ;
			}
			
			if ( this.txtnDeposito.Text.ToString().Length > 21)
			{
				MessageBox.Show("El maximo de numeros es de 20 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtnDeposito.Focus();
				return ;
			}

			if (!Validacion.vbCampoDecimalVacio(this.txtValor, "El Valor del deposito no puede ser 0.00", 0, 2100)) return;
	

			#endregion Validacion

			#region Validadación existencia

			int idDPorLiquidar = Funcion.iEscalarSQL(cdsSeteoF,string.Format("if Exists(select idDepositosPorLiquidar from DepositosPorLiquidar  where Numero = REPLACE('{0}',' ','') )select idDepositosPorLiquidar from DepositosPorLiquidar  where Numero = REPLACE('{0}',' ','') Else Select CONVERT(Int, 0)",
				this.txtnDeposito.Text.ToString()));
			if ( idDPorLiquidar >  0 )  {MessageBox.Show("El número del deposito esta por Liquidar ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); this.txtnDeposito.Focus(); return; }

			int idAsiento = Funcion.iEscalarSQL(cdsSeteoF,string.Format("if Exists(Select a.idAsiento from Asiento a inner join AsientoDetalle ad on a.idAsiento = ad.idAsiento inner join CuentaCorriente cc on cc.idCuenta = ad.idCuenta where Cheque = REPLACE('{0}',' ','')  and a.idTipoAsiento = 2)Select a.idAsiento from Asiento a inner join AsientoDetalle ad on a.idAsiento = ad.idAsiento inner join CuentaCorriente cc on cc.idCuenta = ad.idCuenta where Cheque = REPLACE('{0}',' ','')  and a.idTipoAsiento = 2 Else Select CONVERT(Int, 0)",
				this.txtnDeposito.Text.ToString()));
			if ( idAsiento >  0 )  {MessageBox.Show("El número del deposito esta por Liquidar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); this.txtnDeposito.Focus(); return; }

	#endregion Validadación existencia
       
			#region Guarda Deposito
			string sSQL = string.Format("Exec GuardaDepositos {0},{1},{2},'{3}',{4},{5},{6}",

				(int)this.txtidDeposito.Value, (int)this.cmbCuenta.Value, (int) this.txtidCliente.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
				this.txtnDeposito.Text.ToString(), Convert.ToDecimal(this.txtValor.Value), this.txtEstado.Text.ToString());
			this.txtidDeposito.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			#endregion Guarda Deposito

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 4", idBloqueaTransacciones, (int)this.txtidDeposito.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);

			#region Controles
			
			this.txtnDeposito.Enabled=false;
			this.cmbCuenta.Enabled=false;
			this.txtRuc.Enabled=false;
			this.dtFecha.Enabled=false;
			this.txtnDeposito.Enabled=false;
			this.txtValor.Enabled=false;
			this.txtEstado.Enabled=false;
		

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

			
			bNuevo = false;
			bEdicion = false;
			#endregion Controles	
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cmbCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFecha.Focus();
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{			
			if (e.KeyCode == Keys.Enter) this.cmbCuenta.Focus();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtnDeposito.Focus();
		}

		private void txtnDeposito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtValor.Focus();
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!bEdicion) return;

			if (this.txtRuc.Text.ToString().Length == 0) return;

			#region Extrae informacion del cliente
			int idCliente = Funcion.iEscalarSQL(cdsSeteoF,string.Format("if Exists(select idCliente from Cliente where Ruc= '{0}' )select idCliente from Cliente where Ruc= '{0}'   Else Select CONVERT(Int, 0)",this.txtRuc.Text.ToString()));
 	      
 	      
			if (idCliente > 0)
			{
				this.lblNombre.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Nombre From cliente Where idCliente = {0} ", idCliente));
				this.txtidCliente.Value = idCliente;
				using (frmClientesProveedores cliente  = new frmClientesProveedores(false,1,idCliente))
				{
					
					if (DialogResult.OK == cliente.ShowDialog())
					{
						this.txtidCliente.Value = cliente.txtidCliente.Value;
						this.lblNombre.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Nombre From cliente Where idCliente = {0} ", (int)this.txtidCliente.Value));
						this.txtRuc.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Ruc From cliente Where idCliente = {0} ", (int)this.txtidCliente.Value));
						//							this.txtRuc.Enabled = false;
					}
				}
				//					this.txtRuc.Enabled = false;
			}
			else 
			{
				using (frmClientesProveedores cliente  = new frmClientesProveedores(false,1,idCliente))
				{
					if (DialogResult.OK == cliente.ShowDialog())
					{
						this.txtidCliente.Value = cliente.txtidCliente.Value;
						this.lblNombre.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Nombre From cliente Where idCliente = {0} ", (int)this.txtidCliente.Value));
						this.txtRuc.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Ruc From cliente Where idCliente = {0} ", (int)this.txtidCliente.Value));
						//							this.txtRuc.Enabled = false;
					}
				}
			
			}
			#endregion Extrae informacion del cliente 			
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	
		class validar 
		{
			public static void Solonumeros (KeyPressEventArgs Nu)
			{
				if(char.IsDigit(Nu.KeyChar))
				{
			
					Nu.Handled = false;
			
				}
				else if (char.IsControl(Nu.KeyChar))
				{
			
					Nu.Handled = false;			} 
				else
				{
					Nu.Handled= true;
				}
			}
	
		}
	}
}
