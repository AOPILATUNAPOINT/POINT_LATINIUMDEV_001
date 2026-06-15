using System;
using System.Drawing;
using System.Data.SqlClient;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for VentaCliente.
	/// </summary>
	public class VentaCliente : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdentificacion;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEMail;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNota;
		private C1.Data.C1DataSet cdsCliente;
		private BarraDatoSQL.BarraDatoSQL barraDato1;

		string sParam = "";
		int IdCliente = 0;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private System.Windows.Forms.Label lblIdentificacion;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAgenteDeRetencion;
		private System.Windows.Forms.Label label56;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.GroupBox groupBox1;
		int IdTipoRuc = 0;
		public VentaCliente(string sParametro, int idCliente, int idTipoRuc)
		{
			InitializeComponent();
			sParam = sParametro;
			IdCliente = idCliente;
			IdTipoRuc = idTipoRuc;
		}

		public VentaCliente()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(VentaCliente));
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdentificacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNota = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtEMail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.label3 = new System.Windows.Forms.Label();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cmbAgenteDeRetencion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label56 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAgenteDeRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// txtTelefono
			// 
			appearance1.BackColor = System.Drawing.SystemColors.Window;
			appearance1.BackColor2 = System.Drawing.SystemColors.Window;
			appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtTelefono.Appearance = appearance1;
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Telefono"));
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Location = new System.Drawing.Point(145, 104);
			this.txtTelefono.MaxLength = 9;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(136, 22);
			this.txtTelefono.TabIndex = 6;
			this.txtTelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTelefono_KeyPress);
			this.txtTelefono.Validating += new System.ComponentModel.CancelEventHandler(this.txtTelefono_Validating);
			this.txtTelefono.ValueChanged += new System.EventHandler(this.txtTelefono_ValueChanged);
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.FillOnRequest = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(9, 107);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(59, 16);
			this.label7.TabIndex = 60;
			this.label7.Text = "TELÉFONO";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(9, 75);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(63, 16);
			this.label6.TabIndex = 59;
			this.label6.Text = "DIRECCIÓN";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(9, 43);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(115, 16);
			this.label4.TabIndex = 57;
			this.label4.Text = "APELLIDOS NOMBRES";
			// 
			// txtDireccion
			// 
			appearance2.BackColor = System.Drawing.SystemColors.Window;
			appearance2.BackColor2 = System.Drawing.SystemColors.Window;
			appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtDireccion.Appearance = appearance2;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Direccion"));
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Location = new System.Drawing.Point(145, 72);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(400, 22);
			this.txtDireccion.TabIndex = 5;
			this.txtDireccion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDireccion_KeyPress);
			this.txtDireccion.Validating += new System.ComponentModel.CancelEventHandler(this.txtDireccion_Validating);
			this.txtDireccion.ValueChanged += new System.EventHandler(this.txtDireccion_ValueChanged);
			// 
			// txtIdentificacion
			// 
			appearance3.BackColor = System.Drawing.SystemColors.Window;
			appearance3.BackColor2 = System.Drawing.SystemColors.Window;
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtIdentificacion.Appearance = appearance3;
			this.txtIdentificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtIdentificacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Ruc"));
			this.txtIdentificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdentificacion.Location = new System.Drawing.Point(300, 8);
			this.txtIdentificacion.Name = "txtIdentificacion";
			this.txtIdentificacion.Size = new System.Drawing.Size(152, 22);
			this.txtIdentificacion.TabIndex = 1;
			this.txtIdentificacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtIdentificacion_KeyPress);
			this.txtIdentificacion.Validating += new System.ComponentModel.CancelEventHandler(this.txtIdentificacion_Validating);
			this.txtIdentificacion.Validated += new System.EventHandler(this.txtIdentificacion_Validated);
			this.txtIdentificacion.ValueChanged += new System.EventHandler(this.txtIdentificacion_ValueChanged);
			// 
			// txtNombre
			// 
			appearance4.BackColor = System.Drawing.SystemColors.Window;
			appearance4.BackColor2 = System.Drawing.SystemColors.Window;
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtNombre.Appearance = appearance4;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(145, 40);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(399, 22);
			this.txtNombre.TabIndex = 4;
			this.txtNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombre_KeyPress);
			this.txtNombre.Validating += new System.ComponentModel.CancelEventHandler(this.txtNombre_Validating);
			this.txtNombre.ValueChanged += new System.EventHandler(this.txtNombre_ValueChanged);
			// 
			// txtNota
			// 
			appearance5.BackColor = System.Drawing.SystemColors.Window;
			appearance5.BackColor2 = System.Drawing.SystemColors.Window;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtNota.Appearance = appearance5;
			this.txtNota.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNota.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Notas"));
			this.txtNota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNota.Location = new System.Drawing.Point(145, 200);
			this.txtNota.Multiline = true;
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(400, 88);
			this.txtNota.TabIndex = 8;
			this.txtNota.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNota_KeyPress);
			this.txtNota.ValueChanged += new System.EventHandler(this.txtNota_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(9, 208);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 66;
			this.label1.Text = "NOTAS";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(9, 139);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(40, 16);
			this.label2.TabIndex = 65;
			this.label2.Text = "E-MAIL";
			// 
			// txtEMail
			// 
			appearance6.BackColor = System.Drawing.SystemColors.Window;
			appearance6.BackColor2 = System.Drawing.SystemColors.Window;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtEMail.Appearance = appearance6;
			this.txtEMail.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Email"));
			this.txtEMail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEMail.Location = new System.Drawing.Point(145, 136);
			this.txtEMail.Name = "txtEMail";
			this.txtEMail.Size = new System.Drawing.Size(240, 22);
			this.txtEMail.TabIndex = 7;
			this.txtEMail.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtEMail_KeyPress);
			this.txtEMail.ValueChanged += new System.EventHandler(this.txtEMail_ValueChanged);
			// 
			// txtidCliente
			// 
			this.txtidCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidCliente.Appearance = appearance7;
			this.txtidCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCliente"));
			this.txtidCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.FormatString = "";
			this.txtidCliente.Location = new System.Drawing.Point(528, 320);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(24, 22);
			this.txtidCliente.TabIndex = 9;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Cliente";
			this.barraDato1.DataNombreId = "idCliente";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsCliente;
			this.barraDato1.DataTabla = "Cliente";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(384, 440);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 78;
			this.barraDato1.Visible = false;
			this.barraDato1.VisibleBorrar = false;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = false;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(329, 107);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(50, 16);
			this.label3.TabIndex = 80;
			this.label3.Text = "CELULAR";
			// 
			// txtCelular
			// 
			appearance8.BackColor = System.Drawing.SystemColors.Window;
			appearance8.BackColor2 = System.Drawing.SystemColors.Window;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtCelular.Appearance = appearance8;
			this.txtCelular.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelular.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Celular"));
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Location = new System.Drawing.Point(409, 104);
			this.txtCelular.MaxLength = 10;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(136, 22);
			this.txtCelular.TabIndex = 79;
			this.txtCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCelular_KeyPress);
			this.txtCelular.Validating += new System.ComponentModel.CancelEventHandler(this.txtCelular_Validating);
			this.txtCelular.ValueChanged += new System.EventHandler(this.txtCelular_ValueChanged);
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(9, 11);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(91, 16);
			this.lblIdentificacion.TabIndex = 163;
			this.lblIdentificacion.Text = "IDENTIFICACIÓN";
			// 
			// cmbTipoRuc
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance9;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRuc.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idTipoRuc"));
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 233;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(145, 8);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 21);
			this.cmbTipoRuc.TabIndex = 164;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.Leave += new System.EventHandler(this.cmbTipoRuc_Leave);
			this.cmbTipoRuc.Validated += new System.EventHandler(this.cmbTipoRuc_Validated);
			this.cmbTipoRuc.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoRuc_InitializeLayout);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 312);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(83, 23);
			this.btnGuardar.TabIndex = 251;
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
			this.btnCancelar.Location = new System.Drawing.Point(96, 312);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 250;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// cmbAgenteDeRetencion
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAgenteDeRetencion.Appearance = appearance10;
			this.cmbAgenteDeRetencion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbAgenteDeRetencion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAgenteDeRetencion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idAgenteRetencion"));
			this.cmbAgenteDeRetencion.DataSource = this.ultraDataSource2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 350;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbAgenteDeRetencion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbAgenteDeRetencion.DisplayMember = "Nombre";
			this.cmbAgenteDeRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAgenteDeRetencion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbAgenteDeRetencion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAgenteDeRetencion.Location = new System.Drawing.Point(144, 168);
			this.cmbAgenteDeRetencion.Name = "cmbAgenteDeRetencion";
			this.cmbAgenteDeRetencion.Size = new System.Drawing.Size(328, 21);
			this.cmbAgenteDeRetencion.TabIndex = 253;
			this.cmbAgenteDeRetencion.ValueMember = "idAgenteRetencion";
			this.cmbAgenteDeRetencion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbAgenteDeRetencion_KeyDown);
			this.cmbAgenteDeRetencion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbAgenteDeRetencion_KeyPress);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label56
			// 
			this.label56.AutoSize = true;
			this.label56.BackColor = System.Drawing.Color.Transparent;
			this.label56.Location = new System.Drawing.Point(8, 168);
			this.label56.Name = "label56";
			this.label56.Size = new System.Drawing.Size(127, 16);
			this.label56.TabIndex = 252;
			this.label56.Text = "AGENTE DE RETENCIÓN";
			this.label56.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 296);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(560, 8);
			this.groupBox1.TabIndex = 254;
			this.groupBox1.TabStop = false;
			// 
			// VentaCliente
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(554, 344);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbAgenteDeRetencion);
			this.Controls.Add(this.label56);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.txtCelular);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtidCliente);
			this.Controls.Add(this.txtNota);
			this.Controls.Add(this.txtEMail);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.txtDireccion);
			this.Controls.Add(this.txtIdentificacion);
			this.Controls.Add(this.txtNombre);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "VentaCliente";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "+";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.VentaCliente_KeyDown);
			this.Load += new System.EventHandler(this.VentaCliente_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAgenteDeRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void VentaCliente_Load(object sender, System.EventArgs e)
		{
			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsCliente, "Exec ClienteTipoRucCarga 1");
			this.cmbAgenteDeRetencion.DataSource = Funcion.dvProcedimiento(cdsCliente, "Select idAgenteRetencion, Nombre From AgenteRetencion");
						
			if (this.IdCliente > 0)
			{
				ClienteBuscado(IdCliente);
				barraDato1.EditarRegistro();				
			}
			else
			{
				barraDato1.PosUltima();

				barraDato1.CrearRegistro();

				if (IdTipoRuc > 3) drCliente["idTipoRuc"] = IdTipoRuc;

				else drCliente["idTipoRuc"] = 5;

				if (sParam != "")
				{
					if (Char.IsNumber(sParam, 1))
					{
						this.txtIdentificacion.Text = sParam;
						drCliente["Ruc"] = sParam;
					}
					else
					{
						drCliente["Ruc"] = sParam;
						this.txtNombre.Text = sParam;
						drCliente["Nombre"] = sParam;
					}
				}			
			}

			this.cmbTipoRuc.Enabled = false;
			this.txtIdentificacion.Enabled = false;

			this.txtNombre.Focus();
		}

		private void ClienteBuscado(int id)
		{
			this.barraDato1.IdRegistro = id;
			this.barraDato1.ProximoId(5);
		}

		private void txtIdentificacion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (txtIdentificacion.Text.ToString().Trim() != "")
			{
//				if (barraDato1.bEditar)
//				{
					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtIdentificacion.Text.ToString(), (int)drCliente["idTipoRuc"]);
					string ValCed = Funcion.sEscalarSQL(cdsCliente, stExec);
				
					if (ValCed.StartsWith("Error"))
					{
						MessageBox.Show("Número de identificación incorrecta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtIdentificacion.Focus();
						e.Cancel = true;						
					}

//					string sSelect = string.Format("Exec VentaClienteInsertaV '{0}', '{1}'", this.txtNombre.Text.Trim(), this.txtIdentificacion.Text.Trim());
//					string sMensaje = Funcion.sEscalarSQL(cdsCliente, sSelect);
//				
//					if (sMensaje.Length > 0)
//					{
//						MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						e.Cancel = true;
//						return;
//					}
//				}
			}
		}

		private void txtNombre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (txtNombre.Text.ToString().Trim() == "")
			{
				MessageBox.Show("Ingrese el nombre del cliente");
				this.txtNombre.Focus();
				e.Cancel = true;
			}	
		}

		private void txtDireccion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (txtDireccion.Text.ToString().Trim() == "")
			{
				MessageBox.Show("Ingrese la dirección del cliente");
				this.txtDireccion.Focus();
				e.Cancel = true;
			}
		}

		private void txtTelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsControl(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsSeparator(e.KeyChar))
			{
				e.Handled = false;
			}
			else
			{
				e.Handled = true;
			}

			if (e.KeyChar == 13)
			{
				this.txtCelular.Focus();
			}
		}
		private C1.Data.C1DataRow drCliente;

		private void txtTelefono_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{			
			
		}

		private void VentaCliente_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape)
			{
				Close();
			}
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drCliente = cdsCliente.TableViews["Cliente"].Rows[0];
			}
			catch{};
		}

		private void txtDireccion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtTelefono_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtIdentificacion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtIdentificacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.txtNombre.Focus();
			}
		}

		private void txtNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.txtDireccion.Focus();
			}
		}

		private void txtDireccion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.txtTelefono.Focus();
			}
		}

		private void txtEMail_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.cmbAgenteDeRetencion.Focus();
			}
		}

		private void txtNota_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNota_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.btnGuardar.Focus();
			}
		}

		private void txtCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsControl(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsSeparator(e.KeyChar))
			{
				e.Handled = false;
			}
			else
			{
				e.Handled = true;
			}

			if (e.KeyChar == 13)
			{
				this.txtEMail.Focus();
			}
		}

		private void txtCelular_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCelular_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (txtTelefono.Text.ToString().Trim() == "" && txtCelular.Text.ToString().Trim() == "")
			{
				MessageBox.Show("Ingrese el teléfono del cliente.");
				this.txtTelefono.Focus();
				e.Cancel = true;
			}
		}

		private void cmbTipoRuc_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipoRuc_Validated(object sender, System.EventArgs e)
		{
				
		}

		private void txtEMail_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (txtIdentificacion.Text.ToString().Trim() == "")
			{
				MessageBox.Show("Ingrese el número de identificación del cliente");
				this.txtIdentificacion.Focus();			
			}
			else if (txtNombre.Text.ToString().Trim() == "")
			{
				MessageBox.Show("Ingrese el nombre del cliente");
				this.txtNombre.Focus();				
			}
			else if (txtDireccion.Text.ToString().Trim() == "")
			{
				MessageBox.Show("Ingrese la dirección del cliente");
				this.txtDireccion.Focus();				
			}
			else if (txtTelefono.Text.ToString().Trim() == "" && txtCelular.Text.ToString().Trim() == "")
			{
				MessageBox.Show("Ingrese el teléfono Convencional o Celular del ciente.");
				this.txtTelefono.Focus();			
			}
			else if ((int)this.cmbTipoRuc.Value == 4 && this.cmbAgenteDeRetencion.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Agente de Retención");
				this.cmbAgenteDeRetencion.Focus();			
			}
			else
			{
				drCliente["idGrupoCliente"] = 6;
				drCliente["ContadoCred"] = 1;
				drCliente["Proveedor"] = 0;
				drCliente["InstPublica"] = 0;		
			
				barraDato1.GrabaRegistro();
			
				DialogResult = DialogResult.OK;
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void cmbTipoRuc_Leave(object sender, System.EventArgs e)
		{
			drCliente["idTipoRuc"] = (int)this.cmbTipoRuc.Value;	
			
			if ((int)drCliente["idTipoRuc"] == 4 || (int)drCliente["idTipoRuc"] == 6) this.txtIdentificacion.MaxLength = 13;
			if ((int)drCliente["idTipoRuc"] == 5) this.txtIdentificacion.MaxLength = 10;
		}

		private void txtIdentificacion_Validated(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec BuscaIdCliente '{0}'", this.txtIdentificacion.Text.ToString());
			int idCliente = Funcion.iEscalarSQL(cdsCliente, sSQL);

			if (idCliente > 0) 
			{
				ClienteBuscado(IdCliente);
				barraDato1.EditarRegistro();
			}
		}

		private void cmbAgenteDeRetencion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		
		}

		private void cmbAgenteDeRetencion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.txtNota.Focus();
			}
		}
	}
}

