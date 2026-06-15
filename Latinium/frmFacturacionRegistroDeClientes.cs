using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Net;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFacturacionRegistroDeClientes.
	/// </summary>
	public class frmFacturacionRegistroDeClientes : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label56;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEMail;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdentificacion;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoRuc;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAgenteDeRetencion;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValAgente;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNResolucion;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public int idAgenteRetencionContribuyentesEspeciales = 2;

		public frmFacturacionRegistroDeClientes()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idTipoRuc = 0;
		string sIdentificacion = "";
		bool bCancelar = false;
		bool beditarNombre = false;

		public frmFacturacionRegistroDeClientes(int IdTipoRuc, string SIdentificacion, bool BCancelar)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idTipoRuc = IdTipoRuc;
			sIdentificacion = SIdentificacion;
			bCancelar = BCancelar;

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public frmFacturacionRegistroDeClientes(int IdTipoRuc, string SIdentificacion, bool BCancelar, bool BEditarNombre)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idTipoRuc = IdTipoRuc;
			sIdentificacion = SIdentificacion;
			bCancelar = BCancelar;
			beditarNombre = BEditarNombre;

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFacturacionRegistroDeClientes));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label56 = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEMail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdentificacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdTipoRuc = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbAgenteDeRetencion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtValAgente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNResolucion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAgenteDeRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValAgente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNResolucion)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 295);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(672, 10);
			this.groupBox1.TabIndex = 275;
			this.groupBox1.TabStop = false;
			// 
			// label56
			// 
			this.label56.AutoSize = true;
			this.label56.BackColor = System.Drawing.Color.Transparent;
			this.label56.Location = new System.Drawing.Point(10, 193);
			this.label56.Name = "label56";
			this.label56.Size = new System.Drawing.Size(164, 18);
			this.label56.TabIndex = 273;
			this.label56.Text = "AGENTE DE RETENCIÓN";
			this.label56.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(10, 9);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(113, 18);
			this.lblIdentificacion.TabIndex = 269;
			this.lblIdentificacion.Text = "IDENTIFICACIÓN";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(403, 119);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(67, 18);
			this.label3.TabIndex = 268;
			this.label3.Text = "CELULAR";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(10, 237);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 18);
			this.label1.TabIndex = 266;
			this.label1.Text = "NOTAS";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(10, 156);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(49, 18);
			this.label2.TabIndex = 265;
			this.label2.Text = "E-MAIL";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(10, 119);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(77, 18);
			this.label7.TabIndex = 264;
			this.label7.Text = "TELÉFONO";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(10, 82);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 18);
			this.label6.TabIndex = 263;
			this.label6.Text = "DIRECCIÓN";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(10, 45);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(151, 18);
			this.label4.TabIndex = 262;
			this.label4.Text = "APELLIDOS NOMBRES";
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(10, 313);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(99, 26);
			this.btnGuardar.TabIndex = 8;
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
			this.btnCancelar.Location = new System.Drawing.Point(115, 313);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(90, 26);
			this.btnCancelar.TabIndex = 9;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtCelular
			// 
			appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtCelular.Appearance = appearance1;
			this.txtCelular.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Location = new System.Drawing.Point(499, 117);
			this.txtCelular.MaxLength = 10;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(211, 24);
			this.txtCelular.TabIndex = 4;
			this.txtCelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCelular_KeyDown);
			this.txtCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCelular_KeyPress);
			// 
			// txtidCliente
			// 
			this.txtidCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidCliente.Appearance = appearance2;
			this.txtidCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.FormatString = "";
			this.txtidCliente.Location = new System.Drawing.Point(684, 314);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(29, 24);
			this.txtidCliente.TabIndex = 261;
			// 
			// txtNotas
			// 
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtNotas.Appearance = appearance3;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(182, 231);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(528, 55);
			this.txtNotas.TabIndex = 7;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// txtEMail
			// 
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtEMail.Appearance = appearance4;
			this.txtEMail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEMail.Location = new System.Drawing.Point(182, 153);
			this.txtEMail.Name = "txtEMail";
			this.txtEMail.Size = new System.Drawing.Size(288, 24);
			this.txtEMail.TabIndex = 5;
			this.txtEMail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEMail_KeyDown);
			// 
			// txtTelefono
			// 
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtTelefono.Appearance = appearance5;
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Location = new System.Drawing.Point(182, 117);
			this.txtTelefono.MaxLength = 9;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(164, 24);
			this.txtTelefono.TabIndex = 3;
			this.txtTelefono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTelefono_KeyDown);
			this.txtTelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTelefono_KeyPress);
			// 
			// txtDireccion
			// 
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtDireccion.Appearance = appearance6;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Location = new System.Drawing.Point(182, 80);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(528, 24);
			this.txtDireccion.TabIndex = 2;
			this.txtDireccion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDireccion_KeyDown);
			this.txtDireccion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDireccion_KeyPress);
			// 
			// txtIdentificacion
			// 
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtIdentificacion.Appearance = appearance7;
			this.txtIdentificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtIdentificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdentificacion.Enabled = false;
			this.txtIdentificacion.Location = new System.Drawing.Point(182, 7);
			this.txtIdentificacion.Name = "txtIdentificacion";
			this.txtIdentificacion.Size = new System.Drawing.Size(164, 24);
			this.txtIdentificacion.TabIndex = 0;
			// 
			// txtNombre
			// 
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtNombre.Appearance = appearance8;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(182, 43);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(528, 24);
			this.txtNombre.TabIndex = 1;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombre_KeyDown);
			// 
			// txtIdTipoRuc
			// 
			this.txtIdTipoRuc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdTipoRuc.Appearance = appearance9;
			this.txtIdTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoRuc.Enabled = false;
			this.txtIdTipoRuc.FormatString = "";
			this.txtIdTipoRuc.Location = new System.Drawing.Point(646, 314);
			this.txtIdTipoRuc.Name = "txtIdTipoRuc";
			this.txtIdTipoRuc.PromptChar = ' ';
			this.txtIdTipoRuc.Size = new System.Drawing.Size(28, 24);
			this.txtIdTipoRuc.TabIndex = 276;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbAgenteDeRetencion
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAgenteDeRetencion.Appearance = appearance10;
			this.cmbAgenteDeRetencion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbAgenteDeRetencion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAgenteDeRetencion.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 350;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbAgenteDeRetencion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbAgenteDeRetencion.DisplayMember = "Nombre";
			this.cmbAgenteDeRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAgenteDeRetencion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbAgenteDeRetencion.Enabled = false;
			this.cmbAgenteDeRetencion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAgenteDeRetencion.Location = new System.Drawing.Point(182, 190);
			this.cmbAgenteDeRetencion.Name = "cmbAgenteDeRetencion";
			this.cmbAgenteDeRetencion.Size = new System.Drawing.Size(394, 25);
			this.cmbAgenteDeRetencion.TabIndex = 277;
			this.cmbAgenteDeRetencion.ValueMember = "idAgenteRetencion";
			this.cmbAgenteDeRetencion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbAgenteDeRetencion_KeyDown);
			this.cmbAgenteDeRetencion.ValueChanged += new System.EventHandler(this.cmbAgenteDeRetencion_ValueChanged);
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
			// txtValAgente
			// 
			this.txtValAgente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValAgente.Appearance = appearance11;
			this.txtValAgente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValAgente.Enabled = false;
			this.txtValAgente.FormatString = "";
			this.txtValAgente.Location = new System.Drawing.Point(607, 314);
			this.txtValAgente.Name = "txtValAgente";
			this.txtValAgente.PromptChar = ' ';
			this.txtValAgente.Size = new System.Drawing.Size(29, 24);
			this.txtValAgente.TabIndex = 278;
			// 
			// txtNResolucion
			// 
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtNResolucion.Appearance = appearance12;
			this.txtNResolucion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNResolucion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNResolucion.Location = new System.Drawing.Point(595, 194);
			this.txtNResolucion.MaxLength = 10;
			this.txtNResolucion.Name = "txtNResolucion";
			this.txtNResolucion.Size = new System.Drawing.Size(115, 24);
			this.txtNResolucion.TabIndex = 279;
			this.txtNResolucion.ValueChanged += new System.EventHandler(this.ultraTextEditor1_ValueChanged);
			// 
			// frmFacturacionRegistroDeClientes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(722, 350);
			this.ControlBox = false;
			this.Controls.Add(this.txtNResolucion);
			this.Controls.Add(this.txtValAgente);
			this.Controls.Add(this.txtIdTipoRuc);
			this.Controls.Add(this.groupBox1);
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
			this.Controls.Add(this.txtCelular);
			this.Controls.Add(this.txtidCliente);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.txtEMail);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.txtDireccion);
			this.Controls.Add(this.txtIdentificacion);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.cmbAgenteDeRetencion);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmFacturacionRegistroDeClientes";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Registro De Clientes";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFacturacionRegistroDeClientes_KeyDown);
			this.Load += new System.EventHandler(this.frmFacturacionRegistroDeClientes_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAgenteDeRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValAgente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNResolucion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmFacturacionRegistroDeClientes_Load(object sender, System.EventArgs e)
		{
			
			if (!bCancelar) this.btnCancelar.Enabled = false;

			this.txtIdTipoRuc.Value = idTipoRuc;
			this.txtIdentificacion.Text = sIdentificacion;
			
			string sSQL = string.Format("Exec FacturacionConsultaCliente '{0}'", this.txtIdentificacion.Text.ToString());
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{				 
				this.txtidCliente.Value = dr.GetInt32(0);				
				this.txtIdentificacion.Text = dr.GetString(3);
				this.txtNombre.Text = dr.GetString(4);
				this.txtDireccion.Text = dr.GetString(5);
				this.txtTelefono.Text = dr.GetString(6);
				this.txtCelular.Text = dr.GetString(7);
				this.txtEMail.Text = dr.GetString(8);
				if (dr.GetValue(9) != System.DBNull.Value) this.cmbAgenteDeRetencion.Value = dr.GetInt32(9);
				this.txtNotas.Text = dr.GetString(10);
				this.txtValAgente.Value = dr.GetInt32(11);
				this.txtNResolucion.Value = dr.GetInt32(12);
			}
			dr.Close();

			if ((int)this.txtidCliente.Value > 0)
			{

				int iContFac = 0;

				iContFac = 	Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Exec [ValidaCambioNombreCliente] {0}", (int)this.txtidCliente.Value));

				if (iContFac == 0 && beditarNombre )
				{
					this.txtNombre.Enabled = true;
				}
				else
				{
					this.txtNombre.Enabled = false;
				}
				
			}
			else
			{
				this.txtNombre.Enabled = true;
			}
			// And idFormaPago In (9)
		
			//			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Compra Where idTipoFactura = 1  And Borrar = 0 And Estado In (10, 11) And idCliente = {0}", 
			//				(int)this.txtidCliente.Value)) > 0) this.txtNombre.Enabled = true;	

			if	 ((int)this.txtIdTipoRuc.Value == 4) this.cmbAgenteDeRetencion.Enabled = true;
			this.cmbAgenteDeRetencion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idAgenteRetencion, Nombre From dbo.AgenteRetencion");	
		
			SolicitudNResolucion();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmFacturacionRegistroDeClientes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) if (bCancelar) this.DialogResult = DialogResult.No;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{			
			if (!Validacion.vbTexto(this.txtNombre, 3, 120, "Nombre")) return;

			if (!Validacion.vbTexto(this.txtDireccion, 3, 200, "Dirección")) return;

			if (txtTelefono.Text.ToString().Trim() == "" && txtCelular.Text.ToString().Trim() == "")
			{
				MessageBox.Show("Ingrese el numero de teléfono convencional o celular del cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTelefono.Focus();			
			}

			if (txtCelular.Text.ToString().Trim() == "")
			{
				MessageBox.Show("Ingrese el numero de teléfono convencional o celular del cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCelular.Focus();	
				return;
			}
			
			if (!Validacion.vbValidaTelefono(this.txtTelefono, 1, false)) return;

			if (!Validacion.vbValidaTelefono(this.txtCelular, 2, false)) return;

			//if (!Validacion.vbTexto(this.txtEMail, 1, 200, "E-Mail - Campo obligatorio debe ser el e-mail correcto del cliente, para el envió de documento de facturación electrónica")) return;

			if (!Validacion.bEMail(this.txtEMail.Text.ToString().Trim())) return;

			if ((int)this.txtIdTipoRuc.Value == 4 && this.cmbAgenteDeRetencion.ActiveRow == null)
				if (!Validacion.vbComboVacio(this.cmbAgenteDeRetencion, "Seleccione Agente de Retención")) return;			
	
			if ((int)this.txtIdTipoRuc.Value == 4 && this.cmbAgenteDeRetencion.ActiveRow == null)
			{
				if((int)this.cmbAgenteDeRetencion.Value == idAgenteRetencionContribuyentesEspeciales)
				{
					if (!Validacion.vbTexto(this.txtNResolucion, 3, 10, "Número de resolución")) return;
				}
			}

			int idAgenteRetencion = 0;

			if ((int)this.txtValAgente.Value == 2)
			{
				int iValCE = 0;

				iValCE = 	Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaCECliente {0}", (int)this.txtidCliente.Value));
				
				if (iValCE == 0)
				{
					MessageBox.Show("El Cliente no tiene N° de Resolución", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.btnCancelar.Enabled = true;
					return;
				}
			}
				
			if (this.cmbAgenteDeRetencion.ActiveRow != null) idAgenteRetencion = (int)this.cmbAgenteDeRetencion.Value;

			string sSQL = string.Format("Exec FacturacionGuardaCliente {0}, {1}, {2}, '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', {9}, '{10}', '{11}'", 
				(int)this.txtidCliente.Value, 6,
				(int)this.txtIdTipoRuc.Value, this.txtIdentificacion.Text.ToString(), this.txtNombre.Text.ToString(),
				this.txtDireccion.Text.ToString(), this.txtTelefono.Text.ToString(), this.txtCelular.Text.ToString(), this.txtEMail.Text.ToString(),
				idAgenteRetencion, this.txtNotas.Text.ToString(), this.txtNResolucion.Text.ToString());
			this.txtidCliente.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			MessageBox.Show("Cliente registrado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			DialogResult = DialogResult.OK;
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDireccion.Focus();
		}

		private void txtDireccion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtTelefono.Focus();
		}

		private void txtTelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCelular.Focus();
		}

		private void txtCelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtEMail.Focus();
		}

		private void txtEMail_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if ((int)this.txtIdTipoRuc.Value == 4)	this.cmbAgenteDeRetencion.Focus(); else this.txtNotas.Focus();
			}
		}

		private void cmbAgenteDeRetencion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNotas.Focus();
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void txtTelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtDireccion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.SoloCaracter(e);
		}

		private void ultraTextEditor1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbAgenteDeRetencion_ValueChanged(object sender, System.EventArgs e)
		{
			SolicitudNResolucion();
		}

		public void SolicitudNResolucion()
		{
			#region Solicitud N Resolución Contribuyentes Especiales
			try
			{
				if((int)this.cmbAgenteDeRetencion.Value == idAgenteRetencionContribuyentesEspeciales)
				{
					this.txtNResolucion.Visible = true;
				}
				else
				{
					this.txtNResolucion.Visible = false;
					this.txtNResolucion.Text = "";
				}
			}
			catch(Exception ex)
			{
				this.txtNResolucion.Visible = false;
				this.txtNResolucion.Text = "";
			}
			#endregion Solicitud N Resolución Contribuyentes Especiales
		}

		class validar 
		{
			public static void SoloCaracter (KeyPressEventArgs Nu)
			{
				if(char.IsLetterOrDigit(Nu.KeyChar) || char.IsSeparator(Nu.KeyChar)  || char.IsControl(Nu.KeyChar)  )
				{		
					Nu.Handled = false;			
				}
				else if ( char.IsPunctuation (Nu.KeyChar)|| char.IsSurrogate (Nu.KeyChar) || char.IsSymbol (Nu.KeyChar)    )
				{			
					Nu.Handled = true;		
				} 
				else
				{
					Nu.Handled= true;
				}
			}
		}	

	}
}
