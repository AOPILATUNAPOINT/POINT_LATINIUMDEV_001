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
	/// Descripción breve de frmActualizadatosCliente.
	/// </summary>
	public class frmActualizadatosCliente : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btnGuardar;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdentificacion;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEMail;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelularNuevo;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEMailNuevo;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmActualizadatosCliente()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		string sIdentificacion = "";
		public frmActualizadatosCliente(string SIdentificacion)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

		//	idTipoRuc = IdTipoRuc;
			sIdentificacion = SIdentificacion;
		//	bCancelar = BCancelar;

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmActualizadatosCliente));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEMail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdentificacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.txtCelularNuevo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtEMailNuevo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelularNuevo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMailNuevo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(-1, 257);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(560, 8);
			this.groupBox1.TabIndex = 297;
			this.groupBox1.TabStop = false;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(7, 9);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(95, 16);
			this.lblIdentificacion.TabIndex = 295;
			this.lblIdentificacion.Text = "IDENTIFICACIÓN";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(335, 104);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(56, 16);
			this.label3.TabIndex = 294;
			this.label3.Text = "CELULAR";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(7, 136);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(42, 16);
			this.label2.TabIndex = 292;
			this.label2.Text = "E-MAIL";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(7, 104);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(65, 16);
			this.label7.TabIndex = 291;
			this.label7.Text = "TELÉFONO";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(7, 72);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(68, 16);
			this.label6.TabIndex = 290;
			this.label6.Text = "DIRECCIÓN";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(7, 40);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(127, 16);
			this.label4.TabIndex = 289;
			this.label4.Text = "APELLIDOS NOMBRES";
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(7, 272);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(83, 23);
			this.btnGuardar.TabIndex = 286;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// txtCelular
			// 
			appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtCelular.Appearance = appearance1;
			this.txtCelular.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Enabled = false;
			this.txtCelular.Location = new System.Drawing.Point(415, 102);
			this.txtCelular.MaxLength = 10;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(136, 21);
			this.txtCelular.TabIndex = 283;
			// 
			// txtEMail
			// 
			appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtEMail.Appearance = appearance2;
			this.txtEMail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEMail.Enabled = false;
			this.txtEMail.Location = new System.Drawing.Point(151, 134);
			this.txtEMail.Name = "txtEMail";
			this.txtEMail.Size = new System.Drawing.Size(240, 21);
			this.txtEMail.TabIndex = 284;
			// 
			// txtTelefono
			// 
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtTelefono.Appearance = appearance3;
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Location = new System.Drawing.Point(151, 102);
			this.txtTelefono.MaxLength = 9;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(136, 21);
			this.txtTelefono.TabIndex = 282;
			// 
			// txtDireccion
			// 
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtDireccion.Appearance = appearance4;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Enabled = false;
			this.txtDireccion.Location = new System.Drawing.Point(151, 70);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(400, 21);
			this.txtDireccion.TabIndex = 281;
			// 
			// txtIdentificacion
			// 
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtIdentificacion.Appearance = appearance5;
			this.txtIdentificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtIdentificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdentificacion.Enabled = false;
			this.txtIdentificacion.Location = new System.Drawing.Point(151, 7);
			this.txtIdentificacion.Name = "txtIdentificacion";
			this.txtIdentificacion.Size = new System.Drawing.Size(136, 21);
			this.txtIdentificacion.TabIndex = 279;
			// 
			// txtNombre
			// 
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtNombre.Appearance = appearance6;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(151, 38);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(399, 21);
			this.txtNombre.TabIndex = 280;
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
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(7, 176);
			this.label1.Name = "label1";
			this.label1.TabIndex = 299;
			this.label1.Text = "CELULAR NUEVO";
			// 
			// txtCelularNuevo
			// 
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtCelularNuevo.Appearance = appearance7;
			this.txtCelularNuevo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelularNuevo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelularNuevo.Location = new System.Drawing.Point(151, 176);
			this.txtCelularNuevo.MaxLength = 10;
			this.txtCelularNuevo.Name = "txtCelularNuevo";
			this.txtCelularNuevo.Size = new System.Drawing.Size(136, 21);
			this.txtCelularNuevo.TabIndex = 298;
			this.txtCelularNuevo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCelularNuevo_KeyPress);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(7, 216);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(85, 16);
			this.label5.TabIndex = 301;
			this.label5.Text = "E-MAIL NUEVO";
			this.label5.Click += new System.EventHandler(this.label5_Click);
			// 
			// txtEMailNuevo
			// 
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtEMailNuevo.Appearance = appearance8;
			this.txtEMailNuevo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEMailNuevo.Location = new System.Drawing.Point(151, 208);
			this.txtEMailNuevo.Name = "txtEMailNuevo";
			this.txtEMailNuevo.Size = new System.Drawing.Size(240, 21);
			this.txtEMailNuevo.TabIndex = 300;
			// 
			// txtidCliente
			// 
			this.txtidCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidCliente.Appearance = appearance9;
			this.txtidCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.FormatString = "";
			this.txtidCliente.Location = new System.Drawing.Point(520, 272);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(24, 21);
			this.txtidCliente.TabIndex = 302;
			// 
			// frmActualizadatosCliente
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(558, 302);
			this.ControlBox = false;
			this.Controls.Add(this.txtidCliente);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtEMailNuevo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtCelularNuevo);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.txtCelular);
			this.Controls.Add(this.txtEMail);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.txtDireccion);
			this.Controls.Add(this.txtIdentificacion);
			this.Controls.Add(this.txtNombre);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "frmActualizadatosCliente";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Actualiza Datos del Cliente";
			this.Load += new System.EventHandler(this.frmActualizadatosCliente_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelularNuevo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMailNuevo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmActualizadatosCliente_Load(object sender, System.EventArgs e)
		{
			this.txtIdentificacion.Text = sIdentificacion;
			
			string sSQL = string.Format("Exec CobranzaConsultaCliente '{0}'", this.txtIdentificacion.Text.ToString());
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{		
		 		this.txtidCliente.Value = dr.GetInt32(0);		
				this.txtIdentificacion.Text = dr.GetString(1);
				this.txtNombre.Text = dr.GetString(2);
				this.txtDireccion.Text = dr.GetString(3);
				this.txtTelefono.Text = dr.GetString(4);
				this.txtCelular.Text = dr.GetString(5);
				this.txtEMail.Text = dr.GetString(6);
				this.txtCelularNuevo.Text = dr.GetString(7);
				this.txtEMailNuevo.Text = dr.GetString(8);
			}
			dr.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void label5_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtCelularNuevo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			    if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
    {
        e.Handled = true; // Bloquea la tecla si no es número
    }
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec CobranzaGuardaCliente {0}, '{1}', '{2}'", 
				(int)this.txtidCliente.Value,this.txtCelularNuevo.Text.ToString(), this.txtEMailNuevo.Text.ToString());
			this.txtidCliente.Value = Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			MessageBox.Show("Datos del Cliente Actualizados Correctamentee", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			DialogResult = DialogResult.OK;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
