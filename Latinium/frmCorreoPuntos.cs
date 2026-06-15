using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCorreoPuntos.
	/// </summary>
	public class frmCorreoPuntos : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label lblCliente;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCorreo;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPuntos;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCorreoPuntos()
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
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtCorreo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtPuntos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCorreo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPuntos)).BeginInit();
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
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 39);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(408, 8);
			this.groupBox1.TabIndex = 173;
			this.groupBox1.TabStop = false;
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(8, 55);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 1;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(88, 55);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 2;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(8, 10);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(37, 17);
			this.lblCliente.TabIndex = 172;
			this.lblCliente.Text = "Correo";
			// 
			// txtCorreo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCorreo.Appearance = appearance1;
			this.txtCorreo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCorreo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCorreo.Location = new System.Drawing.Point(64, 7);
			this.txtCorreo.MaxLength = 15;
			this.txtCorreo.Name = "txtCorreo";
			this.txtCorreo.Size = new System.Drawing.Size(200, 22);
			this.txtCorreo.TabIndex = 0;
			this.txtCorreo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCorreo_KeyDown);
			this.txtCorreo.Validated += new System.EventHandler(this.txtCorreo_Validated);
			// 
			// txtPuntos
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPuntos.Appearance = appearance2;
			this.txtPuntos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPuntos.Enabled = false;
			this.txtPuntos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPuntos.FormatString = "#0";
			this.txtPuntos.Location = new System.Drawing.Point(336, 7);
			this.txtPuntos.MaxValue = 100000;
			this.txtPuntos.MinValue = 0;
			this.txtPuntos.Name = "txtPuntos";
			this.txtPuntos.PromptChar = ' ';
			this.txtPuntos.Size = new System.Drawing.Size(60, 22);
			this.txtPuntos.TabIndex = 448;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(280, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(38, 17);
			this.label1.TabIndex = 449;
			this.label1.Text = "Puntos";
			// 
			// frmCorreoPuntos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(408, 84);
			this.ControlBox = false;
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtPuntos);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.txtCorreo);
			this.Controls.Add(this.groupBox1);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmCorreoPuntos";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Ingrese el correo del cliente";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCorreoPuntos_KeyDown);
			this.Load += new System.EventHandler(this.frmCorreoPuntos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCorreo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPuntos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmCorreoPuntos_Load(object sender, System.EventArgs e)
		{
		
		}

		private void txtCorreo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.txtCorreo.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el E-Mail del cliente", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return ;
				}

				if (!Validacion.bEMail(this.txtCorreo.Text)) return;

				this.txtPuntos.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT ISNULL(SUM(Puntos), 0) FROM PuntosPromocionales Where Correo = '{0}' And Estado = 3", 
					this.txtCorreo.Text.ToString()));
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmCorreoPuntos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtPuntos.Value == 0)
			{
				MessageBox.Show("El cliente no tiene puntos a su favor", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.DialogResult = DialogResult.OK;
		}

		private void txtCorreo_Validated(object sender, System.EventArgs e)
		{
		
		}
	}
}
