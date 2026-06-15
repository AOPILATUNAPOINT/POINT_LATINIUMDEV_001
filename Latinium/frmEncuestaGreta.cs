using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmEncuestaGreta.
	/// </summary>
	public class frmEncuestaGreta : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label txtNumEncuesta;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtEncuesta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmEncuestaGreta()
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
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtNumEncuesta = new System.Windows.Forms.Label();
			this.txtEncuesta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtEncuesta)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(352, 8);
			this.groupBox1.TabIndex = 173;
			this.groupBox1.TabStop = false;
			// 
			// btnAceptar
			// 
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(16, 48);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 170;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(96, 48);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 171;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtNumEncuesta
			// 
			this.txtNumEncuesta.AutoSize = true;
			this.txtNumEncuesta.Location = new System.Drawing.Point(8, 8);
			this.txtNumEncuesta.Name = "txtNumEncuesta";
			this.txtNumEncuesta.Size = new System.Drawing.Size(67, 16);
			this.txtNumEncuesta.TabIndex = 172;
			this.txtNumEncuesta.Text = "N° Encuesta";
			// 
			// txtEncuesta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEncuesta.Appearance = appearance1;
			this.txtEncuesta.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtEncuesta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEncuesta.Location = new System.Drawing.Point(88, 8);
			this.txtEncuesta.MaxLength = 15;
			this.txtEncuesta.Name = "txtEncuesta";
			this.txtEncuesta.Size = new System.Drawing.Size(256, 21);
			this.txtEncuesta.TabIndex = 169;
			this.txtEncuesta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtEncuesta_KeyPress);
			// 
			// frmEncuestaGreta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(352, 78);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtNumEncuesta);
			this.Controls.Add(this.txtEncuesta);
			this.Name = "frmEncuestaGreta";
			this.Text = "Ingreso N° Encuesta";
			((System.ComponentModel.ISupportInitialize)(this.txtEncuesta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void txtEncuesta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
