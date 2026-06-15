using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmListaCorreosCotizaciones.
	/// </summary>
	public class frmListaCorreosCotizaciones : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label lblCliente;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmListaCorreosCotizaciones()
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
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			this.SuspendLayout();
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 40);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 1;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 40);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 2;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(8, 8);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(105, 16);
			this.lblCliente.TabIndex = 416;
			this.lblCliente.Text = "Dirección de Correo";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDireccion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccion.Appearance = appearance1;
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Location = new System.Drawing.Point(120, 8);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(208, 21);
			this.txtDireccion.TabIndex = 0;
			this.txtDireccion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDireccion_KeyDown);
			// 
			// frmListaCorreosCotizaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(338, 70);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.txtDireccion);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmListaCorreosCotizaciones";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmListaCorreosCotizaciones_KeyDown);
			this.Load += new System.EventHandler(this.frmListaCorreosCotizaciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmListaCorreosCotizaciones_Load(object sender, System.EventArgs e)
		{
		
		}

		private void frmListaCorreosCotizaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbTexto(this.txtDireccion, 0, 200, "Dirección de Correo")) return;
			if (!Validacion.bEMail(this.txtDireccion.Text.ToString().Trim())) return;

			this.DialogResult = DialogResult.OK;
		}

		private void txtDireccion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAceptar.Focus();
		}
	}
}
