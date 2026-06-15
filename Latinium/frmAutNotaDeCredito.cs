using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmAutNotaDeCredito.
	/// </summary>
	public class frmAutNotaDeCredito : DevExpress.XtraEditors.XtraForm
	{
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		int IEstado = 0;
		private System.Windows.Forms.Button btnAutorizar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnRechazar;
		string SNota = "";
		public frmAutNotaDeCredito(int iEstado, string sNota)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
			IEstado = iEstado;
			SNota = sNota;
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
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnRechazar = new System.Windows.Forms.Button();
			this.btnAutorizar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			this.SuspendLayout();
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance1;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(8, 8);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(536, 168);
			this.txtNotas.TabIndex = 83;
			this.txtNotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotas_KeyPress);
			// 
			// btnRechazar
			// 
			this.btnRechazar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnRechazar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRechazar.Location = new System.Drawing.Point(88, 184);
			this.btnRechazar.Name = "btnRechazar";
			this.btnRechazar.TabIndex = 267;
			this.btnRechazar.Text = "&Rechazar";
			this.btnRechazar.Click += new System.EventHandler(this.btnRechazar_Click);
			// 
			// btnAutorizar
			// 
			this.btnAutorizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAutorizar.CausesValidation = false;
			this.btnAutorizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAutorizar.Location = new System.Drawing.Point(8, 184);
			this.btnAutorizar.Name = "btnAutorizar";
			this.btnAutorizar.TabIndex = 266;
			this.btnAutorizar.Text = "&Autorizar";
			this.btnAutorizar.Click += new System.EventHandler(this.btnAutorizar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(168, 184);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 265;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// frmAutNotaDeCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(552, 216);
			this.Controls.Add(this.btnRechazar);
			this.Controls.Add(this.btnAutorizar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtNotas);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmAutNotaDeCredito";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Autorización de Nota de Crédito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAutNotaDeCredito_KeyDown);
			this.Load += new System.EventHandler(this.frmAutNotaDeCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		public int Estado = 0;
		private void frmAutNotaDeCredito_Load(object sender, System.EventArgs e)
		{
			this.txtNotas.Text = SNota;
			this.txtNotas.Focus();

			if (IEstado == 4)
				this.btnRechazar.Enabled = false;
			else if (IEstado == 5)
				this.btnAutorizar.Enabled = false;
		}

		private void frmAutNotaDeCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void txtNotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnAutorizar.Focus();
		}

		private void btnAutorizar_Click(object sender, System.EventArgs e)
		{
			if (this.txtNotas.Text.ToString() == "")
			{
				MessageBox.Show("Ingrese un comentario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtNotas.Focus();
				return;
			}

			Estado = 5;

			this.DialogResult = DialogResult.OK;
		}

		private void btnRechazar_Click(object sender, System.EventArgs e)
		{
			if (this.txtNotas.Text.ToString() == "")
			{
				MessageBox.Show("Ingrese un comentario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtNotas.Focus();
				return;
			}

			Estado = 4;

			this.DialogResult = DialogResult.OK;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}
	}
}

