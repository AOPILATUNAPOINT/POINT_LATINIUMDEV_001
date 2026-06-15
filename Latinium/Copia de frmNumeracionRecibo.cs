using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmNumeracionRecibo.
	/// </summary>
	public class frmNumeracionRecibo : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label lblMensaje;
		private System.Windows.Forms.Button btnImprimir;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		string SNumero = "";
		public frmNumeracionRecibo(string sNumero)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			SNumero = sNumero;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmNumeracionRecibo));
			this.lblNumero = new System.Windows.Forms.Label();
			this.lblMensaje = new System.Windows.Forms.Label();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Font = new System.Drawing.Font("Tahoma", 45F, System.Drawing.FontStyle.Bold);
			this.lblNumero.ForeColor = System.Drawing.Color.Red;
			this.lblNumero.Location = new System.Drawing.Point(8, 88);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(0, 76);
			this.lblNumero.TabIndex = 154;
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblMensaje
			// 
			this.lblMensaje.AutoSize = true;
			this.lblMensaje.BackColor = System.Drawing.Color.Transparent;
			this.lblMensaje.Font = new System.Drawing.Font("Tahoma", 30F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.lblMensaje.Location = new System.Drawing.Point(8, 8);
			this.lblMensaje.Name = "lblMensaje";
			this.lblMensaje.Size = new System.Drawing.Size(731, 52);
			this.lblMensaje.TabIndex = 153;
			this.lblMensaje.Text = "PREPARE EL RECIBO DE COBRO N:";
			this.lblMensaje.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(624, 192);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(120, 40);
			this.btnImprimir.TabIndex = 251;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// frmNumeracionRecibo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(752, 238);
			this.ControlBox = false;
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.lblMensaje);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmNumeracionRecibo";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "POINT TECHNOLOGY";
			this.Load += new System.EventHandler(this.frmNumeracionRecibo_Load);
			this.ResumeLayout(false);

		}
		#endregion

		private void frmNumeracionRecibo_Load(object sender, System.EventArgs e)
		{
			this.lblNumero.Text = SNumero;
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}
	}
}

