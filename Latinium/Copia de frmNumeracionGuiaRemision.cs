using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmNumeracionGuiaRemision.
	/// </summary>
	public class frmNumeracionGuiaRemision : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label lblMensaje;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmNumeracionGuiaRemision()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		
		string sNumero = "";

		public frmNumeracionGuiaRemision(string SNumero)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			sNumero = SNumero;

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmNumeracionGuiaRemision));
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.lblNumero = new System.Windows.Forms.Label();
			this.lblMensaje = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			appearance1.Image = ((object)(resources.GetObject("appearance1.Image")));
			this.btAceptar.Appearance = appearance1;
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.btAceptar.Location = new System.Drawing.Point(8, 184);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(128, 32);
			this.btAceptar.TabIndex = 152;
			this.btAceptar.Text = "&Imprimir";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// lblNumero
			// 
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblNumero.Font = new System.Drawing.Font("Tahoma", 45F, System.Drawing.FontStyle.Bold);
			this.lblNumero.ForeColor = System.Drawing.Color.Red;
			this.lblNumero.Location = new System.Drawing.Point(8, 72);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(608, 96);
			this.lblNumero.TabIndex = 154;
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblMensaje
			// 
			this.lblMensaje.BackColor = System.Drawing.Color.Transparent;
			this.lblMensaje.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblMensaje.Font = new System.Drawing.Font("Tahoma", 30F, System.Drawing.FontStyle.Bold);
			this.lblMensaje.Location = new System.Drawing.Point(8, 8);
			this.lblMensaje.Name = "lblMensaje";
			this.lblMensaje.Size = new System.Drawing.Size(608, 52);
			this.lblMensaje.TabIndex = 153;
			this.lblMensaje.Text = "PREPARE LA GUÍA NÚMERO";
			this.lblMensaje.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmNumeracionGuiaRemision
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(624, 222);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.lblMensaje);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmNumeracionGuiaRemision";
			this.ShowInTaskbar = false;
			this.Load += new System.EventHandler(this.frmNumeracionGuiaRemision_Load);
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void frmNumeracionGuiaRemision_Load(object sender, System.EventArgs e)
		{
			this.lblNumero.Text = sNumero;
		}
	}
}

