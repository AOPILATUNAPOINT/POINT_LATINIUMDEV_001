using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmAnulaRec.
	/// </summary>
	public class frmAnulaRec : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label lblMensaje;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		string SNumero = "";

		public frmAnulaRec(string sNumero)
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			this.lblNumero = new System.Windows.Forms.Label();
			this.lblMensaje = new System.Windows.Forms.Label();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.SuspendLayout();
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Font = new System.Drawing.Font("Tahoma", 45F, System.Drawing.FontStyle.Bold);
			this.lblNumero.ForeColor = System.Drawing.Color.Red;
			this.lblNumero.Location = new System.Drawing.Point(16, 88);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(0, 76);
			this.lblNumero.TabIndex = 153;
			// 
			// lblMensaje
			// 
			this.lblMensaje.AutoSize = true;
			this.lblMensaje.BackColor = System.Drawing.Color.Transparent;
			this.lblMensaje.Font = new System.Drawing.Font("Tahoma", 30F, System.Drawing.FontStyle.Bold);
			this.lblMensaje.ForeColor = System.Drawing.Color.Black;
			this.lblMensaje.Location = new System.Drawing.Point(14, 8);
			this.lblMensaje.Name = "lblMensaje";
			this.lblMensaje.Size = new System.Drawing.Size(708, 52);
			this.lblMensaje.TabIndex = 152;
			this.lblMensaje.Text = "Esta seguro de anular el recibo N.";
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.Image = 14;
			this.btAceptar.Appearance = appearance1;
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btAceptar.Location = new System.Drawing.Point(8, 192);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 24);
			this.btAceptar.TabIndex = 154;
			this.btAceptar.Text = "&Si";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance2.Image = 14;
			this.ultraButton1.Appearance = appearance2;
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraButton1.Location = new System.Drawing.Point(88, 192);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(75, 24);
			this.ultraButton1.TabIndex = 155;
			this.ultraButton1.Text = "&No";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// frmAnulaRec
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(736, 230);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.lblMensaje);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmAnulaRec";
			this.ShowInTaskbar = false;
			this.Text = "Confirmación";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAnulaRec_KeyDown);
			this.Load += new System.EventHandler(this.frmAnulaRec_Load);
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void frmAnulaRec_Load(object sender, System.EventArgs e)
		{
			this.lblNumero.Text = SNumero;
		}

		private void frmAnulaRec_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
	}
}

