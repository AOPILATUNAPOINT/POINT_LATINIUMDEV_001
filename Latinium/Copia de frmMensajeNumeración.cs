using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmMensajeNumeración.
	/// </summary>
	public class frmMensajeNumeración : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblNumero;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private System.Windows.Forms.Label lblMensaje;

		string SNumero = "";
		private System.Windows.Forms.Label lblGarantiaExtendida;
		bool BAnulado = false;
		int iGarantias = 0;
		bool bDYD = false;

		public frmMensajeNumeración(string sNumero, bool bAnulado, int IGarantias, bool BDYD)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			SNumero = sNumero;
			BAnulado = bAnulado;
			iGarantias = IGarantias;
			BDYD = BDYD;
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
			this.lblMensaje = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.lblGarantiaExtendida = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// lblMensaje
			// 
			this.lblMensaje.AutoSize = true;
			this.lblMensaje.BackColor = System.Drawing.Color.Transparent;
			this.lblMensaje.Font = new System.Drawing.Font("Tahoma", 30F, System.Drawing.FontStyle.Bold);
			this.lblMensaje.Location = new System.Drawing.Point(17, 8);
			this.lblMensaje.Name = "lblMensaje";
			this.lblMensaje.Size = new System.Drawing.Size(689, 52);
			this.lblMensaje.TabIndex = 150;
			this.lblMensaje.Text = "PREPARE LA FACTURA NÚMERO:";
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Font = new System.Drawing.Font("Tahoma", 45F, System.Drawing.FontStyle.Bold);
			this.lblNumero.ForeColor = System.Drawing.Color.Red;
			this.lblNumero.Location = new System.Drawing.Point(17, 72);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(0, 76);
			this.lblNumero.TabIndex = 151;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.Image = 14;
			this.btAceptar.Appearance = appearance1;
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.btAceptar.Location = new System.Drawing.Point(304, 266);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(112, 32);
			this.btAceptar.TabIndex = 0;
			this.btAceptar.Text = "&Imprimir";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// lblGarantiaExtendida
			// 
			this.lblGarantiaExtendida.AutoSize = true;
			this.lblGarantiaExtendida.BackColor = System.Drawing.Color.Transparent;
			this.lblGarantiaExtendida.Font = new System.Drawing.Font("Tahoma", 25F, System.Drawing.FontStyle.Bold);
			this.lblGarantiaExtendida.Location = new System.Drawing.Point(17, 176);
			this.lblGarantiaExtendida.Name = "lblGarantiaExtendida";
			this.lblGarantiaExtendida.Size = new System.Drawing.Size(0, 44);
			this.lblGarantiaExtendida.TabIndex = 152;
			this.lblGarantiaExtendida.Visible = false;
			// 
			// frmMensajeNumeración
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(720, 304);
			this.ControlBox = false;
			this.Controls.Add(this.lblGarantiaExtendida);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.lblMensaje);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmMensajeNumeración";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "POINT TECHNOLOGY";
			this.Load += new System.EventHandler(this.frmMensajeNumeración_Load);
			this.ResumeLayout(false);

		}
		#endregion

		private void frmMensajeNumeración_Load(object sender, System.EventArgs e)
		{
			if (!BAnulado) this.lblMensaje.Text = "PREPARE LA FACTURA NÚMERO:";
			else this.lblMensaje.Text = "FACTURA ANULADA NÚMERO:";

			if (iGarantias > 0)
			{
				this.lblGarantiaExtendida.Visible = true;
				this.lblGarantiaExtendida.Text = iGarantias.ToString() + " HOJA(S) DE GARANTIA EXTENDIDA";
			}

			if (bDYD)
				this.lblGarantiaExtendida.Text = this.lblGarantiaExtendida.Text + "\n SEGURO";

			this.lblNumero.Text = SNumero;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
			Close();
		}
	}
}

