using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmJustDescuento.
	/// </summary>
	public class frmJustDescuento : DevExpress.XtraEditors.XtraForm
	{
		public DevExpress.XtraEditors.SimpleButton btAceptar;
		public DevExpress.XtraEditors.SimpleButton btCancelar;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmJustDescuento()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

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
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(8, 120);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(112, 24);
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			this.btAceptar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.btAceptar_KeyDown);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.Location = new System.Drawing.Point(128, 120);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(112, 24);
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
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
			this.txtNotas.Size = new System.Drawing.Size(584, 104);
			this.txtNotas.TabIndex = 0;
			this.txtNotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotas_KeyPress);
			this.txtNotas.ValueChanged += new System.EventHandler(this.txtNotas_ValueChanged);
			// 
			// frmJustDescuento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(600, 150);
			this.ControlBox = false;
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btCancelar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmJustDescuento";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Nota Descuento";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmJustDescuento_KeyDown);
			this.Load += new System.EventHandler(this.frmJustDescuento_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
			Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (this.txtNotas.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese un Mótivo de DESCUENTO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNotas.Focus();			
			}
			else this.DialogResult = DialogResult.OK;
		}

		private void frmJustDescuento_Load(object sender, System.EventArgs e)
		{
		
		}

		private void txtNotas_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btAceptar.Focus();				
		}

		private void btAceptar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) 
			{
				this.DialogResult = DialogResult.No;
				Close();
			}
		}

		private void frmJustDescuento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
	}
}

