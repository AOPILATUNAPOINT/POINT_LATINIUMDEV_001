using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmImprimeRolesDePago.
	/// </summary>
	public class frmImprimeRolesDePago : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDesde;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHasta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmImprimeRolesDePago()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int iTotal = 0;

		public frmImprimeRolesDePago(int ITotal)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			
			iTotal = ITotal;
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
			this.txtDesde = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).BeginInit();
			this.SuspendLayout();
			// 
			// txtDesde
			// 
			this.txtDesde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDesde.Location = new System.Drawing.Point(64, 8);
			this.txtDesde.Name = "txtDesde";
			this.txtDesde.PromptChar = ' ';
			this.txtDesde.Size = new System.Drawing.Size(72, 21);
			this.txtDesde.TabIndex = 0;
			this.txtDesde.Click += new System.EventHandler(this.txtDesde_Click);
			this.txtDesde.Validated += new System.EventHandler(this.txtDesde_Validated);
			this.txtDesde.Enter += new System.EventHandler(this.txtDesde_Enter);
			// 
			// txtHasta
			// 
			this.txtHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHasta.Location = new System.Drawing.Point(216, 8);
			this.txtHasta.Name = "txtHasta";
			this.txtHasta.PromptChar = ' ';
			this.txtHasta.Size = new System.Drawing.Size(72, 21);
			this.txtHasta.TabIndex = 1;
			this.txtHasta.Click += new System.EventHandler(this.txtHasta_Click);
			this.txtHasta.Validated += new System.EventHandler(this.txtHasta_Validated);
			this.txtHasta.Enter += new System.EventHandler(this.txtHasta_Enter);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 720;
			this.label6.Text = "Desde";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(160, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 721;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(88, 64);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 3;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnAceptar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(8, 64);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(72, 23);
			this.btnAceptar.TabIndex = 2;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(304, 8);
			this.groupBox1.TabIndex = 724;
			this.groupBox1.TabStop = false;
			// 
			// frmImprimeRolesDePago
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(298, 96);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtDesde);
			this.Controls.Add(this.txtHasta);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmImprimeRolesDePago";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Impresión";
			this.Load += new System.EventHandler(this.frmImprimeRolesDePago_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmImprimeRolesDePago_Load(object sender, System.EventArgs e)
		{
			this.txtDesde.MinValue = 1;
			this.txtHasta.MinValue = 1;

//			this.txtDesde.MaxValue = iTotal;
//			this.txtHasta.MaxValue = iTotal;

			this.txtDesde.Value = 1;
			this.txtHasta.Value = 1;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtHasta.Value < (int)this.txtDesde.Value)
			{
				MessageBox.Show("El valor final no puede ser menor al valor inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtHasta.Focus();
				return;
			}

			if ((int)this.txtHasta.Value - (int)this.txtDesde.Value > 75)
			{
				MessageBox.Show("Debe Imprimir hasta 75 Comprobantes de Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtHasta.Focus();
				return;
			}

			this.DialogResult = DialogResult.OK;
		}

		private void txtDesde_Enter(object sender, System.EventArgs e)
		{
			this.txtDesde.SelectAll();
		}

		private void txtDesde_Click(object sender, System.EventArgs e)
		{
			this.txtDesde.SelectAll();
		}

		private void txtHasta_Enter(object sender, System.EventArgs e)
		{
			this.txtHasta.SelectAll();
		}

		private void txtHasta_Click(object sender, System.EventArgs e)
		{
			this.txtHasta.SelectAll();
		}

		private void txtDesde_Validated(object sender, System.EventArgs e)
		{
			if (this.txtDesde.Value == System.DBNull.Value) this.txtDesde.Value = 1;
		}

		private void txtHasta_Validated(object sender, System.EventArgs e)
		{
			if (this.txtHasta.Value == System.DBNull.Value) this.txtDesde.Value = 1;
		}
	}
}
