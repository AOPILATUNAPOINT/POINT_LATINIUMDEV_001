using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmTipoCoberturaCP.
	/// </summary>
	public class frmTipoCoberturaCP : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnCerrar;
		private System.Windows.Forms.Button btnPerdida;
		private System.Windows.Forms.Button btnReparacion;
		private System.Windows.Forms.Label lblSubTotal;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorProforma;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTipoCoberturaCP()
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
			this.btnCerrar = new System.Windows.Forms.Button();
			this.btnReparacion = new System.Windows.Forms.Button();
			this.btnPerdida = new System.Windows.Forms.Button();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.txtValorProforma = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtValorProforma)).BeginInit();
			this.SuspendLayout();
			// 
			// btnCerrar
			// 
			this.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCerrar.Location = new System.Drawing.Point(232, 40);
			this.btnCerrar.Name = "btnCerrar";
			this.btnCerrar.Size = new System.Drawing.Size(104, 24);
			this.btnCerrar.TabIndex = 979;
			this.btnCerrar.Text = "CERRAR";
			this.btnCerrar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
			// 
			// btnReparacion
			// 
			this.btnReparacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnReparacion.Location = new System.Drawing.Point(120, 40);
			this.btnReparacion.Name = "btnReparacion";
			this.btnReparacion.Size = new System.Drawing.Size(104, 24);
			this.btnReparacion.TabIndex = 978;
			this.btnReparacion.Text = "REPARACIÓN";
			this.btnReparacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnReparacion.Click += new System.EventHandler(this.btnReparacion_Click);
			// 
			// btnPerdida
			// 
			this.btnPerdida.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnPerdida.Location = new System.Drawing.Point(8, 40);
			this.btnPerdida.Name = "btnPerdida";
			this.btnPerdida.Size = new System.Drawing.Size(104, 24);
			this.btnPerdida.TabIndex = 977;
			this.btnPerdida.Text = "PÉRDIDA TOTAL";
			this.btnPerdida.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnPerdida.Click += new System.EventHandler(this.btnPerdida_Click);
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(8, 10);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(90, 16);
			this.lblSubTotal.TabIndex = 980;
			this.lblSubTotal.Text = "Valor reparación:";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtValorProforma
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorProforma.Appearance = appearance1;
			this.txtValorProforma.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorProforma.FormatString = "#,##0.00";
			this.txtValorProforma.Location = new System.Drawing.Point(120, 8);
			this.txtValorProforma.Name = "txtValorProforma";
			this.txtValorProforma.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorProforma.PromptChar = ' ';
			this.txtValorProforma.Size = new System.Drawing.Size(104, 21);
			this.txtValorProforma.TabIndex = 981;
			// 
			// frmTipoCoberturaCP
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(346, 72);
			this.ControlBox = false;
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.txtValorProforma);
			this.Controls.Add(this.btnCerrar);
			this.Controls.Add(this.btnReparacion);
			this.Controls.Add(this.btnPerdida);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmTipoCoberturaCP";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Seleccione una opción:";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmTipoCoberturaCP_KeyDown);
			this.Load += new System.EventHandler(this.frmTipoCoberturaCP_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtValorProforma)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmTipoCoberturaCP_Load(object sender, System.EventArgs e)
		{
		
		}

		private void btnCerrar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmTipoCoberturaCP_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		public bool bTipo = false;

		private void btnPerdida_Click(object sender, System.EventArgs e)
		{
			if (Convert.ToDecimal(this.txtValorProforma.Value) == 0.00m)
			{
				MessageBox.Show("Ingrese el valor de pérdida total.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			//this.txtValorProforma.Value = 0.00m;
			if (DialogResult.No == MessageBox.Show("CONFIRMAR COBERTURA POR PERDIDA TOTAL", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;

			bTipo = false;

			this.DialogResult = DialogResult.OK;
		}

		private void btnReparacion_Click(object sender, System.EventArgs e)
		{
			if (Convert.ToDecimal(this.txtValorProforma.Value) == 0.00m)
			{
				MessageBox.Show("Ingrese el valor de la proforma del servicio técnico", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.No == MessageBox.Show("CONFIRMAR COBERTURA POR REPARACION", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;

			bTipo = true;

			this.DialogResult = DialogResult.OK;
		}
	}
}
