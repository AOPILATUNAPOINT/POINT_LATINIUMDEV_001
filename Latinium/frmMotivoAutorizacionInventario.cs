using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmMotivoAutorizacionInventario.
	/// </summary>
	public class frmMotivoAutorizacionInventario : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmMotivoAutorizacionInventario()
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
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
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
			this.txtNotas.Size = new System.Drawing.Size(584, 104);
			this.txtNotas.TabIndex = 1;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			this.txtNotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotas_KeyPress);
			this.txtNotas.ValueChanged += new System.EventHandler(this.txtNotas_ValueChanged);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 120);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 267;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 120);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 268;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// frmMotivoAutorizacionInventario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(598, 148);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtNotas);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmMotivoAutorizacionInventario";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Nota";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmMotivoAutorizacionInventario_KeyDown);
			this.Load += new System.EventHandler(this.frmMotivoAutorizacionInventario_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iLongCadena = 0;

		private void frmMotivoAutorizacionInventario_Load(object sender, System.EventArgs e)
		{
		
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbTexto(this.txtNotas, iLongCadena, 600, "Observacion")) return;
			
			this.DialogResult = DialogResult.OK;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmMotivoAutorizacionInventario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{

		}

		private void txtNotas_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnAceptar.Focus();	
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
	}
}
