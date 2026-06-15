using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmObservacionesVacaciones.
	/// </summary>
	public class frmObservacionesVacaciones : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmObservacionesVacaciones()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		bool bEdicion = true;
		int iNumero = 0;
		string sObservaciones = "";


		public frmObservacionesVacaciones(int INumero, string SObservaciones, bool BEdicion)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			iNumero = INumero;
			sObservaciones = SObservaciones;
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
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 152);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(456, 8);
			this.groupBox1.TabIndex = 721;
			this.groupBox1.TabStop = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(96, 164);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(75, 24);
			this.btnCancelar.TabIndex = 720;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.btnAceptar.Enabled = false;
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(16, 164);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(75, 24);
			this.btnAceptar.TabIndex = 719;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// txtObservaciones
			// 
			appearance1.BackColorDisabled = System.Drawing.Color.White;
			appearance1.BackColorDisabled2 = System.Drawing.Color.White;
			this.txtObservaciones.Appearance = appearance1;
			this.txtObservaciones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Enabled = false;
			this.txtObservaciones.Location = new System.Drawing.Point(8, 8);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(440, 136);
			this.txtObservaciones.TabIndex = 718;
			this.txtObservaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtObservaciones_KeyDown);
			// 
			// frmObservacionesVacaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(456, 198);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.txtObservaciones);
			this.Name = "frmObservacionesVacaciones";
			this.Text = "Observaciones Vacaciones ";
			this.Load += new System.EventHandler(this.frmObservacionesVacaciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmObservacionesVacaciones_Load(object sender, System.EventArgs e)
		{
			if (iNumero == 1) this.Text = this.Text + "";
			if (bEdicion) 
			{
				this.btnAceptar.Enabled = true;
				this.txtObservaciones.Enabled = true;
			}

		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}
		private void btnAceptar_Click(object sender, System.EventArgs e)
		{

			if (this.txtObservaciones.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese una Observación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtObservaciones.Focus();
				return;
			}

			this.DialogResult = DialogResult.OK;
		}

		private void txtObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAceptar.Focus();
		}
	}
}
