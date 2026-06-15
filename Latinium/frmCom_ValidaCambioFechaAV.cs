using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_ValidaCambioFechaAV.
	/// </summary>
	public class frmCom_ValidaCambioFechaAV : System.Windows.Forms.Form
	{
		public System.Windows.Forms.RadioButton rdbNuevoMovimiento;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.RadioButton rdbPFV;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_ValidaCambioFechaAV()
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
			this.rdbNuevoMovimiento = new System.Windows.Forms.RadioButton();
			this.rdbPFV = new System.Windows.Forms.RadioButton();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.SuspendLayout();
			// 
			// rdbNuevoMovimiento
			// 
			this.rdbNuevoMovimiento.Location = new System.Drawing.Point(32, 8);
			this.rdbNuevoMovimiento.Name = "rdbNuevoMovimiento";
			this.rdbNuevoMovimiento.Size = new System.Drawing.Size(184, 24);
			this.rdbNuevoMovimiento.TabIndex = 0;
			this.rdbNuevoMovimiento.Text = "Nuevo Movimiento";
			// 
			// rdbPFV
			// 
			this.rdbPFV.Location = new System.Drawing.Point(32, 40);
			this.rdbPFV.Name = "rdbPFV";
			this.rdbPFV.Size = new System.Drawing.Size(184, 24);
			this.rdbPFV.TabIndex = 1;
			this.rdbPFV.Text = "Permisos Faltas Vacaciones";
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 88);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 2;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(96, 88);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 3;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(248, 8);
			this.groupBox1.TabIndex = 4;
			this.groupBox1.TabStop = false;
			// 
			// frmCom_ValidaCambioFechaAV
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(248, 118);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.rdbPFV);
			this.Controls.Add(this.rdbNuevoMovimiento);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmCom_ValidaCambioFechaAV";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Seleccione una Opción";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCom_ValidaCambioFechaAV_KeyDown);
			this.Load += new System.EventHandler(this.frmCom_ValidaCambioFechaAV_Load);
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCom_ValidaCambioFechaAV_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (!this.rdbNuevoMovimiento.Checked && !this.rdbPFV.Checked)
			{
				MessageBox.Show("Seleccione una Opción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.DialogResult = DialogResult.OK;
		}

		private void frmCom_ValidaCambioFechaAV_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
