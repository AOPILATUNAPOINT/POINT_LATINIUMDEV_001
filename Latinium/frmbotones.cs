using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmbotones.
	/// </summary>
	public class frmbotones : System.Windows.Forms.Form
	{
		public System.Windows.Forms.Button btnOrden;
		public System.Windows.Forms.Button btnProduccion;
		public System.Windows.Forms.Button btnSalir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmbotones()
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
			this.btnOrden = new System.Windows.Forms.Button();
			this.btnProduccion = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// btnOrden
			// 
			this.btnOrden.BackColor = System.Drawing.Color.Red;
			this.btnOrden.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnOrden.ForeColor = System.Drawing.Color.White;
			this.btnOrden.Location = new System.Drawing.Point(9, 9);
			this.btnOrden.Name = "btnOrden";
			this.btnOrden.Size = new System.Drawing.Size(160, 30);
			this.btnOrden.TabIndex = 2;
			this.btnOrden.Text = "Orden De &Compra";
			this.btnOrden.Click += new System.EventHandler(this.btnOrden_Click);
			// 
			// btnProduccion
			// 
			this.btnProduccion.BackColor = System.Drawing.Color.Yellow;
			this.btnProduccion.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnProduccion.Location = new System.Drawing.Point(177, 9);
			this.btnProduccion.Name = "btnProduccion";
			this.btnProduccion.Size = new System.Drawing.Size(160, 30);
			this.btnProduccion.TabIndex = 3;
			this.btnProduccion.Text = "Orden De &Producción";
			this.btnProduccion.Click += new System.EventHandler(this.btnProduccion_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnSalir.Location = new System.Drawing.Point(345, 9);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(160, 30);
			this.btnSalir.TabIndex = 4;
			this.btnSalir.Text = "C&ancelar";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// frmbotones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 13);
			this.BackColor = System.Drawing.SystemColors.Control;
			this.ClientSize = new System.Drawing.Size(514, 48);
			this.ControlBox = false;
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnProduccion);
			this.Controls.Add(this.btnOrden);
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmbotones";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Seleccione Una Opción";
			this.TransparencyKey = System.Drawing.Color.Transparent;
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmbotones_KeyDown);
			this.ResumeLayout(false);

		}
		#endregion

		public int iOrigen = 0;

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;			
		}

		private void btnOrden_Click(object sender, System.EventArgs e)
		{
			iOrigen = 1;

			this.DialogResult = DialogResult.OK;
		}

		private void btnProduccion_Click(object sender, System.EventArgs e)
		{
			iOrigen = 2;

			this.DialogResult = DialogResult.OK;		
		}

		private void frmbotones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
	}
}
