using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de FormProcesando.
	/// </summary>
	public class FormProcesando : System.Windows.Forms.Form
	{
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FormProcesando()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(FormProcesando));
			// 
			// FormProcesando
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(12, 28);
			this.ClientSize = new System.Drawing.Size(186, 0);
			this.ControlBox = false;
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.MinimizeBox = false;
			this.Name = "FormProcesando";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "                  PROCESANDO...";
			this.Load += new System.EventHandler(this.FormProcesando_Load);

		}
		#endregion

		private void FormProcesando_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}