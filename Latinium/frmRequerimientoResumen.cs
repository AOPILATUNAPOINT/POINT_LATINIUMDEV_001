using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRequerimientoResumen.
	/// </summary>
	public class frmRequerimientoResumen : System.Windows.Forms.Form
	{
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoResumen()
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
			// 
			// frmRequerimientoResumen
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(192, 182);
			this.Name = "frmRequerimientoResumen";
			this.Text = "Notificaciones";
			this.Load += new System.EventHandler(this.frmRequerimientoResumen_Load);

		}
		#endregion

		public static int getRight = 0;
		public static int getBottom = 0;

		private void frmRequerimientoResumen_Load(object sender, System.EventArgs e)
		{
			getRight = Screen.PrimaryScreen.WorkingArea.Right;
			getBottom = Screen.PrimaryScreen.WorkingArea.Bottom;
		}
	}
}
