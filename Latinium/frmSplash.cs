using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSplash.
	/// </summary>
	public class frmSplash : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Timer timer1;
		private System.ComponentModel.IContainer components;

		public frmSplash()
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
		int x = 0;
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			// 
			// timer1
			// 
			this.timer1.Interval = 1000;
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// frmSplash
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(456, 208);
			this.ControlBox = false;
			this.Name = "frmSplash";
			this.Load += new System.EventHandler(this.frmSplash_Load);

		}
		#endregion

		private void frmSplash_Load(object sender, System.EventArgs e)
		{
//		this.Location = new Point( Screen.PrimaryScreen.Bounds.Width -01,  Screen.PrimaryScreen.Bounds.Height);
			this.Size = new System.Drawing.Size(
				Bounds.Width-100, Bounds.Height-100);

			// Set the location so the entire form is visible.
			this.Location = new System.Drawing.Point(5, 5);
		}

		private void timer1_Tick(object sender, System.EventArgs e)
		{
//		 if (this.Opacity == 0) this.Close ();
//     if (x == 200) this.Opacity = 0.02;
//		 if ( this.Location.Y == Screen.PrimaryScreen.WorkingArea.Height -100)
//			{
//				 this.Location = new Point (this.Location.X, this.Location.Y -2);
//			}
//			if (x==200 ) x = x + 1;
		}
	}
}
