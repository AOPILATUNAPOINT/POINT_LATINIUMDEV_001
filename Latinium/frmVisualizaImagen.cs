using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmVisualizaImagen.
	/// </summary>
	public class frmVisualizaImagen : System.Windows.Forms.Form
	{
		private System.Windows.Forms.PictureBox miImagen;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		string arhImgn = "";
		public frmVisualizaImagen(string nomArchivo)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			arhImgn = nomArchivo;

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
			this.miImagen = new System.Windows.Forms.PictureBox();
			this.SuspendLayout();
			// 
			// miImagen
			// 
			this.miImagen.Location = new System.Drawing.Point(16, 16);
			this.miImagen.Name = "miImagen";
			this.miImagen.Size = new System.Drawing.Size(896, 448);
			this.miImagen.TabIndex = 0;
			this.miImagen.TabStop = false;
			// 
			// frmVisualizaImagen
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(928, 478);
			this.Controls.Add(this.miImagen);
			this.Name = "frmVisualizaImagen";
			this.Text = "frmVisualizaImagen";
			this.Load += new System.EventHandler(this.frmVisualizaImagen_Load);
			this.ResumeLayout(false);

		}
		#endregion

		private void button1_Click(object sender, System.EventArgs e)
		{
			//Bitmap bm = new Bitmap(@"C:\Latinium\Iconos\Arcade.png");
      //miImagen.Image = (Image)bm;
		}

		private void frmVisualizaImagen_Load(object sender, System.EventArgs e)
		{
			Bitmap bm = new Bitmap(arhImgn);
			miImagen.SizeMode = PictureBoxSizeMode.StretchImage;
			// Para que la imagen se visualize con sus propios tamaños
			//miImagen.ClientSize = new Size(bm.Width, bm.Height);
			miImagen.Image = (Image)bm;
		}
	}
}
