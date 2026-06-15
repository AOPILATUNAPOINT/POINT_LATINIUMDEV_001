using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Juego.
	/// </summary>
	public class Juego : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label Jugador1;
		private System.Windows.Forms.Label Jugador2;
		private System.Windows.Forms.Label Bola;
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.PictureBox pictureBox2;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Label lblAleatorio;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Juego()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Juego));
			this.Jugador1 = new System.Windows.Forms.Label();
			this.Jugador2 = new System.Windows.Forms.Label();
			this.Bola = new System.Windows.Forms.Label();
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.pictureBox2 = new System.Windows.Forms.PictureBox();
			this.button1 = new System.Windows.Forms.Button();
			this.lblAleatorio = new System.Windows.Forms.Label();
			this.button2 = new System.Windows.Forms.Button();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			this.SuspendLayout();
			// 
			// Jugador1
			// 
			this.Jugador1.AutoSize = true;
			this.Jugador1.BackColor = System.Drawing.Color.Transparent;
			this.Jugador1.Location = new System.Drawing.Point(120, 256);
			this.Jugador1.Name = "Jugador1";
			this.Jugador1.Size = new System.Drawing.Size(12, 16);
			this.Jugador1.TabIndex = 0;
			this.Jugador1.Text = "X";
			// 
			// Jugador2
			// 
			this.Jugador2.AutoSize = true;
			this.Jugador2.BackColor = System.Drawing.Color.Transparent;
			this.Jugador2.Location = new System.Drawing.Point(264, 256);
			this.Jugador2.Name = "Jugador2";
			this.Jugador2.Size = new System.Drawing.Size(13, 16);
			this.Jugador2.TabIndex = 1;
			this.Jugador2.Text = "O";
			// 
			// Bola
			// 
			this.Bola.BackColor = System.Drawing.Color.Transparent;
			this.Bola.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Bola.Location = new System.Drawing.Point(192, 112);
			this.Bola.Name = "Bola";
			this.Bola.Size = new System.Drawing.Size(8, 11);
			this.Bola.TabIndex = 2;
			this.Bola.Text = ".";
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
			this.pictureBox1.Location = new System.Drawing.Point(40, 240);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(24, 24);
			this.pictureBox1.TabIndex = 3;
			this.pictureBox1.TabStop = false;
			// 
			// pictureBox2
			// 
			this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
			this.pictureBox2.Location = new System.Drawing.Point(464, 232);
			this.pictureBox2.Name = "pictureBox2";
			this.pictureBox2.TabIndex = 4;
			this.pictureBox2.TabStop = false;
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(344, 264);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(72, 56);
			this.button1.TabIndex = 5;
			this.button1.Text = "button1";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// lblAleatorio
			// 
			this.lblAleatorio.Location = new System.Drawing.Point(24, 48);
			this.lblAleatorio.Name = "lblAleatorio";
			this.lblAleatorio.Size = new System.Drawing.Size(384, 23);
			this.lblAleatorio.TabIndex = 6;
			this.lblAleatorio.Text = "Texto a Memorizar";
			// 
			// button2
			// 
			this.button2.Location = new System.Drawing.Point(440, 8);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(72, 24);
			this.button2.TabIndex = 7;
			this.button2.Text = "Aleatorio";
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.Location = new System.Drawing.Point(24, 88);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(464, 21);
			this.ultraTextEditor1.TabIndex = 8;
			// 
			// ultraNumericEditor1
			// 
			this.ultraNumericEditor1.Location = new System.Drawing.Point(336, 8);
			this.ultraNumericEditor1.MaskInput = "nn";
			this.ultraNumericEditor1.MaxValue = 70;
			this.ultraNumericEditor1.MinValue = 1;
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor1.TabIndex = 9;
			this.ultraNumericEditor1.Value = 10;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(264, 8);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(60, 14);
			this.ultraLabel1.TabIndex = 10;
			this.ultraLabel1.Text = "Caracteres";
			// 
			// Juego
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(536, 365);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.lblAleatorio);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.pictureBox2);
			this.Controls.Add(this.pictureBox1);
			this.Controls.Add(this.Bola);
			this.Controls.Add(this.Jugador2);
			this.Controls.Add(this.Jugador1);
			this.Name = "Juego";
			this.Text = "Juego";
			this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Juego_KeyPress);
			this.Load += new System.EventHandler(this.Juego_Load);
			this.KeyUp += new System.Windows.Forms.KeyEventHandler(this.Juego_KeyUp);
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void Juego_Load(object sender, System.EventArgs e)
		{
		
		}

		private void Juego_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		}

		private bool RangoX(int x)
		{
			int iAncho = 5;
			if ( (x + iAncho > this.Bola.Left)
				&& (x - iAncho < this.Bola.Left))
			{
				return true;
			}
			return false;
		}

		private bool RangoY(int y)
		{
			int iAncho = 5;
			if ( (y + iAncho > this.Bola.Top)
				&& (y - 2*iAncho < this.Bola.Top))
			{
				return true;
			}
			return false;
		}

		private void MueveBola()
		{
			if ( this.RangoX(this.Jugador1.Left) && this.RangoY(this.Jugador1.Top))
			{
				// Movimiento Horizontal
				if (iPosXJ1 < this.Jugador1.Left)
					this.Bola.Left += 10;
				else if (iPosXJ1 > this.Jugador1.Left)
					this.Bola.Left -= 10;

				// Movimiento Vertical
				if (iPosYJ1 < this.Jugador1.Top)
					this.Bola.Top += 10;
				else if (iPosYJ1 > this.Jugador1.Top)
					this.Bola.Top -= 10;
			}

			// Movimientode jugador 2
			if ( this.RangoX(this.Jugador2.Left) && this.RangoY(this.Jugador2.Top))
			{
				// Movimiento Horizontal
				if (iPosXJ2 < this.Jugador2.Left)
					this.Bola.Left += 10;
				else if (iPosXJ2 > this.Jugador2.Left)
					this.Bola.Left -= 10;

				// Movimiento Vertical
				if (iPosYJ2 < this.Jugador2.Top)
					this.Bola.Top += 10;
				else if (iPosYJ2 > this.Jugador2.Top)
					this.Bola.Top -= 10;
			}
		}

		int iPosXJ1 = 0;
		int iPosYJ1 = 0;
		int iPosXJ2 = 0;
		int iPosYJ2 = 0;
		private void Juego_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			int iDelta = 10;
//			MessageBox.Show(e.KeyValue.ToString());
			iPosXJ1 = this.Jugador1.Left;
			iPosYJ1 = this.Jugador1.Top;
			iPosXJ2 = this.Jugador2.Left;
			iPosYJ2 = this.Jugador2.Top;
			switch (e.KeyValue)
			{
				case 37:
					this.Jugador1.Left -= iDelta;
					break;
				case 38:
					this.Jugador1.Top -= iDelta;
					break;
				case 39:
					this.Jugador1.Left += iDelta;
					break;
				case 40:
					this.Jugador1.Top += iDelta;
					break;
				case 65:
					this.Jugador2.Left -= iDelta;
					break;
				case 87:
					this.Jugador2.Top -= iDelta;
					break;
				case 83:
					this.Jugador2.Left += iDelta;
					break;
				case 90:
					this.Jugador2.Top += iDelta;
					break;

			}
			MueveBola();
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			if (Clipboard.GetDataObject() != null)
			{
				IDataObject data = Clipboard.GetDataObject();

				if (data.GetDataPresent(DataFormats.Bitmap))
				{
					Image image = (Image)data.GetData(DataFormats.Bitmap,true);

					image.Save("image.bmp",System.Drawing.Imaging.ImageFormat.Bmp);
					image.Save("image.jpg",System.Drawing.Imaging.ImageFormat.Jpeg);
					image.Save("image.gif",System.Drawing.Imaging.ImageFormat.Gif);
				}
				else
				{
					MessageBox.Show("The Data In Clipboard is not as image format");
				}
			}
			else
			{
				MessageBox.Show("The Clipboard was empty");
			}
			Application.Exit();
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			string stTexto = "ABCDEGHIJKLMNOPQRSTUVWXYZ0123456789";
			int iPos = 0;
			string stAleatorio = "";
			Random rnd = new Random();
			for (int i = 0; i< 10; i++)
			{
        iPos = rnd.Next(35);
				stAleatorio += stTexto.Substring(iPos, 1);
			}
			lblAleatorio.Text = stAleatorio;
		}
	}
}
