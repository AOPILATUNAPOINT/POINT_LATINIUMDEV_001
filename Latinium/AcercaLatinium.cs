using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de AcercaLatinium.
	/// </summary>
	public class AcercaLatinium : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public AcercaLatinium()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(AcercaLatinium));
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label6 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.SuspendLayout();
			// 
			// pictureBox1
			// 
			this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
			this.pictureBox1.Location = new System.Drawing.Point(16, 26);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(120, 112);
			this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox1.TabIndex = 0;
			this.pictureBox1.TabStop = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(201, 26);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(71, 22);
			this.label1.TabIndex = 1;
			this.label1.Text = "Latinium";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(160, 69);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(170, 17);
			this.label2.TabIndex = 2;
			this.label2.Text = "GRUPO INFOELECT Cía. Ltda.";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(32, 336);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(120, 17);
			this.label3.TabIndex = 3;
			this.label3.Text = "www.infoelect.net";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(184, 103);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(105, 16);
			this.label4.TabIndex = 4;
			this.label4.Text = "Version N° 1.0.4450";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(8, 26);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(288, 17);
			this.label5.TabIndex = 5;
			this.label5.Text = "Teléfonos: 2445 435 / 2445 436 / 2450 629 / 2432 934";
			this.label5.Click += new System.EventHandler(this.label5_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.label6);
			this.groupBox1.Controls.Add(this.label5);
			this.groupBox1.Location = new System.Drawing.Point(24, 146);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(320, 87);
			this.groupBox1.TabIndex = 6;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = " Quito ";
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(12, 52);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(284, 17);
			this.label6.TabIndex = 6;
			this.label6.Text = "Dirección: Amazonas N41 56 e Isla Floreana. Of. 101 B";
			this.label6.Click += new System.EventHandler(this.label6_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label7);
			this.groupBox2.Controls.Add(this.label8);
			this.groupBox2.Location = new System.Drawing.Point(24, 233);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(320, 95);
			this.groupBox2.TabIndex = 7;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = " Guayaquil ";
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(12, 52);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(300, 28);
			this.label7.TabIndex = 6;
			this.label7.Text = "Cdla. Kennedy Vieja, Calle Novena Oeste y Av. San Jorge Condominio Delta Piso 2- " +
				"Oficina 210";
			// 
			// label8
			// 
			this.label8.Location = new System.Drawing.Point(8, 26);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(280, 17);
			this.label8.TabIndex = 5;
			this.label8.Text = "Teléfonos: 2296032 - 2295072 - 087477523";
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(216, 336);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(120, 17);
			this.label9.TabIndex = 8;
			this.label9.Text = "correo@infoelect.net";
			// 
			// AcercaLatinium
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(360, 374);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.pictureBox1);
			this.Controls.Add(this.groupBox1);
			this.Name = "AcercaLatinium";
			this.Text = "Acerca de Latinium";
			this.Load += new System.EventHandler(this.AcercaLatinium_Load);
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void AcercaLatinium_Load(object sender, System.EventArgs e)
		{
		}

		private void label5_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label6_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
