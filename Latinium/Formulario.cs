using Infragistics.Shared; 
using Infragistics.Win; 
using Infragistics.Win.UltraWinDataSource; 
using System.Diagnostics; 
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Formulario.
	/// </summary>
	public class Formulario : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton btCargar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Formulario()
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
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.btCargar = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(40, 32);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.TabIndex = 0;
			this.ultraButton1.Text = "Cargar";
			// 
			// btCargar
			// 
			this.btCargar.Location = new System.Drawing.Point(144, 32);
			this.btCargar.Name = "btCargar";
			this.btCargar.TabIndex = 1;
			this.btCargar.Text = "Grabar";
			this.btCargar.Click += new System.EventHandler(this.loadButton_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			this.ultraGrid1.Location = new System.Drawing.Point(16, 72);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(376, 128);
			this.ultraGrid1.TabIndex = 2;
			this.ultraGrid1.Text = "ultraGrid1";
			// 
			// Formulario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 266);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btCargar);
			this.Controls.Add(this.ultraButton1);
			this.Name = "Formulario";
			this.Text = "Formulario";
			this.Load += new System.EventHandler(this.Formulario_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

 
 
		private void saveButton_Click(object sender, System.EventArgs e) 
		{ 
			SaveFileDialog saveFileDlg = new SaveFileDialog( ); 
			saveFileDlg.Filter = "XML File|*.xml"; 
			saveFileDlg.OverwritePrompt = true; 
			DialogResult dialogResult = saveFileDlg.ShowDialog( this ); 
 
			if ( DialogResult.OK == dialogResult ) 
			{ 
				this.ultraDataSource1.SaveAsXml( saveFileDlg.FileName, true ); 
			}
		} 
 
		private void loadButton_Click(object sender, System.EventArgs e) 
		{ 
			// Following code loads previously saved data from a file. 
 
			OpenFileDialog openFileDlg = new OpenFileDialog( ); 
			openFileDlg.Filter = "XML File|*.xml"; 
			DialogResult dialogResult = openFileDlg.ShowDialog( this ); 
 
			if ( DialogResult.OK == dialogResult ) 
			{ 
				// Pass in true for the includeData parameter to load the data 
				// and not just the structure. 
				// 
				this.ultraDataSource1.LoadFromXml( openFileDlg.FileName, true ); 
			} 
		} 

		private void Formulario_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
		}
	}
}
