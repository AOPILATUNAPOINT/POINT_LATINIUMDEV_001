using System;
using System.Drawing.Printing;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ImpresoraPredefinida.
	/// </summary>
	public class ImpresoraPredefinida : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.ComboBox comboInstalledPrinters;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ImpresoraPredefinida()
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
			this.comboInstalledPrinters = new System.Windows.Forms.ComboBox();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.label1 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// comboInstalledPrinters
			// 
			this.comboInstalledPrinters.Location = new System.Drawing.Point(80, 8);
			this.comboInstalledPrinters.Name = "comboInstalledPrinters";
			this.comboInstalledPrinters.Size = new System.Drawing.Size(384, 21);
			this.comboInstalledPrinters.TabIndex = 0;
			this.comboInstalledPrinters.Text = "Ninguna";
			this.comboInstalledPrinters.SelectedIndexChanged += new System.EventHandler(this.comboInstalledPrinters_SelectionChanged);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(8, 40);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(96, 40);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 3;
			this.label1.Text = "Nombre";
			// 
			// ImpresoraPredefinida
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(472, 78);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.comboInstalledPrinters);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ImpresoraPredefinida";
			this.Text = "Configurar Impresora";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ImpresoraPredefinida_Closing);
			this.Load += new System.EventHandler(this.ImpresoraPredefinida_Load);
			this.ResumeLayout(false);

		}
		#endregion

		PrintDocument printDoc = new PrintDocument();
		private void ImpresoraPredefinida_Load(object sender, System.EventArgs e)
		{
			PopulateInstalledPrintersCombo();
			comboInstalledPrinters.Text = printDoc.PrinterSettings.PrinterName;
			Cursor = Cursors.Default;
		}
		private void PopulateInstalledPrintersCombo()
		{
			// Add list of installed printers found to the combo box.
			// The pkInstalledPrinters string will be used to provide the display string.
			String pkInstalledPrinters;
			for (int i = 0; i < PrinterSettings.InstalledPrinters.Count; i++)
			{
				pkInstalledPrinters = PrinterSettings.InstalledPrinters[i];
				comboInstalledPrinters.Items.Add(pkInstalledPrinters);
			}
		}

		private void comboInstalledPrinters_SelectionChanged(object sender, System.EventArgs e)
		{
			// Set the printer to a printer in the combo box when the selection changes.
			if (comboInstalledPrinters.SelectedIndex != -1) 
			{
				// The combo box's Text property returns the selected item's text, which is the printer name.
//				ReportDocument oRpt= new ReportDocument();
//				oRpt.Load(this.crVista.ReportSource.ToString());
				printDoc.PrinterSettings.PrinterName= comboInstalledPrinters.Text;
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			MenuLatinium.ImpresoraNombre = comboInstalledPrinters.Text;
			Close();
		}

		private void ImpresoraPredefinida_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
		}

	}
}
