using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for ArticuloGrupo.
	/// </summary>
	public class ArticuloGrupo : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinTabControl.UltraTabControl tbArticulo;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl6;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl7;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl8;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl9;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ArticuloGrupo()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab6 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab7 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab8 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab9 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.tbArticulo = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl6 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl7 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl8 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl9 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			((System.ComponentModel.ISupportInitialize)(this.tbArticulo)).BeginInit();
			this.tbArticulo.SuspendLayout();
			this.SuspendLayout();
			// 
			// tbArticulo
			// 
			this.tbArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tbArticulo.Controls.Add(this.ultraTabSharedControlsPage1);
			this.tbArticulo.Controls.Add(this.ultraTabPageControl1);
			this.tbArticulo.Controls.Add(this.ultraTabPageControl2);
			this.tbArticulo.Controls.Add(this.ultraTabPageControl3);
			this.tbArticulo.Controls.Add(this.ultraTabPageControl4);
			this.tbArticulo.Controls.Add(this.ultraTabPageControl5);
			this.tbArticulo.Controls.Add(this.ultraTabPageControl6);
			this.tbArticulo.Controls.Add(this.ultraTabPageControl7);
			this.tbArticulo.Controls.Add(this.ultraTabPageControl8);
			this.tbArticulo.Controls.Add(this.ultraTabPageControl9);
			this.tbArticulo.Location = new System.Drawing.Point(16, 32);
			this.tbArticulo.Name = "tbArticulo";
			this.tbArticulo.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.tbArticulo.Size = new System.Drawing.Size(1136, 637);
			this.tbArticulo.TabIndex = 75;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Grupos";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Precios";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Complemento";
			ultraTab4.TabPage = this.ultraTabPageControl4;
			ultraTab4.Text = "Comisión";
			ultraTab5.TabPage = this.ultraTabPageControl5;
			ultraTab5.Text = "Promoción";
			ultraTab6.TabPage = this.ultraTabPageControl6;
			ultraTab6.Text = "Foto";
			ultraTab6.Visible = false;
			ultraTab7.TabPage = this.ultraTabPageControl7;
			ultraTab7.Text = "Activo Fijo";
			ultraTab8.TabPage = this.ultraTabPageControl8;
			ultraTab8.Text = "Reportes";
			ultraTab9.TabPage = this.ultraTabPageControl9;
			ultraTab9.Text = "tab1";
			ultraTab9.Visible = false;
			this.tbArticulo.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																												 ultraTab1,
																																												 ultraTab2,
																																												 ultraTab3,
																																												 ultraTab4,
																																												 ultraTab5,
																																												 ultraTab6,
																																												 ultraTab7,
																																												 ultraTab8,
																																												 ultraTab9});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1132, 611);
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1132, 611);
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(1132, 611);
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(1132, 611);
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(1132, 611);
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(1132, 611);
			// 
			// ultraTabPageControl6
			// 
			this.ultraTabPageControl6.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl6.Name = "ultraTabPageControl6";
			this.ultraTabPageControl6.Size = new System.Drawing.Size(1132, 611);
			// 
			// ultraTabPageControl7
			// 
			this.ultraTabPageControl7.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl7.Name = "ultraTabPageControl7";
			this.ultraTabPageControl7.Size = new System.Drawing.Size(1132, 611);
			// 
			// ultraTabPageControl8
			// 
			this.ultraTabPageControl8.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl8.Name = "ultraTabPageControl8";
			this.ultraTabPageControl8.Size = new System.Drawing.Size(1132, 611);
			// 
			// ultraTabPageControl9
			// 
			this.ultraTabPageControl9.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl9.Name = "ultraTabPageControl9";
			this.ultraTabPageControl9.Size = new System.Drawing.Size(1132, 611);
			// 
			// ArticuloGrupo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1704, 790);
			this.Controls.Add(this.tbArticulo);
			this.Name = "ArticuloGrupo";
			this.Text = "ArticuloGrupo";
			((System.ComponentModel.ISupportInitialize)(this.tbArticulo)).EndInit();
			this.tbArticulo.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		private void btKardex_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btMensual_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btVendedor_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btClienteProv_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
		
		}

		private void chkAgrupado_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}

