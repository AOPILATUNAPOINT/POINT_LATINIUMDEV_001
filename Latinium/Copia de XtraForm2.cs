using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for XtraForm2.
	/// </summary>
	public class XtraForm2 : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraGrid.GridControl grdBalances;
		private DevExpress.XtraGrid.Views.Grid.GridView grvBalances;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView4;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public XtraForm2()
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
			this.grdBalances = new DevExpress.XtraGrid.GridControl();
			this.grvBalances = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView4 = new DevExpress.XtraGrid.Views.Grid.GridView();
			((System.ComponentModel.ISupportInitialize)(this.grdBalances)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBalances)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).BeginInit();
			this.SuspendLayout();
			// 
			// grdBalances
			// 
			// 
			// grdBalances.EmbeddedNavigator
			// 
			this.grdBalances.EmbeddedNavigator.Name = "";
			this.grdBalances.Location = new System.Drawing.Point(80, 104);
			this.grdBalances.MainView = this.grvBalances;
			this.grdBalances.Name = "grdBalances";
			this.grdBalances.Size = new System.Drawing.Size(320, 128);
			this.grdBalances.TabIndex = 18;
			this.grdBalances.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															 this.grvBalances,
																																															 this.gridView4});
			// 
			// grvBalances
			// 
			this.grvBalances.GridControl = this.grdBalances;
			this.grvBalances.Name = "grvBalances";
			this.grvBalances.OptionsSelection.MultiSelect = true;
			this.grvBalances.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvBalances.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView4
			// 
			this.gridView4.GridControl = this.grdBalances;
			this.gridView4.Name = "gridView4";
			// 
			// XtraForm2
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(768, 398);
			this.Controls.Add(this.grdBalances);
			this.Name = "XtraForm2";
			this.Text = "XtraForm2";
			((System.ComponentModel.ISupportInitialize)(this.grdBalances)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBalances)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
	}
}

