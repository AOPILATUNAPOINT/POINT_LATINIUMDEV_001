using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CRMArticulo.
	/// </summary>
	public class CRMArticulo : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Panel panel1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Splitter splitter1;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.Panel panel3;
		private System.Windows.Forms.Splitter splitter2;
		private C1.Data.C1DataSet cdsArticulo;
		private System.Windows.Forms.Panel panel4;
		private System.Windows.Forms.Panel panel5;
		private System.Windows.Forms.Splitter splitter3;
		private System.Windows.Forms.Panel panel6;
		private Infragistics.Win.UltraWinChart.UltraChart ultraChart1;
		private Infragistics.Win.UltraWinChart.UltraChart ultraChart2;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CRMArticulo()
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
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.panel1 = new System.Windows.Forms.Panel();
			this.panel2 = new System.Windows.Forms.Panel();
			this.splitter1 = new System.Windows.Forms.Splitter();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.panel3 = new System.Windows.Forms.Panel();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.splitter2 = new System.Windows.Forms.Splitter();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.panel4 = new System.Windows.Forms.Panel();
			this.panel6 = new System.Windows.Forms.Panel();
			this.ultraChart2 = new Infragistics.Win.UltraWinChart.UltraChart();
			this.splitter3 = new System.Windows.Forms.Splitter();
			this.panel5 = new System.Windows.Forms.Panel();
			this.ultraChart1 = new Infragistics.Win.UltraWinChart.UltraChart();
			this.panel1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.panel3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			this.panel4.SuspendLayout();
			this.panel6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraChart2)).BeginInit();
			this.panel5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraChart1)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Location = new System.Drawing.Point(2, 24);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(396, 74);
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(396, 74);
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(396, 74);
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(396, 74);
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.panel2);
			this.panel1.Controls.Add(this.splitter1);
			this.panel1.Controls.Add(this.ultraGrid1);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(200, 414);
			this.panel1.TabIndex = 0;
			// 
			// panel2
			// 
			this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel2.Location = new System.Drawing.Point(0, 0);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(200, 155);
			this.panel2.TabIndex = 2;
			// 
			// splitter1
			// 
			this.splitter1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.splitter1.Location = new System.Drawing.Point(0, 155);
			this.splitter1.Name = "splitter1";
			this.splitter1.Size = new System.Drawing.Size(200, 3);
			this.splitter1.TabIndex = 1;
			this.splitter1.TabStop = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.ultraGrid1.Location = new System.Drawing.Point(0, 158);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(200, 256);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.Text = "Articulos";
			// 
			// panel3
			// 
			this.panel3.Controls.Add(this.ultraTabControl1);
			this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
			this.panel3.Location = new System.Drawing.Point(200, 0);
			this.panel3.Name = "panel3";
			this.panel3.Size = new System.Drawing.Size(400, 100);
			this.panel3.TabIndex = 1;
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl3);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl4);
			this.ultraTabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.ultraTabControl1.Location = new System.Drawing.Point(0, 0);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(400, 100);
			this.ultraTabControl1.TabIndex = 0;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Articulo";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Transacciones";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Clientes";
			ultraTab4.TabPage = this.ultraTabPageControl4;
			ultraTab4.Text = "Opciones";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2,
																																															 ultraTab3,
																																															 ultraTab4});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(396, 74);
			// 
			// splitter2
			// 
			this.splitter2.Dock = System.Windows.Forms.DockStyle.Top;
			this.splitter2.Location = new System.Drawing.Point(200, 100);
			this.splitter2.Name = "splitter2";
			this.splitter2.Size = new System.Drawing.Size(400, 3);
			this.splitter2.TabIndex = 2;
			this.splitter2.TabStop = false;
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			// 
			// panel4
			// 
			this.panel4.Controls.Add(this.panel6);
			this.panel4.Controls.Add(this.splitter3);
			this.panel4.Controls.Add(this.panel5);
			this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel4.Location = new System.Drawing.Point(200, 103);
			this.panel4.Name = "panel4";
			this.panel4.Size = new System.Drawing.Size(400, 311);
			this.panel4.TabIndex = 3;
			// 
			// panel6
			// 
			this.panel6.Controls.Add(this.ultraChart2);
			this.panel6.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel6.Location = new System.Drawing.Point(203, 0);
			this.panel6.Name = "panel6";
			this.panel6.Size = new System.Drawing.Size(197, 311);
			this.panel6.TabIndex = 2;
			// 
			// ultraChart2
			// 
			this.ultraChart2.Axis.X.Labels.Flip = false;
			this.ultraChart2.Axis.X.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart2.Axis.X.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.ItemLabel;
			this.ultraChart2.Axis.X.Labels.ItemFormatString = "<ITEM_LABEL>";
			this.ultraChart2.Axis.X.Labels.OrientationAngle = 0;
			this.ultraChart2.Axis.X.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart2.Axis.X.ScrollScale.Height = 10;
			this.ultraChart2.Axis.X.ScrollScale.Visible = false;
			this.ultraChart2.Axis.X.ScrollScale.Width = 15;
			this.ultraChart2.Axis.X.TickmarkInterval = 0;
			this.ultraChart2.Axis.X2.Labels.Flip = false;
			this.ultraChart2.Axis.X2.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.ItemLabel;
			this.ultraChart2.Axis.X2.Labels.ItemFormatString = "<ITEM_LABEL>";
			this.ultraChart2.Axis.X2.Labels.OrientationAngle = 0;
			this.ultraChart2.Axis.X2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart2.Axis.X2.ScrollScale.Height = 10;
			this.ultraChart2.Axis.X2.ScrollScale.Visible = false;
			this.ultraChart2.Axis.X2.ScrollScale.Width = 15;
			this.ultraChart2.Axis.X2.TickmarkInterval = 0;
			this.ultraChart2.Axis.Y.Labels.Flip = false;
			this.ultraChart2.Axis.Y.Labels.HorizontalAlign = System.Drawing.StringAlignment.Far;
			this.ultraChart2.Axis.Y.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.DataValue;
			this.ultraChart2.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:00.00>";
			this.ultraChart2.Axis.Y.Labels.OrientationAngle = 0;
			this.ultraChart2.Axis.Y.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart2.Axis.Y.ScrollScale.Height = 10;
			this.ultraChart2.Axis.Y.ScrollScale.Visible = false;
			this.ultraChart2.Axis.Y.ScrollScale.Width = 15;
			this.ultraChart2.Axis.Y.TickmarkInterval = 0;
			this.ultraChart2.Axis.Y2.Labels.Flip = false;
			this.ultraChart2.Axis.Y2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart2.Axis.Y2.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.DataValue;
			this.ultraChart2.Axis.Y2.Labels.ItemFormatString = "<DATA_VALUE:00.00>";
			this.ultraChart2.Axis.Y2.Labels.OrientationAngle = 0;
			this.ultraChart2.Axis.Y2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart2.Axis.Y2.ScrollScale.Height = 10;
			this.ultraChart2.Axis.Y2.ScrollScale.Visible = false;
			this.ultraChart2.Axis.Y2.ScrollScale.Width = 15;
			this.ultraChart2.Axis.Y2.TickmarkInterval = 0;
			this.ultraChart2.Axis.Z.Labels.Flip = false;
			this.ultraChart2.Axis.Z.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart2.Axis.Z.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.None;
			this.ultraChart2.Axis.Z.Labels.ItemFormatString = "";
			this.ultraChart2.Axis.Z.Labels.OrientationAngle = 0;
			this.ultraChart2.Axis.Z.Labels.SeriesFormatString = "";
			this.ultraChart2.Axis.Z.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart2.Axis.Z.ScrollScale.Height = 10;
			this.ultraChart2.Axis.Z.ScrollScale.Visible = false;
			this.ultraChart2.Axis.Z.ScrollScale.Width = 15;
			this.ultraChart2.Axis.Z.TickmarkInterval = 0;
			this.ultraChart2.Axis.Z2.Labels.Flip = false;
			this.ultraChart2.Axis.Z2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart2.Axis.Z2.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.None;
			this.ultraChart2.Axis.Z2.Labels.ItemFormatString = "";
			this.ultraChart2.Axis.Z2.Labels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal;
			this.ultraChart2.Axis.Z2.Labels.OrientationAngle = 0;
			this.ultraChart2.Axis.Z2.Labels.SeriesFormatString = "";
			this.ultraChart2.Axis.Z2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart2.Axis.Z2.ScrollScale.Height = 10;
			this.ultraChart2.Axis.Z2.ScrollScale.Visible = false;
			this.ultraChart2.Axis.Z2.ScrollScale.Width = 15;
			this.ultraChart2.Axis.Z2.TickmarkInterval = 0;
			this.ultraChart2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.ultraChart2.Location = new System.Drawing.Point(0, 0);
			this.ultraChart2.Name = "ultraChart2";
			this.ultraChart2.Size = new System.Drawing.Size(197, 311);
			this.ultraChart2.TabIndex = 0;
			// 
			// splitter3
			// 
			this.splitter3.Location = new System.Drawing.Point(200, 0);
			this.splitter3.Name = "splitter3";
			this.splitter3.Size = new System.Drawing.Size(3, 311);
			this.splitter3.TabIndex = 1;
			this.splitter3.TabStop = false;
			// 
			// panel5
			// 
			this.panel5.Controls.Add(this.ultraChart1);
			this.panel5.Dock = System.Windows.Forms.DockStyle.Left;
			this.panel5.Location = new System.Drawing.Point(0, 0);
			this.panel5.Name = "panel5";
			this.panel5.Size = new System.Drawing.Size(200, 311);
			this.panel5.TabIndex = 0;
			// 
			// ultraChart1
			// 
			this.ultraChart1.Axis.X.Labels.Flip = false;
			this.ultraChart1.Axis.X.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart1.Axis.X.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.ItemLabel;
			this.ultraChart1.Axis.X.Labels.ItemFormatString = "<ITEM_LABEL>";
			this.ultraChart1.Axis.X.Labels.OrientationAngle = 0;
			this.ultraChart1.Axis.X.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart1.Axis.X.ScrollScale.Height = 10;
			this.ultraChart1.Axis.X.ScrollScale.Visible = false;
			this.ultraChart1.Axis.X.ScrollScale.Width = 15;
			this.ultraChart1.Axis.X.TickmarkInterval = 0;
			this.ultraChart1.Axis.X2.Labels.Flip = false;
			this.ultraChart1.Axis.X2.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.ItemLabel;
			this.ultraChart1.Axis.X2.Labels.ItemFormatString = "<ITEM_LABEL>";
			this.ultraChart1.Axis.X2.Labels.OrientationAngle = 0;
			this.ultraChart1.Axis.X2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart1.Axis.X2.ScrollScale.Height = 10;
			this.ultraChart1.Axis.X2.ScrollScale.Visible = false;
			this.ultraChart1.Axis.X2.ScrollScale.Width = 15;
			this.ultraChart1.Axis.X2.TickmarkInterval = 0;
			this.ultraChart1.Axis.Y.Labels.Flip = false;
			this.ultraChart1.Axis.Y.Labels.HorizontalAlign = System.Drawing.StringAlignment.Far;
			this.ultraChart1.Axis.Y.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.DataValue;
			this.ultraChart1.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:00.00>";
			this.ultraChart1.Axis.Y.Labels.OrientationAngle = 0;
			this.ultraChart1.Axis.Y.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart1.Axis.Y.ScrollScale.Height = 10;
			this.ultraChart1.Axis.Y.ScrollScale.Visible = false;
			this.ultraChart1.Axis.Y.ScrollScale.Width = 15;
			this.ultraChart1.Axis.Y.TickmarkInterval = 0;
			this.ultraChart1.Axis.Y2.Labels.Flip = false;
			this.ultraChart1.Axis.Y2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart1.Axis.Y2.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.DataValue;
			this.ultraChart1.Axis.Y2.Labels.ItemFormatString = "<DATA_VALUE:00.00>";
			this.ultraChart1.Axis.Y2.Labels.OrientationAngle = 0;
			this.ultraChart1.Axis.Y2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart1.Axis.Y2.ScrollScale.Height = 10;
			this.ultraChart1.Axis.Y2.ScrollScale.Visible = false;
			this.ultraChart1.Axis.Y2.ScrollScale.Width = 15;
			this.ultraChart1.Axis.Y2.TickmarkInterval = 0;
			this.ultraChart1.Axis.Z.Labels.Flip = false;
			this.ultraChart1.Axis.Z.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart1.Axis.Z.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.None;
			this.ultraChart1.Axis.Z.Labels.ItemFormatString = "";
			this.ultraChart1.Axis.Z.Labels.OrientationAngle = 0;
			this.ultraChart1.Axis.Z.Labels.SeriesFormatString = "";
			this.ultraChart1.Axis.Z.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart1.Axis.Z.ScrollScale.Height = 10;
			this.ultraChart1.Axis.Z.ScrollScale.Visible = false;
			this.ultraChart1.Axis.Z.ScrollScale.Width = 15;
			this.ultraChart1.Axis.Z.TickmarkInterval = 0;
			this.ultraChart1.Axis.Z2.Labels.Flip = false;
			this.ultraChart1.Axis.Z2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near;
			this.ultraChart1.Axis.Z2.Labels.ItemFormat = Infragistics.UltraChart.Shared.Styles.AxisItemLabelFormat.None;
			this.ultraChart1.Axis.Z2.Labels.ItemFormatString = "";
			this.ultraChart1.Axis.Z2.Labels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal;
			this.ultraChart1.Axis.Z2.Labels.OrientationAngle = 0;
			this.ultraChart1.Axis.Z2.Labels.SeriesFormatString = "";
			this.ultraChart1.Axis.Z2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center;
			this.ultraChart1.Axis.Z2.ScrollScale.Height = 10;
			this.ultraChart1.Axis.Z2.ScrollScale.Visible = false;
			this.ultraChart1.Axis.Z2.ScrollScale.Width = 15;
			this.ultraChart1.Axis.Z2.TickmarkInterval = 0;
			this.ultraChart1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.ultraChart1.Location = new System.Drawing.Point(0, 0);
			this.ultraChart1.Name = "ultraChart1";
			this.ultraChart1.Size = new System.Drawing.Size(200, 311);
			this.ultraChart1.TabIndex = 0;
			// 
			// CRMArticulo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(600, 414);
			this.Controls.Add(this.panel4);
			this.Controls.Add(this.splitter2);
			this.Controls.Add(this.panel3);
			this.Controls.Add(this.panel1);
			this.Name = "CRMArticulo";
			this.Text = "CRM Articulo";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.CRMArticulo_Closing);
			this.Load += new System.EventHandler(this.CRMArticulo_Load);
			this.panel1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.panel3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			this.panel4.ResumeLayout(false);
			this.panel6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraChart2)).EndInit();
			this.panel5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraChart1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void CRMArticulo_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'CRM Articulo'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void CRMArticulo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'CRM Articulo'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}
	}
}
