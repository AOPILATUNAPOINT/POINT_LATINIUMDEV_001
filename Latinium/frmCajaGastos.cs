using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCajaGastos.
	/// </summary>
	public class frmCajaGastos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblBodega;
		public DevExpress.XtraEditors.SimpleButton btVer;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		public DevExpress.XtraEditors.SimpleButton btExportar;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugCajaGastos;
		private System.ComponentModel.IContainer components;

		public frmCajaGastos()
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCajaGastos));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorGasto");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorGasto");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.lblBodega = new System.Windows.Forms.Label();
			this.btVer = new DevExpress.XtraEditors.SimpleButton();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btExportar = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ugCajaGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugCajaGastos)).BeginInit();
			this.SuspendLayout();
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 276;
			this.lblBodega.Text = "LOCAL";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// btVer
			// 
			this.btVer.Appearance.Options.UseTextOptions = true;
			this.btVer.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btVer.Image = ((System.Drawing.Image)(resources.GetObject("btVer.Image")));
			this.btVer.ImageAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.btVer.ImageIndex = 4;
			this.btVer.Location = new System.Drawing.Point(432, 8);
			this.btVer.Name = "btVer";
			this.btVer.Size = new System.Drawing.Size(48, 32);
			this.btVer.TabIndex = 277;
			this.btVer.Click += new System.EventHandler(this.btVer_Click);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 30;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 30;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 180;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 30;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Codigo";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(8, 24);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(112, 21);
			this.cmbBodega.TabIndex = 274;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// btExportar
			// 
			this.btExportar.Appearance.Options.UseTextOptions = true;
			this.btExportar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btExportar.Image = ((System.Drawing.Image)(resources.GetObject("btExportar.Image")));
			this.btExportar.ImageAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.btExportar.ImageIndex = 2;
			this.btExportar.Location = new System.Drawing.Point(480, 8);
			this.btExportar.Name = "btExportar";
			this.btExportar.Size = new System.Drawing.Size(48, 32);
			this.btExportar.TabIndex = 275;
			this.btExportar.Click += new System.EventHandler(this.btExportar_Click);
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn4.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// ugCajaGastos
			// 
			this.ugCajaGastos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ugCajaGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugCajaGastos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugCajaGastos.DisplayLayout.Appearance = appearance1;
			this.ugCajaGastos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 133;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 130;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 130;
			ultraGridColumn8.Header.Caption = "Valor";
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Width = 106;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand2.SummaryFooterCaption = "";
			this.ugCajaGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ugCajaGastos.DisplayLayout.GroupByBox.Hidden = true;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugCajaGastos.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ugCajaGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugCajaGastos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ugCajaGastos.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugCajaGastos.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugCajaGastos.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugCajaGastos.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugCajaGastos.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ugCajaGastos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugCajaGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugCajaGastos.Location = new System.Drawing.Point(8, 56);
			this.ugCajaGastos.Name = "ugCajaGastos";
			this.ugCajaGastos.Size = new System.Drawing.Size(520, 392);
			this.ugCajaGastos.TabIndex = 278;
			// 
			// frmCajaGastos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(536, 461);
			this.Controls.Add(this.ugCajaGastos);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.btVer);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btExportar);
			this.Name = "frmCajaGastos";
			this.Text = "Caja de Gastos";
			this.Load += new System.EventHandler(this.frmCajaGastos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugCajaGastos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCajaGastos_Load(object sender, System.EventArgs e)
		{
			string stExecBodega = string.Format("Select idBodega, Bodega, Codigo, Nombre From Bodega Where Activo = 1");			
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExecBodega);
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btExportar_Click(object sender, System.EventArgs e)
		{
			if (this.ugCajaGastos.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.ugCajaGastos, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void btVer_Click(object sender, System.EventArgs e)
		{
			int Bodg = 0;
			if(this.cmbBodega.ActiveRow != null)
			{
				Bodg = (int)this.cmbBodega.Value;
			}
			/*Visuliza cajas con gastos*/
			string sSQL = string.Format("Exec ReporteCaja_Gastos {0}", Bodg);
			this.ugCajaGastos.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);

			this.ugCajaGastos.DisplayLayout.Bands[0].SortedColumns.Add("Fecha", false, true);			
			this.ugCajaGastos.Rows.CollapseAll(true);
		}
	}
}
