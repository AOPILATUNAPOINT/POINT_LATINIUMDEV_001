using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteTarjetasVentas.
	/// </summary>
	public class frmReporteTarjetasVentas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAño;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optConsulta;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optVista;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optReporte;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPorcentajes;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteTarjetasVentas()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteTarjetasVentas));
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Porcentaje", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Porcentaje", 1, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.optConsulta = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtAño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnVer = new System.Windows.Forms.Button();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btnExcel = new System.Windows.Forms.Button();
			this.optVista = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optReporte = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.uGridPorcentajes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.optConsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optVista)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPorcentajes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// optConsulta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optConsulta.Appearance = appearance1;
			this.optConsulta.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optConsulta.CheckedIndex = 0;
			this.optConsulta.ItemAppearance = appearance2;
			this.optConsulta.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Valores";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Porcentajes";
			this.optConsulta.Items.Add(valueListItem1);
			this.optConsulta.Items.Add(valueListItem2);
			this.optConsulta.ItemSpacingVertical = 10;
			this.optConsulta.Location = new System.Drawing.Point(176, 40);
			this.optConsulta.Name = "optConsulta";
			this.optConsulta.Size = new System.Drawing.Size(216, 24);
			this.optConsulta.TabIndex = 749;
			this.optConsulta.Text = "Valores";
			this.optConsulta.ValueChanged += new System.EventHandler(this.optVista_ValueChanged);
			// 
			// txtAño
			// 
			this.txtAño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAño.Location = new System.Drawing.Point(48, 8);
			this.txtAño.MinValue = 2014;
			this.txtAño.Name = "txtAño";
			this.txtAño.PromptChar = ' ';
			this.txtAño.Size = new System.Drawing.Size(88, 21);
			this.txtAño.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAño.TabIndex = 744;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 8);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(24, 16);
			this.label7.TabIndex = 748;
			this.label7.Text = "Año";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance3;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridInformacion.DisplayLayout.GroupByBox.Prompt = " ";
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridInformacion.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 80);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(872, 416);
			this.uGridInformacion.TabIndex = 747;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(8, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 745;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(88, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 746;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// optVista
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.optVista.Appearance = appearance9;
			this.optVista.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optVista.CheckedIndex = 0;
			this.optVista.Enabled = false;
			this.optVista.ItemAppearance = appearance10;
			this.optVista.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Normal";
			valueListItem4.DataValue = 2;
			valueListItem4.DisplayText = "Agrupado";
			this.optVista.Items.Add(valueListItem3);
			this.optVista.Items.Add(valueListItem4);
			this.optVista.ItemSpacingVertical = 10;
			this.optVista.Location = new System.Drawing.Point(408, 40);
			this.optVista.Name = "optVista";
			this.optVista.Size = new System.Drawing.Size(176, 24);
			this.optVista.TabIndex = 750;
			this.optVista.Text = "Normal";
			// 
			// optReporte
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.optReporte.Appearance = appearance11;
			this.optReporte.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optReporte.CheckedIndex = 0;
			this.optReporte.ItemAppearance = appearance12;
			this.optReporte.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Detallado";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Totales Por Tarjeta";
			this.optReporte.Items.Add(valueListItem5);
			this.optReporte.Items.Add(valueListItem6);
			this.optReporte.ItemSpacingVertical = 10;
			this.optReporte.Location = new System.Drawing.Point(176, 8);
			this.optReporte.Name = "optReporte";
			this.optReporte.Size = new System.Drawing.Size(216, 24);
			this.optReporte.TabIndex = 751;
			this.optReporte.Text = "Detallado";
			this.optReporte.ValueChanged += new System.EventHandler(this.optReporte_ValueChanged);
			// 
			// uGridPorcentajes
			// 
			this.uGridPorcentajes.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPorcentajes.DataSource = this.ultraDataSource1;
			appearance13.BackColor = System.Drawing.Color.White;
			this.uGridPorcentajes.DisplayLayout.Appearance = appearance13;
			this.uGridPorcentajes.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 193;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.CellAppearance = appearance14;
			ultraGridColumn2.Format = "#,##0.00";
			ultraGridColumn2.Header.Caption = "%";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 50;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance15;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridPorcentajes.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridPorcentajes.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridPorcentajes.DisplayLayout.GroupByBox.Prompt = " ";
			this.uGridPorcentajes.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPorcentajes.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPorcentajes.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridPorcentajes.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorcentajes.DisplayLayout.Override.CellAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPorcentajes.DisplayLayout.Override.HeaderAppearance = appearance18;
			this.uGridPorcentajes.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorcentajes.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorcentajes.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.uGridPorcentajes.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridPorcentajes.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPorcentajes.Location = new System.Drawing.Point(888, 80);
			this.uGridPorcentajes.Name = "uGridPorcentajes";
			this.uGridPorcentajes.Size = new System.Drawing.Size(264, 416);
			this.uGridPorcentajes.TabIndex = 752;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// frmReporteTarjetasVentas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1160, 502);
			this.Controls.Add(this.uGridPorcentajes);
			this.Controls.Add(this.optReporte);
			this.Controls.Add(this.optVista);
			this.Controls.Add(this.txtAño);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.optConsulta);
			this.KeyPreview = true;
			this.Name = "frmReporteTarjetasVentas";
			this.Text = "Reporte de Tarjetas Ventas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteTarjetasVentas_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteTarjetasVentas_Load);
			((System.ComponentModel.ISupportInitialize)(this.optConsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optVista)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPorcentajes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.uGridInformacion.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.uGridInformacion, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void frmReporteTarjetasVentas_Load(object sender, System.EventArgs e)
		{
			this.txtAño.MaxValue = DateTime.Today.Year;
			this.txtAño.Value = DateTime.Today.Year;

			this.optVista.Enabled = true;
		}

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridInformacion.Rows.Count > 0)
			{
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Tarjeta"].Width = 250;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Tarjeta"].CellActivation = Activation.ActivateOnly;

//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Banco"].Width = 150;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Banco"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Planes"].Width = 100;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Planes"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Plazos"].Width = 100;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Plazos"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].Width = 90;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.NoEdit;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor = Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0.00}";				
								
				if (this.uGridInformacion.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 4; j < iColumnas; j++)
					{
						nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 80;

						this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
					}	
					
					string sNCSuma = "";
					if ((int)this.optConsulta.Value == 1)
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
						{
							for (int j = 4; j < iColumnas; j++)
							{
								sNCSuma = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
								if (dr.Cells[sNCSuma].Value != System.DBNull.Value) dr.Cells["Total"].Value = Convert.ToDecimal(dr.Cells["Total"].Value) + Convert.ToDecimal(dr.Cells[sNCSuma].Value);
							}
						}
					}
				}				
			}
			#endregion Diseño Grid
		}

		private void DiseñoGridTotalPorTarjetas()
		{
			#region Diseño Grid
			if (this.uGridInformacion.Rows.Count > 0)
			{
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Planes_Plazos"].Width = 200;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Planes_Plazos"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].Width = 90;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.NoEdit;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor = Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0.00}";				
								
				if (this.uGridInformacion.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 2; j < iColumnas; j++)
					{
						nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 150;

						this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
					}	
					
					string sNCSuma = "";
					if ((int)this.optConsulta.Value == 1)
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
						{
							for (int j = 2; j < iColumnas; j++)
							{
								sNCSuma = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
								if (dr.Cells[sNCSuma].Value != System.DBNull.Value) dr.Cells["Total"].Value = Convert.ToDecimal(dr.Cells["Total"].Value) + Convert.ToDecimal(dr.Cells[sNCSuma].Value);
							}
						}
					}
				}				
			}
			#endregion Diseño Grid
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.uGridInformacion.DataSource = null;
			
			if ((int)this.optReporte.Value == 1)
			{
				string sProcedimiento = "ReporteTarjetasValores";
				if ((int)this.optConsulta.Value == 2) sProcedimiento = "ReporteTarjetasPorcentajes";

				string sSQL = string.Format("Exec {0} {1}", sProcedimiento, (int)this.txtAño.Value);
				this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
			
				this.DiseñoGrid();

				if ((int)this.optVista.Value == 2) 
				{
					this.uGridInformacion.DisplayLayout.Bands[0].SortedColumns.Add("Tarjeta", false, true);
					this.uGridInformacion.Rows.ExpandAll(true);
				}
			}
			else
			{
				string sProcedimiento = "ReporteVentasPlanesTarjetas";
				if ((int)this.optConsulta.Value == 2) sProcedimiento = "ReporteVentasPlanesTarjetasPorcentajes";

				string sSQL = string.Format("Exec {0} {1}", sProcedimiento, (int)this.txtAño.Value);
				this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

				DiseñoGridTotalPorTarjetas();
			}
			
			string sSQLP = string.Format("Exec ReporteVentasPlanesTarjetasPorcentajesTotales {0}", (int)this.txtAño.Value);
			this.uGridPorcentajes.DataSource = FuncionesProcedimientos.dtGeneral(sSQLP);
		}

		private void frmReporteTarjetasVentas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void optVista_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void optReporte_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int)this.optReporte.Value == 1) this.optVista.Enabled = true;
			else 
			{
				this.optVista.Enabled = false;
				this.optVista.Value = 1;
			}
		}
	}
}
