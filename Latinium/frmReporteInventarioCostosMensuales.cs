using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteInventarioCostosMensuales.
	/// </summary>
	public class frmReporteInventarioCostosMensuales : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAño;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteInventarioCostosMensuales()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteInventarioCostosMensuales));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.txtAño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.label7 = new System.Windows.Forms.Label();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			this.SuspendLayout();
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
			this.txtAño.TabIndex = 0;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(152, 6);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 1;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(232, 6);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 2;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 10);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(24, 16);
			this.label7.TabIndex = 743;
			this.label7.Text = "Año";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 56);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1008, 520);
			this.uGridInformacion.TabIndex = 742;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1024, 8);
			this.groupBox1.TabIndex = 744;
			this.groupBox1.TabStop = false;
			// 
			// frmReporteInventarioCostosMensuales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1024, 582);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.txtAño);
			this.Controls.Add(this.btnVer);
			this.Name = "frmReporteInventarioCostosMensuales";
			this.Text = "InventarioCostosMensuales";
			this.Load += new System.EventHandler(this.frmReporteInventarioCostosMensuales_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
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

		private void frmReporteInventarioCostosMensuales_Load(object sender, System.EventArgs e)
		{
			this.txtAño.MaxValue = DateTime.Today.Year;
			this.txtAño.Value = DateTime.Today.Year;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec  {1}", (int)this.txtAño.Value);
			this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			this.DiseñoGrid();
		}

		private void DiseñoGrid()
		{
			#region Diseño Grid
//			if (this.uGridInformacion.Rows.Count > 0)
//			{
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Codigo"].Width = 100;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;
//
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 300;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Descripcion"].CellActivation = Activation.ActivateOnly;
//
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].Width = 90;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.NoEdit;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor = Color.Silver;
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0.00}";
//
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"].Width = 90;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"].Hidden = true;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"].CellActivation = Activation.NoEdit;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Semestre", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Semestre"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Semestre"].DisplayFormat = "{0: #,##0.00}";
//
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].Width = 90;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellActivation = Activation.NoEdit;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellAppearance.BackColor = Color.Silver;
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Promedio", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Promedio"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Promedio"].DisplayFormat = "{0: #,##0.00}";
//
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].Width = 90;
//				if ((int)this.optVista.Value == 2 && (int)this.optGastoMes.Value == 2) this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].Hidden = false;
//				else this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].Hidden = true;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].CellActivation = Activation.NoEdit;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].CellAppearance.BackColor = Color.Silver;
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("PromedioMes", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["PromedioMes"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["PromedioMes"].DisplayFormat = "{0: #,##0.00}";
//				
//				if (this.uGridInformacion.Rows.Count > 0)
//				{
//					string nombrCol = "";
//							
//					int iColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;
//				
//					for (int j = 6; j < iColumnas; j++)
//					{
//						nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
//												
//						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
//						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
//						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 80;
//
//						this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
//						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
//						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";										
//					}
//				}				
//			}

//			this.uGridInformacion.DisplayLayout.Bands[0].Header.Caption = this.uGridInformacion.Rows.Count + " ARTÍCULOS ENCONTRADOS.";
			#endregion Diseño Grid
		}
	}
}
