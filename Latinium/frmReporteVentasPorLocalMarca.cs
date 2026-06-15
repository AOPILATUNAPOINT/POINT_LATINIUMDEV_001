using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteVentasPorLocalMarca.
	/// </summary>
	public class frmReporteVentasPorLocalMarca : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private DevExpress.XtraEditors.SimpleButton btnGenerar;
		public DevExpress.XtraEditors.SimpleButton btExportar;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridVentasPorMarca;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteVentasPorLocalMarca()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteVentasPorLocalMarca));
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.uGridVentasPorMarca = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.btExportar = new DevExpress.XtraEditors.SimpleButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasPorMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridVentasPorMarca
			// 
			this.uGridVentasPorMarca.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVentasPorMarca.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridVentasPorMarca.DisplayLayout.Appearance = appearance1;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance2;
			this.uGridVentasPorMarca.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridVentasPorMarca.DisplayLayout.ColumnScrollbarSmallChange = 50;
			this.uGridVentasPorMarca.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridVentasPorMarca.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentasPorMarca.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridVentasPorMarca.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVentasPorMarca.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridVentasPorMarca.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPorMarca.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPorMarca.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPorMarca.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridVentasPorMarca.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVentasPorMarca.Location = new System.Drawing.Point(7, 56);
			this.uGridVentasPorMarca.Name = "uGridVentasPorMarca";
			this.uGridVentasPorMarca.Size = new System.Drawing.Size(1194, 264);
			this.uGridVentasPorMarca.TabIndex = 121;
			this.uGridVentasPorMarca.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridVentasPorMarca_InitializeRow);
			this.uGridVentasPorMarca.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVentasPorMarca_KeyDown);
			this.uGridVentasPorMarca.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVentasPorMarca_InitializeLayout);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(208, 8);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(37, 16);
			this.label17.TabIndex = 649;
			this.label17.Text = "Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(16, 8);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(40, 16);
			this.label18.TabIndex = 648;
			this.label18.Text = "Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance8;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(264, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 646;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			// 
			// dtDesde
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance9;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(72, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 645;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
			// 
			// btnGenerar
			// 
			this.btnGenerar.Location = new System.Drawing.Point(592, 0);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(80, 21);
			this.btnGenerar.TabIndex = 647;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// btExportar
			// 
			this.btExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btExportar.Location = new System.Drawing.Point(1120, 8);
			this.btExportar.Name = "btExportar";
			this.btExportar.Size = new System.Drawing.Size(80, 21);
			this.btExportar.TabIndex = 644;
			this.btExportar.Text = "&Exportar";
			this.btExportar.Click += new System.EventHandler(this.btExportar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1208, 3);
			this.groupBox1.TabIndex = 650;
			this.groupBox1.TabStop = false;
			// 
			// frmReporteVentasPorLocalMarca
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1208, 326);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.btExportar);
			this.Controls.Add(this.uGridVentasPorMarca);
			this.Name = "frmReporteVentasPorLocalMarca";
			this.Text = "Ventas Por Marca y Local";
			this.Load += new System.EventHandler(this.frmReporteVentasPorLocalMarca_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasPorMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;	
			this.uGridVentasPorMarca.DataSource = null;

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size=75000;"))
			{
				oConexion.Open();

				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;				
				oCmdActualiza.CommandTimeout = 0;			
				oCmdActualiza.CommandType = CommandType.StoredProcedure;
				oCmdActualiza.CommandText = "ReporteVentasPorLocalMarca";

				int idMarca = 64;

				oCmdActualiza.Parameters.Add("@Desde", SqlDbType.DateTime).Value = (DateTime)this.dtDesde.Value;
				oCmdActualiza.Parameters.Add("@Hasta", SqlDbType.DateTime).Value = (DateTime)this.dtHasta.Value;
				oCmdActualiza.Parameters.Add("@idMarca", SqlDbType.Int).Value = idMarca;

				SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
				DataTable oDTDetalle = new DataTable();
				oDADetalle.Fill(oDTDetalle);
				this.uGridVentasPorMarca.DataSource = oDTDetalle;

				oConexion.Close();
			}			

			this.DiseñoGrid();

			this.uGridVentasPorMarca.DisplayLayout.Bands[0].SortedColumns.Add("Codigo", false, false);

			this.Cursor = Cursors.Default;
		}

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridVentasPorMarca.Rows.Count > 0)
			{
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Codigo"].Width = 160;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Modelo"].Width = 90;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Modelo"].CellActivation = Activation.ActivateOnly;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["NumeroParte"].Width = 90;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["NumeroParte"].CellActivation = Activation.ActivateOnly;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["NumeroParte"].Header.Caption = "N. Parte";
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Costo"].Width = 70;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Costo"].CellActivation = Activation.ActivateOnly;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Costo"].Format = "#,##0.00";
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Costo"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Existencia"].CellActivation = Activation.ActivateOnly;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Existencia"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Existencia"].Width = 40;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Ventas"].CellActivation = Activation.ActivateOnly;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Ventas"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Ventas"].Width = 40;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Total"].Width = 40;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.ActivateOnly;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridVentasPorMarca.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0}";

				if (this.uGridVentasPorMarca.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns.Count - 1;
				
					for (int j = 7; j < iColumnas; j++)
					{
						nombrCol = this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						
						this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns[nombrCol].Width = 30;

						this.uGridVentasPorMarca.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						this.uGridVentasPorMarca.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridVentasPorMarca.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";				
					}
				}
			}			
			#endregion Diseño Grid
		}

		private void frmReporteVentasPorLocalMarca_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
		}

		private void btExportar_Click(object sender, System.EventArgs e)
		{
			if (this.uGridVentasPorMarca.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.uGridVentasPorMarca, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGenerar_Click(sender, e);
		}

		private void uGridVentasPorMarca_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridVentasPorMarca_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridVentasPorMarca_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			string nombrCol = "";
							
			int iColumnas = (int)this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns.Count - 1;
				
			for (int j = 7; j < iColumnas; j++)
			{
				nombrCol = this.uGridVentasPorMarca.DisplayLayout.Bands[0].Columns[j].Header.Caption;						
				e.Row.Cells["Total"].Value = (int)e.Row.Cells["Total"].Value + (int)e.Row.Cells[nombrCol].Value;				
			}		
		}
	}
}
