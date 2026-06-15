using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteDeVentasPorMes.
	/// </summary>
	public class frmReporteDeVentasPorMes : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAño;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugridGarantia;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.ContextMenu mnuExportar;
		private System.Windows.Forms.MenuItem mnuExportarVentaGeneral;
		private System.Windows.Forms.MenuItem mnuExportarGEx;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteDeVentasPorMes()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteDeVentasPorMes));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.btnVer = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtAño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.ugridGarantia = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnExportar = new System.Windows.Forms.Button();
			this.mnuExportar = new System.Windows.Forms.ContextMenu();
			this.mnuExportarVentaGeneral = new System.Windows.Forms.MenuItem();
			this.mnuExportarGEx = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridGarantia)).BeginInit();
			this.SuspendLayout();
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
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 10);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(23, 17);
			this.label7.TabIndex = 734;
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
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 40);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1072, 192);
			this.uGridInformacion.TabIndex = 735;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1088, 8);
			this.groupBox1.TabIndex = 736;
			this.groupBox1.TabStop = false;
			// 
			// txtAño
			// 
			this.txtAño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAño.Location = new System.Drawing.Point(48, 8);
			this.txtAño.MinValue = 2014;
			this.txtAño.Name = "txtAño";
			this.txtAño.PromptChar = ' ';
			this.txtAño.Size = new System.Drawing.Size(88, 22);
			this.txtAño.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAño.TabIndex = 0;
			this.txtAño.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAño_KeyDown);
			this.txtAño.Click += new System.EventHandler(this.txtAño_Click);
			this.txtAño.Validated += new System.EventHandler(this.txtAño_Validated);
			this.txtAño.Enter += new System.EventHandler(this.txtAño_Enter);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(552, 16);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 738;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ugridGarantia
			// 
			this.ugridGarantia.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugridGarantia.Cursor = System.Windows.Forms.Cursors.Default;
			appearance7.BackColor = System.Drawing.Color.White;
			this.ugridGarantia.DisplayLayout.Appearance = appearance7;
			ultraGridBand2.SummaryFooterCaption = "";
			this.ugridGarantia.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ugridGarantia.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugridGarantia.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugridGarantia.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.ugridGarantia.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridGarantia.DisplayLayout.Override.CellAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugridGarantia.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.ugridGarantia.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridGarantia.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridGarantia.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ugridGarantia.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugridGarantia.Location = new System.Drawing.Point(8, 240);
			this.ugridGarantia.Name = "ugridGarantia";
			this.ugridGarantia.Size = new System.Drawing.Size(1072, 200);
			this.ugridGarantia.TabIndex = 739;
			this.ugridGarantia.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugridGarantia_InitializeLayout);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(248, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 745;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// mnuExportar
			// 
			this.mnuExportar.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuExportarVentaGeneral,
																																								this.mnuExportarGEx});
			// 
			// mnuExportarVentaGeneral
			// 
			this.mnuExportarVentaGeneral.Index = 0;
			this.mnuExportarVentaGeneral.Text = "Venta General";
			this.mnuExportarVentaGeneral.Click += new System.EventHandler(this.mnuExportarVentaGeneral_Click);
			// 
			// mnuExportarGEx
			// 
			this.mnuExportarGEx.Index = 1;
			this.mnuExportarGEx.Text = "Garantia Extendida";
			this.mnuExportarGEx.Click += new System.EventHandler(this.mnuExportarGEx_Click);
			// 
			// frmReporteDeVentasPorMes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1088, 454);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.ugridGarantia);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.txtAño);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.btnVer);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "frmReporteDeVentasPorMes";
			this.Text = "Ventas Anuales Por Local";
			this.Load += new System.EventHandler(this.frmReporteDeVentasPorMes_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridGarantia)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmReporteDeVentasPorMes_Load(object sender, System.EventArgs e)
		{
			this.txtAño.MaxValue = DateTime.Today.Year;
			this.txtAño.Value = DateTime.Today.Year;
		}

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridInformacion.Rows.Count > 0)
			{
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Bodega"].Width = 160;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Bodega"].CellActivation = Activation.ActivateOnly;

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
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 80;

						this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
					}

					string sNCSuma = "";
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
					{
						for (int j = 1; j < iColumnas; j++)
						{
							sNCSuma = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
							if (dr.Cells[sNCSuma].Value != System.DBNull.Value) dr.Cells["Total"].Value = Convert.ToDecimal(dr.Cells["Total"].Value) + Convert.ToDecimal(dr.Cells[sNCSuma].Value);
						}
					}
				}				
			}

			this.uGridInformacion.DisplayLayout.Bands[0].Header.Caption = this.uGridInformacion.Rows.Count + " ARTÍCULOS ENCONTRADOS.";
			#endregion Diseño Grid

		}




		private void diseñoGridG()
		{
		
		
			#region Diseño Grid1
			if (this.ugridGarantia.Rows.Count > 0)
			{
				this.ugridGarantia.DisplayLayout.Bands[0].Columns["Bodega"].Width = 160;
				this.ugridGarantia.DisplayLayout.Bands[0].Columns["Bodega"].CellActivation = Activation.ActivateOnly;

				this.ugridGarantia.DisplayLayout.Bands[0].Columns["Total"].Width = 90;
				this.ugridGarantia.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.NoEdit;
				this.ugridGarantia.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ugridGarantia.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor = Color.Silver;
				this.ugridGarantia.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.ugridGarantia.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.ugridGarantia.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ugridGarantia.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0.00}";


				if (this.ugridGarantia.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.ugridGarantia.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 2; j < iColumnas; j++)
					{
						nombrCol = this.ugridGarantia.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
						this.ugridGarantia.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						this.ugridGarantia.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.ugridGarantia.DisplayLayout.Bands[0].Columns[nombrCol].Width = 80;

						this.ugridGarantia.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.ugridGarantia.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						this.ugridGarantia.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.ugridGarantia.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
					}

					string sNCSuma = "";
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugridGarantia.Rows.All)
					{
						for (int j = 1; j < iColumnas; j++)
						{
							sNCSuma = this.ugridGarantia.DisplayLayout.Bands[0].Columns[j].Header.Caption;
							if (dr.Cells[sNCSuma].Value != System.DBNull.Value) dr.Cells["Total"].Value = Convert.ToDecimal(dr.Cells["Total"].Value) + Convert.ToDecimal(dr.Cells[sNCSuma].Value);
						}
					}
				}				
			}

			this.ugridGarantia.DisplayLayout.Bands[0].Header.Caption = this.ugridGarantia.Rows.Count + " ARTÍCULOS ENCONTRADOS.";
			#endregion Diseño Grid1
		
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.uGridInformacion.DataSource = null;
		
			string sSQL = string.Format("Exec ReporteDeVentasPorMes {0}", Convert.ToInt32(this.txtAño.Value));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInformacion);

			this.DiseñoGrid();

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";

			this.ugridGarantia.DataSource=null;
			string sSQLG = string.Format("Exec ReporteDeVentasPorMesGEx {0}",Convert.ToInt32(this.txtAño.Value));
			FuncionesProcedimientos.dsGeneral(sSQLG, this.ugridGarantia);
			this.diseñoGridG();
		}

		private void txtAño_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void txtAño_Click(object sender, System.EventArgs e)
		{
			this.txtAño.SelectAll();
		}

		private void txtAño_Enter(object sender, System.EventArgs e)
		{
			this.txtAño.SelectAll();
		}

		private void txtAño_Validated(object sender, System.EventArgs e)
		{
			if (this.txtAño.Value == System.DBNull.Value) this.txtAño.Value = DateTime.Today.Year;
		}



		private void ugridGarantia_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			this.mnuExportar.Show(this.btnExportar, new Point(1, 20));
		}

		private void mnuExportarVentaGeneral_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void mnuExportarGEx_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ugridGarantia);
		}

		
	}
}
