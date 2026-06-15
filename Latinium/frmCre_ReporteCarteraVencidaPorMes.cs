using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_ReporteCarteraVencidaPorMes.
	/// </summary>
	public class frmCre_ReporteCarteraVencidaPorMes : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDatos;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_ReporteCarteraVencidaPorMes()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_ReporteCarteraVencidaPorMes));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridDatos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDatos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(992, 8);
			this.groupBox1.TabIndex = 765;
			this.groupBox1.TabStop = false;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(896, 7);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 762;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(384, 6);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 761;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(200, 10);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(34, 16);
			this.label17.TabIndex = 760;
			this.label17.Text = "Hasta";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 10);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(37, 16);
			this.label18.TabIndex = 759;
			this.label18.Text = "Desde";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 758;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 757;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// uGridDatos
			// 
			this.uGridDatos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDatos.Cursor = System.Windows.Forms.Cursors.Default;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridDatos.DisplayLayout.Appearance = appearance3;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridDatos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDatos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDatos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDatos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridDatos.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDatos.DisplayLayout.Override.CellAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 7.75F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDatos.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridDatos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDatos.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDatos.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridDatos.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDatos.Location = new System.Drawing.Point(8, 56);
			this.uGridDatos.Name = "uGridDatos";
			this.uGridDatos.Size = new System.Drawing.Size(976, 304);
			this.uGridDatos.TabIndex = 756;
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// frmCre_ReporteCarteraVencidaPorMes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(992, 366);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridDatos);
			this.KeyPreview = true;
			this.Name = "frmCre_ReporteCarteraVencidaPorMes";
			this.Text = "Cartera Vencida Por Mes";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_ReporteCarteraVencidaPorMes_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_ReporteCarteraVencidaPorMes_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDatos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}


		private void frmCre_ReporteCarteraVencidaPorMes_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "081218");

			if (!Funcion.Permiso("944", cdsSeteoF))
			{		
				MessageBox.Show("No tiene Acceso CARTERA VENCIDA POR MES", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if(miAcceso.BExportar)this.btnExportar.Enabled = true;

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = DateTime.Parse("01/01/" + DateTime.Today.Year.ToString());
			this.dtHasta.Value = DateTime.Today;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if(!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la Fecha Inicial", false,"", cdsSeteoF)) return;
			if(!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la Fecha Final", false,"", cdsSeteoF)) return;

			if(!Validacion.vbComparaFechas((DateTime) this.dtDesde.Value, (DateTime) this.dtHasta.Value, this.dtHasta,"La fecha inicial no puede ser posterior a la fecha final",false))return;
			#endregion Valida Fechas
			
			this.uGridDatos.DataSource = null;

			string sSQL = string.Format("Exec Cre_ReporteCarteraVencidaPorMes '{0}', '{1}', 0", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDatos);

			this.DiseñoGrid(this.uGridDatos);
		}

		private void DiseñoGrid(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			#region Diseño Grid
			if (uGrid.Rows.Count > 0)
			{
				int iColIni = 1;

				uGrid.DisplayLayout.Bands[0].Columns["RangoMorosidad"].Width = 120;
				uGrid.DisplayLayout.Bands[0].Columns["RangoMorosidad"].CellActivation = Activation.ActivateOnly;
				
				string nombrCol = "";
							
				int iColumnas = (int)uGrid.DisplayLayout.Bands[0].Columns.Count;
				
				for (int j = iColIni; j < iColumnas; j++)
				{
					nombrCol = uGrid.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
					uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
					uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					uGrid.DisplayLayout.Bands[0].Columns[nombrCol].Width = 90;

					uGrid.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, uGrid.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
					uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
				}					
			}						
			#endregion Diseño Grid
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDatos);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmCre_ReporteCarteraVencidaPorMes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
