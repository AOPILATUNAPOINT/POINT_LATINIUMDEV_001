using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRolReportesNominaIngresos.
	/// </summary>
	public class frmRolReportesNominaIngresos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDatos;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEDepartamento;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRolReportesNominaIngresos()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRolReportesNominaIngresos));
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.uGridDatos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label12 = new System.Windows.Forms.Label();
			this.cmbNEDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.uGridDatos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridDatos
			// 
			this.uGridDatos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDatos.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDatos.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridDatos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDatos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDatos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDatos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.uGridDatos.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDatos.DisplayLayout.Override.CellAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 7.75F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDatos.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridDatos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDatos.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDatos.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.uGridDatos.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDatos.Location = new System.Drawing.Point(8, 88);
			this.uGridDatos.Name = "uGridDatos";
			this.uGridDatos.Size = new System.Drawing.Size(876, 272);
			this.uGridDatos.TabIndex = 737;
			this.uGridDatos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDatos_ClickCellButton);
			this.uGridDatos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDatos_InitializeLayout);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(796, 7);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 752;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(392, 39);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 751;
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
			this.label17.TabIndex = 750;
			this.label17.Text = "Hasta";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 10);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(37, 16);
			this.label18.TabIndex = 749;
			this.label18.Text = "Desde";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance7;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 748;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance8;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 747;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(8, 43);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(76, 16);
			this.label12.TabIndex = 753;
			this.label12.Text = "Departamento";
			// 
			// cmbNEDepartamento
			// 
			this.cmbNEDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridBand2.AddButtonCaption = "RolTipo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 250;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbNEDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNEDepartamento.DisplayMember = "Descripcion";
			this.cmbNEDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEDepartamento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEDepartamento.Location = new System.Drawing.Point(104, 40);
			this.cmbNEDepartamento.Name = "cmbNEDepartamento";
			this.cmbNEDepartamento.Size = new System.Drawing.Size(264, 22);
			this.cmbNEDepartamento.TabIndex = 754;
			this.cmbNEDepartamento.ValueMember = "idDepartamento";
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(892, 8);
			this.groupBox1.TabIndex = 755;
			this.groupBox1.TabStop = false;
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
			// frmRolReportesNominaIngresos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(892, 366);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.cmbNEDepartamento);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridDatos);
			this.KeyPreview = true;
			this.Name = "frmRolReportesNominaIngresos";
			this.Text = "Ingresos Por Departamento";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRolReportesNominaIngresos_KeyDown);
			this.Load += new System.EventHandler(this.frmRolReportesNominaIngresos_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDatos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;

		private void frmRolReportesNominaIngresos_Load(object sender, System.EventArgs e)
		{
			

		
			miAcceso = new Acceso (cdsSeteoF,"200307");

			if (miAcceso.BExportar) this.btnExportar.Enabled = true;
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = DateTime.Parse("01/01/" + DateTime.Today.Year.ToString());
			this.dtHasta.Value = DateTime.Today;

			this.cmbNEDepartamento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 48, 0");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDatos);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if(!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la Fecha Inicial", false,"", cdsSeteoF)) return;
			if(!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la Fecha Final", false,"", cdsSeteoF)) return;

			if(!Validacion.vbComparaFechas((DateTime) this.dtDesde.Value, (DateTime) this.dtHasta.Value, this.dtHasta,"La fecha inicial no puede ser posterior a la fecha final",false))return;
			#endregion Valida Fechas

			int idDepartamento = 0;

			if (this.cmbNEDepartamento.ActiveRow != null) idDepartamento = (int)this.cmbNEDepartamento.Value;

			this.uGridDatos.DataSource = null;

			string sSQL = string.Format("Exec RolReporteIngresosPorDepartamento '{0}', '{1}', {2}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), idDepartamento);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDatos);

			this.DiseñoGrid(this.uGridDatos);
		}

		private void DiseñoGrid(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			#region Diseño Grid
			if (uGrid.Rows.Count > 0)
			{
				int iColIni = 2;

				uGrid.DisplayLayout.Bands[0].Columns["idDepartamento"].Hidden = true;
				uGrid.DisplayLayout.Bands[0].Columns["Departamento"].Width = 140;
				uGrid.DisplayLayout.Bands[0].Columns["Departamento"].CellActivation = Activation.ActivateOnly;
				
				string nombrCol = "";
							
				int iColumnas = (int)uGrid.DisplayLayout.Bands[0].Columns.Count;
				
				for (int j = iColIni; j < iColumnas; j++)
				{
					nombrCol = uGrid.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
					uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.NoEdit;
					uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					uGrid.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
					uGrid.DisplayLayout.Bands[0].Columns[nombrCol].Width = 90;

					uGrid.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, uGrid.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
					uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
				}					
			}						
			#endregion Diseño Grid
		}

		private void uGridDatos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridDatos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.Key;
			
			if (e.Cell.Column.ToString() == sColumna)
			{
				if (e.Cell.Row.Cells[sColumna].Value == System.DBNull.Value)
				{
					MessageBox.Show("No existen registros en este periodo y departamento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				string iPosicion = Convert.ToString(sColumna.IndexOf("-", 0));
				string sMes = sColumna.Substring(0, int.Parse(iPosicion.ToString()));
				int iAño = int.Parse(sColumna.Substring(int.Parse(iPosicion.ToString()) + 1, 4));
				int iMes = 1;
				int iDia = 31;

				if (sMes == "FEBRERO") iMes = 2;
				if (sMes == "MARZO") iMes = 3;
				if (sMes == "ABRIL") iMes = 4;
				if (sMes == "MAYO") iMes = 5;
				if (sMes == "JUNIO") iMes = 6;
				if (sMes == "JULIO") iMes = 7;
				if (sMes == "AGOSTO") iMes = 8;
				if (sMes == "SEPTIEMBRE") iMes = 9;
				if (sMes == "OCTUBRE") iMes = 10;
				if (sMes == "NOVIEMBRE") iMes = 11;
				if (sMes == "DICIEMBRE") iMes = 12;				
				
				iDia = DateTime.DaysInMonth(iAño, iMes);

				DateTime dtPeriodo = DateTime.Parse(iDia.ToString() + "/" + iMes.ToString() + "/" + iAño.ToString());
				
				using (frmRolReporteDetalleIngresosPorDepartamento RRDID = new frmRolReporteDetalleIngresosPorDepartamento((int)e.Cell.Row.Cells["idDepartamento"].Value, dtPeriodo))
				{
					if (DialogResult.OK == RRDID.ShowDialog())
					{

					}
				}				
			}
		}

		private void frmRolReportesNominaIngresos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
