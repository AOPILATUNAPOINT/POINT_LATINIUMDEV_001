using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_ResumenRevisionDeCedulas.
	/// </summary>
	public class frmCre_ResumenRevisionDeCedulas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label53;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Label lblBodega;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.Label lblContador;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_ResumenRevisionDeCedulas()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Consultas");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SiAplica");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoAplica");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Llena");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoFacturado");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturado");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectividad");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Consultas", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Consultas", 1, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SiAplica", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SiAplica", 2, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NoAplica", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NoAplica", 3, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Llena", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Llena", 4, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NoFacturado", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NoFacturado", 5, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Facturado", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Facturado", 6, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Efectividad", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectividad", 7, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Consultas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SiAplica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NoAplica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Llena");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NoFacturado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Facturado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectividad");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_ResumenRevisionDeCedulas));
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label53 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExcel = new System.Windows.Forms.Button();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.lblContador = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			this.uGridInformacion.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 180;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.CellAppearance = appearance2;
			ultraGridColumn2.Format = "#,##0";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 100;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance3;
			ultraGridColumn3.Format = "#,##0";
			ultraGridColumn3.Header.Caption = "Aplica";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 100;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Format = "#,##0";
			ultraGridColumn4.Header.Caption = "No Aplica";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 100;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance5;
			ultraGridColumn5.Format = "#,##0";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance6;
			ultraGridColumn6.Format = "#,##0";
			ultraGridColumn6.Header.Caption = "No Facturado";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance7;
			ultraGridColumn7.Format = "#,##0";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 100;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance8;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "% Efectividad";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 100;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand1.Header.Caption = "";
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance9;
			summarySettings1.DisplayFormat = "{0: #,##0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance10;
			summarySettings2.DisplayFormat = "{0: #,##0}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance11;
			summarySettings3.DisplayFormat = "{0: #,##0}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance12;
			summarySettings4.DisplayFormat = "{0: #,##0}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance13;
			summarySettings5.DisplayFormat = "{0: #,##0}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance14;
			summarySettings6.DisplayFormat = "{0: #,##0}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance15;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 10F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance19;
			appearance20.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance20.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowAlternateAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 56);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(901, 360);
			this.uGridInformacion.TabIndex = 677;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(200, 8);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 681;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance23;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 6);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 678;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// dtHasta
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance24;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 6);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 679;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 16);
			this.lblBodega.TabIndex = 680;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(392, 4);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 682;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(923, 7);
			this.groupBox1.TabIndex = 683;
			this.groupBox1.TabStop = false;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(838, 4);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 684;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// saveFileDialog1
			// 
			this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
			// 
			// ultraGridExcelExporter1
			// 
			this.ultraGridExcelExporter1.BeginExport += new Infragistics.Win.UltraWinGrid.ExcelExport.BeginExportEventHandler(this.ultraGridExcelExporter1_BeginExport);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(512, 6);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 685;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmCre_ResumenRevisionDeCedulas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(917, 422);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.uGridInformacion);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCre_ResumenRevisionDeCedulas";
			this.Text = "Resumen Revision De Cedulas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_ResumenRevisionDeCedulas_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_ResumenRevisionDeCedulas_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmCre_ResumenRevisionDeCedulas_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0827");

			if (!Funcion.Permiso("639", cdsSeteoF))
			{				
				MessageBox.Show("No tiene acceso a Resumen Revisin de Cdula", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if(miAcceso.BExportar)this.btnExcel.Enabled = true;

			DateTime dtFecha = DateTime.Parse("21/04/2016");
			this.dtDesde.CalendarInfo.MinDate = dtFecha.Date;
			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MinDate = dtFecha.Date;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			string sSQL = string.Format("Exec Cre_ReporteHistorialRevisionCedulas '{0}', '{1}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));			
			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " LOCALES ACTIVOS";
		}

		private void frmCre_ResumenRevisionDeCedulas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void saveFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void ultraGridExcelExporter1_BeginExport(object sender, Infragistics.Win.UltraWinGrid.ExcelExport.BeginExportEventArgs e)
		{
		
		}
	}
}
