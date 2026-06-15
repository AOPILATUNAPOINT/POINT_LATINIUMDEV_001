using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Threading;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_ReporteCargas.
	/// </summary>
	public class frmCre_ReporteCargas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnImportarCartera;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPagos;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnGenerarPagos;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_ReporteCargas()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_ReporteCargas));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Documento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Haber");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoFinal");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Documento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoFinal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, "", "Debe", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debe", 6, true);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Haber", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Haber", 7, true);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoFinal", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoFinal", 8, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.label7 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnImportarCartera = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnExportar = new System.Windows.Forms.Button();
			this.uGridPagos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnGenerarPagos = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.SuspendLayout();
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(16, 64);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(37, 17);
			this.label7.TabIndex = 969;
			this.label7.Text = "Buscar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(56, 64);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(216, 20);
			this.txtBuscar.TabIndex = 0;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(288, 64);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 1;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-272, 96);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1016, 8);
			this.groupBox1.TabIndex = 972;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(776, 16);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 973;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnImportarCartera
			// 
			this.btnImportarCartera.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImportarCartera.Enabled = false;
			this.btnImportarCartera.Image = ((System.Drawing.Image)(resources.GetObject("btnImportarCartera.Image")));
			this.btnImportarCartera.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportarCartera.Location = new System.Drawing.Point(384, 16);
			this.btnImportarCartera.Name = "btnImportarCartera";
			this.btnImportarCartera.Size = new System.Drawing.Size(104, 24);
			this.btnImportarCartera.TabIndex = 1018;
			this.btnImportarCartera.Text = "Importar Carga";
			this.btnImportarCartera.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportarCartera.Click += new System.EventHandler(this.btnImportarCartera_Click);
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
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(384, 64);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 24);
			this.btnExportar.TabIndex = 1026;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// uGridPagos
			// 
			this.uGridPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPagos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPagos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridPagos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.BackColorAlpha = Infragistics.Win.Alpha.Opaque;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			appearance2.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellButtonAppearance = appearance2;
			ultraGridColumn10.Header.Caption = "Ver";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn10.Width = 26;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance3;
			summarySettings1.DisplayFormat = "{0:#,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance4;
			summarySettings2.DisplayFormat = "{0:#,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance5;
			summarySettings3.DisplayFormat = "{0:#,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.uGridPagos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridPagos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPagos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPagos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridPagos.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPagos.DisplayLayout.Override.CellAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8.5F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPagos.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.uGridPagos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagos.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagos.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridPagos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPagos.Location = new System.Drawing.Point(8, 104);
			this.uGridPagos.Name = "uGridPagos";
			this.uGridPagos.Size = new System.Drawing.Size(1000, 320);
			this.uGridPagos.TabIndex = 966;
			this.uGridPagos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridPagos_InitializeLayout);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(176, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 17);
			this.label4.TabIndex = 1030;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(16, 16);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(35, 17);
			this.label5.TabIndex = 1029;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance11;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(224, 16);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1028;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance12;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 16);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1027;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// btnGenerarPagos
			// 
			this.btnGenerarPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerarPagos.Enabled = false;
			this.btnGenerarPagos.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerarPagos.Image")));
			this.btnGenerarPagos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerarPagos.Location = new System.Drawing.Point(528, 16);
			this.btnGenerarPagos.Name = "btnGenerarPagos";
			this.btnGenerarPagos.Size = new System.Drawing.Size(104, 24);
			this.btnGenerarPagos.TabIndex = 1031;
			this.btnGenerarPagos.Text = "Generar Archivo ";
			this.btnGenerarPagos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerarPagos.Click += new System.EventHandler(this.btnGenerarPagos_Click);
			// 
			// frmCre_ReporteCargas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1016, 430);
			this.Controls.Add(this.btnGenerarPagos);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnImportarCartera);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.uGridPagos);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmCre_ReporteCargas";
			this.Text = "Reporte Cargas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_ReporteCargas_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_ReporteCargas_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;
		private void frmCre_ReporteCargas_Load(object sender, System.EventArgs e)
		{
			string stTipo = "47001";
			string stPermiso = "1997";
							
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			
			if (!Funcion.Permiso(stPermiso, cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Libro Mayores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				if (miAcceso.BImportar) this.btnImportarCartera.Enabled = true;
				if (miAcceso.BExportar) this.btnGenerarPagos.Enabled = true;
				if (miAcceso.BExportar) this.btnExportar.Enabled = true;
				DateTime dtFecha = DateTime.Parse("01/06/2016");

				this.dtDesde.CalendarInfo.MinDate = dtFecha;
				this.dtHasta.CalendarInfo.MinDate = dtFecha;

				this.dtDesde.Value = "01/01/2020";
				this.dtHasta.Value = "31/12/2020";

//				this.dtDesde.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
//				this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
				//FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
				
				//			this.txtIMoraP.MinValue = -540;
				//			this.txtIMoraP.MaxValue = 105;
				//			this.txtFMoraP.MinValue = -540;
				//			this.txtFMoraP.MaxValue = 105;
			}
			
		}
		private void UnloadMe()
		{
			this.Close();
		}

//		private void optEstado_ValueChanged(object sender, System.EventArgs e)
//		{
//			Consulta();
//		}
//
//		private void btnVer_Click(object sender, System.EventArgs e)
//		{
//			Consulta();
//		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Consulta()
		{
			string sSQL = string.Format("Exec Cre_ConsultaCargaMayores '{0}', '{1}','{2}'", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), this.txtBuscar.Text.ToString());
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridPagos);

			this.lblContador.Text = this.uGridPagos.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnImportarCartera_Click(object sender, System.EventArgs e)
		{
			#region Validacion
		
	
			
//			iContRegistros = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select COUNT(*) From Cre_Cartera_Vigente_Pichincha Where CONVERT(DATE, Fecha) = '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));		
//
//			if (iContRegistros > 0)
//			{
//				MessageBox.Show("Existen datos importados a la fecha actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				return;
//			}

//			if (Funcion.iEscalarSQL(cdsSeteoF, "Exec Cre_ValidaPagosPichinchaPorLiquidar") > 0)
//			{
//				MessageBox.Show("Existen pagos por liquidar de dias anteriores a Hoy ('01-11-2018')\n\nLiquide los pagos antes de generar un nuevo envio", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				return;
//			}			
			#endregion Validacion
			
			int iProceso = 0;
			
			try
			{				
				#region Archivo Cartera Vigente
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;
					
					#region Variables
					int iContVigente = 0;
					int iCont = 0;
					string sSQLVD = "";
					string FECHA = "" ; 
					string	CODIGO	="";
					string	NOMBRE	 ="";
					string	NUMERO	="";
					string	DETALLE	="";
					string	DOCUMENTO	="";
					decimal	DEBE	= 0;
					decimal	HABER	= 0;
					decimal	SALDOFINAL	= 0;
					DateTime FECHASUBIDA=DateTime.Today ; 
					string USUARIO="";
					#endregion Variables
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [Hoja1$]", "Hoja1").Tables[0].Rows)
					{	
						if (row["FECHA"] != System.DBNull.Value)  FECHA	 = row["FECHA"].ToString();
						if (row["CODIGO"] != System.DBNull.Value) CODIGO = row["CODIGO"].ToString();
						if (row["NUMERO"] != System.DBNull.Value) NUMERO = row["NUMERO"].ToString(); 
						if (row["DETALLE"] != System.DBNull.Value)DETALLE = row["DETALLE"].ToString();
						if (row["DOCUMENTO"] != System.DBNull.Value)DOCUMENTO= row["DOCUMENTO"].ToString();
						#region Duplicados
						 if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT COUNT(*) FROM CargaMayoresImp WHERE DOCUMENTO = '{0}' AND NUMERO = '{1}'", DOCUMENTO, NUMERO)) > 0)
						{
							sSQLVD = string.Format("Exec ValidaDuplicadosCargaMayores '{0}', '{1}', '{2}', '{3}', '{4}',{5},'REGISTRO DUPLICADO'", 
								FECHA,CODIGO,NUMERO,DETALLE,DOCUMENTO,iCont);					
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}
							iCont++;
						#endregion Duplicados
					
					}
					#region REPETIDOS
					if (Funcion.iEscalarSQL(cdsSeteoF, "select count(*) from  ValidaCargaMayoresImp") > 0)
					{
						frmValidacionCargaMayores VRB = new frmValidacionCargaMayores();
						VRB.ShowDialog();

						Funcion.EjecutaSQL(cdsSeteoF, "Delete From ValidaCargaMayoresImp");

						this.Cursor = Cursors.Default;

						return;
					}
					
					iCont = 0;
					#endregion REPETIDOS
					#region Sube informacion
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [Hoja1$]", "Hoja1").Tables[0].Rows)
					{	
						if (row["FECHA"] != System.DBNull.Value)  FECHA	 = row["FECHA"].ToString();
						if (row["CODIGO"] != System.DBNull.Value) CODIGO = row["CODIGO"].ToString();
						if (row["NOMBRE"] != System.DBNull.Value) NOMBRE = row["NOMBRE"].ToString();
						if (row["NUMERO"] != System.DBNull.Value) NUMERO = row["NUMERO"].ToString(); 
						if (row["DETALLE"] != System.DBNull.Value)DETALLE = row["DETALLE"].ToString();
						if (row["DOCUMENTO"] != System.DBNull.Value)DOCUMENTO= row["DOCUMENTO"].ToString();
						if (row["DEBE"] != System.DBNull.Value) DEBE = decimal.Parse( row["DEBE"].ToString());
						if (row["HABER"] != System.DBNull.Value) HABER = decimal.Parse( row["HABER"].ToString());
						if (row["SALDOFINAL"] != System.DBNull.Value) SALDOFINAL= decimal.Parse( row["SALDOFINAL"].ToString());

//						// Verifica si el documento ya existe en la base de datos
//						string sSQLVerificar = string.Format("SELECT COUNT(*) FROM CargaMayoresImp WHERE DOCUMENTO = '{0}'", DOCUMENTO);
//						int existeDocumento = Funcion.iEscalarSQL(cdsSeteoF, sSQLVerificar); // Asumiendo que esta función ejecuta una consulta SQL y devuelve un valor entero.
				
							// El documento no existe en la base de datos, puedes insertar el registro
							string sSQLAct = string.Format("Exec Importa_Carga_Mayores '{0}', '{1}', '{2}', '{3}', '{4}', '{5}',{6},{7}, {8},'{9}'", 
								FECHA,CODIGO,NOMBRE,NUMERO,DETALLE,DOCUMENTO,DEBE,HABER,SALDOFINAL,MenuLatinium.stUsuario);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
							iContVigente++;		
						
							
					}
					MessageBox.Show(string.Format("Proceso Finalizado Correctamente, {0} Importadas", iContVigente), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					Consulta();
					#endregion Sube informacion
					this.Cursor = Cursors.Default;
					#endregion Archivo Cartera Vigente
				}
				}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Demo", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				this.Cursor = Cursors.Default;							
			}
		}

		

//		private void btnGenerarPagos_Click(object sender, System.EventArgs e)
//		{
//			if (DialogResult.OK == MessageBox.Show("¿Esta seguro de GENERAR EL ARCHIVO DE PAGOS?", "Demo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
//			{
//				SaveFileDialog sArchivo = new SaveFileDialog();
//			
//				sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
//				sArchivo.FileName = string.Format("TRAPNT_{0}.txt", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
//				sArchivo.OverwritePrompt = true;
//				sArchivo.Title = "Genera Archivo de Pagos Banco Pichincha";
//			
//				#region Pagos
//				if (sArchivo.ShowDialog() == DialogResult.OK)
//				{
//					using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
//					{
//						#region Genera Archivo
//						int i = 0;
//
//						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_Tramas_Pagos_Pichincha '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"))).Rows)
//						{			
//							string iLinea = dr["Fecha"].ToString() + "\t" + dr["Peticion"].ToString() + "\t" + dr["NumeroLote"].ToString() + "\t" + dr["Producto"].ToString() + "\t"
//								+ dr["Referencia"].ToString() + "\t" + dr["AgenciaIngreso"].ToString() + "\t" + dr["Usuario"].ToString() + "\t"
//								+ dr["FechaIngreso"].ToString() + "\t" + dr["HoraIngreso"].ToString() + "\t" + dr["SistemaIngreso"].ToString() + "\t"
//								+ dr["Campo_11"].ToString();			
//							file.WriteLine(iLinea);
//			
//							i++;					
//						}
//						#endregion Genera Archivo
//			
////						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_Actualiza_Estado_Envio_Pagos '{0}', 2, 0",
////							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
//
//						MessageBox.Show(string.Format("{0} Tramas de Pagos Generadas Correctamente", i), "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
//
//						this.btnImportarCartera.Enabled = false;
//						this.btnActualizaPagos.Enabled = false;
//						this.btnGenerarPagos.Enabled = false;
//					}
//				}
//				#endregion Pagos
//			}
//		}
//
//		


		private void uGridPagos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridPagos);
		}

//		private void btnSeleccionar_Click(object sender, System.EventArgs e)
//		{
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPagos.Rows.All)
//				dr.Cells["Sel"].Value = !(bool)dr.Cells["Sel"].Value;
//		}

		private void frmCre_ReporteCargas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		

	

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtIMoraP_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos_1(sender, e);
		}

		private void txtFMoraP_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos_1(sender, e);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnGenerarPagos_Click(object sender, System.EventArgs e)
		{
			try
			{
				if (uGridPagos.Rows.Count == 0)
				{
					MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}
				SaveFileDialog sArchivo = new SaveFileDialog();
				sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivo.FileName = sArchivo.FileName;
				sArchivo.OverwritePrompt = true;
				sArchivo.Title = "Guardar CargasM Archivo TXT";

				if (sArchivo.ShowDialog() == DialogResult.OK)
				{
					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
					{
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandTimeout = 0;

						string sSQL =string.Format("Exec Cre_ConsultaCargaMayores '{0}', '{1}','{2}'", 
							Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), this.txtBuscar.Text.ToString());
							oCmdActualiza.CommandText = sSQL;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);	


						using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
						{
							int i = 1;
							string sEspacio  =" ";
							foreach (DataRow dr in oDTDetalle.Rows)
							{
								string fecha = dr["Fecha"].ToString();
								string codigo = dr["Codigo"].ToString();
								string nombre = dr["Nombre"].ToString();
								string numero = dr["Numero"].ToString();
								string detalle = dr["Detalle"].ToString();
								string documento = dr["Documento"].ToString();
								string debe = dr["Debe"].ToString();
								string haber = dr["Haber"].ToString();
								string saldoFinal = dr["SaldoFinal"].ToString();

								// Establece un ancho fijo para el campo "Detalle" (por ejemplo, 20 caracteres)
								detalle = detalle.PadRight(35);
								numero = numero.PadRight(20);
								documento = documento.PadRight(20);
								debe = debe.PadRight(20);
								haber = haber.PadRight(20);
								saldoFinal = saldoFinal.PadRight(20);


								 string iLinea = fecha + "\t" + codigo + "\t" + nombre + "\t" + numero + "\t" + detalle + "\t" + documento + "\t" + debe + "\t" + haber + "\t" + saldoFinal;
    
								file.WriteLine(iLinea);
								i++;
							}

						}
						MessageBox.Show(string.Format("Archivo Exportado Correctamente"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
			}
			catch (Exception exc)
			{
				MessageBox.Show(exc.Message);
			}
		}

	}
}
