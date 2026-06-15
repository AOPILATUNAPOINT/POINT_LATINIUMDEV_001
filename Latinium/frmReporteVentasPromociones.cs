using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteVentasPromociones.
	/// </summary>
	public class frmReporteVentasPromociones : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaFin;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIni;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label2;
		public DevExpress.XtraEditors.SimpleButton simpleButton1;
		public DevExpress.XtraEditors.SimpleButton simpleButton2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugResVentas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCategoria;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticulo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDescripcion;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteVentasPromociones()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteVentasPromociones));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaInicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaFin");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cantidad", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cantidad", 5, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 6, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaInicio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaFin");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoBP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaBP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorBP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoLF");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaLF");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorLF");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoPL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaPL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorPL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VENTOTAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValTotal");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Promocion");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.label4 = new System.Windows.Forms.Label();
			this.dtFechaFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label10 = new System.Windows.Forms.Label();
			this.dtFechaIni = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton2 = new DevExpress.XtraEditors.SimpleButton();
			this.ugResVentas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.label3 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbCategoria = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbArticulo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbDescripcion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugResVentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDescripcion)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(16, 48);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(72, 16);
			this.label4.TabIndex = 295;
			this.label4.Text = "Descripción";
			// 
			// dtFechaFin
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaFin.Appearance = appearance1;
			this.dtFechaFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaFin.DateButtons.Add(dateButton1);
			this.dtFechaFin.Format = "dd/MM/yyyy";
			this.dtFechaFin.Location = new System.Drawing.Point(312, 14);
			this.dtFechaFin.Name = "dtFechaFin";
			this.dtFechaFin.NonAutoSizeHeight = 23;
			this.dtFechaFin.Size = new System.Drawing.Size(120, 21);
			this.dtFechaFin.SpinButtonsVisible = true;
			this.dtFechaFin.TabIndex = 292;
			this.dtFechaFin.Value = ((object)(resources.GetObject("dtFechaFin.Value")));
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(16, 16);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(53, 16);
			this.label10.TabIndex = 291;
			this.label10.Text = "Fecha INI";
			// 
			// dtFechaIni
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaIni.Appearance = appearance2;
			this.dtFechaIni.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaIni.DateButtons.Add(dateButton2);
			this.dtFechaIni.Format = "dd/MM/yyyy";
			this.dtFechaIni.Location = new System.Drawing.Point(96, 14);
			this.dtFechaIni.Name = "dtFechaIni";
			this.dtFechaIni.NonAutoSizeHeight = 23;
			this.dtFechaIni.Size = new System.Drawing.Size(120, 21);
			this.dtFechaIni.SpinButtonsVisible = true;
			this.dtFechaIni.TabIndex = 290;
			this.dtFechaIni.Value = ((object)(resources.GetObject("dtFechaIni.Value")));
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
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold);
			this.lblBodega.Location = new System.Drawing.Point(16, 80);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 635;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 8;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 128;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 8;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(96, 78);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(128, 21);
			this.cmbBodega.TabIndex = 634;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold);
			this.label2.Location = new System.Drawing.Point(240, 80);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(54, 16);
			this.label2.TabIndex = 638;
			this.label2.Text = "Categoria";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// simpleButton1
			// 
			this.simpleButton1.Appearance.Options.UseTextOptions = true;
			this.simpleButton1.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.simpleButton1.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton1.Image")));
			this.simpleButton1.ImageAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.simpleButton1.ImageIndex = 2;
			this.simpleButton1.Location = new System.Drawing.Point(704, 56);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(48, 32);
			this.simpleButton1.TabIndex = 297;
			this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
			// 
			// simpleButton2
			// 
			this.simpleButton2.Appearance.Options.UseTextOptions = true;
			this.simpleButton2.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.simpleButton2.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton2.Image")));
			this.simpleButton2.ImageAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.simpleButton2.ImageIndex = 4;
			this.simpleButton2.Location = new System.Drawing.Point(704, 16);
			this.simpleButton2.Name = "simpleButton2";
			this.simpleButton2.Size = new System.Drawing.Size(48, 32);
			this.simpleButton2.TabIndex = 294;
			this.simpleButton2.Click += new System.EventHandler(this.simpleButton2_Click);
			// 
			// ugResVentas
			// 
			this.ugResVentas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugResVentas.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugResVentas.DataSource = this.ultraDataSource2;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ugResVentas.DisplayLayout.Appearance = appearance3;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 200;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 150;
			ultraGridColumn8.Header.VisiblePosition = 4;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance4;
			ultraGridColumn9.Header.VisiblePosition = 5;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance5;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance6;
			summarySettings1.DisplayFormat = "{0: #,##0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance7;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand2.SummaryFooterCaption = "TOTAL GENERAL";
			this.ugResVentas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ugResVentas.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugResVentas.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.ugResVentas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugResVentas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugResVentas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ugResVentas.DisplayLayout.Override.CardAreaAppearance = appearance9;
			this.ugResVentas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugResVentas.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.ugResVentas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugResVentas.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugResVentas.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ugResVentas.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ugResVentas.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ugResVentas.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ugResVentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugResVentas.Location = new System.Drawing.Point(8, 120);
			this.ugResVentas.Name = "ugResVentas";
			this.ugResVentas.Size = new System.Drawing.Size(776, 416);
			this.ugResVentas.TabIndex = 288;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(240, 16);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(57, 16);
			this.label3.TabIndex = 293;
			this.label3.Text = "Fecha FIN";
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(-192, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1167, 8);
			this.groupBox1.TabIndex = 639;
			this.groupBox1.TabStop = false;
			// 
			// cmbCategoria
			// 
			appearance13.FontData.Name = "Tahoma";
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCategoria.Appearance = appearance13;
			this.cmbCategoria.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCategoria.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCategoria.DataSource = this.ultraDataSource3;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Width = 128;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11});
			this.cmbCategoria.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCategoria.DisplayMember = "";
			this.cmbCategoria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCategoria.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCategoria.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCategoria.Location = new System.Drawing.Point(304, 77);
			this.cmbCategoria.Name = "cmbCategoria";
			this.cmbCategoria.Size = new System.Drawing.Size(128, 22);
			this.cmbCategoria.TabIndex = 641;
			this.cmbCategoria.ValueMember = "";
			// 
			// ultraDataSource3
			// 
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold);
			this.label1.Location = new System.Drawing.Point(456, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 642;
			this.label1.Text = "Articulo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbArticulo
			// 
			appearance14.FontData.Name = "Tahoma";
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArticulo.Appearance = appearance14;
			this.cmbArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticulo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 8;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 160;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbArticulo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbArticulo.DisplayMember = "Articulo";
			this.cmbArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArticulo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArticulo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticulo.Location = new System.Drawing.Point(512, 13);
			this.cmbArticulo.MaxDropDownItems = 50;
			this.cmbArticulo.Name = "cmbArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(160, 22);
			this.cmbArticulo.TabIndex = 643;
			this.cmbArticulo.ValueMember = "idArticulo";
			// 
			// ultraDataSource4
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// ultraDataSource5
			// 
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25});
			// 
			// cmbDescripcion
			// 
			appearance15.FontData.Name = "Tahoma";
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDescripcion.Appearance = appearance15;
			this.cmbDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDescripcion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDescripcion.DataSource = this.ultraDataSource5;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Width = 336;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn14});
			this.cmbDescripcion.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbDescripcion.DisplayMember = "";
			this.cmbDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDescripcion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDescripcion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDescripcion.Location = new System.Drawing.Point(96, 45);
			this.cmbDescripcion.Name = "cmbDescripcion";
			this.cmbDescripcion.Size = new System.Drawing.Size(336, 22);
			this.cmbDescripcion.TabIndex = 644;
			this.cmbDescripcion.ValueMember = "";
			this.cmbDescripcion.ValueChanged += new System.EventHandler(this.cmbDescripcion_ValueChanged);
			// 
			// frmReporteVentasPromociones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(792, 542);
			this.Controls.Add(this.cmbDescripcion);
			this.Controls.Add(this.cmbArticulo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbCategoria);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.ugResVentas);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtFechaFin);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.dtFechaIni);
			this.Controls.Add(this.simpleButton1);
			this.Controls.Add(this.simpleButton2);
			this.Controls.Add(this.label3);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmReporteVentasPromociones";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Ventas Promociones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteVentasPromociones_KeyDown);
			this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.frmReporteVentasPromociones_KeyPress);
			this.Load += new System.EventHandler(this.frmReporteVentasPromociones_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugResVentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDescripcion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmReporteVentasPromociones_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteo, "111714");

			if (!Funcion.Permiso("1487", cdsSeteo))
			{					
				MessageBox.Show("No tiene acceso Reporte Ventas Promociones", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			
			this.dtFechaIni.Value = DateTime.Today;
			this.dtFechaFin.Value = DateTime.Today;

			if (MenuLatinium.stDirFacturacion == "GRETASANTORINI")this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteo, "select bodega = '10', nombre = 'GENERAL', CODIGO = 'GN' union Select Bodega, Nombre, Codigo From Bodega Where Bodega In (1, 2, 10)");
			if (MenuLatinium.stDirFacturacion == "PROSERVICIOS")this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Select Bodega, Nombre, Codigo From Bodega Where Bodega In (4)");
			this.cmbCategoria.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Select Promocion from ListaPromocionCategoria");
			this.cmbDescripcion.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Select Descripcion from CompraPromocion where Estado = 'ACTIVO'");
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			if (this.ugResVentas.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.ugResVentas, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void simpleButton2_Click(object sender, System.EventArgs e)
		{
			string Prodct = "";
			if(this.cmbDescripcion.Text.ToString() != "") Prodct = this.cmbDescripcion.Text.ToString();

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;
			if (this.dtFechaIni.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione Fecha Inicial para REPORTE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaIni.Focus();
			}
			else if (this.dtFechaFin.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione Fecha Inicial para REPORTE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaFin.Focus();
			}
			else
			{
				DateTime dtFeIni = (DateTime) this.dtFechaIni.Value;
				DateTime dtFeFin = (DateTime) this.dtFechaFin.Value;
				if(this.cmbDescripcion.Text.ToString() != "") Prodct = this.cmbDescripcion.Text.ToString();

				if(cmbArticulo.Text == "")
				{
					string sSQL = string.Format("Exec ResumenVentasXPromocion '{0}', '{1}', '{2}', {3}, '{4}'", 
						dtFeIni.ToString("yyyyMMdd"), dtFeFin.ToString("yyyyMMdd"), Prodct, (int)this.cmbBodega.Value, this.cmbCategoria.Text.ToString());
					this.ugResVentas.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
				} 
				else {
					string sSQL = string.Format("Exec ResumenVentasXPromocionDetalle '{0}', '{1}', '{2}', {3}, '{4}', {5}", 
						dtFeIni.ToString("yyyyMMdd"), dtFeFin.ToString("yyyyMMdd"), Prodct, (int)this.cmbBodega.Value, this.cmbCategoria.Text.ToString(), (int)this.cmbArticulo.Value);
					this.ugResVentas.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
				}
			}
		}

		private void frmReporteVentasPromociones_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{

		}

		private void frmReporteVentasPromociones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cmbDescripcion_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDescripcion.ActiveRow != null)
			{
				string Prodct = "";
				if(this.cmbDescripcion.Text.ToString() != "") Prodct = this.cmbDescripcion.Text.ToString();

				int idpromocion = Funcion.iEscalarSQL(cdsSeteo, string.Format("select idPromocion from CompraPromocion where Descripcion = '{0}'", Prodct));
				this.cmbArticulo.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("select a.idArticulo, a.Articulo from DetallePromocion dp inner join Articulo a on dp.idArticulo = a.idArticulo where dp.idPromocion = {0}", idpromocion));
			}
		}
	}
}
