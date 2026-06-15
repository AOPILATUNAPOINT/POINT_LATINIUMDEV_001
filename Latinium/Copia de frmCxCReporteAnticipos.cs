using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCxCReporteAnticipos.
	/// </summary>
	public class frmCxCReporteAnticipos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private DevExpress.XtraEditors.RadioGroup rdgEstado;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTipos;
		private System.Windows.Forms.Button btnIrA;
		private System.Windows.Forms.ContextMenu mnuIrA;
		private System.Windows.Forms.MenuItem mnuIrADocumento;
		private System.Windows.Forms.MenuItem mnuIrAEstadoDeCuenta;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCxCReporteAnticipos()
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
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCxCReporteAnticipos));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abono");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CXC");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPOINT");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ReposicionSolidario");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Devuelto");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Abono", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Abono", 9, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CXC", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CXC", 10, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPOINT", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPOINT", 11, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ReposicionSolidario", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ReposicionSolidario", 12, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Devuelto", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Devuelto", 13, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 14, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Abono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CXC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPOINT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ReposicionSolidario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Devuelto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.btnVer = new System.Windows.Forms.Button();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnExcel = new System.Windows.Forms.Button();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.label6 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.rdgEstado = new DevExpress.XtraEditors.RadioGroup();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnIrA = new System.Windows.Forms.Button();
			this.mnuIrA = new System.Windows.Forms.ContextMenu();
			this.mnuIrADocumento = new System.Windows.Forms.MenuItem();
			this.mnuIrAEstadoDeCuenta = new System.Windows.Forms.MenuItem();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgEstado.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(352, 70);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(77, 24);
			this.btnVer.TabIndex = 691;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 51;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 88;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 70;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 75;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Cédula";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 90;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 250;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 15;
			ultraGridColumn8.Width = 500;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance2;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 40;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance3;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 75;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance4;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 75;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance5;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 75;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance6;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Solidario";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 75;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance7;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 75;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance8;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 75;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.Width = 100;
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
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			ultraGridBand1.Header.Caption = "";
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance9;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance10;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance11;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance12;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance13;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance14;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.CellAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance18;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance19.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance19.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.uGridDocumentos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(8, 128);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(1256, 272);
			this.uGridDocumentos.TabIndex = 688;
			this.uGridDocumentos.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridDocumentos_InitializeRow);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn12.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn13.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn14.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
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
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// dtHasta
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance22;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(240, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 680;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance23;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 679;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// txtBuscar
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance24;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(64, 72);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(280, 21);
			this.txtBuscar.TabIndex = 681;
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1192, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 692;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 42);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(26, 16);
			this.label6.TabIndex = 689;
			this.label6.Text = "Tipo";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(184, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 684;
			this.label1.Text = "Hasta";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 683;
			this.label2.Text = "Desde";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 74);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(39, 16);
			this.label3.TabIndex = 682;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// rdgEstado
			// 
			this.rdgEstado.EditValue = 1;
			this.rdgEstado.Location = new System.Drawing.Point(440, 8);
			this.rdgEstado.Name = "rdgEstado";
			// 
			// rdgEstado.Properties
			// 
			this.rdgEstado.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgEstado.Properties.Appearance.Options.UseBackColor = true;
			this.rdgEstado.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
			this.rdgEstado.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Pendientes"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Anulados"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Cancelados")});
			this.rdgEstado.Size = new System.Drawing.Size(88, 88);
			this.rdgEstado.TabIndex = 687;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(544, 72);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 693;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1272, 8);
			this.groupBox1.TabIndex = 694;
			this.groupBox1.TabStop = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// cmbTipos
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipos.Appearance = appearance25;
			this.cmbTipos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipos.DataSource = this.ultraDataSource2;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 280;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbTipos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipos.DisplayMember = "Descripcion";
			this.cmbTipos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipos.Location = new System.Drawing.Point(64, 40);
			this.cmbTipos.Name = "cmbTipos";
			this.cmbTipos.Size = new System.Drawing.Size(280, 21);
			this.cmbTipos.TabIndex = 696;
			this.cmbTipos.ValueMember = "Tipo";
			// 
			// btnIrA
			// 
			this.btnIrA.Image = ((System.Drawing.Image)(resources.GetObject("btnIrA.Image")));
			this.btnIrA.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnIrA.Location = new System.Drawing.Point(1192, 40);
			this.btnIrA.Name = "btnIrA";
			this.btnIrA.Size = new System.Drawing.Size(75, 24);
			this.btnIrA.TabIndex = 713;
			this.btnIrA.Text = "Ir a";
			this.btnIrA.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnIrA.Click += new System.EventHandler(this.btnIrA_Click);
			// 
			// mnuIrA
			// 
			this.mnuIrA.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																					 this.mnuIrADocumento,
																																					 this.mnuIrAEstadoDeCuenta});
			// 
			// mnuIrADocumento
			// 
			this.mnuIrADocumento.Index = 0;
			this.mnuIrADocumento.Text = "Documento";
			this.mnuIrADocumento.Click += new System.EventHandler(this.mnuIrADocumento_Click);
			// 
			// mnuIrAEstadoDeCuenta
			// 
			this.mnuIrAEstadoDeCuenta.Index = 1;
			this.mnuIrAEstadoDeCuenta.Text = "Estado de Cuenta";
			this.mnuIrAEstadoDeCuenta.Click += new System.EventHandler(this.mnuIrAEstadoDeCuenta_Click);
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
			// frmCxCReporteAnticipos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1272, 406);
			this.Controls.Add(this.btnIrA);
			this.Controls.Add(this.cmbTipos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.uGridDocumentos);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.rdgEstado);
			this.Controls.Add(this.btnVer);
			this.KeyPreview = true;
			this.Name = "frmCxCReporteAnticipos";
			this.Text = "Reporte De Anticipos - Abonos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCxCReporteAnticipos_KeyDown);
			this.Load += new System.EventHandler(this.frmCxCReporteAnticipos_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgEstado.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCxCReporteAnticipos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "081211");

			if (!Funcion.Permiso("937", cdsSeteoF))
			{		
		
				MessageBox.Show("No tiene Acceso ANTICIPOS-ABONOS", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;

			}

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;

			this.cmbTipos.DataSource = FuncionesProcedimientos.dtGeneral("Select Tipo, Descripcion From TiposDeAnticipos Where Activo = 1");
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int iTipo = 0;
			if (this.cmbTipos.ActiveRow != null) iTipo = (int)this.cmbTipos.Value;

			string sSQL = string.Format("Exec CXCReporteAnticipos '{0}', '{1}', {2}, {3}, '{4}'", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), iTipo, (int)this.rdgEstado.EditValue, this.txtBuscar.Text.ToString());
			this.uGridDocumentos.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			this.lblContador.Text = this.uGridDocumentos.Rows.Count + " DOCUMENTOS";
		}

		private void uGridDocumentos_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (Convert.ToDecimal(e.Row.Cells["Saldo"].Value) < 0)
			{
				e.Row.Appearance.BackColor = Color.Red;			
				e.Row.Appearance.BackColor2 = Color.Red;
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDocumentos);
		}

		private void btnIrA_Click(object sender, System.EventArgs e)
		{
			this.mnuIrA.Show(this.btnIrA, new Point(1, 20));
		}

		private void mnuIrADocumento_Click(object sender, System.EventArgs e)
		{
			if (this.uGridDocumentos.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una fila", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			frmAnticiposClientes Anticipos = new frmAnticiposClientes((int)this.uGridDocumentos.ActiveRow.Cells["idAnticipo"].Value);
			Anticipos.MdiParent = MdiParent;
			Anticipos.Show();
		}

		private void mnuIrAEstadoDeCuenta_Click(object sender, System.EventArgs e)
		{
			if (this.uGridDocumentos.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una fila", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			frmCuentasPorCobrar CXC = new frmCuentasPorCobrar(this.uGridDocumentos.ActiveRow.Cells["Ruc"].Value.ToString());
			CXC.MdiParent = MdiParent;
			CXC.Show();
		}

		private void frmCxCReporteAnticipos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
		  Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
	}
}
