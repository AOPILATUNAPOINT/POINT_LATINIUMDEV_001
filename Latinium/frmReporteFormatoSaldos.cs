using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteFormatoSaldos.
	/// </summary>
	public class frmReporteFormatoSaldos : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid UgridDetalle;
		private System.Windows.Forms.Label lblEstado;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private System.Windows.Forms.Button btnConsultar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Label label17;
		public Infragistics.Win.UltraWinGrid.UltraGrid UgriCXP;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinGrid.UltraGrid UgriNVENTARIO;
		private System.Windows.Forms.ContextMenu mnuBI;
		private System.Windows.Forms.MenuItem mnuBIPedidos;
		private System.Windows.Forms.MenuItem mnuBICotizaciones;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.Button btnExcel;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteFormatoSaldos()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoMayor");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoModulo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoMayor", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoMayor", 2, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoModulo", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoModulo", 3, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Diferencia", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Diferencia", 4, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoMayor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoModulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencia");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteFormatoSaldos));
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoMayor");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoModulo");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoMayor", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoMayor", 2, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoModulo", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoModulo", 3, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Diferencia", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Diferencia", 4, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoMayor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoModulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencia");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoMayor");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoModulo");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoMayor", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoMayor", 2, true);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoModulo", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoModulo", 3, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Diferencia", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Diferencia", 4, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			this.UgridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label17 = new System.Windows.Forms.Label();
			this.UgriCXP = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.UgriNVENTARIO = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.mnuBI = new System.Windows.Forms.ContextMenu();
			this.mnuBIPedidos = new System.Windows.Forms.MenuItem();
			this.mnuBICotizaciones = new System.Windows.Forms.MenuItem();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.btnExcel = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.UgridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.UgriCXP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.UgriNVENTARIO)).BeginInit();
			this.SuspendLayout();
			// 
			// UgridDetalle
			// 
			this.UgridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.UgridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.UgridDetalle.DataSource = this.ultraDataSource8;
			appearance1.BackColor = System.Drawing.Color.White;
			this.UgridDetalle.DisplayLayout.Appearance = appearance1;
			this.UgridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 222;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 217;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance2;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.Caption = "Saldo Mayor";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 175;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance3;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "Saldo Modulo";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 175;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance4;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 178;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance5;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance6;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance7;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand1.SummaryFooterCaption = "";
			this.UgridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.UgridDetalle.DisplayLayout.GroupByBox.Hidden = true;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.UgridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.UgridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.UgridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.UgridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.UgridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.UgridDetalle.DisplayLayout.Override.CellAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.UgridDetalle.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.UgridDetalle.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UgridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UgridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UgridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.UgridDetalle.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.UgridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.UgridDetalle.Location = new System.Drawing.Point(8, 40);
			this.UgridDetalle.Name = "UgridDetalle";
			this.UgridDetalle.Size = new System.Drawing.Size(988, 136);
			this.UgridDetalle.TabIndex = 772;
			this.UgridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.UgridDetalle_InitializeLayout);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(448, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(144, 22);
			this.lblEstado.TabIndex = 868;
			this.lblEstado.Text = "CARTERA POINT";
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
			// btnConsultar
			// 
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(216, 8);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(79, 24);
			this.btnConsultar.TabIndex = 869;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// dtHasta
			// 
			appearance15.FontData.Name = "Tahoma";
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance15;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(64, 10);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(144, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 870;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(8, 12);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(34, 16);
			this.label17.TabIndex = 871;
			this.label17.Text = "Hasta";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// UgriCXP
			// 
			this.UgriCXP.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.UgriCXP.Cursor = System.Windows.Forms.Cursors.Default;
			this.UgriCXP.DataSource = this.ultraDataSource1;
			appearance16.BackColor = System.Drawing.Color.White;
			this.UgriCXP.DisplayLayout.Appearance = appearance16;
			this.UgriCXP.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Width = 222;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 217;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance17;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Saldo Mayor";
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Width = 175;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance18;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Saldo Modulo";
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Width = 175;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance19;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 178;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance20;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance21;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance22;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6});
			ultraGridBand2.SummaryFooterCaption = "";
			this.UgriCXP.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.UgriCXP.DisplayLayout.GroupByBox.Hidden = true;
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.UgriCXP.DisplayLayout.Override.ActiveRowAppearance = appearance23;
			this.UgriCXP.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.UgriCXP.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.UgriCXP.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance24.BackColor = System.Drawing.Color.Transparent;
			this.UgriCXP.DisplayLayout.Override.CardAreaAppearance = appearance24;
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.UgriCXP.DisplayLayout.Override.CellAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 8F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.UgriCXP.DisplayLayout.Override.HeaderAppearance = appearance26;
			this.UgriCXP.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance27.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance27.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UgriCXP.DisplayLayout.Override.RowAlternateAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UgriCXP.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UgriCXP.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.UgriCXP.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.UgriCXP.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.UgriCXP.Location = new System.Drawing.Point(8, 216);
			this.UgriCXP.Name = "UgriCXP";
			this.UgriCXP.Size = new System.Drawing.Size(988, 144);
			this.UgriCXP.TabIndex = 872;
			this.UgriCXP.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.UgriCXP_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(456, 184);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(137, 22);
			this.label1.TabIndex = 873;
			this.label1.Text = "PROVEEDORES";
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.Firebrick;
			this.label2.Location = new System.Drawing.Point(456, 368);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(110, 22);
			this.label2.TabIndex = 875;
			this.label2.Text = "INVENTARIO";
			// 
			// UgriNVENTARIO
			// 
			this.UgriNVENTARIO.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.UgriNVENTARIO.Cursor = System.Windows.Forms.Cursors.Default;
			this.UgriNVENTARIO.DataSource = this.ultraDataSource1;
			appearance30.BackColor = System.Drawing.Color.White;
			this.UgriNVENTARIO.DisplayLayout.Appearance = appearance30;
			this.UgriNVENTARIO.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Width = 222;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 217;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance31;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Saldo Mayor";
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Width = 175;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance32;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Saldo Modulo";
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn14.Width = 175;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance33;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Width = 178;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance34;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance35;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance36;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings7,
																																															summarySettings8,
																																															summarySettings9});
			ultraGridBand3.SummaryFooterCaption = "";
			this.UgriNVENTARIO.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.UgriNVENTARIO.DisplayLayout.GroupByBox.Hidden = true;
			appearance37.ForeColor = System.Drawing.Color.Black;
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.UgriNVENTARIO.DisplayLayout.Override.ActiveRowAppearance = appearance37;
			this.UgriNVENTARIO.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.UgriNVENTARIO.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.UgriNVENTARIO.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance38.BackColor = System.Drawing.Color.Transparent;
			this.UgriNVENTARIO.DisplayLayout.Override.CardAreaAppearance = appearance38;
			appearance39.ForeColor = System.Drawing.Color.Black;
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.UgriNVENTARIO.DisplayLayout.Override.CellAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance40.FontData.BoldAsString = "True";
			appearance40.FontData.Name = "Arial";
			appearance40.FontData.SizeInPoints = 8F;
			appearance40.ForeColor = System.Drawing.Color.White;
			appearance40.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.UgriNVENTARIO.DisplayLayout.Override.HeaderAppearance = appearance40;
			this.UgriNVENTARIO.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance41.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance41.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UgriNVENTARIO.DisplayLayout.Override.RowAlternateAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UgriNVENTARIO.DisplayLayout.Override.RowSelectorAppearance = appearance42;
			appearance43.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance43.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UgriNVENTARIO.DisplayLayout.Override.SelectedRowAppearance = appearance43;
			this.UgriNVENTARIO.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.UgriNVENTARIO.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.UgriNVENTARIO.Location = new System.Drawing.Point(8, 400);
			this.UgriNVENTARIO.Name = "UgriNVENTARIO";
			this.UgriNVENTARIO.Size = new System.Drawing.Size(988, 152);
			this.UgriNVENTARIO.TabIndex = 874;
			// 
			// mnuBI
			// 
			this.mnuBI.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																					this.mnuBIPedidos,
																																					this.mnuBICotizaciones,
																																					this.menuItem1});
			// 
			// mnuBIPedidos
			// 
			this.mnuBIPedidos.Index = 0;
			this.mnuBIPedidos.Text = "Cartera POINT";
			this.mnuBIPedidos.Click += new System.EventHandler(this.mnuBIPedidos_Click);
			// 
			// mnuBICotizaciones
			// 
			this.mnuBICotizaciones.Index = 1;
			this.mnuBICotizaciones.Text = "Proveedores";
			this.mnuBICotizaciones.Click += new System.EventHandler(this.mnuBICotizaciones_Click);
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 2;
			this.menuItem1.Text = "Inventario";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(920, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 994;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// frmReporteFormatoSaldos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1032, 566);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.UgriNVENTARIO);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.UgriCXP);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.UgridDetalle);
			this.KeyPreview = true;
			this.Name = "frmReporteFormatoSaldos";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Formato Saldos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteFormatoSaldos_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteFormatoSaldos_Load);
			((System.ComponentModel.ISupportInitialize)(this.UgridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.UgriCXP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.UgriNVENTARIO)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmReporteFormatoSaldos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200220");

			if (!Funcion.Permiso("2033", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a FORMATO SALDOS", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
	    if (miAcceso.BExportar) this.btnExcel.Visible =  true;

			DateTime dtFecha = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);

			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2019");
			this.dtHasta.CalendarInfo.MaxDate = dtFecha.AddDays(-1);
			this.dtHasta.Value = dtFecha.AddDays(-1);

		
		}



		private void UgridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "SaldoMayor", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SaldoModulo", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Diferencia", 2);
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione La Fecha De Corte", false, "", cdsSeteoF)) return;

			decimal SaldosPoint = 0.00m;
			decimal SaldosRefinanciadas = 0.00m;
			decimal SaldosGarantia = 0.00m;
			decimal SaldosRescates = 0.00m;
			decimal SaldosFDG = 0.00m;
			decimal SaldosCorporativo = 0.00m;
			
			decimal dANTICIPOS = 0.00m;
			decimal dPROVEEDORESMERCADERIA	= 0.00m;
			decimal dPROVEEDORESSERVICIOS	= 0.00m;
			decimal dPROVEEDORESUSA	 = 0.00m;
			decimal dPROVEEDORESASIA = 0.00m;

			decimal dInventario = 0.00m;

			string sSQL = string.Format("Exec [ReporteFormatoSaldos] '{0}', '{1}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneralModulo (sSQL, this.UgridDetalle);

			string sSQLCxP = string.Format("Exec [ReporteFormatoProveedores] '{0}', '{1}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneralModulo (sSQLCxP, this.UgriCXP);

			string sSQLiNV= string.Format("Exec [ReporteFormatoInventario] '{0}', '{1}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneralModulo (sSQLiNV, this.UgriNVENTARIO);

			SaldosPoint = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ReporteFormatoSaldosPoint '{0}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")));
			SaldosRefinanciadas = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ReporteFormatoSaldosRefinanciadas '{0}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")));
			SaldosGarantia = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ReporteFormatoSaldosGarantia '{0}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")));
			SaldosRescates = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ReporteFormatoSaldosRescates '{0}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")));
			SaldosFDG = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ReporteFormatoSaldosFDG '{0}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")));
			SaldosCorporativo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ReporteFormatoSaldosCorporativo '{0}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")));
	   
			this.ReporteFormatoSaldos ( SaldosPoint,  SaldosRefinanciadas,  SaldosGarantia ,  SaldosRescates,  SaldosFDG,   SaldosCorporativo );

			string sSQLGru = string.Format("Exec [ReporteFormatoProveedoresGrupos] '{0}', 1, 0, -1, 0", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLGru, true);
			dr.Read();
			if (dr.HasRows)
			{
				if (dr.GetValue(0) != System.DBNull.Value) dANTICIPOS = dr.GetDecimal(0);
				if (dr.GetValue(1) != System.DBNull.Value) dPROVEEDORESMERCADERIA = dr.GetDecimal(1);
				if (dr.GetValue(2) != System.DBNull.Value) dPROVEEDORESSERVICIOS = dr.GetDecimal(2);
				if (dr.GetValue(3) != System.DBNull.Value) dPROVEEDORESUSA = dr.GetDecimal(3); 
				if (dr.GetValue(4) != System.DBNull.Value) dPROVEEDORESASIA= dr.GetDecimal(4); 
			}
			dr.Close();


			string sSQLIn = string.Format("Exec [ReporteFormatoInventarioTotal] '{0}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			SqlDataReader dri = Funcion.rEscalarSQL(cdsSeteoF, sSQLIn, true);
			dri.Read();
			if (dri.HasRows)
			{
				if (dri.GetValue(0) != System.DBNull.Value) dInventario = dri.GetDecimal(0);

			}
			dri.Close();


			this.ReporteFormatoCxP ( dANTICIPOS,  dPROVEEDORESMERCADERIA,  dPROVEEDORESSERVICIOS ,  dPROVEEDORESUSA,  dPROVEEDORESASIA );
			this.ReporteFormatoInventario ( dInventario );

		}

		private void ReporteFormatoInventario (decimal SaldosPoint )
		{ 


			decimal DSaldoMayor = 0.00m;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dri in this.UgriNVENTARIO.Rows.All)
			{
				if ( Convert.ToString( dri.Cells["Cuenta"].Value) == "INVENTARIO DE MERCADERIAS") dri.Cells["SaldoModulo"].Value =  Math.Round(SaldosPoint,2);    
   
				DSaldoMayor = Convert.ToDecimal( dri.Cells["SaldoMayor"].Value ) + DSaldoMayor ;
				
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dri in this.UgriNVENTARIO.Rows.All)
			{
			if ( Convert.ToString( dri.Cells["Cuenta"].Value) == "INVENTARIO DE MERCADERIAS") dri.Cells["Diferencia"].Value =  DSaldoMayor - Convert.ToDecimal( dri.Cells["SaldoModulo"].Value )  ;

			}
		}


		private void ReporteFormatoSaldos (decimal SaldosPoint, decimal SaldosRefinanciadas, decimal SaldosGarantia , decimal SaldosRescates, decimal SaldosFDG, decimal  SaldosCorporativo )
		{ 

			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.UgridDetalle.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Cuenta"].Value) == "CLIENTES CREDIPOINT") dr.Cells["SaldoModulo"].Value =  Math.Round(SaldosPoint,2);
				if ( Convert.ToString( dr.Cells["Cuenta"].Value) == "CLIENTES REFINANCIADOS") dr.Cells["SaldoModulo"].Value =  Math.Round(SaldosRefinanciadas,2);
				if ( Convert.ToString( dr.Cells["Cuenta"].Value) == "CLIENTES GARANTIA") dr.Cells["SaldoModulo"].Value =  Math.Round(SaldosGarantia,2);
				if ( Convert.ToString( dr.Cells["Cuenta"].Value) == "RESCATES CARTERA PICHINCHA") dr.Cells["SaldoModulo"].Value =  Math.Round(SaldosRescates,2);
				if ( Convert.ToString( dr.Cells["Cuenta"].Value) == "CLIENTES FDG GARANTIA PICHINCHA") dr.Cells["SaldoModulo"].Value =  Math.Round(SaldosFDG,2);
				if ( Convert.ToString( dr.Cells["Cuenta"].Value) == "CLIENTES CORPORATIVOS") dr.Cells["SaldoModulo"].Value =  Math.Round(SaldosCorporativo,2);
        
				dr.Cells["Diferencia"].Value = Convert.ToDecimal( dr.Cells["SaldoMayor"].Value) - Convert.ToDecimal( dr.Cells["SaldoModulo"].Value );
        

			}
		}

		private void ReporteFormatoCxP (decimal dANTICIPOS, decimal dPROVEEDORESMERCADERIA, decimal dPROVEEDORESSERVICIOS , decimal dPROVEEDORESUSA, decimal dPROVEEDORESASIA )
		{ 

			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.UgriCXP.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Cuenta"].Value) == "ANTICIPOS PROVEEDORES") dr.Cells["SaldoModulo"].Value =  Math.Round(dANTICIPOS,2);
				if ( Convert.ToString( dr.Cells["Cuenta"].Value) == "PROVEEDORES MERCADERIA") dr.Cells["SaldoModulo"].Value =  Math.Round(dPROVEEDORESMERCADERIA,2);
				if ( Convert.ToString( dr.Cells["Cuenta"].Value) == "PROVEEDORES SERVICIOS") dr.Cells["SaldoModulo"].Value =  Math.Round(dPROVEEDORESSERVICIOS,2);
				if ( Convert.ToString( dr.Cells["Cuenta"].Value) == "PROVEEDORES USA") dr.Cells["SaldoModulo"].Value =  Math.Round(dPROVEEDORESUSA,2);
				if ( Convert.ToString( dr.Cells["Cuenta"].Value) == "PROVEEDORES ASIA") dr.Cells["SaldoModulo"].Value =  Math.Round(dPROVEEDORESASIA,2);

				dr.Cells["Diferencia"].Value = Convert.ToDecimal( dr.Cells["SaldoMayor"].Value) - Convert.ToDecimal( dr.Cells["SaldoModulo"].Value );
        

			}
		}


		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
	
		}

		private void UgriCXP_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "SaldoMayor", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SaldoModulo", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Diferencia", 2);
		}

		private void frmReporteFormatoSaldos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void mnuBIPedidos_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.UgridDetalle);
		}

		private void mnuBICotizaciones_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.UgriCXP);
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.UgriNVENTARIO);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
		this.mnuBI.Show(this.btnExcel, new Point(1, 20));
		}
	}
}
