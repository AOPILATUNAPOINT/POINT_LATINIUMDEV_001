using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReportesAsientos.
	/// </summary>
	public class frmReportesAsientos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpDesde;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridReporteAsientos;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpHasta;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalleAsiento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private System.Windows.Forms.Button btnAnularCheques;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReportesAsientos()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReportesAsientos));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Documento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioAutoriza");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAutorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasAutorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAutorizar", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 9, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Documento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioAutoriza");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAutorizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasAutorizacion");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Centro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debe", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debe", 4, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Haber", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Haber", 5, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Centro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Haber");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			this.uGridReporteAsientos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtpDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.dtpHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.uGridDetalleAsiento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnAnularCheques = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridReporteAsientos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridReporteAsientos
			// 
			this.uGridReporteAsientos.AccessibleDescription = resources.GetString("uGridReporteAsientos.AccessibleDescription");
			this.uGridReporteAsientos.AccessibleName = resources.GetString("uGridReporteAsientos.AccessibleName");
			this.uGridReporteAsientos.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("uGridReporteAsientos.Anchor")));
			this.uGridReporteAsientos.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("uGridReporteAsientos.BackgroundImage")));
			this.uGridReporteAsientos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridReporteAsientos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridReporteAsientos.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = resources.GetString("ultraGridBand1.AddButtonCaption");
			ultraGridBand1.AddButtonToolTipText = resources.GetString("ultraGridBand1.AddButtonToolTipText");
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Format = resources.GetString("ultraGridColumn1.Format");
			ultraGridColumn1.Header.Caption = resources.GetString("resource.Caption");
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.MaskInput = resources.GetString("ultraGridColumn1.MaskInput");
			ultraGridColumn1.NullText = resources.GetString("ultraGridColumn1.NullText");
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Format = resources.GetString("ultraGridColumn2.Format");
			ultraGridColumn2.Header.Caption = resources.GetString("resource.Caption1");
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.MaskInput = resources.GetString("ultraGridColumn2.MaskInput");
			ultraGridColumn2.NullText = resources.GetString("ultraGridColumn2.NullText");
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Format = resources.GetString("ultraGridColumn3.Format");
			ultraGridColumn3.Header.Caption = resources.GetString("resource.Caption2");
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.MaskInput = resources.GetString("ultraGridColumn3.MaskInput");
			ultraGridColumn3.NullText = resources.GetString("ultraGridColumn3.NullText");
			ultraGridColumn3.Width = 120;
			ultraGridColumn4.Format = resources.GetString("ultraGridColumn4.Format");
			ultraGridColumn4.Header.Caption = resources.GetString("resource.Caption3");
			ultraGridColumn4.Header.VisiblePosition = 16;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.MaskInput = resources.GetString("ultraGridColumn4.MaskInput");
			ultraGridColumn4.NullText = resources.GetString("ultraGridColumn4.NullText");
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Format = resources.GetString("ultraGridColumn5.Format");
			ultraGridColumn5.Header.Caption = resources.GetString("resource.Caption4");
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.MaskInput = resources.GetString("ultraGridColumn5.MaskInput");
			ultraGridColumn5.NullText = resources.GetString("ultraGridColumn5.NullText");
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Format = resources.GetString("ultraGridColumn6.Format");
			ultraGridColumn6.Header.Caption = resources.GetString("resource.Caption5");
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.MaskInput = resources.GetString("ultraGridColumn6.MaskInput");
			ultraGridColumn6.NullText = resources.GetString("ultraGridColumn6.NullText");
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Format = resources.GetString("ultraGridColumn7.Format");
			ultraGridColumn7.Header.Caption = resources.GetString("resource.Caption6");
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.MaskInput = resources.GetString("ultraGridColumn7.MaskInput");
			ultraGridColumn7.NullText = resources.GetString("ultraGridColumn7.NullText");
			ultraGridColumn7.Width = 90;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Format = resources.GetString("ultraGridColumn8.Format");
			ultraGridColumn8.Header.Caption = resources.GetString("resource.Caption7");
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.MaskInput = resources.GetString("ultraGridColumn8.MaskInput");
			ultraGridColumn8.NullText = resources.GetString("ultraGridColumn8.NullText");
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Format = resources.GetString("ultraGridColumn9.Format");
			ultraGridColumn9.Header.Caption = resources.GetString("resource.Caption8");
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.MaskInput = resources.GetString("ultraGridColumn9.MaskInput");
			ultraGridColumn9.NullText = resources.GetString("ultraGridColumn9.NullText");
			ultraGridColumn9.Width = 300;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance2;
			ultraGridColumn10.Format = resources.GetString("ultraGridColumn10.Format");
			ultraGridColumn10.Header.Caption = resources.GetString("resource.Caption9");
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.MaskInput = resources.GetString("ultraGridColumn10.MaskInput");
			ultraGridColumn10.NullText = resources.GetString("ultraGridColumn10.NullText");
			ultraGridColumn10.Width = 90;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Format = resources.GetString("ultraGridColumn11.Format");
			ultraGridColumn11.Header.Caption = resources.GetString("resource.Caption10");
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.MaskInput = resources.GetString("ultraGridColumn11.MaskInput");
			ultraGridColumn11.NullText = resources.GetString("ultraGridColumn11.NullText");
			ultraGridColumn11.Width = 120;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Format = resources.GetString("ultraGridColumn12.Format");
			ultraGridColumn12.Header.Caption = resources.GetString("resource.Caption11");
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.MaskInput = resources.GetString("ultraGridColumn12.MaskInput");
			ultraGridColumn12.NullText = resources.GetString("ultraGridColumn12.NullText");
			ultraGridColumn12.Width = 110;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Format = resources.GetString("ultraGridColumn13.Format");
			ultraGridColumn13.Header.Caption = resources.GetString("resource.Caption12");
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.MaskInput = resources.GetString("ultraGridColumn13.MaskInput");
			ultraGridColumn13.NullText = resources.GetString("ultraGridColumn13.NullText");
			ultraGridColumn13.Width = 400;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Format = resources.GetString("ultraGridColumn14.Format");
			ultraGridColumn14.Header.Caption = resources.GetString("resource.Caption13");
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.MaskInput = resources.GetString("ultraGridColumn14.MaskInput");
			ultraGridColumn14.NullText = resources.GetString("ultraGridColumn14.NullText");
			ultraGridColumn14.Width = 120;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Format = resources.GetString("ultraGridColumn15.Format");
			ultraGridColumn15.Header.Caption = resources.GetString("resource.Caption14");
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.MaskInput = resources.GetString("ultraGridColumn15.MaskInput");
			ultraGridColumn15.NullText = resources.GetString("ultraGridColumn15.NullText");
			ultraGridColumn15.Width = 140;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Format = resources.GetString("ultraGridColumn16.Format");
			ultraGridColumn16.Header.Caption = resources.GetString("resource.Caption15");
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.MaskInput = resources.GetString("ultraGridColumn16.MaskInput");
			ultraGridColumn16.NullText = resources.GetString("ultraGridColumn16.NullText");
			ultraGridColumn16.Width = 300;
			ultraGridColumn17.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			appearance3.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance3.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn17.CellButtonAppearance = appearance3;
			ultraGridColumn17.Format = resources.GetString("ultraGridColumn17.Format");
			ultraGridColumn17.Header.Caption = resources.GetString("resource.Caption16");
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.MaskInput = resources.GetString("ultraGridColumn17.MaskInput");
			ultraGridColumn17.NullText = resources.GetString("ultraGridColumn17.NullText");
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn17.Width = 40;
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
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			appearance4.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance4;
			ultraGridBand1.Header.Caption = resources.GetString("resource.Caption17");
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.GroupByRowDescriptionMask = resources.GetString("resource.GroupByRowDescriptionMask");
			ultraGridBand1.Override.NullText = resources.GetString("resource.NullText");
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance5;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = resources.GetString("ultraGridBand1.SummaryFooterCaption");
			this.uGridReporteAsientos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridReporteAsientos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridReporteAsientos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridReporteAsientos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridReporteAsientos.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridReporteAsientos.DisplayLayout.Override.CellAppearance = appearance7;
			this.uGridReporteAsientos.DisplayLayout.Override.GroupByRowDescriptionMask = resources.GetString("resource.GroupByRowDescriptionMask1");
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = resources.GetString("resource.BoldAsString");
			appearance8.FontData.ItalicAsString = resources.GetString("resource.ItalicAsString");
			appearance8.FontData.Name = resources.GetString("resource.Name");
			appearance8.FontData.SizeInPoints = ((System.Single)(resources.GetObject("resource.SizeInPoints")));
			appearance8.FontData.StrikeoutAsString = resources.GetString("resource.StrikeoutAsString");
			appearance8.FontData.UnderlineAsString = resources.GetString("resource.UnderlineAsString");
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridReporteAsientos.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.uGridReporteAsientos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.uGridReporteAsientos.DisplayLayout.Override.NullText = resources.GetString("resource.NullText1");
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridReporteAsientos.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridReporteAsientos.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridReporteAsientos.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("uGridReporteAsientos.Dock")));
			this.uGridReporteAsientos.Enabled = ((bool)(resources.GetObject("uGridReporteAsientos.Enabled")));
			this.uGridReporteAsientos.Font = ((System.Drawing.Font)(resources.GetObject("uGridReporteAsientos.Font")));
			this.uGridReporteAsientos.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("uGridReporteAsientos.ImeMode")));
			this.uGridReporteAsientos.Location = ((System.Drawing.Point)(resources.GetObject("uGridReporteAsientos.Location")));
			this.uGridReporteAsientos.Name = "uGridReporteAsientos";
			this.uGridReporteAsientos.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("uGridReporteAsientos.RightToLeft")));
			this.uGridReporteAsientos.Size = ((System.Drawing.Size)(resources.GetObject("uGridReporteAsientos.Size")));
			this.uGridReporteAsientos.TabIndex = ((int)(resources.GetObject("uGridReporteAsientos.TabIndex")));
			this.uGridReporteAsientos.Text = resources.GetString("uGridReporteAsientos.Text");
			this.uGridReporteAsientos.Visible = ((bool)(resources.GetObject("uGridReporteAsientos.Visible")));
			this.uGridReporteAsientos.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridReporteAsientos_BeforeSelectChange);
			this.uGridReporteAsientos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridReporteAsientos_ClickCellButton);
			this.uGridReporteAsientos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridReporteAsientos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.DateTime);
			ultraDataColumn15.DataType = typeof(System.DateTime);
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
			// label3
			// 
			this.label3.AccessibleDescription = resources.GetString("label3.AccessibleDescription");
			this.label3.AccessibleName = resources.GetString("label3.AccessibleName");
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("label3.Anchor")));
			this.label3.AutoSize = ((bool)(resources.GetObject("label3.AutoSize")));
			this.label3.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("label3.Dock")));
			this.label3.Enabled = ((bool)(resources.GetObject("label3.Enabled")));
			this.label3.Font = ((System.Drawing.Font)(resources.GetObject("label3.Font")));
			this.label3.Image = ((System.Drawing.Image)(resources.GetObject("label3.Image")));
			this.label3.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label3.ImageAlign")));
			this.label3.ImageIndex = ((int)(resources.GetObject("label3.ImageIndex")));
			this.label3.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("label3.ImeMode")));
			this.label3.Location = ((System.Drawing.Point)(resources.GetObject("label3.Location")));
			this.label3.Name = "label3";
			this.label3.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("label3.RightToLeft")));
			this.label3.Size = ((System.Drawing.Size)(resources.GetObject("label3.Size")));
			this.label3.TabIndex = ((int)(resources.GetObject("label3.TabIndex")));
			this.label3.Text = resources.GetString("label3.Text");
			this.label3.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label3.TextAlign")));
			this.label3.Visible = ((bool)(resources.GetObject("label3.Visible")));
			// 
			// label6
			// 
			this.label6.AccessibleDescription = resources.GetString("label6.AccessibleDescription");
			this.label6.AccessibleName = resources.GetString("label6.AccessibleName");
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("label6.Anchor")));
			this.label6.AutoSize = ((bool)(resources.GetObject("label6.AutoSize")));
			this.label6.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("label6.Dock")));
			this.label6.Enabled = ((bool)(resources.GetObject("label6.Enabled")));
			this.label6.Font = ((System.Drawing.Font)(resources.GetObject("label6.Font")));
			this.label6.Image = ((System.Drawing.Image)(resources.GetObject("label6.Image")));
			this.label6.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label6.ImageAlign")));
			this.label6.ImageIndex = ((int)(resources.GetObject("label6.ImageIndex")));
			this.label6.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("label6.ImeMode")));
			this.label6.Location = ((System.Drawing.Point)(resources.GetObject("label6.Location")));
			this.label6.Name = "label6";
			this.label6.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("label6.RightToLeft")));
			this.label6.Size = ((System.Drawing.Size)(resources.GetObject("label6.Size")));
			this.label6.TabIndex = ((int)(resources.GetObject("label6.TabIndex")));
			this.label6.Text = resources.GetString("label6.Text");
			this.label6.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label6.TextAlign")));
			this.label6.Visible = ((bool)(resources.GetObject("label6.Visible")));
			// 
			// dtpDesde
			// 
			this.dtpDesde.AccessibleDescription = resources.GetString("dtpDesde.AccessibleDescription");
			this.dtpDesde.AccessibleName = resources.GetString("dtpDesde.AccessibleName");
			this.dtpDesde.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("dtpDesde.Anchor")));
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpDesde.Appearance = appearance11;
			this.dtpDesde.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("dtpDesde.BackgroundImage")));
			this.dtpDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpDesde.DateButtons.Add(dateButton1);
			this.dtpDesde.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("dtpDesde.Dock")));
			this.dtpDesde.Enabled = ((bool)(resources.GetObject("dtpDesde.Enabled")));
			this.dtpDesde.Font = ((System.Drawing.Font)(resources.GetObject("dtpDesde.Font")));
			this.dtpDesde.Format = "dd/MM/yyyy";
			this.dtpDesde.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("dtpDesde.ImeMode")));
			this.dtpDesde.Location = ((System.Drawing.Point)(resources.GetObject("dtpDesde.Location")));
			this.dtpDesde.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpDesde.Name = "dtpDesde";
			this.dtpDesde.NonAutoSizeHeight = 23;
			this.dtpDesde.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("dtpDesde.RightToLeft")));
			this.dtpDesde.Size = ((System.Drawing.Size)(resources.GetObject("dtpDesde.Size")));
			this.dtpDesde.SpinButtonsVisible = true;
			this.dtpDesde.TabIndex = ((int)(resources.GetObject("dtpDesde.TabIndex")));
			this.dtpDesde.Value = ((object)(resources.GetObject("dtpDesde.Value")));
			this.dtpDesde.Visible = ((bool)(resources.GetObject("dtpDesde.Visible")));
			// 
			// lblEstado
			// 
			this.lblEstado.AccessibleDescription = resources.GetString("lblEstado.AccessibleDescription");
			this.lblEstado.AccessibleName = resources.GetString("lblEstado.AccessibleName");
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("lblEstado.Anchor")));
			this.lblEstado.AutoSize = ((bool)(resources.GetObject("lblEstado.AutoSize")));
			this.lblEstado.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("lblEstado.Dock")));
			this.lblEstado.Enabled = ((bool)(resources.GetObject("lblEstado.Enabled")));
			this.lblEstado.Font = ((System.Drawing.Font)(resources.GetObject("lblEstado.Font")));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Image = ((System.Drawing.Image)(resources.GetObject("lblEstado.Image")));
			this.lblEstado.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("lblEstado.ImageAlign")));
			this.lblEstado.ImageIndex = ((int)(resources.GetObject("lblEstado.ImageIndex")));
			this.lblEstado.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("lblEstado.ImeMode")));
			this.lblEstado.Location = ((System.Drawing.Point)(resources.GetObject("lblEstado.Location")));
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("lblEstado.RightToLeft")));
			this.lblEstado.Size = ((System.Drawing.Size)(resources.GetObject("lblEstado.Size")));
			this.lblEstado.TabIndex = ((int)(resources.GetObject("lblEstado.TabIndex")));
			this.lblEstado.Text = resources.GetString("lblEstado.Text");
			this.lblEstado.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("lblEstado.TextAlign")));
			this.lblEstado.Visible = ((bool)(resources.GetObject("lblEstado.Visible")));
			// 
			// btnVer
			// 
			this.btnVer.AccessibleDescription = resources.GetString("btnVer.AccessibleDescription");
			this.btnVer.AccessibleName = resources.GetString("btnVer.AccessibleName");
			this.btnVer.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnVer.Anchor")));
			this.btnVer.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnVer.BackgroundImage")));
			this.btnVer.CausesValidation = false;
			this.btnVer.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnVer.Dock")));
			this.btnVer.Enabled = ((bool)(resources.GetObject("btnVer.Enabled")));
			this.btnVer.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnVer.FlatStyle")));
			this.btnVer.Font = ((System.Drawing.Font)(resources.GetObject("btnVer.Font")));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnVer.ImageAlign")));
			this.btnVer.ImageIndex = ((int)(resources.GetObject("btnVer.ImageIndex")));
			this.btnVer.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnVer.ImeMode")));
			this.btnVer.Location = ((System.Drawing.Point)(resources.GetObject("btnVer.Location")));
			this.btnVer.Name = "btnVer";
			this.btnVer.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnVer.RightToLeft")));
			this.btnVer.Size = ((System.Drawing.Size)(resources.GetObject("btnVer.Size")));
			this.btnVer.TabIndex = ((int)(resources.GetObject("btnVer.TabIndex")));
			this.btnVer.Text = resources.GetString("btnVer.Text");
			this.btnVer.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnVer.TextAlign")));
			this.btnVer.Visible = ((bool)(resources.GetObject("btnVer.Visible")));
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.AccessibleDescription = resources.GetString("btnExportar.AccessibleDescription");
			this.btnExportar.AccessibleName = resources.GetString("btnExportar.AccessibleName");
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnExportar.Anchor")));
			this.btnExportar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnExportar.BackgroundImage")));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnExportar.Dock")));
			this.btnExportar.Enabled = ((bool)(resources.GetObject("btnExportar.Enabled")));
			this.btnExportar.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnExportar.FlatStyle")));
			this.btnExportar.Font = ((System.Drawing.Font)(resources.GetObject("btnExportar.Font")));
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnExportar.ImageAlign")));
			this.btnExportar.ImageIndex = ((int)(resources.GetObject("btnExportar.ImageIndex")));
			this.btnExportar.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnExportar.ImeMode")));
			this.btnExportar.Location = ((System.Drawing.Point)(resources.GetObject("btnExportar.Location")));
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnExportar.RightToLeft")));
			this.btnExportar.Size = ((System.Drawing.Size)(resources.GetObject("btnExportar.Size")));
			this.btnExportar.TabIndex = ((int)(resources.GetObject("btnExportar.TabIndex")));
			this.btnExportar.Text = resources.GetString("btnExportar.Text");
			this.btnExportar.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnExportar.TextAlign")));
			this.btnExportar.Visible = ((bool)(resources.GetObject("btnExportar.Visible")));
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// dtpHasta
			// 
			this.dtpHasta.AccessibleDescription = resources.GetString("dtpHasta.AccessibleDescription");
			this.dtpHasta.AccessibleName = resources.GetString("dtpHasta.AccessibleName");
			this.dtpHasta.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("dtpHasta.Anchor")));
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpHasta.Appearance = appearance12;
			this.dtpHasta.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("dtpHasta.BackgroundImage")));
			this.dtpHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtpHasta.DateButtons.Add(dateButton2);
			this.dtpHasta.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("dtpHasta.Dock")));
			this.dtpHasta.Enabled = ((bool)(resources.GetObject("dtpHasta.Enabled")));
			this.dtpHasta.Font = ((System.Drawing.Font)(resources.GetObject("dtpHasta.Font")));
			this.dtpHasta.Format = "dd/MM/yyyy";
			this.dtpHasta.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("dtpHasta.ImeMode")));
			this.dtpHasta.Location = ((System.Drawing.Point)(resources.GetObject("dtpHasta.Location")));
			this.dtpHasta.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpHasta.Name = "dtpHasta";
			this.dtpHasta.NonAutoSizeHeight = 23;
			this.dtpHasta.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("dtpHasta.RightToLeft")));
			this.dtpHasta.Size = ((System.Drawing.Size)(resources.GetObject("dtpHasta.Size")));
			this.dtpHasta.SpinButtonsVisible = true;
			this.dtpHasta.TabIndex = ((int)(resources.GetObject("dtpHasta.TabIndex")));
			this.dtpHasta.Value = ((object)(resources.GetObject("dtpHasta.Value")));
			this.dtpHasta.Visible = ((bool)(resources.GetObject("dtpHasta.Visible")));
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
			// uGridDetalleAsiento
			// 
			this.uGridDetalleAsiento.AccessibleDescription = resources.GetString("uGridDetalleAsiento.AccessibleDescription");
			this.uGridDetalleAsiento.AccessibleName = resources.GetString("uGridDetalleAsiento.AccessibleName");
			this.uGridDetalleAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("uGridDetalleAsiento.Anchor")));
			this.uGridDetalleAsiento.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("uGridDetalleAsiento.BackgroundImage")));
			this.uGridDetalleAsiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalleAsiento.DataSource = this.ultraDataSource2;
			appearance13.BackColor = System.Drawing.Color.White;
			this.uGridDetalleAsiento.DisplayLayout.Appearance = appearance13;
			this.uGridDetalleAsiento.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = resources.GetString("ultraGridBand2.AddButtonCaption");
			ultraGridBand2.AddButtonToolTipText = resources.GetString("ultraGridBand2.AddButtonToolTipText");
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Format = resources.GetString("ultraGridColumn18.Format");
			ultraGridColumn18.Header.Caption = resources.GetString("resource.Caption18");
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.MaskInput = resources.GetString("ultraGridColumn18.MaskInput");
			ultraGridColumn18.NullText = resources.GetString("ultraGridColumn18.NullText");
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn19.Format = resources.GetString("ultraGridColumn19.Format");
			ultraGridColumn19.Header.Caption = resources.GetString("resource.Caption19");
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.MaskInput = resources.GetString("ultraGridColumn19.MaskInput");
			ultraGridColumn19.NullText = resources.GetString("ultraGridColumn19.NullText");
			ultraGridColumn19.Width = 120;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn20.Format = resources.GetString("ultraGridColumn20.Format");
			ultraGridColumn20.Header.Caption = resources.GetString("resource.Caption20");
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridColumn20.MaskInput = resources.GetString("ultraGridColumn20.MaskInput");
			ultraGridColumn20.NullText = resources.GetString("ultraGridColumn20.NullText");
			ultraGridColumn20.Width = 200;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn21.Format = resources.GetString("ultraGridColumn21.Format");
			ultraGridColumn21.Header.Caption = resources.GetString("resource.Caption21");
			ultraGridColumn21.Header.VisiblePosition = 3;
			ultraGridColumn21.MaskInput = resources.GetString("ultraGridColumn21.MaskInput");
			ultraGridColumn21.NullText = resources.GetString("ultraGridColumn21.NullText");
			ultraGridColumn21.Width = 160;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance14;
			ultraGridColumn22.Format = resources.GetString("ultraGridColumn22.Format");
			ultraGridColumn22.Header.Caption = resources.GetString("resource.Caption22");
			ultraGridColumn22.Header.VisiblePosition = 4;
			ultraGridColumn22.MaskInput = resources.GetString("ultraGridColumn22.MaskInput");
			ultraGridColumn22.NullText = resources.GetString("ultraGridColumn22.NullText");
			ultraGridColumn22.Width = 100;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance15;
			ultraGridColumn23.Format = resources.GetString("ultraGridColumn23.Format");
			ultraGridColumn23.Header.Caption = resources.GetString("resource.Caption23");
			ultraGridColumn23.Header.VisiblePosition = 5;
			ultraGridColumn23.MaskInput = resources.GetString("ultraGridColumn23.MaskInput");
			ultraGridColumn23.NullText = resources.GetString("ultraGridColumn23.NullText");
			ultraGridColumn23.Width = 100;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn24.Format = resources.GetString("ultraGridColumn24.Format");
			ultraGridColumn24.Header.Caption = resources.GetString("resource.Caption24");
			ultraGridColumn24.Header.VisiblePosition = 6;
			ultraGridColumn24.MaskInput = resources.GetString("ultraGridColumn24.MaskInput");
			ultraGridColumn24.NullText = resources.GetString("ultraGridColumn24.NullText");
			ultraGridColumn24.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			appearance16.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridBand2.Header.Appearance = appearance16;
			ultraGridBand2.Header.Caption = resources.GetString("resource.Caption25");
			ultraGridBand2.HeaderVisible = true;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance17;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance18;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand2.SummaryFooterCaption = resources.GetString("ultraGridBand2.SummaryFooterCaption");
			this.uGridDetalleAsiento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridDetalleAsiento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDetalleAsiento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDetalleAsiento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance19.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalleAsiento.DisplayLayout.Override.CardAreaAppearance = appearance19;
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleAsiento.DisplayLayout.Override.CellAppearance = appearance20;
			this.uGridDetalleAsiento.DisplayLayout.Override.GroupByRowDescriptionMask = resources.GetString("resource.GroupByRowDescriptionMask2");
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance21.FontData.BoldAsString = resources.GetString("resource.BoldAsString1");
			appearance21.FontData.ItalicAsString = resources.GetString("resource.ItalicAsString1");
			appearance21.FontData.Name = resources.GetString("resource.Name1");
			appearance21.FontData.SizeInPoints = ((System.Single)(resources.GetObject("resource.SizeInPoints1")));
			appearance21.FontData.StrikeoutAsString = resources.GetString("resource.StrikeoutAsString1");
			appearance21.FontData.UnderlineAsString = resources.GetString("resource.UnderlineAsString1");
			appearance21.ForeColor = System.Drawing.Color.White;
			appearance21.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalleAsiento.DisplayLayout.Override.HeaderAppearance = appearance21;
			this.uGridDetalleAsiento.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.uGridDetalleAsiento.DisplayLayout.Override.NullText = resources.GetString("resource.NullText2");
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleAsiento.DisplayLayout.Override.RowSelectorAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleAsiento.DisplayLayout.Override.SelectedRowAppearance = appearance23;
			this.uGridDetalleAsiento.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("uGridDetalleAsiento.Dock")));
			this.uGridDetalleAsiento.Enabled = ((bool)(resources.GetObject("uGridDetalleAsiento.Enabled")));
			this.uGridDetalleAsiento.Font = ((System.Drawing.Font)(resources.GetObject("uGridDetalleAsiento.Font")));
			this.uGridDetalleAsiento.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("uGridDetalleAsiento.ImeMode")));
			this.uGridDetalleAsiento.Location = ((System.Drawing.Point)(resources.GetObject("uGridDetalleAsiento.Location")));
			this.uGridDetalleAsiento.Name = "uGridDetalleAsiento";
			this.uGridDetalleAsiento.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("uGridDetalleAsiento.RightToLeft")));
			this.uGridDetalleAsiento.Size = ((System.Drawing.Size)(resources.GetObject("uGridDetalleAsiento.Size")));
			this.uGridDetalleAsiento.TabIndex = ((int)(resources.GetObject("uGridDetalleAsiento.TabIndex")));
			this.uGridDetalleAsiento.Text = resources.GetString("uGridDetalleAsiento.Text");
			this.uGridDetalleAsiento.Visible = ((bool)(resources.GetObject("uGridDetalleAsiento.Visible")));
			// 
			// ultraDataSource2
			// 
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23});
			// 
			// optEstados
			// 
			this.optEstados.AccessibleDescription = resources.GetString("optEstados.AccessibleDescription");
			this.optEstados.AccessibleName = resources.GetString("optEstados.AccessibleName");
			this.optEstados.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("optEstados.Anchor")));
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance24;
			this.optEstados.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("optEstados.BackgroundImage")));
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 1;
			this.optEstados.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("optEstados.Dock")));
			this.optEstados.Enabled = ((bool)(resources.GetObject("optEstados.Enabled")));
			this.optEstados.Font = ((System.Drawing.Font)(resources.GetObject("optEstados.Font")));
			this.optEstados.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("optEstados.ImeMode")));
			this.optEstados.ItemAppearance = appearance25;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = resources.GetString("valueListItem1.DisplayText");
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = resources.GetString("valueListItem2.DisplayText");
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = resources.GetString("valueListItem3.DisplayText");
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = resources.GetString("valueListItem4.DisplayText");
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = ((System.Drawing.Point)(resources.GetObject("optEstados.Location")));
			this.optEstados.Name = "optEstados";
			this.optEstados.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("optEstados.RightToLeft")));
			this.optEstados.Size = ((System.Drawing.Size)(resources.GetObject("optEstados.Size")));
			this.optEstados.TabIndex = ((int)(resources.GetObject("optEstados.TabIndex")));
			this.optEstados.Text = resources.GetString("optEstados.Text");
			this.optEstados.Visible = ((bool)(resources.GetObject("optEstados.Visible")));
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// btnAnularCheques
			// 
			this.btnAnularCheques.AccessibleDescription = resources.GetString("btnAnularCheques.AccessibleDescription");
			this.btnAnularCheques.AccessibleName = resources.GetString("btnAnularCheques.AccessibleName");
			this.btnAnularCheques.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnAnularCheques.Anchor")));
			this.btnAnularCheques.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnAnularCheques.BackgroundImage")));
			this.btnAnularCheques.CausesValidation = false;
			this.btnAnularCheques.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnAnularCheques.Dock")));
			this.btnAnularCheques.Enabled = ((bool)(resources.GetObject("btnAnularCheques.Enabled")));
			this.btnAnularCheques.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnAnularCheques.FlatStyle")));
			this.btnAnularCheques.Font = ((System.Drawing.Font)(resources.GetObject("btnAnularCheques.Font")));
			this.btnAnularCheques.Image = ((System.Drawing.Image)(resources.GetObject("btnAnularCheques.Image")));
			this.btnAnularCheques.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnAnularCheques.ImageAlign")));
			this.btnAnularCheques.ImageIndex = ((int)(resources.GetObject("btnAnularCheques.ImageIndex")));
			this.btnAnularCheques.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnAnularCheques.ImeMode")));
			this.btnAnularCheques.Location = ((System.Drawing.Point)(resources.GetObject("btnAnularCheques.Location")));
			this.btnAnularCheques.Name = "btnAnularCheques";
			this.btnAnularCheques.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnAnularCheques.RightToLeft")));
			this.btnAnularCheques.Size = ((System.Drawing.Size)(resources.GetObject("btnAnularCheques.Size")));
			this.btnAnularCheques.TabIndex = ((int)(resources.GetObject("btnAnularCheques.TabIndex")));
			this.btnAnularCheques.Text = resources.GetString("btnAnularCheques.Text");
			this.btnAnularCheques.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnAnularCheques.TextAlign")));
			this.btnAnularCheques.Visible = ((bool)(resources.GetObject("btnAnularCheques.Visible")));
			this.btnAnularCheques.Click += new System.EventHandler(this.btnAnularCheques_Click);
			// 
			// frmReportesAsientos
			// 
			this.AccessibleDescription = resources.GetString("$this.AccessibleDescription");
			this.AccessibleName = resources.GetString("$this.AccessibleName");
			this.AutoScaleBaseSize = ((System.Drawing.Size)(resources.GetObject("$this.AutoScaleBaseSize")));
			this.AutoScroll = ((bool)(resources.GetObject("$this.AutoScroll")));
			this.AutoScrollMargin = ((System.Drawing.Size)(resources.GetObject("$this.AutoScrollMargin")));
			this.AutoScrollMinSize = ((System.Drawing.Size)(resources.GetObject("$this.AutoScrollMinSize")));
			this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
			this.ClientSize = ((System.Drawing.Size)(resources.GetObject("$this.ClientSize")));
			this.Controls.Add(this.btnAnularCheques);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.uGridDetalleAsiento);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtpDesde);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.dtpHasta);
			this.Controls.Add(this.uGridReporteAsientos);
			this.Enabled = ((bool)(resources.GetObject("$this.Enabled")));
			this.Font = ((System.Drawing.Font)(resources.GetObject("$this.Font")));
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("$this.ImeMode")));
			this.KeyPreview = true;
			this.Location = ((System.Drawing.Point)(resources.GetObject("$this.Location")));
			this.MaximumSize = ((System.Drawing.Size)(resources.GetObject("$this.MaximumSize")));
			this.MinimumSize = ((System.Drawing.Size)(resources.GetObject("$this.MinimumSize")));
			this.Name = "frmReportesAsientos";
			this.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("$this.RightToLeft")));
			this.StartPosition = ((System.Windows.Forms.FormStartPosition)(resources.GetObject("$this.StartPosition")));
			this.Text = resources.GetString("$this.Text");
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReportesAsientos_KeyDown);
			this.Load += new System.EventHandler(this.frmReportesAsientos_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridReporteAsientos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmReportesAsientos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0403");
			
			if (!Funcion.Permiso("1261", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Autorizar Asientos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExportar.Enabled = true;			
			if (!miAcceso.BAPAutorizar) this.uGridReporteAsientos.DisplayLayout.Bands[0].Columns["btnAutorizar"].CellActivation = Activation.Disabled;
			if (miAcceso.Anular) this.btnAnularCheques.Enabled = true;

			this.dtpDesde.Value = DateTime.Today;
			this.dtpHasta.Value = DateTime.Today;	
		}

		private void Consulta()
		{
			if(!Validacion.vbComparaFechas((DateTime)this.dtpDesde.Value, (DateTime)this.dtpHasta.Value, this.dtpDesde, "La Fecha de inicial no puede ser mayor a la Fecha actual", false))	return;			

			this.uGridDetalleAsiento.DataSource = ultraDataSource2;

			string sSQL = string.Format("Exec ReporteEBZAsiento '{0}', '{1}', {2}", 
				Convert.ToDateTime(this.dtpDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtpHasta.Value).ToString("yyyyMMdd"), (int)this.optEstados.Value);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridReporteAsientos);			

			this.lblEstado.Text = this.uGridReporteAsientos.Rows.Count + " Registros Encontrados";
				
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.Consulta();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void uGridReporteAsientos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString().Equals("btnAutorizar"))
			{				
				if(Convert.ToDateTime(e.Cell.Row.Cells["Fecha"].Value) < DateTime.Today)
				{
					if (!miAcceso.Editar) 
					{
						MessageBox.Show("Solo Puede Autorizar Movimientos Asientos Del Día Actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}

				if ((int)this.uGridReporteAsientos.ActiveRow.Cells["idTipoAsiento"].Value == 4)
				{
					if(Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value) > 5000000.00m)
					{
						if (!miAcceso.BAprobarValorMayor) 
						{
							MessageBox.Show("No puede autorizar Movimientos con valor mayor a 500.00", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
					}
				}
				
				using (frmNCAR Nota = new frmNCAR())
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{	
						string sSQL = string.Format("Exec AsientosActualizaEstado {0}, {1}, '{2}'", 
							(int)e.Cell.Row.Cells["idAsiento"].Value, (int)Nota.optEstado.Value, Nota.txtNotas.Text.ToString().Trim());
						Funcion.EjecutaSQL(cdsSeteoF, sSQL);

						this.Consulta();
					}
				}				
			}
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridReporteAsientos);	
		}

		private void uGridReporteAsientos_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (this.uGridReporteAsientos.ActiveRow == null) return;
			if (this.uGridReporteAsientos.ActiveRow.Cells["idAsiento"].Value == DBNull.Value) return;

			string sSQL = string.Format("Exec ReportesDetallesAsiento {0}", (int)this.uGridReporteAsientos.ActiveRow.Cells["idAsiento"].Value);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridDetalleAsiento);
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			this.Consulta();

			this.uGridReporteAsientos.DisplayLayout.Bands[0].Columns["btnAutorizar"].CellActivation = Activation.Disabled;

			if ((int)this.optEstados.Value == 1) 
				if (miAcceso.BAPAutorizar) this.uGridReporteAsientos.DisplayLayout.Bands[0].Columns["btnAutorizar"].CellActivation = Activation.AllowEdit;
		}

		private void uGridReporteAsientos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmReportesAsientos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnAnularCheques_Click(object sender, System.EventArgs e)
		{
			using (frmAnulaCheques Anula = new frmAnulaCheques())
			{
				if (DialogResult.OK == Anula.ShowDialog())
				{	
					string sSQL = string.Format("Exec AsientoChequeAnulado '{0}', '{1}', {2}, '{3}'",
						Convert.ToDateTime(Anula.dtFecha.Value).ToString("yyyyMMdd"), Anula.txtCheque.Text.ToString(), (int)Anula.cmbCuentas.Value, Anula.txtNotas.Text.ToString());
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);

					MessageBox.Show("Cheque Anulado Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
		}
	}
}
