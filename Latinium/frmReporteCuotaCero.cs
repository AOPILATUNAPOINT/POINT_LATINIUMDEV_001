using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteCuotaCero.
	/// </summary>
	public class frmReporteCuotaCero : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optRegion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkLocal;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocal;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnExcel;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteCuotaCero()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Region");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaCuotaCero");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaCarteraFlujo");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobradoCuotaCero");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobradoCarteraFlujo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCuotaCero");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCarteraFlujo");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionCuotaCero");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionCarteraFlujo");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PenalizacionCartera");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CobradoCuotaCero", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CobradoCuotaCero", 5, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ComisionCuotaCero", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ComisionCuotaCero", 9, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ComisionCarteraFlujo", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ComisionCarteraFlujo", 10, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PenalizacionCartera", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PenalizacionCartera", 11, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CobradoCarteraFlujo", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CobradoCarteraFlujo", 6, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteCuotaCero));
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaCuotaCero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaCarteraFlujo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CobradoCuotaCero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CobradoCarteraFlujo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCuotaCero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCarteraFlujo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionCuotaCero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionCarteraFlujo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PenalizacionCartera");
			this.optRegion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.chkLocal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridLocal = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnVer = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExcel = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.optRegion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// optRegion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optRegion.Appearance = appearance1;
			this.optRegion.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optRegion.CheckedIndex = 0;
			this.optRegion.ItemAppearance = appearance2;
			this.optRegion.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Sierra";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Costa";
			this.optRegion.Items.Add(valueListItem1);
			this.optRegion.Items.Add(valueListItem2);
			this.optRegion.Items.Add(valueListItem3);
			this.optRegion.ItemSpacingVertical = 10;
			this.optRegion.Location = new System.Drawing.Point(16, 40);
			this.optRegion.Name = "optRegion";
			this.optRegion.Size = new System.Drawing.Size(192, 26);
			this.optRegion.TabIndex = 1031;
			this.optRegion.Text = "Todos";
			this.optRegion.ValueChanged += new System.EventHandler(this.optRegion_ValueChanged);
			// 
			// chkLocal
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkLocal.Appearance = appearance3;
			this.chkLocal.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkLocal.Checked = true;
			this.chkLocal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkLocal.Location = new System.Drawing.Point(184, 80);
			this.chkLocal.Name = "chkLocal";
			this.chkLocal.Size = new System.Drawing.Size(13, 16);
			this.chkLocal.TabIndex = 1030;
			this.chkLocal.CheckedChanged += new System.EventHandler(this.chkLocal_CheckedChanged);
			// 
			// uGridLocal
			// 
			this.uGridLocal.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocal.DataSource = this.ultraDataSource6;
			appearance4.BackColor = System.Drawing.Color.White;
			this.uGridLocal.DisplayLayout.Appearance = appearance4;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.Caption = "Local";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 150;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 28;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand1.SummaryFooterCaption = "TOTALES";
			this.uGridLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridLocal.DisplayLayout.GroupByBox.Hidden = true;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocal.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.uGridLocal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLocal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLocal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocal.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocal.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.uGridLocal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocal.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocal.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocal.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridLocal.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridLocal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocal.Location = new System.Drawing.Point(8, 80);
			this.uGridLocal.Name = "uGridLocal";
			this.uGridLocal.Size = new System.Drawing.Size(213, 240);
			this.uGridLocal.TabIndex = 1029;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource1;
			appearance11.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance11;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 180;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 180;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 250;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance12;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Meta Cuota Cero";
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance13;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Meta Flujo";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance14;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Cob Cuota Cero";
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance15;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Cob Cartera Flujo";
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance16;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "% Cuota Cero";
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance17;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "% Cartera Flujo";
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance18;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Com Cuota Cero";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance19;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Com Flujo";
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance20;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Pen Cartera";
			ultraGridColumn16.Header.VisiblePosition = 11;
			ultraGridBand2.Columns.AddRange(new object[] {
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
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance21;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance22;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance23;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance24;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance25;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5});
			ultraGridBand2.SummaryFooterCaption = "Totales :";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance26.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance26;
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance28.FontData.BoldAsString = "True";
			appearance28.FontData.Name = "Arial";
			appearance28.FontData.SizeInPoints = 8F;
			appearance28.ForeColor = System.Drawing.Color.White;
			appearance28.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance28;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance30;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(232, 80);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1008, 240);
			this.uGridInformacion.TabIndex = 1032;
			this.uGridInformacion.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridInformacion_InitializeRow);
			this.uGridInformacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInformacion_InitializeLayout);
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
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(176, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 1036;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 16);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(37, 16);
			this.label5.TabIndex = 1035;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance31;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(224, 14);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1034;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance32;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 14);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1033;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(240, 42);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(88, 23);
			this.btnVer.TabIndex = 1037;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1168, 48);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 1038;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// frmReporteCuotaCero
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1256, 326);
			this.ControlBox = false;
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.optRegion);
			this.Controls.Add(this.chkLocal);
			this.Controls.Add(this.uGridLocal);
			this.KeyPreview = true;
			this.Name = "frmReporteCuotaCero";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cuota Cero - Flujo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteCuotaCero_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteCuotaCero_Load);
			((System.ComponentModel.ISupportInitialize)(this.optRegion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmReporteCuotaCero_Load(object sender, System.EventArgs e)
		{
			DateTime dtFecha = DateTime.Parse("01/07/2018");

			this.dtDesde.CalendarInfo.MinDate = dtFecha;
			this.dtHasta.CalendarInfo.MinDate = dtFecha;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.dtHasta.Value = DateTime.Today;
	
			this.dtDesde.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
			this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);

			string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegion.Value);		
			
			this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
		}

		private void chkLocal_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
			{
				dr.Cells["Sel"].Value = this.chkLocal.Checked;
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value != System.DBNull.Value)
			{
				this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/07/2017");
				this.dtHasta.CalendarInfo.MaxDate = DateTime.Parse("31/12/2030");

				if (Convert.ToDateTime(this.dtDesde.Value).Month == DateTime.Today.Month) 
				{
					this.dtHasta.Value = DateTime.Today;
					this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;
				}
				else
				{
					this.dtHasta.Value = DateTime.DaysInMonth(Convert.ToDateTime(this.dtDesde.Value).Year, Convert.ToDateTime(this.dtDesde.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Year.ToString();
					this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
				}

				this.dtHasta.CalendarInfo.MinDate = Convert.ToDateTime(this.dtDesde.Value);						
			}
			else this.dtHasta.Value = System.DBNull.Value;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.Text = "Resumen Mensual De Cuota Cero - Flujo ";
			string sLocales = "";
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocales.Length == 0) sLocales = dr.Cells["Bodega"].Value.ToString();
					else if (sLocales.Length > 0) sLocales = sLocales + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}
			string sCargos = "1";

			#region Valida Fechas
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
			#endregion Valida Fechas

			string sSQL = "";

				sSQL = string.Format("Exec Cob_ResumenDeComisionesCuotaCero '{0}', '{1}', '{2}','{3}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, sCargos);

			  FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridInformacion);

			this.Text = this.Text + " - " + this.uGridInformacion.Rows.Count + " Registros Encontrados";
		}

		private void optRegion_ValueChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
			{
				dr.Cells["Sel"].Value = false;
			}
			
			if ((int)this.optRegion.Value == 0 )
			{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegion.Value);		
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				{
					if ( (int) dr.Cells["Region"].Value == 1)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
			}


			if ((int)this.optRegion.Value == 1)
			{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegion.Value);		
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				{
					if ( (int) dr.Cells["Region"].Value == 1)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
			}

			if ((int)this.optRegion.Value == 2)
			{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegion.Value);		
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				{
					if ( (int) dr.Cells["Region"].Value == 2)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
			}
		}

		private void frmReporteCuotaCero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
								
		}

		private void uGridInformacion_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (Convert.ToDecimal(e.Row.Cells["PorcentajeCuotaCero"].Value) >= 100.00m) e.Row.Cells["PorcentajeCuotaCero"].Appearance.BackColor = Color.LimeGreen;
			else if (Convert.ToDecimal(e.Row.Cells["PorcentajeCuotaCero"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeCuotaCero"].Value) < 100.00m) e.Row.Cells["PorcentajeCuotaCero"].Appearance.BackColor = Color.Yellow;
			else if (Convert.ToDecimal(e.Row.Cells["PorcentajeCuotaCero"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeCuotaCero"].Value) < 85.00m) e.Row.Cells["PorcentajeCuotaCero"].Appearance.BackColor = Color.Orange;			
			else if (Convert.ToDecimal(e.Row.Cells["PorcentajeCuotaCero"].Value) < 70.00m) e.Row.Cells["PorcentajeCuotaCero"].Appearance.BackColor = Color.Red;
			
			if (Convert.ToDecimal(e.Row.Cells["PorcentajeCarteraFlujo"].Value) >= 100.00m) e.Row.Cells["PorcentajeCarteraFlujo"].Appearance.BackColor = Color.LimeGreen;
			else if (Convert.ToDecimal(e.Row.Cells["PorcentajeCarteraFlujo"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeCarteraFlujo"].Value) < 100.00m) e.Row.Cells["PorcentajeCarteraFlujo"].Appearance.BackColor = Color.Yellow;
			else if (Convert.ToDecimal(e.Row.Cells["PorcentajeCarteraFlujo"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeCarteraFlujo"].Value) < 85.00m) e.Row.Cells["PorcentajeCarteraFlujo"].Appearance.BackColor = Color.Orange;			
			else if (Convert.ToDecimal(e.Row.Cells["PorcentajeCarteraFlujo"].Value) < 70.00m) e.Row.Cells["PorcentajeCarteraFlujo"].Appearance.BackColor = Color.Red;
			
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
		FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}
	}
}
