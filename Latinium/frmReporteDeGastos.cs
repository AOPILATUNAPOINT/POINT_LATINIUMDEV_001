using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmReporteDeGastos.
	/// </summary>
	public class frmReporteDeGastos : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCentrosDeCosto;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGastos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCentrosDeCosto;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentas;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optGastos;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteDeGastos()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteDeGastos));
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GASTO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GastoEfectivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("GASTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("GastoEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoCC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridCentrosDeCosto = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label53 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.chkGastos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCentrosDeCosto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCuentas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.optFiltro = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optGastos = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCentrosDeCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// uGridCentrosDeCosto
			// 
			this.uGridCentrosDeCosto.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridCentrosDeCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCentrosDeCosto.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridCentrosDeCosto.DisplayLayout.Appearance = appearance1;
			this.uGridCentrosDeCosto.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.Caption = "N.";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 32;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 292;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn3.Width = 23;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			ultraGridBand1.Header.Caption = "Centros de Costo";
			ultraGridBand1.HeaderVisible = true;
			this.uGridCentrosDeCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCentrosDeCosto.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridCentrosDeCosto.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridCentrosDeCosto.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridCentrosDeCosto.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridCentrosDeCosto.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCentrosDeCosto.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridCentrosDeCosto.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCentrosDeCosto.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCentrosDeCosto.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCentrosDeCosto.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridCentrosDeCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCentrosDeCosto.Location = new System.Drawing.Point(8, 368);
			this.uGridCentrosDeCosto.Name = "uGridCentrosDeCosto";
			this.uGridCentrosDeCosto.Size = new System.Drawing.Size(368, 258);
			this.uGridCentrosDeCosto.TabIndex = 652;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(208, 8);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(33, 16);
			this.label53.TabIndex = 651;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(35, 16);
			this.lblBodega.TabIndex = 650;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance8;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(264, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 649;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance9;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 648;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// uGridGastos
			// 
			this.uGridGastos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource2;
			appearance10.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance10;
			this.uGridGastos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 52;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 104;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Cuenta";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 181;
			ultraGridColumn7.Header.Caption = "G";
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 29;
			ultraGridColumn8.Header.Caption = "G E";
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn9.Header.Caption = "...";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 22;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand2.Header.Caption = "Gastos";
			ultraGridBand2.HeaderVisible = true;
			this.uGridGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance11;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 10F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.uGridGastos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance14.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance14.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(8, 88);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(368, 224);
			this.uGridGastos.TabIndex = 653;
			this.uGridGastos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_CellChange);
			this.uGridGastos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridGastos_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn7.DefaultValue = false;
			ultraDataColumn8.DataType = typeof(bool);
			ultraDataColumn8.DefaultValue = false;
			ultraDataColumn9.DataType = typeof(bool);
			ultraDataColumn9.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance17.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance17;
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance18;
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 8F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance20;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(384, 56);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(856, 568);
			this.uGridInformacion.TabIndex = 654;
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton);
			// 
			// saveFileDialog1
			// 
			this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
			// 
			// ultraGridExcelExporter1
			// 
			this.ultraGridExcelExporter1.BeginExport += new Infragistics.Win.UltraWinGrid.ExcelExport.BeginExportEventHandler(this.ultraGridExcelExporter1_BeginExport);
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
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1160, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(72, 24);
			this.btnExcel.TabIndex = 655;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(400, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 656;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// chkGastos
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkGastos.Appearance = appearance23;
			this.chkGastos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkGastos.Location = new System.Drawing.Point(344, 96);
			this.chkGastos.Name = "chkGastos";
			this.chkGastos.Size = new System.Drawing.Size(13, 13);
			this.chkGastos.TabIndex = 662;
			this.chkGastos.Visible = false;
			this.chkGastos.CheckedChanged += new System.EventHandler(this.chkGastos_CheckedChanged);
			// 
			// chkCentrosDeCosto
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkCentrosDeCosto.Appearance = appearance24;
			this.chkCentrosDeCosto.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCentrosDeCosto.Location = new System.Drawing.Point(352, 376);
			this.chkCentrosDeCosto.Name = "chkCentrosDeCosto";
			this.chkCentrosDeCosto.Size = new System.Drawing.Size(13, 13);
			this.chkCentrosDeCosto.TabIndex = 663;
			this.chkCentrosDeCosto.CheckedChanged += new System.EventHandler(this.chkCentrosDeCosto_CheckedChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1248, 8);
			this.groupBox1.TabIndex = 664;
			this.groupBox1.TabStop = false;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 16);
			this.label2.TabIndex = 665;
			this.label2.Text = "Nombre";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn10.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// cmbCuentas
			// 
			this.cmbCuentas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentas.Cursor = System.Windows.Forms.Cursors.IBeam;
			this.cmbCuentas.DataSource = this.ultraDataSource3;
			this.cmbCuentas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 65;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 50;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Width = 219;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbCuentas.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCuentas.DisplayMember = "Descripcion";
			this.cmbCuentas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentas.Location = new System.Drawing.Point(64, 56);
			this.cmbCuentas.Name = "cmbCuentas";
			this.cmbCuentas.Size = new System.Drawing.Size(288, 21);
			this.cmbCuentas.TabIndex = 670;
			this.cmbCuentas.ValueMember = "Tipo";
			this.cmbCuentas.ValueChanged += new System.EventHandler(this.cmbCuentas_ValueChanged);
			this.cmbCuentas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCuentas_InitializeLayout);
			// 
			// optFiltro
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro.Appearance = appearance25;
			this.optFiltro.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro.CheckedIndex = 0;
			this.optFiltro.ItemAppearance = appearance26;
			this.optFiltro.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Centro De Costo";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Por Mes";
			this.optFiltro.Items.Add(valueListItem1);
			this.optFiltro.Items.Add(valueListItem2);
			this.optFiltro.ItemSpacingVertical = 10;
			this.optFiltro.Location = new System.Drawing.Point(496, 8);
			this.optFiltro.Name = "optFiltro";
			this.optFiltro.Size = new System.Drawing.Size(192, 26);
			this.optFiltro.TabIndex = 671;
			this.optFiltro.Text = "Centro De Costo";
			// 
			// optGastos
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.optGastos.Appearance = appearance27;
			this.optGastos.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optGastos.CheckedIndex = 0;
			this.optGastos.ItemAppearance = appearance28;
			this.optGastos.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Gastos";
			valueListItem4.DataValue = 2;
			valueListItem4.DisplayText = "Efectivo";
			this.optGastos.Items.Add(valueListItem3);
			this.optGastos.Items.Add(valueListItem4);
			this.optGastos.ItemSpacingVertical = 10;
			this.optGastos.Location = new System.Drawing.Point(704, 8);
			this.optGastos.Name = "optGastos";
			this.optGastos.Size = new System.Drawing.Size(144, 26);
			this.optGastos.TabIndex = 672;
			this.optGastos.Text = "Gastos";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(864, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 16);
			this.label1.TabIndex = 674;
			this.label1.Text = "Grupo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataSource = this.ultraDataSource4;
			this.cmbGrupo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 84;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 197;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(912, 8);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(216, 21);
			this.cmbGrupo.TabIndex = 675;
			this.cmbGrupo.ValueMember = "idGrupoCC";
			// 
			// frmReporteDeGastos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1248, 630);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.optGastos);
			this.Controls.Add(this.optFiltro);
			this.Controls.Add(this.cmbCuentas);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.chkCentrosDeCosto);
			this.Controls.Add(this.chkGastos);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.uGridGastos);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridCentrosDeCosto);
			this.KeyPreview = true;
			this.Name = "frmReporteDeGastos";
			this.Text = "Reporte De Gastos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteDeGastos_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteDeGastos_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCentrosDeCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);			
		}
		private Acceso miAcceso;

		private void frmReporteDeGastos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso (cdsSeteoF,"200204");

			if (miAcceso.Nuevo) this.btnExcel.Enabled = true;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			
			this.cmbCuentas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Tipo, Codigo, Descripcion From Cuenta Where Grupo = 1 And NumPuntos = 1 Order By Codigo");
			this.cmbCuentas.Value = 5;
			this.cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select 0 AS idGrupoCC , 'TODOS' AS Grupo UNION select idGrupoCC, Grupo as Estado from GrupoCC");
		

			this.uGridCentrosDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre, CONVERT(Bit, 0) As Sel From Proyecto Order By Nombre");
		}

		private void chkGastos_CheckedChanged(object sender, System.EventArgs e)
		{
			string sNombreCol = "Gasto";

			if (Convert.ToInt32(this.optGastos.Value) == 2) sNombreCol = "GastoEfectivo";

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if (this.chkGastos.Checked) dr.Cells[sNombreCol].Value = 1; else dr.Cells[sNombreCol].Value = 0;
			}
		}

		private void chkCentrosDeCosto_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCentrosDeCosto.Rows.All)
			{
				if (this.chkCentrosDeCosto.Checked) dr.Cells["Sel"].Value = 1; else dr.Cells["Sel"].Value = 0;
			}
		}

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridInformacion.Rows.Count > 0)
			{
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Codigo"].Width = 100;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 300;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Descripcion"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].Width = 100;
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
				
					for (int j = 3; j < iColumnas; j++)
					{
						nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 120;
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
					}

					string sNCSuma = "";
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
					{
						for (int j = 3; j < iColumnas; j++)
						{
							sNCSuma = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
							if (dr.Cells[sNCSuma].Value != System.DBNull.Value) dr.Cells["Total"].Value = Convert.ToDecimal(dr.Cells["Total"].Value) + Convert.ToDecimal(dr.Cells[sNCSuma].Value);
						}
					}
				}	
	

				if (this.uGridInformacion.Rows.Count > 0)
				{
					string nombrCols = "";
							
					int iColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 0; j < iColumnas; j++)
					{
						nombrCols = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
						if ( nombrCols 	=="Codigo")
						{
							this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCols].Style = ColumnStyle.Button;
							this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCols].Width = 100;
						}
						if ( nombrCols 	=="Descripcion")
						{
							this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCols].Style = ColumnStyle.Button;
							this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCols].Width = 100;
						}
					}

				}

			}

			this.uGridInformacion.DisplayLayout.Bands[0].Header.Caption = this.uGridInformacion.Rows.Count + " ARTÍCULOS ENCONTRADOS.";
			#endregion Diseño Grid
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			string sGastos = "";
			string sCentrosDeCosto = "";
			string sNombreCol = "Gasto";
			this.uGridInformacion.DataSource = null;	
			if (Convert.ToInt32(this.optGastos.Value) == 2) sNombreCol = "GastoEfectivo";
			int iGrupo = -1;

			if (this.cmbGrupo.ActiveRow != null) iGrupo = (int)this.cmbGrupo.Value;

			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
				if ((bool) dr.Cells[sNombreCol].Value)
					if (sGastos.Length == 0) sGastos = dr.Cells["idCuenta"].Value.ToString(); else if (sGastos.Length > 0) sGastos = sGastos + ", " + dr.Cells["idCuenta"].Value.ToString();

			//			if ((int)this.optFiltro.Value == 0)
			//			{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCentrosDeCosto.Rows.All)
				if ((bool) dr.Cells["Sel"].Value)
					if (sCentrosDeCosto.Length == 0) sCentrosDeCosto = dr.Cells["idProyecto"].Value.ToString(); else if (sCentrosDeCosto.Length > 0) sCentrosDeCosto = sCentrosDeCosto + ", " + dr.Cells["idProyecto"].Value.ToString();
			//			}

			if (sGastos.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (sCentrosDeCosto.Length == 0)// && (int)this.optFiltro.Value == 0)
			{
				MessageBox.Show("Seleccione al menos un Centro de Costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
						
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;
			this.uGridInformacion.DataSource = null;
			if(  iGrupo != -1 )
			{
				string sSQLa = string.Format("Exec ReporteDeGastosGrupo '{0}', '{1}', '{2}', '{3}', {4}, {5}", 
					dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), sGastos, sCentrosDeCosto, (int)this.optFiltro.Value, iGrupo);
				FuncionesProcedimientos.dsGeneral(sSQLa, this.uGridInformacion);
			}
			else
			{
				string sSQL = string.Format("Exec ReporteDeGastos '{0}', '{1}', '{2}', '{3}', {4}", 
					dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), sGastos, sCentrosDeCosto, (int)this.optFiltro.Value);
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInformacion);
			}
			

			

			this.DiseñoGrid();
			this.DiseñoGridTurno (this.uGridInformacion, true, false, "");

			

		}
		private void DiseñoGridTurno(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, bool bIniciaMes, bool bSumatorias, string sColumna)
		{
			try
			{			
				//				nombrCol = this.uGridCartera.DisplayLayout.Bands[0].Columns[j].Header.Caption;
				//				this.uGridCartera.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
				//				this.uGridCartera.DisplayLayout.Bands[0].Columns[nombrCol].Width = 130;
				//				this.uGridCartera.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				//				
				#region Diseño Grid
				//				if (uGrid.Rows.Count > 0)
				//				{
				//					int iColIni = 3;
				//
				//				
				//					if (uGrid.Rows.Count > 0)
				//					{
				//						string nombrCol = "";
				//							
				//						int iColumnas = (int)uGrid.DisplayLayout.Bands[0].Columns.Count;
				//				
				//						for (int j = iColIni; j < iColumnas; j++)
				//						{
				//							
				//							nombrCol = uGrid.DisplayLayout.Bands[0].Columns[j].Header.Caption;			
				//							uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
				//							uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				//	
				////							uGrid.DisplayLayout.Bands[0].Columns[nombrCol].Width = 80;
				////							uGrid.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, uGrid.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);		
				////							uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				////							uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
				////							
				//						}					
				//					}				
				//				}	
				#endregion Diseño Grid
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void frmReporteDeGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridGastos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridGastos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idCuenta"].Value != System.DBNull.Value)
			{
				if (e.Cell.Column.ToString() == "GASTO")
				{
					this.uGridGastos.UpdateData();
					
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaGastoCuentas {0}, {1}", (int)e.Cell.Row.Cells["idCuenta"].Value, (bool)e.Cell.Row.Cells["GASTO"].Value));
				}

				if (e.Cell.Column.ToString() == "GastoEfectivo")  //GASTOEFECTIVO
				{
					this.uGridGastos.UpdateData();
					
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaGastoEfectivoCuentas {0}, {1}", (int)e.Cell.Row.Cells["idCuenta"].Value, (bool)e.Cell.Row.Cells["GastoEfectivo"].Value));
				}
			}
		}

		private void cmbNombre_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
		
		private void cmbCuentas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbCuentas_ValueChanged(object sender, System.EventArgs e)
		{
			this.uGridGastos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idCuenta, Codigo, Descripcion, GASTO, GastoEfectivo, CONVERT(Bit, 0) As Sel From Cuenta Where Tipo = {0} And Grupo = 0 Order By Descripcion", (int)this.cmbCuentas.Value));
		}

		private void ultraGridExcelExporter1_BeginExport(object sender, Infragistics.Win.UltraWinGrid.ExcelExport.BeginExportEventArgs e)
		{
		
		}

		private void saveFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Validacion
			string sGastos = "";
			string sCentrosDeCosto = "";
			string sNombreCol = "Gasto";

			if (Convert.ToInt32(this.optGastos.Value) == 2) sNombreCol = "GastoEfectivo";
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
				if ((bool) dr.Cells[sNombreCol].Value)
					if (sGastos.Length == 0) sGastos = dr.Cells["idCuenta"].Value.ToString(); else if (sGastos.Length > 0) sGastos = sGastos + ", " + dr.Cells["idCuenta"].Value.ToString();

			//			if ((int)this.optFiltro.Value == 0)
			//			{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCentrosDeCosto.Rows.All)
				if ((bool) dr.Cells["Sel"].Value)
					if (sCentrosDeCosto.Length == 0) sCentrosDeCosto = dr.Cells["idProyecto"].Value.ToString(); else if (sCentrosDeCosto.Length > 0) sCentrosDeCosto = sCentrosDeCosto + ", " + dr.Cells["idProyecto"].Value.ToString();
			//			}

			if (sGastos.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (sCentrosDeCosto.Length == 0)// && (int)this.optFiltro.Value == 0)
			{
				MessageBox.Show("Seleccione al menos un Centro de Costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
						
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;


			string sColumna = e.Cell.Column.Key;

			if (e.Cell.Row.Cells[sColumna].Value == System.DBNull.Value)
			{
				MessageBox.Show(string.Format("No hay  registrado en la Columna : {0}", sColumna), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion

			if (e.Cell.Column.ToString() == sColumna)
			{
				//				int Numcuotas = Convert.ToInt32(sColumna);
								
				using (frmReporteDeGastosDetalle CCJ = new frmReporteDeGastosDetalle (dtDesde, dtHasta, Convert.ToString( e.Cell.Row.Cells["Codigo"].Value) , sCentrosDeCosto ))
				{
					if (DialogResult.OK == CCJ.ShowDialog())
					{

					}
				}			 				
			}

		}

		private void optTipoGrupo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}

