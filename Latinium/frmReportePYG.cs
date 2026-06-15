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
	/// Descripción breve de frmReportePYG.
	/// </summary>
	public class frmReportePYG : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCentrosDeCosto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGastos;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCentrosDeCosto;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optActivas;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReportePYG()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReportePYG));
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.optFiltro = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label53 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkCentrosDeCosto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkGastos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridCentrosDeCosto = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.optActivas = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCentrosDeCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optActivas)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(bool);
			ultraDataColumn4.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
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
			// ultraDataSource1
			// 
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn7.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// optFiltro
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro.Appearance = appearance1;
			this.optFiltro.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro.CheckedIndex = 0;
			this.optFiltro.ItemAppearance = appearance2;
			this.optFiltro.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Centro De Costo";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Por Mes";
			this.optFiltro.Items.Add(valueListItem1);
			this.optFiltro.Items.Add(valueListItem2);
			this.optFiltro.ItemSpacingVertical = 10;
			this.optFiltro.Location = new System.Drawing.Point(648, 8);
			this.optFiltro.Name = "optFiltro";
			this.optFiltro.Size = new System.Drawing.Size(192, 26);
			this.optFiltro.TabIndex = 687;
			this.optFiltro.Text = "Centro De Costo";
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(208, 6);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 676;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 6);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 16);
			this.lblBodega.TabIndex = 675;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 38);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(992, 8);
			this.groupBox1.TabIndex = 684;
			this.groupBox1.TabStop = false;
			// 
			// chkCentrosDeCosto
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkCentrosDeCosto.Appearance = appearance3;
			this.chkCentrosDeCosto.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCentrosDeCosto.Location = new System.Drawing.Point(352, 328);
			this.chkCentrosDeCosto.Name = "chkCentrosDeCosto";
			this.chkCentrosDeCosto.Size = new System.Drawing.Size(13, 16);
			this.chkCentrosDeCosto.TabIndex = 683;
			this.chkCentrosDeCosto.CheckedChanged += new System.EventHandler(this.chkCentrosDeCosto_CheckedChanged);
			// 
			// chkGastos
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkGastos.Appearance = appearance4;
			this.chkGastos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkGastos.Location = new System.Drawing.Point(352, 88);
			this.chkGastos.Name = "chkGastos";
			this.chkGastos.Size = new System.Drawing.Size(13, 13);
			this.chkGastos.TabIndex = 682;
			this.chkGastos.CheckedChanged += new System.EventHandler(this.chkGastos_CheckedChanged);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(400, 6);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 681;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(912, 6);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 680;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance5.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance5;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(384, 54);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(600, 418);
			this.uGridInformacion.TabIndex = 679;
			// 
			// uGridGastos
			// 
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource2;
			appearance11.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance11;
			this.uGridGastos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 52;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 120;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Cuenta";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 205;
			ultraGridColumn4.Header.Caption = "...";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 22;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand2.Header.Caption = "Gastos";
			ultraGridBand2.HeaderVisible = true;
			this.uGridGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 10F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance14;
			this.uGridGastos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance15.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance15.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(8, 56);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(368, 240);
			this.uGridGastos.TabIndex = 678;
			// 
			// dtHasta
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance18;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(264, 6);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 674;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance19;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 6);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 673;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// uGridCentrosDeCosto
			// 
			this.uGridCentrosDeCosto.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridCentrosDeCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCentrosDeCosto.DataSource = this.ultraDataSource1;
			appearance20.BackColor = System.Drawing.Color.White;
			this.uGridCentrosDeCosto.DisplayLayout.Appearance = appearance20;
			this.uGridCentrosDeCosto.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "N.";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 30;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 289;
			ultraGridColumn7.Header.Caption = "...";
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn7.Width = 28;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand3.Header.Caption = "Centros de Costo";
			ultraGridBand3.HeaderVisible = true;
			this.uGridCentrosDeCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance21.ForeColor = System.Drawing.Color.Black;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCentrosDeCosto.DisplayLayout.Override.ActiveRowAppearance = appearance21;
			this.uGridCentrosDeCosto.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridCentrosDeCosto.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.uGridCentrosDeCosto.DisplayLayout.Override.CardAreaAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance23.FontData.BoldAsString = "True";
			appearance23.FontData.Name = "Arial";
			appearance23.FontData.SizeInPoints = 10F;
			appearance23.ForeColor = System.Drawing.Color.White;
			appearance23.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCentrosDeCosto.DisplayLayout.Override.HeaderAppearance = appearance23;
			this.uGridCentrosDeCosto.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance24.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance24.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCentrosDeCosto.DisplayLayout.Override.RowAlternateAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCentrosDeCosto.DisplayLayout.Override.RowSelectorAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCentrosDeCosto.DisplayLayout.Override.SelectedRowAppearance = appearance26;
			this.uGridCentrosDeCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCentrosDeCosto.Location = new System.Drawing.Point(8, 304);
			this.uGridCentrosDeCosto.Name = "uGridCentrosDeCosto";
			this.uGridCentrosDeCosto.Size = new System.Drawing.Size(368, 168);
			this.uGridCentrosDeCosto.TabIndex = 677;
			// 
			// optActivas
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.optActivas.Appearance = appearance27;
			this.optActivas.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optActivas.CheckedIndex = 0;
			this.optActivas.ItemAppearance = appearance28;
			this.optActivas.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = 0;
			valueListItem3.DisplayText = "Activas";
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "Inactivas";
			this.optActivas.Items.Add(valueListItem3);
			this.optActivas.Items.Add(valueListItem4);
			this.optActivas.ItemSpacingVertical = 10;
			this.optActivas.Location = new System.Drawing.Point(488, 8);
			this.optActivas.Name = "optActivas";
			this.optActivas.Size = new System.Drawing.Size(152, 26);
			this.optActivas.TabIndex = 688;
			this.optActivas.Text = "Activas";
			this.optActivas.ValueChanged += new System.EventHandler(this.optActivas_ValueChanged);
			// 
			// frmReportePYG
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(992, 478);
			this.Controls.Add(this.optActivas);
			this.Controls.Add(this.optFiltro);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.lblBodega);
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
			this.Name = "frmReportePYG";
			this.Text = "P&G";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReportePYG_KeyDown);
			this.Load += new System.EventHandler(this.frmReportePYG_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCentrosDeCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optActivas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void frmReportePYG_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200810");

			if (!Funcion.Permiso("1870", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a reporte P&G", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExcel.Visible = true;
			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();	
			this.uGridGastos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idCuenta, Codigo, Descripcion,  CONVERT(Bit, 0) As Sel From Cuenta  where ActivoPYG = 1 Order By Descripcion"));
			this.uGridCentrosDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre, CONVERT(Bit, 0) As Sel From Proyecto Where Visible = 1 Order By Nombre");
		
		}

		private void chkGastos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if (this.chkGastos.Checked)
				{
					dr.Cells["Sel"].Value = 1; 
				}
				else
				{
          dr.Cells["Sel"].Value = 0; 
				}
			}
		}

		private void chkCentrosDeCosto_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCentrosDeCosto.Rows.All)
			{
				if (this.chkCentrosDeCosto.Checked) dr.Cells["Sel"].Value = 1; else dr.Cells["Sel"].Value = 0;
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
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


			string sGastos = "";
			string sCentrosDeCosto = "";


			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sGastos.Length == 0) sGastos = dr.Cells["idCuenta"].Value.ToString();
					else if (sGastos.Length > 0) sGastos = sGastos + ", " + dr.Cells["idCuenta"].Value.ToString();
				}
			}
			if (sGastos.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Gasto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCentrosDeCosto.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sCentrosDeCosto.Length == 0) sCentrosDeCosto = dr.Cells["idProyecto"].Value.ToString();
					else if (sCentrosDeCosto.Length > 0) sCentrosDeCosto = sCentrosDeCosto + ", " + dr.Cells["idProyecto"].Value.ToString();
				}
			}
			if (sCentrosDeCosto.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un centro de costos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			this.uGridInformacion.DataSource = null;

			string sSQL = string.Format("Exec [ReporteDeGastosPYG] '{0}', '{1}', '{2}', '{3}', {4}", 
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), sGastos, sCentrosDeCosto, (int)this.optFiltro.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInformacion);

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
		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridInformacion.Rows.Count > 0)
			{
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Codigo"].Width = 100;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 300;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Descripcion"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Estructura"].Width = 300;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Estructura"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Grupo"].Width = 300;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Grupo"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Tipo"].Width = 300;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Tipo"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cuenta"].Width = 300;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cuenta"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Clase"].Width = 300;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Clase"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor = Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0.00}";


				if (this.uGridInformacion.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 8; j < iColumnas; j++)
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
						for (int j = 8; j < iColumnas; j++)
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
						if ( nombrCols 	=="Estructura")
						{
							this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCols].Style = ColumnStyle.Button;
							this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCols].Width = 100;
						}
						if ( nombrCols 	=="Grupo")
						{
							this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCols].Style = ColumnStyle.Button;
							this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCols].Width = 100;
						}
						if ( nombrCols 	=="Tipo")
						{
							this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCols].Style = ColumnStyle.Button;
							this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCols].Width = 100;
						}
						if ( nombrCols 	=="Cuenta")
						{
							this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCols].Style = ColumnStyle.Button;
							this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCols].Width = 100;
						}
						if ( nombrCols 	=="Clase")
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

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
		
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
		FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void frmReportePYG_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void optActivas_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int) this.optActivas.Value == 0 )
			{
				this.uGridGastos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idCuenta, Codigo, Descripcion,  CONVERT(Bit, 0) As Sel From Cuenta  where ActivoPYG = 1 Order By Codigo"));
				
			}
			else
			{
				this.uGridGastos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idCuenta, Codigo, Descripcion,  CONVERT(Bit, 0) As Sel From Cuenta  where Tipo = 5 And Grupo = 0 And isnull(ActivoPYG,0) = 0 Order By Codigo"));
				this.chkGastos.Checked = false;
			}
		}
		
	}
}
