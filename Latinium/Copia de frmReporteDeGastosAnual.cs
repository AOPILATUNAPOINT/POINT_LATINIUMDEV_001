using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteDeGastosAnual.
	/// </summary>
	public class frmReporteDeGastosAnual : System.Windows.Forms.Form
	{
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optPromedio;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optVista;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optGastoMes;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaGrupo;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optConsulta;
		private System.Windows.Forms.Button btnGrupos;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteDeGastosAnual()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteDeGastosAnual));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gasto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAgrupa", -1, "cmbCuentaGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoGasto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gasto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoGasto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.optPromedio = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optVista = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optGastoMes = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cmbCuentaGrupo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optConsulta = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnGrupos = new System.Windows.Forms.Button();
			this.label53 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optPromedio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optVista)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optGastoMes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optConsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			this.SuspendLayout();
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(200, 104);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 1;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(280, 104);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 2;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(552, 8);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(696, 480);
			this.uGridInformacion.TabIndex = 658;
			// 
			// uGridGastos
			// 
			this.uGridGastos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource1;
			appearance7.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance7;
			this.uGridGastos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 52;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 90;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Cuenta";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 185;
			ultraGridColumn4.Header.Caption = "G";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 21;
			ultraGridColumn5.Header.Caption = "Grupo";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn5.Width = 140;
			ultraGridColumn6.Header.Caption = "F-V-E";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 35;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance8;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "%";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 44;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand2.Header.Caption = "Gastos";
			ultraGridBand2.HeaderVisible = true;
			this.uGridGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.uGridGastos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.uGridGastos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(8, 136);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(536, 352);
			this.uGridGastos.TabIndex = 3;
			this.uGridGastos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_CellChange);
			this.uGridGastos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridGastos_KeyPress);
			this.uGridGastos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridGastos_KeyDown);
			this.uGridGastos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_AfterCellUpdate);
			this.uGridGastos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridGastos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(bool);
			ultraDataColumn4.DefaultValue = false;
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
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
			// optPromedio
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.optPromedio.Appearance = appearance15;
			this.optPromedio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optPromedio.CheckedIndex = 0;
			this.optPromedio.ItemAppearance = appearance16;
			this.optPromedio.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Anual";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Ultimo semestre";
			this.optPromedio.Items.Add(valueListItem1);
			this.optPromedio.Items.Add(valueListItem2);
			this.optPromedio.ItemSpacingVertical = 10;
			this.optPromedio.Location = new System.Drawing.Point(8, 104);
			this.optPromedio.Name = "optPromedio";
			this.optPromedio.Size = new System.Drawing.Size(176, 24);
			this.optPromedio.TabIndex = 739;
			this.optPromedio.Text = "Anual";
			// 
			// optVista
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.optVista.Appearance = appearance17;
			this.optVista.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optVista.CheckedIndex = 0;
			this.optVista.ItemAppearance = appearance18;
			this.optVista.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Valores";
			valueListItem4.DataValue = 2;
			valueListItem4.DisplayText = "Porcentajes";
			this.optVista.Items.Add(valueListItem3);
			this.optVista.Items.Add(valueListItem4);
			this.optVista.ItemSpacingVertical = 10;
			this.optVista.Location = new System.Drawing.Point(8, 72);
			this.optVista.Name = "optVista";
			this.optVista.Size = new System.Drawing.Size(176, 24);
			this.optVista.TabIndex = 740;
			this.optVista.Text = "Valores";
			// 
			// optGastoMes
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.optGastoMes.Appearance = appearance19;
			this.optGastoMes.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optGastoMes.CheckedIndex = 0;
			this.optGastoMes.ItemAppearance = appearance20;
			this.optGastoMes.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Por Gasto";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Por Mes";
			this.optGastoMes.Items.Add(valueListItem5);
			this.optGastoMes.Items.Add(valueListItem6);
			this.optGastoMes.ItemSpacingVertical = 10;
			this.optGastoMes.Location = new System.Drawing.Point(192, 72);
			this.optGastoMes.Name = "optGastoMes";
			this.optGastoMes.Size = new System.Drawing.Size(176, 24);
			this.optGastoMes.TabIndex = 741;
			this.optGastoMes.Text = "Por Gasto";
			// 
			// cmbCuentaGrupo
			// 
			this.cmbCuentaGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaGrupo.DataSource = this.ultraDataSource2;
			ultraGridBand3.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 138;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 278;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbCuentaGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCuentaGrupo.DisplayMember = "Grupo";
			this.cmbCuentaGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaGrupo.Location = new System.Drawing.Point(40, 264);
			this.cmbCuentaGrupo.Name = "cmbCuentaGrupo";
			this.cmbCuentaGrupo.Size = new System.Drawing.Size(280, 112);
			this.cmbCuentaGrupo.TabIndex = 742;
			this.cmbCuentaGrupo.ValueMember = "idCuentaAgrupa";
			this.cmbCuentaGrupo.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// optConsulta
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.optConsulta.Appearance = appearance21;
			this.optConsulta.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optConsulta.CheckedIndex = 0;
			this.optConsulta.ItemAppearance = appearance22;
			this.optConsulta.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem7.DataValue = 1;
			valueListItem7.DisplayText = "Detallado";
			valueListItem8.DataValue = 2;
			valueListItem8.DisplayText = "Agrupado";
			this.optConsulta.Items.Add(valueListItem7);
			this.optConsulta.Items.Add(valueListItem8);
			this.optConsulta.ItemSpacingVertical = 10;
			this.optConsulta.Location = new System.Drawing.Point(8, 40);
			this.optConsulta.Name = "optConsulta";
			this.optConsulta.Size = new System.Drawing.Size(176, 24);
			this.optConsulta.TabIndex = 743;
			this.optConsulta.Text = "Detallado";
			// 
			// btnGrupos
			// 
			this.btnGrupos.Enabled = false;
			this.btnGrupos.Image = ((System.Drawing.Image)(resources.GetObject("btnGrupos.Image")));
			this.btnGrupos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGrupos.Location = new System.Drawing.Point(360, 104);
			this.btnGrupos.Name = "btnGrupos";
			this.btnGrupos.Size = new System.Drawing.Size(75, 24);
			this.btnGrupos.TabIndex = 744;
			this.btnGrupos.Text = "Grupos";
			this.btnGrupos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGrupos.Click += new System.EventHandler(this.btnGrupos_Click);
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(200, 10);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 748;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label53.Click += new System.EventHandler(this.label53_Click);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 16);
			this.lblBodega.TabIndex = 747;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblBodega.Click += new System.EventHandler(this.lblBodega_Click);
			// 
			// dtDesde
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance23;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 745;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtDesde_BeforeDropDown);
			// 
			// dtHasta
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance24;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 746;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtHasta_BeforeDropDown);
			// 
			// frmReporteDeGastosAnual
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1256, 494);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.btnGrupos);
			this.Controls.Add(this.optConsulta);
			this.Controls.Add(this.cmbCuentaGrupo);
			this.Controls.Add(this.optGastoMes);
			this.Controls.Add(this.optVista);
			this.Controls.Add(this.optPromedio);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.uGridGastos);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnExcel);
			this.KeyPreview = true;
			this.Name = "frmReporteDeGastosAnual";
			this.Text = "Reporte De Gastos Anual";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteDeGastosAnual_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteDeGastosAnual_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optPromedio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optVista)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optGastoMes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optConsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
   	private Acceso miAcceso;
		private void frmReporteDeGastosAnual_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso (cdsSeteoF,"200207");

			if (miAcceso.BExportar) this.btnExcel.Enabled = true;
			if (miAcceso.Nuevo) this.btnGrupos.Enabled = true;
			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");


			this.uGridGastos.DataSource = FuncionesProcedimientos.dtGeneral("Select idCuenta, Codigo, Descripcion, Gasto, idCuentaAgrupa, TipoGasto, Porcentaje From Cuenta Where Tipo = 5 And Grupo = 0 Order By Codigo");
			this.cmbCuentaGrupo.DataSource = FuncionesProcedimientos.dtGeneral("Select idCuentaAgrupa, Grupo From CuentaGrupo Order By Grupo");
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
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

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].Width = 90;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.NoEdit;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor = Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0.00}";

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"].Width = 90;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"].Hidden = true;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"].CellActivation = Activation.NoEdit;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Semestre", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Semestre"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Semestre"].DisplayFormat = "{0: #,##0.00}";

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].Width = 90;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellActivation = Activation.NoEdit;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellAppearance.BackColor = Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Promedio", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Promedio"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Promedio"].DisplayFormat = "{0: #,##0.00}";

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].Width = 90;
				if ((int)this.optVista.Value == 2 && (int)this.optGastoMes.Value == 2) this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].Hidden = false;
				else this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].Hidden = true;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].CellActivation = Activation.NoEdit;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].CellAppearance.BackColor = Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("PromedioMes", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["PromedioMes"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["PromedioMes"].DisplayFormat = "{0: #,##0.00}";
				
				if (this.uGridInformacion.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 6; j < iColumnas; j++)
					{
						nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 80;

						this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
					}

					string sNCSuma = "";
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
					{
						int iPromedioAnual = 0;
						int iPromedioSemestral = 0;

						for (int i = 6; i < iColumnas; i++)
						{
							sNCSuma = this.uGridInformacion.DisplayLayout.Bands[0].Columns[i].Header.Caption;
							if (dr.Cells[sNCSuma].Value != System.DBNull.Value) iPromedioAnual++;
							if (i > 11) if (dr.Cells[sNCSuma].Value != System.DBNull.Value) iPromedioSemestral++;
						}

						for (int j = 6; j < iColumnas; j++)
						{
							sNCSuma = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
							if (dr.Cells[sNCSuma].Value != System.DBNull.Value) dr.Cells["Total"].Value = Convert.ToDecimal(dr.Cells["Total"].Value) + Convert.ToDecimal(dr.Cells[sNCSuma].Value);
							if (j > 11)
								if (dr.Cells[sNCSuma].Value != System.DBNull.Value) 
									dr.Cells["Semestre"].Value = Convert.ToDecimal(dr.Cells["Semestre"].Value) + Convert.ToDecimal(dr.Cells[sNCSuma].Value);
							              	
							if ((int)this.optPromedio.Value == 1)
								dr.Cells["Promedio"].Value = Decimal.Round(Convert.ToDecimal(dr.Cells["Total"].Value) / iPromedioAnual, 2);
							if ((int)this.optPromedio.Value == 2)
								dr.Cells["Promedio"].Value = Decimal.Round(Convert.ToDecimal(dr.Cells["Semestre"].Value) / iPromedioSemestral, 2);
						}
					}

					sNCSuma = "";
					if ((int)optVista.Value == 2)
					{
						int iPromedioMes = 0;
						decimal dSumaTotalMes = 0.00m;

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
						{
							iPromedioMes = 0;
							dSumaTotalMes = 0.00m;

							for (int j = 6; j < iColumnas; j++)
							{
								sNCSuma = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;

								#region Porcentaje Por Gasto
								if ((int)optGastoMes.Value == 1)
								{
									if (dr.Cells[sNCSuma].Value != System.DBNull.Value) 
									{
										if (Convert.ToDecimal(dr.Cells["Promedio"].Value) > 0.00m)
											if (Convert.ToDecimal(dr.Cells[sNCSuma].Value) > 0.00m)
												dr.Cells[sNCSuma].Value = Decimal.Round(Convert.ToDecimal(dr.Cells[sNCSuma].Value) * 100/Convert.ToDecimal(dr.Cells["Promedio"].Value), 2);
									}
								}
								#endregion Porcentaje Por Gasto

								#region Porcentaje Por Mes
								if ((int)optGastoMes.Value == 2)
								{
									decimal dTotalMes = 0.00m;
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drTotal in this.uGridInformacion.Rows.All)
									{
										if (drTotal.Cells[sNCSuma].Value != System.DBNull.Value) 
											dTotalMes = dTotalMes + Convert.ToDecimal(drTotal.Cells[sNCSuma].Value);
									}
									
									if (dr.Cells[sNCSuma].Value != System.DBNull.Value) 
									{
										if (dTotalMes > 0.00m)
											if (Convert.ToDecimal(dr.Cells[sNCSuma].Value) > 0.00m)
												dr.Cells[sNCSuma].Value = Decimal.Round(Convert.ToDecimal(dr.Cells[sNCSuma].Value) * 100/dTotalMes, 2);

										iPromedioMes++;
										dSumaTotalMes = dSumaTotalMes + Convert.ToDecimal(dr.Cells[sNCSuma].Value);
									}
								}

								if (dSumaTotalMes > 0.00m)
									if (iPromedioMes > 0)
										dr.Cells["PromedioMes"].Value = Decimal.Round(dSumaTotalMes / iPromedioMes, 2);
								#endregion Porcentaje Por Mes
							}
						}
					}
				}				
			}

			this.uGridInformacion.DisplayLayout.Bands[0].Header.Caption = this.uGridInformacion.Rows.Count + " ARTÍCULOS ENCONTRADOS.";
			#endregion Diseño Grid
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

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{	
				if ((bool)dr.Cells["Gasto"].Value)
				{
					if (sGastos.Length == 0) sGastos = dr.Cells["idCuenta"].Value.ToString(); 
					else if (sGastos.Length > 0) sGastos = sGastos + ", " + dr.Cells["idCuenta"].Value.ToString();
				}				
			}

			this.uGridInformacion.DataSource = null;
			
			string sProcedimiento = "ReporteDeGastosAnual";
      if ((int)this.optConsulta.Value == 2) sProcedimiento = "ReporteDeGastosAnualAgrupado";

			string sSQL = string.Format("Exec {0} '{1}', '{2}', '{3}'", sProcedimiento, Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sGastos);
			this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			this.DiseñoGrid();
		}

		private void uGridGastos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idCuenta"].Value != System.DBNull.Value)
			{
				if (e.Cell.Column.ToString() == "GASTO")
				{
					this.uGridGastos.UpdateData();
					
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaGastoCuentas {0}, {1}", (int)e.Cell.Row.Cells["idCuenta"].Value, (bool)e.Cell.Row.Cells["Gasto"].Value));
				}				
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void uGridGastos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
		}

		private void uGridGastos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "idCuentaAgrupa")
			{									
				int idCuentaAgrupa = 0;
				if (e.Cell.Row.Cells["idCuentaAgrupa"].Value != System.DBNull.Value) idCuentaAgrupa = (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value;

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set idCuentaAgrupa = {0} Where idCuenta = {1}", idCuentaAgrupa, (int)e.Cell.Row.Cells["idCuenta"].Value));
			}	
		
			if (e.Cell.Column.ToString() == "TipoGasto")
				if (e.Cell.Row.Cells["TipoGasto"].Value.ToString() != "V")
					e.Cell.Row.Cells["Porcentaje"].Value = 0.00m;

			if (e.Cell.Column.ToString() == "TipoGasto")
			{					
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set TipoGasto = '{0}' Where idCuenta = {1}", e.Cell.Row.Cells["TipoGasto"].Value.ToString(), (bool)e.Cell.Row.Cells["Fijo"].Value));
			}

			if (e.Cell.Column.ToString() == "Porcentaje")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set Porcentaje = {0} Where idCuenta = {1}", Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
			}		
		}

		private void frmReporteDeGastosAnual_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
		
		private void uGridGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridGastos);
		}

		private void uGridGastos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.uGridGastos.ActiveRow.Cells["TipoGasto"].Value.ToString() != "V" && 
				this.uGridGastos.ActiveCell.Column.Key.ToString().ToUpper() == "PORCENTAJE") e.Handled = true;
		}

		private void btnGrupos_Click(object sender, System.EventArgs e)
		{
			using (frmCuentaGrupo CG = new frmCuentaGrupo())
			{				
				if (DialogResult.OK == CG.ShowDialog())
				{
					this.cmbCuentaGrupo.DataSource = FuncionesProcedimientos.dtGeneral("Select idCuentaAgrupa, Grupo From CuentaGrupo Order By Grupo");
				}
			}			
		}

		private void lblBodega_Click(object sender, System.EventArgs e)
		{
		
		}

		private void dtDesde_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtHasta_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void label53_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
