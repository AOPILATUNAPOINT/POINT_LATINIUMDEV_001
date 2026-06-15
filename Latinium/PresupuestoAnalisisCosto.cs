using System;
using System.Globalization;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Data;
using DevExpress.XtraPivotGrid;
using System.IO;
using CrystalDecisions.Shared;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de PresupuestoAnalisisCosto.
	/// </summary>
	public class PresupuestoAnalisisCosto : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbObra;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsPresupuestoC;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaDesde;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private DevExpress.XtraEditors.CheckEdit chkTotalF;
		private DevExpress.XtraEditors.CheckEdit chkTotalC;
		private DevExpress.XtraEditors.SimpleButton btCarga;
		private DevExpress.XtraEditors.SimpleButton btExcel;
		private DevExpress.XtraEditors.SimpleButton btGrabaC;
		private DevExpress.XtraEditors.SimpleButton btCargaC;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField2;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField3;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField4;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField5;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField6;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField7;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField8;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField9;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField10;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField12;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField13;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField14;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField15;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField16;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField17;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField18;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField20;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField19;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField21;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField22;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField23;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField24;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField25;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField1;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField11;
		private DevExpress.XtraPivotGrid.PivotGridControl pgrBalance;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public PresupuestoAnalisisCosto()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.cmbObra = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.cdsPresupuestoC = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.dtFechaHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.chkTotalF = new DevExpress.XtraEditors.CheckEdit();
			this.chkTotalC = new DevExpress.XtraEditors.CheckEdit();
			this.btCarga = new DevExpress.XtraEditors.SimpleButton();
			this.btExcel = new DevExpress.XtraEditors.SimpleButton();
			this.btGrabaC = new DevExpress.XtraEditors.SimpleButton();
			this.btCargaC = new DevExpress.XtraEditors.SimpleButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.pivotGridField2 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField3 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField4 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField5 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField6 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField7 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField8 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField9 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField10 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField12 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField13 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField14 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField15 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField16 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField17 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField18 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField20 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField19 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField21 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField22 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField23 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField24 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField25 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField1 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField11 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pgrBalance = new DevExpress.XtraPivotGrid.PivotGridControl();
			((System.ComponentModel.ISupportInitialize)(this.cmbObra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPresupuestoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTotalF.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTotalC.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.pgrBalance)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbObra
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbObra.Appearance = appearance1;
			this.cmbObra.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbObra.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbObra.DisplayMember = "";
			this.cmbObra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbObra.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbObra.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbObra.Location = new System.Drawing.Point(96, 10);
			this.cmbObra.Name = "cmbObra";
			this.cmbObra.Size = new System.Drawing.Size(202, 25);
			this.cmbObra.TabIndex = 3;
			this.cmbObra.ValueMember = "";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(29, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 28);
			this.label1.TabIndex = 14;
			this.label1.Text = "OBRA:";
			// 
			// cdsPresupuestoC
			// 
			this.cdsPresupuestoC.BindingContextCtrl = this;
			this.cdsPresupuestoC.DataLibrary = "LibFacturacion";
			this.cdsPresupuestoC.DataLibraryUrl = "";
			this.cdsPresupuestoC.DataSetDef = "dsPresupuestoC";
			this.cdsPresupuestoC.EnforceConstraints = false;
			this.cdsPresupuestoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPresupuestoC.Name = "cdsPresupuestoC";
			this.cdsPresupuestoC.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPresupuestoC.SchemaDef = null;
			this.cdsPresupuestoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsPresupuestoC_BeforeFill);
			// 
			// dtFechaHasta
			// 
			this.dtFechaHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaHasta.DateButtons.Add(dateButton1);
			this.dtFechaHasta.Format = "dd/MMM/yyyy";
			this.dtFechaHasta.Location = new System.Drawing.Point(595, 10);
			this.dtFechaHasta.Name = "dtFechaHasta";
			this.dtFechaHasta.NonAutoSizeHeight = 23;
			this.dtFechaHasta.Size = new System.Drawing.Size(135, 25);
			this.dtFechaHasta.SpinButtonsVisible = true;
			this.dtFechaHasta.TabIndex = 87;
			this.dtFechaHasta.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.AllowMonthSelection = true;
			this.dtFechaDesde.AllowWeekSelection = true;
			this.dtFechaDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaDesde.DateButtons.Add(dateButton2);
			this.dtFechaDesde.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.dtFechaDesde.Format = "dd/MMM/yyyy";
			this.dtFechaDesde.Location = new System.Drawing.Point(394, 10);
			this.dtFechaDesde.Name = "dtFechaDesde";
			this.dtFechaDesde.NonAutoSizeHeight = 23;
			this.dtFechaDesde.Size = new System.Drawing.Size(134, 25);
			this.dtFechaDesde.SpinButtonsVisible = true;
			this.dtFechaDesde.TabIndex = 86;
			this.dtFechaDesde.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(538, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(43, 18);
			this.label2.TabIndex = 85;
			this.label2.Text = "Hasta:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(317, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(46, 18);
			this.label3.TabIndex = 84;
			this.label3.Text = "Desde:";
			// 
			// saveFileDialog1
			// 
			this.saveFileDialog1.DefaultExt = "Xml";
			this.saveFileDialog1.Filter = "Archivo de Diseño | *.Xml";
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.DefaultExt = "Xml";
			this.openFileDialog1.Filter = "Archivo de Diseño | *.Xml";
			// 
			// chkTotalF
			// 
			this.chkTotalF.Location = new System.Drawing.Point(24, 48);
			this.chkTotalF.Name = "chkTotalF";
			// 
			// chkTotalF.Properties
			// 
			this.chkTotalF.Properties.Caption = "Total Filas";
			this.chkTotalF.Size = new System.Drawing.Size(125, 21);
			this.chkTotalF.TabIndex = 91;
			this.chkTotalF.CheckedChanged += new System.EventHandler(this.chkTotalF_CheckedChanged);
			// 
			// chkTotalC
			// 
			this.chkTotalC.EditValue = true;
			this.chkTotalC.Location = new System.Drawing.Point(152, 48);
			this.chkTotalC.Name = "chkTotalC";
			// 
			// chkTotalC.Properties
			// 
			this.chkTotalC.Properties.Caption = "Total Columnas";
			this.chkTotalC.Size = new System.Drawing.Size(125, 21);
			this.chkTotalC.TabIndex = 90;
			this.chkTotalC.CheckedChanged += new System.EventHandler(this.chkTotalC_CheckedChanged);
			// 
			// btCarga
			// 
			this.btCarga.Location = new System.Drawing.Point(740, 8);
			this.btCarga.Name = "btCarga";
			this.btCarga.Size = new System.Drawing.Size(96, 26);
			this.btCarga.TabIndex = 92;
			this.btCarga.Text = "&Filtro";
			this.btCarga.Click += new System.EventHandler(this.btCarga_Click);
			// 
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(868, 8);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(96, 26);
			this.btExcel.TabIndex = 93;
			this.btExcel.Text = "&Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// btGrabaC
			// 
			this.btGrabaC.Location = new System.Drawing.Point(304, 40);
			this.btGrabaC.Name = "btGrabaC";
			this.btGrabaC.Size = new System.Drawing.Size(96, 26);
			this.btGrabaC.TabIndex = 94;
			this.btGrabaC.Text = "&Graba Diseño";
			this.btGrabaC.Click += new System.EventHandler(this.btGrabaC_Click);
			// 
			// btCargaC
			// 
			this.btCargaC.Location = new System.Drawing.Point(408, 40);
			this.btCargaC.Name = "btCargaC";
			this.btCargaC.Size = new System.Drawing.Size(96, 26);
			this.btCargaC.TabIndex = 95;
			this.btCargaC.Text = "&Carga Diseño";
			this.btCargaC.Click += new System.EventHandler(this.btCargaC_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// pivotGridField2
			// 
			this.pivotGridField2.Area = DevExpress.XtraPivotGrid.PivotArea.RowArea;
			this.pivotGridField2.AreaIndex = 0;
			this.pivotGridField2.Caption = "Obra";
			this.pivotGridField2.FieldName = "Obra";
			this.pivotGridField2.Name = "pivotGridField2";
			this.pivotGridField2.Options.AllowExpand = DevExpress.Utils.DefaultBoolean.False;
			// 
			// pivotGridField3
			// 
			this.pivotGridField3.Area = DevExpress.XtraPivotGrid.PivotArea.RowArea;
			this.pivotGridField3.AreaIndex = 1;
			this.pivotGridField3.FieldName = "Capitulo";
			this.pivotGridField3.Name = "pivotGridField3";
			// 
			// pivotGridField4
			// 
			this.pivotGridField4.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField4.AreaIndex = 1;
			this.pivotGridField4.Caption = "NotasCapitulo";
			this.pivotGridField4.FieldName = "NotasCapitulo";
			this.pivotGridField4.Name = "pivotGridField4";
			// 
			// pivotGridField5
			// 
			this.pivotGridField5.Area = DevExpress.XtraPivotGrid.PivotArea.RowArea;
			this.pivotGridField5.AreaIndex = 2;
			this.pivotGridField5.Caption = "Sub-Capitulo";
			this.pivotGridField5.FieldName = "SubCapitulo";
			this.pivotGridField5.Name = "pivotGridField5";
			this.pivotGridField5.Width = 200;
			// 
			// pivotGridField6
			// 
			this.pivotGridField6.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField6.AreaIndex = 4;
			this.pivotGridField6.Caption = "NotasSubCapitulo";
			this.pivotGridField6.CellFormat.FormatString = "n2";
			this.pivotGridField6.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField6.FieldName = "NotasSubCapitulo";
			this.pivotGridField6.Name = "pivotGridField6";
			// 
			// pivotGridField7
			// 
			this.pivotGridField7.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField7.AreaIndex = 5;
			this.pivotGridField7.Caption = "Rubro";
			this.pivotGridField7.CellFormat.FormatString = "n2";
			this.pivotGridField7.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField7.FieldName = "Rubro";
			this.pivotGridField7.Name = "pivotGridField7";
			// 
			// pivotGridField8
			// 
			this.pivotGridField8.Area = DevExpress.XtraPivotGrid.PivotArea.RowArea;
			this.pivotGridField8.AreaIndex = 3;
			this.pivotGridField8.Caption = "Descripción";
			this.pivotGridField8.FieldName = "Descripcion";
			this.pivotGridField8.Name = "pivotGridField8";
			// 
			// pivotGridField9
			// 
			this.pivotGridField9.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField9.AreaIndex = 2;
			this.pivotGridField9.Caption = "Mes";
			this.pivotGridField9.FieldName = "Fecha";
			this.pivotGridField9.GroupInterval = DevExpress.XtraPivotGrid.PivotGroupInterval.DateMonth;
			this.pivotGridField9.Name = "pivotGridField9";
			// 
			// pivotGridField10
			// 
			this.pivotGridField10.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField10.AreaIndex = 3;
			this.pivotGridField10.Caption = "Año";
			this.pivotGridField10.FieldName = "Fecha";
			this.pivotGridField10.GroupInterval = DevExpress.XtraPivotGrid.PivotGroupInterval.DateYear;
			this.pivotGridField10.Name = "pivotGridField10";
			// 
			// pivotGridField12
			// 
			this.pivotGridField12.Area = DevExpress.XtraPivotGrid.PivotArea.RowArea;
			this.pivotGridField12.AreaIndex = 4;
			this.pivotGridField12.Caption = "Unidad";
			this.pivotGridField12.FieldName = "Unidad";
			this.pivotGridField12.Name = "pivotGridField12";
			// 
			// pivotGridField13
			// 
			this.pivotGridField13.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
			this.pivotGridField13.AreaIndex = 1;
			this.pivotGridField13.Caption = "Cant. Presupuesto";
			this.pivotGridField13.FieldName = "CantPresuspuesto";
			this.pivotGridField13.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField13.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField13.Name = "pivotGridField13";
			this.pivotGridField13.ValueFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			// 
			// pivotGridField14
			// 
			this.pivotGridField14.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
			this.pivotGridField14.AreaIndex = 0;
			this.pivotGridField14.Caption = "Cant. Ejecutada";
			this.pivotGridField14.FieldName = "CantEjecutada";
			this.pivotGridField14.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField14.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField14.Name = "pivotGridField14";
			// 
			// pivotGridField15
			// 
			this.pivotGridField15.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
			this.pivotGridField15.AreaIndex = 3;
			this.pivotGridField15.Caption = "Precio Unitario";
			this.pivotGridField15.FieldName = "PUnitario";
			this.pivotGridField15.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField15.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField15.Name = "pivotGridField15";
			// 
			// pivotGridField16
			// 
			this.pivotGridField16.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
			this.pivotGridField16.AreaIndex = 4;
			this.pivotGridField16.Caption = "Precio Total";
			this.pivotGridField16.EmptyCellText = "0";
			this.pivotGridField16.FieldName = "PTotal";
			this.pivotGridField16.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField16.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField16.Name = "pivotGridField16";
			// 
			// pivotGridField17
			// 
			this.pivotGridField17.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField17.AreaIndex = 11;
			this.pivotGridField17.Caption = "MATERIALES EJECUTADOS";
			this.pivotGridField17.FieldName = "MATERIALES";
			this.pivotGridField17.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField17.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField17.Name = "pivotGridField17";
			this.pivotGridField17.Width = 112;
			// 
			// pivotGridField18
			// 
			this.pivotGridField18.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField18.AreaIndex = 12;
			this.pivotGridField18.Caption = "MAQUINARIA EJECUTADA";
			this.pivotGridField18.FieldName = "MAQUINARIA";
			this.pivotGridField18.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField18.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField18.Name = "pivotGridField18";
			// 
			// pivotGridField20
			// 
			this.pivotGridField20.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField20.AreaIndex = 7;
			this.pivotGridField20.Caption = "INDIRECTOS EJECUTADOS";
			this.pivotGridField20.FieldName = "INDIRECTOS";
			this.pivotGridField20.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField20.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField20.Name = "pivotGridField20";
			// 
			// pivotGridField19
			// 
			this.pivotGridField19.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField19.AreaIndex = 13;
			this.pivotGridField19.Caption = "MANO DE OBRA EJECUTADA";
			this.pivotGridField19.FieldName = "MANO_OBRA";
			this.pivotGridField19.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField19.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField19.Name = "pivotGridField19";
			// 
			// pivotGridField21
			// 
			this.pivotGridField21.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField21.AreaIndex = 6;
			this.pivotGridField21.Caption = "MATERIALES PRESUPUESTO";
			this.pivotGridField21.FieldName = "MATERIALESP";
			this.pivotGridField21.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField21.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField21.Name = "pivotGridField21";
			this.pivotGridField21.SummaryType = DevExpress.Data.PivotGrid.PivotSummaryType.Max;
			// 
			// pivotGridField22
			// 
			this.pivotGridField22.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField22.AreaIndex = 8;
			this.pivotGridField22.Caption = "MAQUINARIA PRESUPUESTO";
			this.pivotGridField22.FieldName = "MAQUINARIAP";
			this.pivotGridField22.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField22.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField22.Name = "pivotGridField22";
			this.pivotGridField22.SummaryType = DevExpress.Data.PivotGrid.PivotSummaryType.Max;
			// 
			// pivotGridField23
			// 
			this.pivotGridField23.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField23.AreaIndex = 10;
			this.pivotGridField23.Caption = "INDIRECTOS PRESUPUESTO";
			this.pivotGridField23.FieldName = "INDIRECTOSP";
			this.pivotGridField23.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField23.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField23.Name = "pivotGridField23";
			this.pivotGridField23.SummaryType = DevExpress.Data.PivotGrid.PivotSummaryType.Max;
			// 
			// pivotGridField24
			// 
			this.pivotGridField24.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField24.AreaIndex = 9;
			this.pivotGridField24.Caption = "MANO DE OBRA PRESUPUESTO";
			this.pivotGridField24.FieldName = "MANO_OBRAP";
			this.pivotGridField24.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField24.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField24.Name = "pivotGridField24";
			this.pivotGridField24.SummaryType = DevExpress.Data.PivotGrid.PivotSummaryType.Max;
			// 
			// pivotGridField25
			// 
			this.pivotGridField25.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField25.AreaIndex = 0;
			this.pivotGridField25.Caption = "N. Compra";
			this.pivotGridField25.FieldName = "numero";
			this.pivotGridField25.Name = "pivotGridField25";
			// 
			// pivotGridField1
			// 
			this.pivotGridField1.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
			this.pivotGridField1.AreaIndex = 2;
			this.pivotGridField1.Caption = "Total Ejecutados";
			this.pivotGridField1.FieldName = "SumaEjecutados";
			this.pivotGridField1.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField1.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField1.Name = "pivotGridField1";
			// 
			// pivotGridField11
			// 
			this.pivotGridField11.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
			this.pivotGridField11.AreaIndex = 5;
			this.pivotGridField11.FieldName = "Diferencia";
			this.pivotGridField11.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField11.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField11.GrandTotalText = "Diferencia";
			this.pivotGridField11.Name = "pivotGridField11";
			// 
			// pgrBalance
			// 
			this.pgrBalance.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.pgrBalance.Appearance.GrandTotalCell.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			this.pgrBalance.Appearance.GrandTotalCell.Options.UseBackColor = true;
			this.pgrBalance.Appearance.TotalCell.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(192)), ((System.Byte)(255)));
			this.pgrBalance.Appearance.TotalCell.Options.UseBackColor = true;
			this.pgrBalance.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.pgrBalance.Cursor = System.Windows.Forms.Cursors.Default;
			this.pgrBalance.Fields.AddRange(new DevExpress.XtraPivotGrid.PivotGridField[] {
																																											this.pivotGridField2,
																																											this.pivotGridField3,
																																											this.pivotGridField4,
																																											this.pivotGridField5,
																																											this.pivotGridField6,
																																											this.pivotGridField7,
																																											this.pivotGridField8,
																																											this.pivotGridField9,
																																											this.pivotGridField10,
																																											this.pivotGridField12,
																																											this.pivotGridField13,
																																											this.pivotGridField14,
																																											this.pivotGridField15,
																																											this.pivotGridField16,
																																											this.pivotGridField17,
																																											this.pivotGridField18,
																																											this.pivotGridField20,
																																											this.pivotGridField19,
																																											this.pivotGridField21,
																																											this.pivotGridField22,
																																											this.pivotGridField23,
																																											this.pivotGridField24,
																																											this.pivotGridField25,
																																											this.pivotGridField1,
																																											this.pivotGridField11});
			this.pgrBalance.ForeColor = System.Drawing.Color.Black;
			this.pgrBalance.Location = new System.Drawing.Point(8, 80);
			this.pgrBalance.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Style3D;
			this.pgrBalance.Name = "pgrBalance";
			this.pgrBalance.OptionsView.ShowColumnTotals = false;
			this.pgrBalance.OptionsView.ShowRowGrandTotals = false;
			this.pgrBalance.OptionsView.ShowRowTotals = false;
			this.pgrBalance.Size = new System.Drawing.Size(1040, 464);
			this.pgrBalance.TabIndex = 88;
			// 
			// PresupuestoAnalisisCosto
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 16);
			this.ClientSize = new System.Drawing.Size(1056, 552);
			this.Controls.Add(this.btCargaC);
			this.Controls.Add(this.btGrabaC);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.btCarga);
			this.Controls.Add(this.chkTotalF);
			this.Controls.Add(this.chkTotalC);
			this.Controls.Add(this.pgrBalance);
			this.Controls.Add(this.dtFechaHasta);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbObra);
			this.Name = "PresupuestoAnalisisCosto";
			this.Text = "Analisis Costos";
			this.Load += new System.EventHandler(this.PresupuestoAnalisisCosto_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbObra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPresupuestoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTotalF.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTotalC.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.pgrBalance)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void button1_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void cdsPresupuestoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPresupuestoC.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void PresupuestoAnalisisCosto_Load(object sender, System.EventArgs e)
		{
			string stSelect ="Select idPresupuestoC,Obra from P_Presupuesto";
			cmbObra.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);
			
			cmbObra.Rows.Band.Columns["idPresupuestoC"].Hidden = true;
			cmbObra.DisplayMember = "Obra";
			cmbObra.ValueMember = "idPresupuestoC";

			DateTime dtDesde = DateTime.Today.AddMonths(-1).AddDays(-DateTime.Today.Day+1);
			
			dtFechaDesde.Value = dtDesde;
			dtFechaHasta.Value = dtDesde.AddMonths(1).AddDays(-1);

			#region Configuracion de Impresion de Cubos
			pgrBalance.OptionsPrint.PrintColumnHeaders = DevExpress.Utils.DefaultBoolean.False;
			pgrBalance.OptionsPrint.PrintDataHeaders = DevExpress.Utils.DefaultBoolean.False;;
			pgrBalance.OptionsPrint.PrintFilterHeaders = DevExpress.Utils.DefaultBoolean.False;;
			#endregion

		}

		
		string stTitulo = "Titulo Generico";
		
		private void printableComponentLink1_CreateReportHeaderArea(object sender, DevExpress.XtraPrinting.CreateAreaEventArgs e)
		{
			string reportHeader = stTitulo;
			e.Graph.StringFormat = new DevExpress.XtraPrinting.BrickStringFormat(StringAlignment.Near);
			e.Graph.Font = new Font("Tahoma", 14, FontStyle.Bold);
			RectangleF rec = new RectangleF(0, 0, e.Graph.ClientPageSize.Width, e.Graph.Font.Height * 4 + e.Graph.Font.Height/4);
			e.Graph.DrawString(reportHeader, rec);
		}
	
		

		private void chkTotalC_CheckedChanged(object sender, System.EventArgs e)
		{
			pgrBalance.OptionsView.ShowColumnGrandTotals = chkTotalC.Checked;
			pgrBalance.OptionsView.ShowColumnTotals = chkTotalC.Checked;
		}

		private void chkTotalF_CheckedChanged(object sender, System.EventArgs e)
		{
			pgrBalance.OptionsView.ShowRowTotals = chkTotalF.Checked;
			pgrBalance.OptionsView.ShowRowGrandTotals = chkTotalF.Checked;
		}

		private void btCarga_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbObra, "");
			if(cmbObra.Value == null || cmbObra.Value == DBNull.Value || cmbObra.Value.ToString().Length <=0)
			{
				errorProvider.SetError(cmbObra, "Seleccione Obra...");
				return;
			}
			DateTime dtDesde = (DateTime) dtFechaDesde.Value;
			DateTime dtHasta = (DateTime) dtFechaHasta.Value;
 
			string stSelect ="Exec P_AnalisCostosCubos "+ cmbObra.Value +",0,0,0, '"+dtDesde.ToString("yyyyMMdd")+"','"+dtHasta.ToString("yyyyMMdd")+"'";
			pgrBalance.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) dtFechaDesde.Value;
			DateTime dtHasta = (DateTime) dtFechaHasta.Value;
			string stEmpresa = Funcion.sEscalarSQL(cdsPresupuestoC, "Select Empresa From seteo");
			stTitulo = string.Format("{3}\nAnálisis de Costos \ndel {0} al {1}", 
				dtDesde.ToString("dd/MMM/yyyy"), 
				dtHasta.ToString("dd/MMM/yyyy"), cmbObra.Text, stEmpresa);

			DevExpress.XtraPrinting.PrintingSystem ps = new DevExpress.XtraPrinting.PrintingSystem();
			DevExpress.XtraPrinting.PrintableComponentLink link = new DevExpress.XtraPrinting.PrintableComponentLink(ps);
			link.Component = pgrBalance;
			link.CreateReportHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(
				printableComponentLink1_CreateReportHeaderArea);
			link.CreateDocument();
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo en Excel?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				string stDir = MenuLatinium.stDirInicio + "\\AnalisisCosto.xls";
				ps.ExportToXls(stDir, true);
				System.Diagnostics.Process.Start(stDir);
			}
			else
				link.ShowPreview();

		}

		private void btGrabaC_Click(object sender, System.EventArgs e)
		{
			saveFileDialog1.InitialDirectory = Application.StartupPath;
			if (saveFileDialog1.ShowDialog() == DialogResult.OK)
			{
				pgrBalance.SaveLayoutToXml(saveFileDialog1.FileName);
			}
		}

		private void btCargaC_Click(object sender, System.EventArgs e)
		{
			openFileDialog1.InitialDirectory = Application.StartupPath;
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				pgrBalance.RestoreLayoutFromXml(openFileDialog1.FileName);
			}
		}

		}
}
