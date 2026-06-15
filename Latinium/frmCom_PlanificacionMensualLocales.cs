using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_PlanificacionMensualLocales.
	/// </summary>
	public class frmCom_PlanificacionMensualLocales : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocales;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodos;
		private System.Windows.Forms.Button btnIrA;
		private System.Windows.Forms.Button btnExcel;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_PlanificacionMensualLocales()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_PlanificacionMensualLocales));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasTrabajados");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega", -1, "cmbBodega");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Promedio", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Promedio", 8, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Promedio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasTrabajados");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.btnBuscar = new System.Windows.Forms.Button();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkTodos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnIrA = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			this.SuspendLayout();
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(184, 7);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(77, 23);
			this.btnBuscar.TabIndex = 301;
			this.btnBuscar.Text = "&Ver";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 299;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 300;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance2;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 200;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 160;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 60;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 50;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 75;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance6;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 75;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance7;
			ultraGridColumn10.Header.Caption = "Dias";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 50;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.Caption = "Vacaciones - Faltas - Permisos";
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Width = 200;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 200;
			ultraGridColumn13.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn13.Header.Caption = "Local Destino";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn13.Width = 160;
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
																										 ultraGridColumn13});
			ultraGridBand1.Header.Caption = "";
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance8;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance9;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridInformacion.DisplayLayout.GroupByBox.Prompt = " ";
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 7.5F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.uGridInformacion.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(192, 56);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1016, 376);
			this.uGridInformacion.TabIndex = 302;
			this.uGridInformacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridInformacion_KeyDown);
			this.uGridInformacion.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
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
																																 ultraDataColumn13});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1216, 8);
			this.groupBox1.TabIndex = 303;
			this.groupBox1.TabStop = false;
			// 
			// uGridLocales
			// 
			this.uGridLocales.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocales.DataSource = this.ultraDataSource2;
			appearance16.BackColor = System.Drawing.Color.White;
			this.uGridLocales.DisplayLayout.Appearance = appearance16;
			this.uGridLocales.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 100;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 140;
			ultraGridColumn16.Header.Caption = "...";
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn16.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			ultraGridBand2.Header.Caption = "LOCALES";
			ultraGridBand2.HeaderVisible = true;
			appearance17.FontData.SizeInPoints = 7.5F;
			ultraGridBand2.Override.HeaderAppearance = appearance17;
			this.uGridLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocales.DisplayLayout.Override.ActiveRowAppearance = appearance18;
			this.uGridLocales.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridLocales.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridLocales.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance19.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocales.DisplayLayout.Override.CardAreaAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 10F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocales.DisplayLayout.Override.HeaderAppearance = appearance20;
			appearance21.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance21.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowAlternateAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowSelectorAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.SelectedRowAppearance = appearance23;
			this.uGridLocales.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocales.Location = new System.Drawing.Point(8, 56);
			this.uGridLocales.Name = "uGridLocales";
			this.uGridLocales.Size = new System.Drawing.Size(181, 376);
			this.uGridLocales.TabIndex = 647;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(bool);
			ultraDataColumn16.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// chkTodos
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkTodos.Appearance = appearance24;
			this.chkTodos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkTodos.Checked = true;
			this.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodos.Location = new System.Drawing.Point(171, 79);
			this.chkTodos.Name = "chkTodos";
			this.chkTodos.Size = new System.Drawing.Size(13, 13);
			this.chkTodos.TabIndex = 662;
			this.chkTodos.CheckedChanged += new System.EventHandler(this.chkTodos_CheckedChanged);
			// 
			// btnIrA
			// 
			this.btnIrA.Enabled = false;
			this.btnIrA.Image = ((System.Drawing.Image)(resources.GetObject("btnIrA.Image")));
			this.btnIrA.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnIrA.Location = new System.Drawing.Point(1136, 6);
			this.btnIrA.Name = "btnIrA";
			this.btnIrA.Size = new System.Drawing.Size(75, 24);
			this.btnIrA.TabIndex = 714;
			this.btnIrA.Text = "Ir a";
			this.btnIrA.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnIrA.Click += new System.EventHandler(this.btnIrA_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1056, 6);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 715;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
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
			// cmbBodega
			// 
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource3;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Width = 50;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 60;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Width = 180;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			ultraGridBand3.SummaryFooterCaption = "Grand Summaries";
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbBodega.Location = new System.Drawing.Point(558, 167);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(338, 161);
			this.cmbBodega.TabIndex = 716;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// frmCom_PlanificacionMensualLocales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1216, 438);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnIrA);
			this.Controls.Add(this.chkTodos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.uGridLocales);
			this.KeyPreview = true;
			this.Name = "frmCom_PlanificacionMensualLocales";
			this.Text = "Planificacion Mensual Locales Para Alcanzar la Meta";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCom_PlanificacionMensualLocales_KeyDown);
			this.Load += new System.EventHandler(this.frmCom_PlanificacionMensualLocales_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void frmCom_PlanificacionMensualLocales_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0723");

			this.chkTodos.Checked = true;
			this.dtFecha.Value = DateTime.Parse(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());

			this.AutoScroll = true;

			if (miAcceso.BExportar) this.btnExcel.Enabled = true;
			if (miAcceso.BIrA) this.btnIrA.Enabled = true;
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			string sLocales = "";

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocales.Length == 0) sLocales = dr.Cells["Bodega"].Value.ToString();
					else if (sLocales.Length > 0) sLocales = sLocales + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}

			if (sLocales.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', '{1}'", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), sLocales), this.uGridInformacion);

			this.uGridInformacion.DisplayLayout.Bands[0].SortedColumns.Add("Local", false, true);			
			this.uGridInformacion.Rows.ExpandAll(true);
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			this.uGridLocales.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Select b.Nombre, b.Bodega, Convert(Bit, 1) As Sel From Bodega b Inner Join UsuariosBodegas ub On ub.Bodega = b.Bodega Inner Join Usuario u On u.idUsuario = ub.idUsuario Where u.idUsuario = {0} And '{1}' Between FechaApertura And ISNULL(FechaCierre, '20301231') And Factura = 1  And Credito = 1 Order By b.Nombre Asc", MenuLatinium.IdUsuario, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

			this.cmbBodega.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec FacturacionListaBodegas {0}, {1}, '{2}'", MenuLatinium.IdUsuario, 43, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));				
		}

		private void chkTodos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
				dr.Cells["Sel"].Value = this.chkTodos.Checked;
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void btnIrA_Click(object sender, System.EventArgs e)
		{
			this.uGridInformacion.DisplayLayout.Bands[0].SortedColumns.Add("Local", false, false);			

			if (this.uGridInformacion.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridInformacion.DisplayLayout.Bands[0].SortedColumns.Add("Local", false, true);
				this.uGridInformacion.Rows.ExpandAll(true);
				return;
			}
						
			frmReporteVendedoresPorFecha RVF = new frmReporteVendedoresPorFecha(this.uGridInformacion.ActiveRow.Cells["Vendedor"].Value.ToString(), 
				(int)this.uGridInformacion.ActiveRow.Cells["idPersonal"].Value);
			RVF.ShowDialog();	

			this.uGridInformacion.DisplayLayout.Bands[0].SortedColumns.Add("Local", false, true);
			this.uGridInformacion.Rows.ExpandAll(true);
		}

		private void frmCom_PlanificacionMensualLocales_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void uGridInformacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridInformacion);
		}
		
		private void uGridInformacion_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
//			if (e.Cell.Column.ToString() == "Notas" || e.Cell.Column.ToString() == "Bodega")
//			{
//				string sSQL = string.Format("Exec Com_GuardaPlanificacionMensual '{0}', {1}, {2}, '{3}'", 
//					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)e.Cell.Row.Cells["idPersonal"].Value, (int)e.Cell.Row.Cells["Bodega"].Value, e.Cell.Row.Cells["Notas"].Value.ToString());				
//				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
//			}
			if (e.Cell.Column.ToString() == "Bodega")
			{
				this.uGridInformacion.UpdateData();
				MessageBox.Show(e.Cell.Row.Cells["Bodega"].Value.ToString());
				MessageBox.Show(this.cmbBodega.ActiveRow.Cells["Codigo"].Value.ToString());
			}
		}
	}
}
