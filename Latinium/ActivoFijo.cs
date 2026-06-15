using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ActivoFijo.
	/// </summary>
	public class ActivoFijo : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIni;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaFin;
		private DevExpress.XtraEditors.SimpleButton btGenerar;
		private DevExpress.XtraEditors.SimpleButton btContabilizar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsDepreciacion;
		private C1.Data.C1DataSet cdsGrupo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbGrupo;
		private DevExpress.XtraEditors.CheckEdit chkCentroCosto;
		private C1.Data.C1DataSet cdsSeguridad;
		private C1.Data.C1DataSet cdsAsiento;
		private DevExpress.XtraEditors.SimpleButton btImprimir;
		private DevExpress.XtraEditors.SimpleButton btFiltrar;
		private C1.Data.C1DataView cdvDepreciacion;
		private System.Windows.Forms.HelpProvider helpProvider1;
		private C1.Data.C1DataSet cdsProyecto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private System.Windows.Forms.GroupBox groupBox1;
		private DevExpress.XtraEditors.SimpleButton btn_EnviarCI;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ActivoFijo()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Depreciacion", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DepPeriodo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DepTotal");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anios");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo", -1, "cmbGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DepPeriodo", 6, true, "Depreciacion", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DepPeriodo", 6, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DepTotal", 7, true, "Depreciacion", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DepTotal", 7, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Precio", 5, true, "Depreciacion", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Precio", 5, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Residual", 9, true, "Depreciacion", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Residual", 9, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas8");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			this.dtFechaIni = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.dtFechaFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.btContabilizar = new DevExpress.XtraEditors.SimpleButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdvDepreciacion = new C1.Data.C1DataView();
			this.cdsDepreciacion = new C1.Data.C1DataSet();
			this.cdsGrupo = new C1.Data.C1DataSet();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.chkCentroCosto = new DevExpress.XtraEditors.CheckEdit();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.btImprimir = new DevExpress.XtraEditors.SimpleButton();
			this.btFiltrar = new DevExpress.XtraEditors.SimpleButton();
			this.helpProvider1 = new System.Windows.Forms.HelpProvider();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btn_EnviarCI = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvDepreciacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDepreciacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCentroCosto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// dtFechaIni
			// 
			this.dtFechaIni.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaIni.DateButtons.Add(dateButton1);
			this.dtFechaIni.Format = "dd/MMM/yyyy";
			this.dtFechaIni.Location = new System.Drawing.Point(64, 9);
			this.dtFechaIni.Name = "dtFechaIni";
			this.dtFechaIni.NonAutoSizeHeight = 23;
			this.dtFechaIni.Size = new System.Drawing.Size(112, 21);
			this.dtFechaIni.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtFechaIni.SpinButtonsVisible = true;
			this.dtFechaIni.TabIndex = 0;
			this.dtFechaIni.Value = new System.DateTime(2005, 3, 17, 0, 0, 0, 0);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Desde:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(192, 9);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Hasta:";
			// 
			// dtFechaFin
			// 
			this.dtFechaFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaFin.DateButtons.Add(dateButton2);
			this.dtFechaFin.Format = "dd/MMM/yyyy";
			this.dtFechaFin.Location = new System.Drawing.Point(232, 9);
			this.dtFechaFin.Name = "dtFechaFin";
			this.dtFechaFin.NonAutoSizeHeight = 23;
			this.dtFechaFin.Size = new System.Drawing.Size(112, 21);
			this.dtFechaFin.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtFechaFin.SpinButtonsVisible = true;
			this.dtFechaFin.TabIndex = 2;
			this.dtFechaFin.Value = new System.DateTime(2005, 3, 17, 0, 0, 0, 0);
			// 
			// btGenerar
			// 
			this.btGenerar.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btGenerar.Location = new System.Drawing.Point(16, 43);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(75, 25);
			this.btGenerar.TabIndex = 4;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// btContabilizar
			// 
			this.btContabilizar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btContabilizar.Location = new System.Drawing.Point(144, 43);
			this.btContabilizar.Name = "btContabilizar";
			this.btContabilizar.Size = new System.Drawing.Size(75, 25);
			this.btContabilizar.TabIndex = 5;
			this.btContabilizar.Text = "Contabilizar";
			this.btContabilizar.Click += new System.EventHandler(this.btContabilizar_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvDepreciacion;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn1.Header.VisiblePosition = 11;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 74;
			ultraGridColumn3.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 122;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Format = "dd/MMM/yyyy";
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 46;
			ultraGridColumn5.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance1.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance1;
			ultraGridColumn5.Header.Caption = "Cant.";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 46;
			ultraGridColumn6.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 51;
			ultraGridColumn7.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Dep. Periodo";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 62;
			ultraGridColumn8.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Dep. Total";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 62;
			ultraGridColumn9.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance5;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Años";
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 71;
			ultraGridColumn10.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance6;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Width = 56;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn12.Header.Caption = "Grupo";
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn12.Width = 61;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 76;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 100;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 86;
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
																										 ultraGridColumn15});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4});
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre la Columna de Grupo o Centro de Costo para Agrupar";
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 78);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(848, 310);
			this.ultraGrid1.TabIndex = 6;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// cdvDepreciacion
			// 
			this.cdvDepreciacion.BindingContextCtrl = this;
			this.cdvDepreciacion.DataSet = this.cdsDepreciacion;
			this.cdvDepreciacion.TableName = "";
			this.cdvDepreciacion.TableViewName = "Depreciacion";
			// 
			// cdsDepreciacion
			// 
			this.cdsDepreciacion.BindingContextCtrl = this;
			this.cdsDepreciacion.DataLibrary = "LibFacturacion";
			this.cdsDepreciacion.DataLibraryUrl = "";
			this.cdsDepreciacion.DataSetDef = "dsDepreciacion";
			this.cdsDepreciacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsDepreciacion.Name = "cdsDepreciacion";
			this.cdsDepreciacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsDepreciacion.SchemaDef = null;
			this.cdsDepreciacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsDepreciacion_BeforeFill);
			// 
			// cdsGrupo
			// 
			this.cdsGrupo.BindingContextCtrl = this;
			this.cdsGrupo.DataLibrary = "LibFacturacion";
			this.cdsGrupo.DataLibraryUrl = "";
			this.cdsGrupo.DataSetDef = "dsArticuloTabla";
			this.cdsGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupo.Name = "cdsGrupo";
			this.cdsGrupo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupo.SchemaDef = null;
			this.cdsGrupo.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupo_BeforeFill);
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataMember = "ArticuloGrupo";
			this.cmbGrupo.DataSource = this.cdsGrupo;
			ultraGridBand2.AddButtonCaption = "CompraComprobante";
			ultraGridColumn16.Header.VisiblePosition = 6;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Width = 216;
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 8;
			ultraGridColumn24.Header.VisiblePosition = 9;
			ultraGridColumn25.Header.VisiblePosition = 10;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 11;
			ultraGridColumn28.Header.VisiblePosition = 12;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(184, 207);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(232, 86);
			this.cmbGrupo.TabIndex = 7;
			this.cmbGrupo.ValueMember = "idGrupoArticulo";
			this.cmbGrupo.Visible = false;
			// 
			// chkCentroCosto
			// 
			this.chkCentroCosto.EditValue = true;
			this.chkCentroCosto.Location = new System.Drawing.Point(360, 9);
			this.chkCentroCosto.Name = "chkCentroCosto";
			// 
			// chkCentroCosto.Properties
			// 
			this.chkCentroCosto.Properties.Appearance.BackColor = System.Drawing.SystemColors.ActiveBorder;
			this.chkCentroCosto.Properties.Appearance.Options.UseBackColor = true;
			this.chkCentroCosto.Properties.Caption = "Todo Centro Costo";
			this.chkCentroCosto.Size = new System.Drawing.Size(120, 18);
			this.chkCentroCosto.TabIndex = 53;
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			this.cdsSeguridad.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeguridad_BeforeFill);
			// 
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsAsiento";
			this.cdsAsiento.FillOnRequest = false;
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			this.cdsAsiento.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsiento_BeforeFill);
			// 
			// btImprimir
			// 
			this.btImprimir.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btImprimir.Location = new System.Drawing.Point(272, 43);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(75, 25);
			this.btImprimir.TabIndex = 55;
			this.btImprimir.Text = "Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// btFiltrar
			// 
			this.btFiltrar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btFiltrar.Location = new System.Drawing.Point(496, 43);
			this.btFiltrar.Name = "btFiltrar";
			this.btFiltrar.Size = new System.Drawing.Size(75, 25);
			this.btFiltrar.TabIndex = 56;
			this.btFiltrar.Text = "Filtrar";
			this.btFiltrar.Click += new System.EventHandler(this.btFiltrar_Click);
			// 
			// helpProvider1
			// 
			this.helpProvider1.HelpNamespace = "C:\\Latinium\\Latinium.chm";
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataMember = "Proyecto";
			this.cmbProyecto.DataSource = this.cdsProyecto;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 2;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn32.Header.VisiblePosition = 3;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 4;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 5;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 6;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 7;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 8;
			ultraGridColumn37.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbProyecto.DisplayMember = "";
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(496, 9);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(168, 21);
			this.cmbProyecto.TabIndex = 58;
			this.cmbProyecto.ValueMember = "";
			this.cmbProyecto.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbProyecto_InitializeLayout);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btn_EnviarCI);
			this.groupBox1.Location = new System.Drawing.Point(704, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(128, 69);
			this.groupBox1.TabIndex = 111;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Envio Costo Indirecto";
			this.groupBox1.Visible = false;
			// 
			// btn_EnviarCI
			// 
			this.btn_EnviarCI.Location = new System.Drawing.Point(32, 26);
			this.btn_EnviarCI.Name = "btn_EnviarCI";
			this.btn_EnviarCI.Size = new System.Drawing.Size(64, 26);
			this.btn_EnviarCI.TabIndex = 112;
			this.btn_EnviarCI.Text = "Enviar";
			this.btn_EnviarCI.Click += new System.EventHandler(this.btn_EnviarCI_Click);
			// 
			// ActivoFijo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(864, 401);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.btFiltrar);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.chkCentroCosto);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btContabilizar);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFechaFin);
			this.Controls.Add(this.dtFechaIni);
			this.helpProvider1.SetHelpKeyword(this, "ActivosFijos.htm");
			this.helpProvider1.SetHelpNavigator(this, System.Windows.Forms.HelpNavigator.Topic);
			this.Name = "ActivoFijo";
			this.helpProvider1.SetShowHelp(this, true);
			this.Text = "Activos Fijos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ActivoFijo_Closing);
			this.Load += new System.EventHandler(this.ActivoFijo_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvDepreciacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDepreciacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCentroCosto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsDepreciacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsDepreciacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFechaI = (DateTime) this.dtFechaIni.Value;
			DateTime dtFechaF = (DateTime) this.dtFechaFin.Value;
			string stExec = "Exec DepreciacionGenera '" + dtFechaI.ToString("yyyyMMdd")
				+ "', '" + dtFechaF.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsDepreciacion, stExec, true);
			cdsDepreciacion.Clear();
			cdsDepreciacion.Fill();
			Cursor = Cursors.Default;
			MessageBox.Show("Depreciación Generada");
		}

		private void cdsGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsDepreciacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void btContabilizar_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("88", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a la Contabilización de Activos Fijos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			this.Cursor = Cursors.WaitCursor;
			DateTime dtFechaI = (DateTime) this.dtFechaIni.Value;
			DateTime dtFechaF = (DateTime) this.dtFechaFin.Value;
			string stExec = "Exec DepreciacionAsiento '" + dtFechaI.ToString("yyyyMMdd")
				+ "', '" + dtFechaF.ToString("yyyyMMdd")+ "'";
			Funcion.EjecutaSQL(cdsAsiento, stExec, true);
			this.Cursor = Cursors.Default;
			MessageBox.Show("Asiento Generado");

			return;

//			C1.Data.C1DataRow drAsiento = this.cdsAsiento.TableViews["Asiento"].AddNew();
//			drAsiento["Fecha"] = dtFechaF;
//			drAsiento["Descripcion"] = "Depreciación desde el " + dtFechaI.ToString("dd/MMM/yyyy") 
//				+ " Hasta el " + dtFechaF.ToString("dd/MMM/yyyy");
//			this.cdsAsiento.Update();
//			int iIdPrinc = (int) drAsiento["IdAsiento"];
//
//			this.cdsAsiento.Schema.Connections[0].Open();
//			SqlDataReader miReader;
//			C1.Data.SchemaObjects.Connection miConeccion = cdsAsiento.Schema.Connections[0];
//
//			string stFiltroCC = "";
//			if (!chkCentroCosto.Checked)
//				stFiltroCC= " Where Depreciacion.CentroCosto = " + cmbProyecto.Text + " ";
//			string stSelect = "SELECT Depreciacion.CentroCosto, "
//				+ "ArticuloGrupo.idGrupoArticulo, "
//				+ "Max(ArticuloGrupo.Grupo) AS GrupoArt, "
//				+ "Max(ArticuloGrupo.idCuentas4) AS idCuentaDebe, "
//				+ "Max(ArticuloGrupo.idCuentas3) AS idCuentaHaber, "
//				+ "Sum(Depreciacion.DepPeriodo) AS SumaDepPeriodo "
//				+ "FROM (Depreciacion INNER JOIN ArticuloGrupo "
//				+ "ON Depreciacion.idGrupoArticulo = ArticuloGrupo.idGrupoArticulo) "
//				+ "LEFT JOIN Proyecto ON Depreciacion.idProyecto = Proyecto.idProyecto "
//				+ stFiltroCC
//				+ "GROUP BY Depreciacion.CentroCosto, ArticuloGrupo.idGrupoArticulo "
//				+ "ORDER BY Depreciacion.CentroCosto, ArticuloGrupo.idGrupoArticulo";
//
//			SqlCommand cmdCant = new SqlCommand (stSelect,(SqlConnection) miConeccion.DbConnection);
//			miReader = cmdCant.ExecuteReader();
//
//			C1.Data.C1DataRow drDetAsiento;
//			string stCentroC;
//			string stGrupo;
//			try
//			{
//				while (miReader.Read())
//				{
//					stCentroC = "";
//					stGrupo = "";
//					if (miReader.GetValue(0) != System.DBNull.Value)
//						stCentroC = miReader.GetString(0);
//					if (miReader.GetValue(2) != System.DBNull.Value)
//						stGrupo = miReader.GetString(2);
//					int iCuentaDebe = miReader.GetInt32(3);
//					int iCuentaHaber = miReader.GetInt32(4);
//					double dDepGrupo = miReader.GetDouble(5);
//				
//					if (dDepGrupo > 0)
//					{
//						// Debe
//						drDetAsiento = this.cdsAsiento.TableViews["AsientoDetalle"].AddNew();
//						drDetAsiento["IdAsiento"] = iIdPrinc;
//						drDetAsiento["IdCuenta"] = iCuentaDebe;
//						drDetAsiento["Debe"] = dDepGrupo;
//						drDetAsiento["Detalle"] = stGrupo;
//						drDetAsiento["CentroCosto"] = stCentroC;
//		
//						// Haber
//						drDetAsiento = this.cdsAsiento.TableViews["AsientoDetalle"].AddNew();
//						drDetAsiento["IdAsiento"] = iIdPrinc;
//						drDetAsiento["IdCuenta"] = iCuentaHaber;
//						drDetAsiento["Haber"] = dDepGrupo;
//						drDetAsiento["Detalle"] = stGrupo;
//						drDetAsiento["CentroCosto"] = stCentroC;
//					}
//				}
//			}
//			catch (System.Exception ex)
//			{
//				MessageBox.Show(ex.Message);
//			}
//			miReader.Close();
//			this.cdsAsiento.Schema.Connections[0].Close();
//
//			drAsiento["CodAsiento"] = Contabilidad.NumeroAsiento(this.cdsAsiento, 1);
//			this.cdsAsiento.Update();
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsAsiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAsiento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) dtFechaIni.Value;
			DateTime dtHasta = (DateTime) dtFechaFin.Value;
			string stTitulo = "Depreciacion desde " + dtDesde.ToString("dd/MMM/yyyy")
				+ " hasta " + dtHasta.ToString("dd/MMM/yyyy");
			Reporte miReporte = new Reporte("Depreciacion.Rpt");
			miReporte.Titulo(stTitulo);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			this.Cursor = Cursors.Default;
		}

		private void btFiltrar_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			if (this.cmbProyecto.Value == null)
				this.cdvDepreciacion.RowFilter = "CentroCosto like '*'";
			else
				this.cdvDepreciacion.RowFilter = "CentroCosto like '" + 
					this.cmbProyecto.Text.ToString().Trim() + "*'";
			this.Cursor = Cursors.Default;
		}

		private void ActivoFijo_Load(object sender, System.EventArgs e)
		{
			if (Funcion.bEjecutaSQL(cdsAsiento, "Exec SeteoGExiste 'FTRPB'")) 
			{
					groupBox1.Visible=true;
			}
			dtFechaIni.Value = DateTime.Today.AddMonths(-1).AddDays(-DateTime.Today.Day+1);
			dtFechaFin.Value = DateTime.Today.AddDays(-DateTime.Today.Day);
			Cursor = Cursors.Default;
			string stAudita = "Exec AuditaCrear 53, 6, 'Activo Fijo'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
		}

		private void ActivoFijo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Act. Fijo'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//				MessageBox.Show(ex.Message, "Tecla");
			}

		}

		private void cmbProyecto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btn_EnviarCI_Click(object sender, System.EventArgs e)
		{
					
			if (MessageBox.Show("¿Desea Enviar Depreciaciones?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No)
			{
				return;
			}
			else
			{
				enviaCostosIndirectos("Depreciacion_CI");				
			}
			


		}
		private void enviaCostosIndirectos(string tipo)
		{
			DateTime dtHasta;
			DateTime dtDesde;
			try
			{
				dtDesde = (DateTime) dtFechaIni.Value;
				dtHasta = (DateTime) dtFechaFin.Value;
			}
			catch(Exception ex)
			{
				MessageBox.Show("Error: "+ex);
				return;
			}
			
			string stExec = "Exec Agro_MovimientoCI '"+tipo+"', '"				
				+dtDesde.ToString("yyyyMMdd")+"', '"
				+dtHasta.ToString("yyyyMMdd")+"',''";
						
			string mesj = Funcion.sEscalarSQL(cdsAsiento, stExec);

			if (mesj.Length > 0)
			{
				MessageBox.Show(mesj,"Información",MessageBoxButtons.OK,MessageBoxIcon.Information);
			}
		}





	}
}
