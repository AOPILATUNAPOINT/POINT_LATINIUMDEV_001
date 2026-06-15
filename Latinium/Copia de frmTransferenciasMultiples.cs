using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmTransferenciasMultiples.
	/// </summary>
	public class frmTransferenciasMultiples : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label61;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.SimpleButton btnCalcular;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRangoDias;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridTransferencias;
		private DevExpress.XtraEditors.SimpleButton btTransferir;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridExistencias;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVentas;
		private DevExpress.XtraEditors.SimpleButton btnExcel;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasAbastecimiento;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridResumen;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblEstado;
		private DevExpress.XtraEditors.SimpleButton btnBuscar;
		private System.Windows.Forms.ToolTip toolTip1;
		private DevExpress.XtraEditors.SimpleButton btnActualizaSaldo;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocales;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodos;
		private System.ComponentModel.IContainer components;

		public frmTransferenciasMultiples()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTransferenciasMultiples));
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UltimaTransferencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadTransferido");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasUltTransf");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExistenciaAnterior");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Existencia");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaDiaria");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Requerido");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RequeridoExist");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Requerido", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Requerido", 11, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RequeridoExist", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RequeridoExist", 12, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cantidad", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cantidad", 13, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nuevo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Peso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UltimaTransferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadTransferido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasUltTransf");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ExistenciaAnterior");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Existencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaDiaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Requerido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RequeridoExist");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			this.uGridTransferencias = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label1 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label61 = new System.Windows.Forms.Label();
			this.txtRangoDias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnCalcular = new DevExpress.XtraEditors.SimpleButton();
			this.btTransferir = new DevExpress.XtraEditors.SimpleButton();
			this.uGridExistencias = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridVentas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnExcel = new DevExpress.XtraEditors.SimpleButton();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.label3 = new System.Windows.Forms.Label();
			this.txtDiasAbastecimiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.uGridResumen = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblEstado = new System.Windows.Forms.Label();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btnActualizaSaldo = new DevExpress.XtraEditors.SimpleButton();
			this.uGridLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkTodos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.uGridTransferencias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRangoDias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridExistencias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasAbastecimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridResumen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridTransferencias
			// 
			this.uGridTransferencias.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridTransferencias.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTransferencias.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridTransferencias.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTransferencias.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridTransferencias.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridTransferencias.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridTransferencias.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.uGridTransferencias.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 6.75F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTransferencias.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridTransferencias.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransferencias.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransferencias.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransferencias.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridTransferencias.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTransferencias.Location = new System.Drawing.Point(240, 80);
			this.uGridTransferencias.Name = "uGridTransferencias";
			this.uGridTransferencias.Size = new System.Drawing.Size(920, 224);
			this.uGridTransferencias.TabIndex = 160;
			this.uGridTransferencias.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridTransferencias_BeforeSelectChange);
			this.uGridTransferencias.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridTransferencias_KeyDown);
			this.uGridTransferencias.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridTransferencias_AfterCellUpdate);
			this.uGridTransferencias.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridTransferencias_InitializeLayout);
			this.uGridTransferencias.Click += new System.EventHandler(this.uGridTransferencias_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 16);
			this.label1.TabIndex = 180;
			this.label1.Text = "Articulo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance8;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(64, 40);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(304, 22);
			this.txtBuscar.TabIndex = 0;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// btnBuscar
			// 
			this.btnBuscar.Location = new System.Drawing.Point(376, 41);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(80, 21);
			this.btnBuscar.TabIndex = 1;
			this.btnBuscar.Text = "Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 183;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance9;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 182;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// label61
			// 
			this.label61.AutoSize = true;
			this.label61.BackColor = System.Drawing.Color.Transparent;
			this.label61.Location = new System.Drawing.Point(200, 10);
			this.label61.Name = "label61";
			this.label61.Size = new System.Drawing.Size(76, 16);
			this.label61.TabIndex = 185;
			this.label61.Text = "Rango de Dias";
			this.label61.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRangoDias
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRangoDias.Appearance = appearance10;
			this.txtRangoDias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRangoDias.FormatString = "#,##0";
			this.txtRangoDias.Location = new System.Drawing.Point(288, 7);
			this.txtRangoDias.Name = "txtRangoDias";
			this.txtRangoDias.PromptChar = ' ';
			this.txtRangoDias.Size = new System.Drawing.Size(72, 22);
			this.txtRangoDias.TabIndex = 2;
			this.txtRangoDias.Click += new System.EventHandler(this.txtRangoDias_Click);
			this.txtRangoDias.Enter += new System.EventHandler(this.txtRangoDias_Enter);
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
			// btnCalcular
			// 
			this.btnCalcular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCalcular.Location = new System.Drawing.Point(995, 8);
			this.btnCalcular.Name = "btnCalcular";
			this.btnCalcular.Size = new System.Drawing.Size(80, 21);
			this.btnCalcular.TabIndex = 4;
			this.btnCalcular.Text = "Calcular";
			this.btnCalcular.Click += new System.EventHandler(this.btnCalcular_Click);
			// 
			// btTransferir
			// 
			this.btTransferir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btTransferir.Location = new System.Drawing.Point(1083, 41);
			this.btTransferir.Name = "btTransferir";
			this.btTransferir.Size = new System.Drawing.Size(80, 21);
			this.btTransferir.TabIndex = 7;
			this.btTransferir.Text = "&Transferir";
			this.btTransferir.Click += new System.EventHandler(this.btTransferir_Click);
			// 
			// uGridExistencias
			// 
			this.uGridExistencias.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridExistencias.Cursor = System.Windows.Forms.Cursors.Default;
			appearance11.BackColor = System.Drawing.Color.White;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridExistencias.DisplayLayout.Appearance = appearance11;
			ultraGridBand2.Header.Caption = "Existencia";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridExistencias.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridExistencias.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridExistencias.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.uGridExistencias.DisplayLayout.Override.CardAreaAppearance = appearance12;
			this.uGridExistencias.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 10F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridExistencias.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.uGridExistencias.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance14.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance14.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridExistencias.DisplayLayout.Override.RowAlternateAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridExistencias.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridExistencias.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.uGridExistencias.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridExistencias.Location = new System.Drawing.Point(248, 384);
			this.uGridExistencias.Name = "uGridExistencias";
			this.uGridExistencias.Size = new System.Drawing.Size(896, 88);
			this.uGridExistencias.TabIndex = 190;
			// 
			// uGridVentas
			// 
			this.uGridVentas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVentas.Cursor = System.Windows.Forms.Cursors.Default;
			appearance17.BackColor = System.Drawing.Color.White;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentas.DisplayLayout.Appearance = appearance17;
			ultraGridBand3.Header.Caption = "Ventas";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentas.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridVentas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.uGridVentas.DisplayLayout.Override.CardAreaAppearance = appearance18;
			this.uGridVentas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 10F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVentas.DisplayLayout.Override.HeaderAppearance = appearance19;
			this.uGridVentas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance20.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance20.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.RowAlternateAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.uGridVentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVentas.Location = new System.Drawing.Point(240, 360);
			this.uGridVentas.Name = "uGridVentas";
			this.uGridVentas.Size = new System.Drawing.Size(904, 368);
			this.uGridVentas.TabIndex = 191;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Location = new System.Drawing.Point(995, 41);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(80, 21);
			this.btnExcel.TabIndex = 6;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(376, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(105, 16);
			this.label3.TabIndex = 194;
			this.label3.Text = "Dias Abastecimiento";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasAbastecimiento
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasAbastecimiento.Appearance = appearance23;
			this.txtDiasAbastecimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasAbastecimiento.FormatString = "#,##0";
			this.txtDiasAbastecimiento.Location = new System.Drawing.Point(496, 7);
			this.txtDiasAbastecimiento.Name = "txtDiasAbastecimiento";
			this.txtDiasAbastecimiento.PromptChar = ' ';
			this.txtDiasAbastecimiento.Size = new System.Drawing.Size(72, 22);
			this.txtDiasAbastecimiento.TabIndex = 3;
			this.txtDiasAbastecimiento.Click += new System.EventHandler(this.txtDiasAbastecimiento_Click);
			this.txtDiasAbastecimiento.Enter += new System.EventHandler(this.txtDiasAbastecimiento_Enter);
			// 
			// uGridResumen
			// 
			this.uGridResumen.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridResumen.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridResumen.DataSource = this.ultraDataSource1;
			appearance24.BackColor = System.Drawing.Color.White;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridResumen.DisplayLayout.Appearance = appearance24;
			this.uGridResumen.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 164;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Estado";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 67;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Local";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 134;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance25;
			ultraGridColumn4.Format = "#,##0";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 38;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "U Transf";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 59;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance26;
			ultraGridColumn6.Format = "#,##0";
			ultraGridColumn6.Header.Caption = "C Transf";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 47;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance27;
			ultraGridColumn7.Format = "#,##0";
			ultraGridColumn7.Header.Caption = "D. U. T.";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 39;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance28;
			ultraGridColumn8.Format = "#,##0";
			ultraGridColumn8.Header.Caption = "E Ant";
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Width = 48;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance29;
			ultraGridColumn9.Format = "#,##0";
			ultraGridColumn9.Header.Caption = "Exist";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 39;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance30;
			ultraGridColumn10.Format = "#,##0";
			ultraGridColumn10.Header.Caption = "Venta";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 39;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance31;
			ultraGridColumn11.Format = "#,##0.00000000000";
			ultraGridColumn11.Header.Caption = "Venta Diaria";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 58;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance32;
			ultraGridColumn12.Format = "#,##0";
			ultraGridColumn12.Header.Caption = "Necesidad";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 59;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance33;
			ultraGridColumn13.Format = "#,##0";
			ultraGridColumn13.Header.Caption = "Requer";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 51;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance34;
			ultraGridColumn14.Format = "#,##0";
			ultraGridColumn14.Header.Caption = "Transf";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 57;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn14});
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance35;
			summarySettings1.DisplayFormat = "{0: #,##0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance36;
			summarySettings2.DisplayFormat = "{0: #,##0}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance37;
			summarySettings3.DisplayFormat = "{0: #,##0}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridResumen.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance38.ForeColor = System.Drawing.Color.Black;
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridResumen.DisplayLayout.Override.ActiveRowAppearance = appearance38;
			this.uGridResumen.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridResumen.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance39.BackColor = System.Drawing.Color.Transparent;
			this.uGridResumen.DisplayLayout.Override.CardAreaAppearance = appearance39;
			this.uGridResumen.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance40.FontData.BoldAsString = "True";
			appearance40.FontData.Name = "Arial";
			appearance40.FontData.SizeInPoints = 8.25F;
			appearance40.ForeColor = System.Drawing.Color.White;
			appearance40.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridResumen.DisplayLayout.Override.HeaderAppearance = appearance40;
			this.uGridResumen.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance41.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance41.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumen.DisplayLayout.Override.RowAlternateAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumen.DisplayLayout.Override.RowSelectorAppearance = appearance42;
			appearance43.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance43.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumen.DisplayLayout.Override.SelectedRowAppearance = appearance43;
			this.uGridResumen.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridResumen.Location = new System.Drawing.Point(240, 312);
			this.uGridResumen.Name = "uGridResumen";
			this.uGridResumen.Size = new System.Drawing.Size(920, 416);
			this.uGridResumen.TabIndex = 195;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn4.DataType = typeof(System.Double);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn6.DataType = typeof(System.Double);
			ultraDataColumn7.DataType = typeof(System.Double);
			ultraDataColumn8.DataType = typeof(System.Double);
			ultraDataColumn9.DataType = typeof(System.Double);
			ultraDataColumn10.DataType = typeof(System.Double);
			ultraDataColumn11.DataType = typeof(System.Double);
			ultraDataColumn12.DataType = typeof(System.Double);
			ultraDataColumn13.DataType = typeof(System.Double);
			ultraDataColumn14.DataType = typeof(System.Double);
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
																																 ultraDataColumn14});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 64);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1171, 8);
			this.groupBox1.TabIndex = 196;
			this.groupBox1.TabStop = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(480, 42);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 19);
			this.lblEstado.TabIndex = 241;
			// 
			// btnActualizaSaldo
			// 
			this.btnActualizaSaldo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizaSaldo.Location = new System.Drawing.Point(1083, 8);
			this.btnActualizaSaldo.Name = "btnActualizaSaldo";
			this.btnActualizaSaldo.Size = new System.Drawing.Size(80, 21);
			this.btnActualizaSaldo.TabIndex = 5;
			this.btnActualizaSaldo.Text = "Actualizar";
			this.btnActualizaSaldo.Click += new System.EventHandler(this.btnActualizaSaldo_Click);
			// 
			// uGridLocales
			// 
			this.uGridLocales.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocales.DataSource = this.ultraDataSource2;
			appearance44.BackColor = System.Drawing.Color.White;
			this.uGridLocales.DisplayLayout.Appearance = appearance44;
			this.uGridLocales.DisplayLayout.AutoFitColumns = true;
			ultraGridBand5.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 97;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 178;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Width = 25;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			ultraGridBand5.Header.Caption = "LOCALES";
			ultraGridBand5.HeaderVisible = true;
			this.uGridLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance45.ForeColor = System.Drawing.Color.Black;
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocales.DisplayLayout.Override.ActiveRowAppearance = appearance45;
			this.uGridLocales.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLocales.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLocales.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance46.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocales.DisplayLayout.Override.CardAreaAppearance = appearance46;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance47.FontData.BoldAsString = "True";
			appearance47.FontData.Name = "Arial";
			appearance47.FontData.SizeInPoints = 8.5F;
			appearance47.ForeColor = System.Drawing.Color.White;
			appearance47.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocales.DisplayLayout.Override.HeaderAppearance = appearance47;
			appearance48.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance48.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowAlternateAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowSelectorAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.SelectedRowAppearance = appearance50;
			this.uGridLocales.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocales.Location = new System.Drawing.Point(8, 80);
			this.uGridLocales.Name = "uGridLocales";
			this.uGridLocales.Size = new System.Drawing.Size(224, 648);
			this.uGridLocales.TabIndex = 647;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(bool);
			ultraDataColumn17.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// chkTodos
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkTodos.Appearance = appearance51;
			this.chkTodos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkTodos.Enabled = false;
			this.chkTodos.Location = new System.Drawing.Point(208, 88);
			this.chkTodos.Name = "chkTodos";
			this.chkTodos.Size = new System.Drawing.Size(13, 13);
			this.chkTodos.TabIndex = 654;
			this.chkTodos.CheckedChanged += new System.EventHandler(this.chkTodos_CheckedChanged);
			// 
			// frmTransferenciasMultiples
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1168, 734);
			this.Controls.Add(this.chkTodos);
			this.Controls.Add(this.uGridLocales);
			this.Controls.Add(this.btnActualizaSaldo);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.uGridResumen);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label61);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtDiasAbastecimiento);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.uGridVentas);
			this.Controls.Add(this.uGridExistencias);
			this.Controls.Add(this.btTransferir);
			this.Controls.Add(this.btnCalcular);
			this.Controls.Add(this.txtRangoDias);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.uGridTransferencias);
			this.KeyPreview = true;
			this.Name = "frmTransferenciasMultiples";
			this.Text = "Transferencias Multiples";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmTransferenciasMultiples_KeyDown);
			this.Load += new System.EventHandler(this.frmTransferenciasMultiples_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridTransferencias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRangoDias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridExistencias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasAbastecimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridResumen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmTransferenciasMultiples_Load(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Today;
			this.txtRangoDias.Value = 90;
			this.txtDiasAbastecimiento.Value = 15;

			this.uGridLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select b.Bodega, b.Nombre, Convert(Bit, 0) As Sel From Bodega b Where Inventario = 1 And TransferenciaMultiple = 1 And Sel = 0 Order By Nombre	");

			if (this.uGridLocales.Rows.Count > 1) 
			{
				this.chkTodos.Enabled = true;
				this.uGridLocales.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}

			this.chkTodos.Checked = true;			

		}

		private void btnCalcular_Click(object sender, System.EventArgs e)
		{
			string sLocales = "";		
			this.Cursor = Cursors.WaitCursor;
			
			Funcion.TiempoInicia();

			Resumen(0);

			DateTime dtFecha = (DateTime)this.dtFecha.Value;

			this.uGridTransferencias.DataSource = null;

//			string sSQL = string.Format("Exec CalculoTransferenciasAutomaticas '{0}', {1}, {2}", dtFecha.ToString("yyyyMMdd"), (int)this.txtRangoDias.Value, (int)this.txtDiasAbastecimiento.Value);

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

			string sSQL = string.Format("Exec CalculoTransferenciasAutomaticas '{0}', {1}, {2},'{3}'", dtFecha.ToString("yyyyMMdd"), (int)this.txtRangoDias.Value, (int)this.      txtDiasAbastecimiento.Value, sLocales);	
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridTransferencias);

			DiseñoGrid();

			MessageBox.Show("Calculo Finalizado " + Funcion.Tiempo("EN"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.lblEstado.Text = this.uGridTransferencias.Rows.Count + " ARTICULOS ENCONTRADOS";

			this.Cursor = Cursors.Default;
		}

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridTransferencias.Rows.Count > 0)
			{
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["idCalculoTransferencias"].Hidden = true;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Fecha"].Hidden = true;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Usuario"].Hidden = true;

				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Sel"].Width = 20;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Activation.AllowEdit;

				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Codigo"].Width = 150;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;
				
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["ODC"].Width = 40;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["ODC"].CellActivation = Activation.Disabled;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["ODC"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				//this.uGridTransferencias.DisplayLayout.Bands[0].Columns["ODC"].Hidden = true;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries.Add("ODC", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridTransferencias.DisplayLayout.Bands[0].Columns["ODC"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["ODC"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["ODC"].DisplayFormat = "{0: #,##0}";
				//this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["ODC"].Hidden = true;
				
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["EBI"].Width = 40;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["EBI"].CellActivation = Activation.Disabled;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["EBI"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;				
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries.Add("EBI", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridTransferencias.DisplayLayout.Bands[0].Columns["EBI"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["EBI"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["EBI"].DisplayFormat = "{0: #,##0}";

				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Requerimiento"].Width = 40;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Requerimiento"].Header.Caption = "Req";
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Requerimiento"].CellActivation = Activation.Disabled;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Requerimiento"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;				
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Requerimiento"].Hidden = true;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries.Add("Requerimiento", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Requerimiento"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["Requerimiento"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["Requerimiento"].DisplayFormat = "{0: #,##0}";
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["Requerimiento"].Hidden = true;

				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Total"].Width = 40;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Total"].Header.Caption = "TF";
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.Disabled;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0}";

				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Existencia"].Width = 40;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Existencia"].Header.Caption = "Exis";
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Existencia"].CellActivation = Activation.Disabled;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Existencia"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries.Add("Existencia", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Existencia"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["Existencia"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["Existencia"].DisplayFormat = "{0: #,##0}";
				
				if (this.uGridTransferencias.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.uGridTransferencias.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 11; j < iColumnas; j++)
					{
						nombrCol = this.uGridTransferencias.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						
						if (nombrCol.Length == 2)
						{
							this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].Width = 25;
							this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.AllowEdit;													
						}
						else
						{
							this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].Width = 25;
//							this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol.Substring(0, 2);
							this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.Disabled;							
							this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LightBlue;
							this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor2 = Color.LightBlue;
							this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColorDisabled = Color.LightBlue;
							this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColorDisabled2 = Color.LightBlue;
						}

						this.uGridTransferencias.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						this.uGridTransferencias.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridTransferencias.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
					}
				}				
			}
			#endregion Diseño Grid
		}

		private void uGridTransferencias_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void Resumen(int idArticulo)
		{
			string sSQLVentas = string.Format("Exec ResumenCalculoTransfAut {0}", idArticulo);
			this.uGridResumen.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLVentas);			
		}

		private void uGridTransferencias_Click(object sender, System.EventArgs e)
		{
			Resumen((int)this.uGridTransferencias.ActiveRow.Cells["idArticulo"].Value);
			//			DateTime dtFecha = (DateTime)this.dtFecha.Value;
			//
			//			string sSQLExistencias = string.Format("Exec ListaTransferenciasMultiples TransferenciasExistenciaPorArticulo {0}, {1}, '{2}'", (int)this.uGridTransferencias.ActiveRow.Cells["idArticulo"].Value, (int)this.txtRangoDias.Value, dtFecha.ToString("yyyyMMdd"));
			//			this.uGridExistencias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLExistencias);
			//
			//			string sSQLVentas = string.Format("Exec TransferenciasVendidoPorArticulo {0}, {1}, '{2}'", (int)this.uGridTransferencias.ActiveRow.Cells["idArticulo"].Value, (int)this.txtRangoDias.Value, dtFecha.ToString("yyyyMMdd"));
			//			this.uGridVentas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLVentas);																 
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.uGridTransferencias.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.uGridTransferencias, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

	

		private void uGridTransferencias_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			Resumen((int)this.uGridTransferencias.ActiveRow.Cells["idArticulo"].Value);
		}

		private void frmTransferenciasMultiples_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			this.uGridTransferencias.DataSource = null;

			string sSQL = string.Format("Exec TFMBuscaCodigo '{0}'", this.txtBuscar.Text.ToString());
			this.uGridTransferencias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			DiseñoGrid();	
		
			this.lblEstado.Text = this.uGridTransferencias.Rows.Count + " ARTICULOS ENCONTRADOS";
		}

		private void btTransferir_Click(object sender, System.EventArgs e)
		{
			bool bValidaExistencia = true;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridTransferencias.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					int iSaldo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaExistencia('{0}', {1}, {2}, {3}, 8, 0)", 
						DateTime.Today.ToString("yyyyMMdd"), 0, 1, (int)dr.Cells["idArticulo"].Value));

					if ((int)dr.Cells["Total"].Value > iSaldo) // (int)dr.Cells["EBI"].Value)
					{
						MessageBox.Show(string.Format("La cantidad a transferir {0} del articulo '{1}' es mayor a la existencia en Bodega Principal {2}", (int)dr.Cells["Total"].Value, dr.Cells["Codigo"].Value.ToString(), iSaldo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.uGridTransferencias.ActiveRow = dr;
						this.uGridTransferencias.ActiveRow.Selected = true;										
						bValidaExistencia = false;
						break;					
					}
				}
			}

			if (!bValidaExistencia)return;

			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de generar las transferencias?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				this.uGridTransferencias.DataSource = null;

				string sSQL = string.Format("Exec TFMBuscaCodigo ''");
				this.uGridTransferencias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				DiseñoGrid();	
		
				this.lblEstado.Text = this.uGridTransferencias.Rows.Count + " ARTICULOS ENCONTRADOS";

				Funcion.TiempoInicia();
				
				Cursor = Cursors.WaitCursor;			

				this.btTransferir.Enabled = false;
				int nColsGrid = (int)this.uGridTransferencias.DisplayLayout.Bands[0].Columns.Count;
				int cini = 11;
				string coltran = "";
				int totColumna = 0;
				int iContFila = 1;
				int swpaso = 1;
				int iContTransf = 1;
				int iContFilas = 1;
				bool bIE = false;
				//string sMensaje = "RESUMEN DE TRANSFERENCIA MULTIPLE : ";

				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
				{
					oConexion.Open();

					SqlTransaction oTransaction = null;
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					oCmdActualiza.Connection = oConexion;
					oCmdActualiza.CommandType = CommandType.Text;
					oCmdActualiza.CommandTimeout = 120;
						
					try
					{							
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;

						#region Crea Transferencias
						for (int j = cini; j < nColsGrid; j++)
						{	
							coltran = this.uGridTransferencias.DisplayLayout.Bands[0].Columns[j].Header.Caption;
							
							if (coltran.Length == 2)
							{
								swpaso = 1;
								iContFila = 1;
								bIE = false;																		
								totColumna = 0;

								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridTransferencias.Rows.All)
									totColumna = totColumna + (int)dr.Cells[coltran].Value;

								if (totColumna > 0)
								{
									int idTranferencia = 0;
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridTransferencias.Rows.All)
									{
										if ((int)dr.Cells[coltran].Value > 0 && (bool)dr.Cells["Sel"].Value)
										{
											#region Encabezado
											if (swpaso == 1)
											{									
												string stExec = string.Format("Exec TransformaIE_TM_Cabecera 'BI', '{0}'", coltran);
												oCmdActualiza.CommandText = stExec;
												idTranferencia = (int)oCmdActualiza.ExecuteScalar();

												swpaso = 0;
												bIE = false;

												//													SqlDataReader drMens = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select ct.Codigo, b.Codigo From CompraTransf ct Inner Join Bodega b On b.Bodega = ct.Bodega1 Where idCompraTransf = {0}", idTranferencia), true);
												//													drMens.Read();
												//													if (drMens.HasRows)
												//													{
												//														sMensaje = sMensaje + "\n\n " + iContTransf.ToString() + " Número : " + drMens.GetValue(0).ToString() + " - Bodega de Destino : " + drMens.GetValue(1).ToString();	
												//													}
												//													drMens.Close();

												iContTransf ++;
											}
											#endregion Encabezado
								
											#region Detalle
											if (swpaso == 0)
											{													
												string sqlTran = string.Format("Exec TransformaIE_TM_Detalle {0}, {1}, {2}", 
													idTranferencia, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells[coltran].Value);
												oCmdActualiza.CommandText = sqlTran;
												oCmdActualiza.ExecuteNonQuery();
																								
												iContFilas ++;

												iContFila ++;

												if (iContFila == 26)
												{
													string stTIE = string.Format("Exec CreaIEBodegaTM {0}", idTranferencia);
													oCmdActualiza.CommandText = stTIE;
													oCmdActualiza.ExecuteNonQuery();
												
													iContFila = iContFila - 1;
													//														sMensaje = sMensaje + " - Cantidad de Artículos : " + iContFila.ToString();
													bIE = true;
													swpaso = 1;
													iContFila = 1;
												}
											}
											#endregion Detalle 
										}
									}

									if (!bIE)
									{
										iContFila = iContFila - 1;
										//											sMensaje = sMensaje + " Cantidad de Artículos : " + iContFila.ToString();
										string stTIE = string.Format("Exec CreaIEBodegaTM {0}", idTranferencia);
										oCmdActualiza.CommandText = stTIE;
										oCmdActualiza.ExecuteNonQuery();											
									}
								}
							}
						}						

						oTransaction.Commit();	
					
						oConexion.Close();
						#endregion Crea Transferencias

						iContTransf = iContTransf - 1;
						iContFilas = iContFilas - 1;

						//sMensaje = sMensaje + "\n\n Total Transferencias : " + iContTransf.ToString() + " - Con " + iContFilas.ToString() + " Filas de Articulos";
 
						MessageBox.Show(string.Format("Proceso Finalizado,\n\n Se Generarón {0} Transferencias {1}", iContTransf.ToString(), Funcion.Tiempo("EN").ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

						Cursor = Cursors.Default;
					}
					catch (Exception ex)
					{
						oTransaction.Rollback();

						Cursor = Cursors.Default;

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}				
				}
			}
		}

		bool bEditaCantidad = false;

		private void uGridTransferencias_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int contador = 11;
			string nombreColumna = "";
			int TColumnas = this.uGridTransferencias.DisplayLayout.Bands[0].Columns.Count;
			int nExistencias = int.Parse(e.Cell.Row.Cells["EBI"].Value.ToString());
			int sumaCol = 0;

			if (e.Cell.Column.ToString() == "Sel")
			{				
				string sSQL = string.Format("Exec ActualizaSelCalculoTransferencias {0}, {1}",
					(int)e.Cell.Row.Cells["idCalculoTransferencias"].Value, (bool)e.Cell.Row.Cells["Sel"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			}
			
			string sColumna = e.Cell.Column.Header.Caption;

			if (bEditaCantidad) return;

			if (e.Cell.Column.ToString() == sColumna)
			{
				bEditaCantidad = true;

				if (e.Cell.Value == System.DBNull.Value) e.Cell.Value = 0;

				for (int j = contador; j < TColumnas; j++)
				{
					nombreColumna = this.uGridTransferencias.DisplayLayout.Bands[0].Columns[j].Header.Caption;
					if (nombreColumna.Length == 2)						
						sumaCol = sumaCol + int.Parse(e.Cell.Row.Cells[nombreColumna].Value.ToString());
				}

				if (sumaCol > (int)e.Cell.Row.Cells["EBI"].Value)
				{
					if ((int)e.Cell.Value > 0)
					{
						e.Cell.Value = (int)e.Cell.OriginalValue;
//						return;
					}
				}

				sumaCol = 0;
				for (int j = contador; j < TColumnas; j++)
				{
					nombreColumna = this.uGridTransferencias.DisplayLayout.Bands[0].Columns[j].Header.Caption;
					if (nombreColumna.Length == 2)
						sumaCol = sumaCol + int.Parse(e.Cell.Row.Cells[nombreColumna].Value.ToString());
				}

				e.Cell.Row.Cells["Total"].Value = sumaCol;

				e.Cell.Row.Cells["Existencia"].Value = (int)e.Cell.Row.Cells["EBI"].Value - sumaCol;

//				if (sumaCol == 0) e.Cell.Row.Cells["Sel"].Value = false;					
//				else e.Cell.Row.Cells["Sel"].Value = true;

				string sSQLActualiza = string.Format("Exec ActualizaCantidadTransferencias {0}, '{1}', {2}, {3}", 
					int.Parse(e.Cell.Row.Cells["idCalculoTransferencias"].Value.ToString()), sColumna, int.Parse(e.Cell.Row.Cells[sColumna].Value.ToString()), int.Parse(e.Cell.Row.Cells["Total"].Value.ToString()));
				Funcion.EjecutaSQL(cdsSeteoF, sSQLActualiza);			

				bEditaCantidad = false;
			}			
		}

		private void btnActualizaSaldo_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show("Esta Seguro de Actualizar Existencias en la Bodega Principal", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				this.uGridTransferencias.DataSource = null;
				
				this.uGridTransferencias.DataSource = Funcion.dvProcedimiento (cdsSeteoF, "Exec TMActualizaExistenciaBI");

				DiseñoGrid();	
			}
		}

		private void uGridTransferencias_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridTransferencias);
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnBuscar_Click(sender, e);
		}

		private void txtRangoDias_Click(object sender, System.EventArgs e)
		{
			this.txtRangoDias.SelectAll();
		}

		private void txtRangoDias_Enter(object sender, System.EventArgs e)
		{
			this.txtRangoDias.SelectAll();
		}

		private void txtDiasAbastecimiento_Click(object sender, System.EventArgs e)
		{
			this.txtDiasAbastecimiento.SelectAll();
		}

		private void txtDiasAbastecimiento_Enter(object sender, System.EventArgs e)
		{
			this.txtDiasAbastecimiento.SelectAll();
		}

		private void chkTodos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if (this.chkTodos.Checked) dr.Cells["Sel"].Value = 1;
				else dr.Cells["Sel"].Value = 0;
			}
		}
	}
}

