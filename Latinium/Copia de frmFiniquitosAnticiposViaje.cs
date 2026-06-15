using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Globalization;
using System.Data.SqlClient;
using C1.Data;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmFiniquitosAnticiposViaje.
	/// </summary>
	public class frmFiniquitosAnticiposViaje : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaEntrega;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;

		bool bAnulado = false;
		bool bNuevo = false;
		bool bEdicion = false;
		private C1DataRow drSeteoF;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdRolPrestamo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Button btnConsignacion;
		bool bEliminaAlValidar = false;

		public frmFiniquitosAnticiposViaje()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDesglose");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debe", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debe", 5, true);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Haber", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Haber", 6, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDesglose");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRolPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Haber");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFiniquitosAnticiposViaje));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipoPrest");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRolTipoPrest");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPrestamo");
			this.label8 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbFechaEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdRolPrestamo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnConsignacion = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(641, 11);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(44, 16);
			this.label8.TabIndex = 59;
			this.label8.Text = "Número";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(713, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(96, 22);
			this.txtNumero.TabIndex = 58;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(49, 11);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(26, 16);
			this.label7.TabIndex = 56;
			this.label7.Text = "Tipo";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(49, 74);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 49;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(305, 74);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(30, 16);
			this.label3.TabIndex = 47;
			this.label3.Text = "Valor";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(49, 42);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 16);
			this.label2.TabIndex = 44;
			this.label2.Text = "Nombre";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(49, 129);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 16);
			this.label1.TabIndex = 42;
			this.label1.Text = "Concepto";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource3;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 63;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 79;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 54;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 121;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 260;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.PromptChar = ' ';
			ultraGridColumn6.Width = 80;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.PromptChar = ' ';
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 280;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance4;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance5;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8.5F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 184);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(842, 296);
			this.ultraGrid1.TabIndex = 54;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Double);
			ultraDataColumn7.DataType = typeof(System.Double);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// cmbFechaEntrega
			// 
			this.cmbFechaEntrega.AllowMonthSelection = true;
			this.cmbFechaEntrega.AutoSelectionUpdate = true;
			this.cmbFechaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFechaEntrega.DateButtons.Add(dateButton1);
			this.cmbFechaEntrega.DayDisplayStyle = Infragistics.Win.UltraWinSchedule.DayDisplayStyle.DayNumberAndImage;
			this.cmbFechaEntrega.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.cmbFechaEntrega.Enabled = false;
			this.cmbFechaEntrega.Format = "dd/MM/yyyy";
			this.cmbFechaEntrega.Location = new System.Drawing.Point(137, 72);
			this.cmbFechaEntrega.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.cmbFechaEntrega.Name = "cmbFechaEntrega";
			this.cmbFechaEntrega.NonAutoSizeHeight = 23;
			this.cmbFechaEntrega.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaEntrega.SpinButtonsVisible = true;
			this.cmbFechaEntrega.TabIndex = 48;
			this.cmbFechaEntrega.Value = ((object)(resources.GetObject("cmbFechaEntrega.Value")));
			this.cmbFechaEntrega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbFechaEntrega_KeyDown);
			this.cmbFechaEntrega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFechaEntrega_KeyPress);
			// 
			// txtValor
			// 
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(377, 71);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(96, 22);
			this.txtValor.TabIndex = 46;
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			this.txtValor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValor_KeyPress);
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			// 
			// cmbNombre
			// 
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombre.DataSource = this.ultraDataSource1;
			this.cmbNombre.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Width = 317;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombre.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNombre.Enabled = false;
			this.cmbNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombre.Location = new System.Drawing.Point(137, 40);
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.Size = new System.Drawing.Size(336, 21);
			this.cmbNombre.TabIndex = 45;
			this.cmbNombre.ValueMember = "idPersonal";
			this.cmbNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbNombre_KeyPress);
			this.cmbNombre.Validating += new System.ComponentModel.CancelEventHandler(this.cmbNombre_Validating);
			this.cmbNombre.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNombre_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			this.ultraDataSource1.Band.Key = "Personal";
			// 
			// txtConcepto
			// 
			this.txtConcepto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Enabled = false;
			this.txtConcepto.Location = new System.Drawing.Point(137, 104);
			this.txtConcepto.Multiline = true;
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(672, 66);
			this.txtConcepto.TabIndex = 43;
			this.txtConcepto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtConcepto_KeyPress);
			// 
			// cmbTipo
			// 
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.ultraDataSource2;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 88;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 245;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipo.DisplayMember = "TipoPrestamo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Enabled = false;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(137, 9);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(264, 21);
			this.cmbTipo.TabIndex = 61;
			this.cmbTipo.ValueMember = "idRolTipoPrest";
			this.cmbTipo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTipo_KeyPress);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			this.ultraDataSource2.Band.Key = "Personal";
			// 
			// groupBox1
			// 
			this.groupBox1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.groupBox1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.groupBox1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.groupBox1.Location = new System.Drawing.Point(0, 488);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(864, 8);
			this.groupBox1.TabIndex = 60;
			this.groupBox1.TabStop = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 504);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 260;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(84, 504);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 259;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(166, 504);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 258;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(246, 504);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 257;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(326, 504);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 256;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(482, 504);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 255;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(558, 504);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 254;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
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
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(513, 69);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 264;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(816, 504);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsiento.TabIndex = 267;
			this.txtIdAsiento.Visible = false;
			// 
			// txtIdRolPrestamo
			// 
			this.txtIdRolPrestamo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdRolPrestamo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdRolPrestamo.Enabled = false;
			this.txtIdRolPrestamo.Location = new System.Drawing.Point(800, 504);
			this.txtIdRolPrestamo.Name = "txtIdRolPrestamo";
			this.txtIdRolPrestamo.PromptChar = ' ';
			this.txtIdRolPrestamo.Size = new System.Drawing.Size(16, 22);
			this.txtIdRolPrestamo.TabIndex = 266;
			this.txtIdRolPrestamo.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(832, 504);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 265;
			this.txtEstado.Visible = false;
			// 
			// btnConsignacion
			// 
			this.btnConsignacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsignacion.CausesValidation = false;
			this.btnConsignacion.Enabled = false;
			this.btnConsignacion.Image = ((System.Drawing.Image)(resources.GetObject("btnConsignacion.Image")));
			this.btnConsignacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsignacion.Location = new System.Drawing.Point(406, 504);
			this.btnConsignacion.Name = "btnConsignacion";
			this.btnConsignacion.Size = new System.Drawing.Size(72, 23);
			this.btnConsignacion.TabIndex = 268;
			this.btnConsignacion.Text = "&Consignar";
			this.btnConsignacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsignacion.Click += new System.EventHandler(this.btnConsignacion_Click);
			// 
			// frmFiniquitosAnticiposViaje
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(858, 534);
			this.Controls.Add(this.btnConsignacion);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtIdRolPrestamo);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.cmbFechaEntrega);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.txtConcepto);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmFiniquitosAnticiposViaje";
			this.Text = "Finiquitos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFiniquitosAnticiposViaje_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmFiniquitosAnticiposViaje_Closing);
			this.Load += new System.EventHandler(this.frmFiniquitosAnticiposViaje_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		int idBloqueaTransacciones = 0;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}
		
		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (!bNuevo && bEdicion) 
				FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, false, (int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF);

			bAnulado = false;
			bNuevo = false;
			bEdicion = false;

			this.txtIdRolPrestamo.Value = 0;
			this.txtIdAsiento.Value = 0;
			this.cmbTipo.Value = System.DBNull.Value;
			this.txtNumero.Text = "";
			this.cmbNombre.Value = System.DBNull.Value;
			this.cmbFechaEntrega.Value = System.DBNull.Value;
			this.txtValor.Value = 0.00m;
			this.txtConcepto.Text = "";
			this.txtEstado.Value = 0;
			this.lblEstado.Text = "";

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDesgloseFiniquitos 0");

			this.cmbNombre.Enabled = false;
			this.cmbFechaEntrega.Enabled = false;
			this.txtValor.Enabled = false;
			this.txtConcepto.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnConsignacion.Enabled = false;
			this.btnCancelar.Enabled = false;

			idBloqueaTransacciones = 0;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			bNuevo = true;
			bEdicion = true;

			this.cmbTipo.Value = 9;

			this.cmbFechaEntrega.Value = DateTime.Today;

			this.cmbNombre.Enabled = true;
			if (miAcceso.CambiarFecha) this.cmbFechaEntrega.Enabled = true;
			this.txtValor.Enabled = true;
			this.txtConcepto.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;

			this.cmbNombre.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaEmpleadosAnticiposPrestamos '20301231', 9, 1");

			this.cmbNombre.Focus();
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmFiniquitosAnticiposViaje_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0520");
				
			if (!Funcion.Permiso("790", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Finiquitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
				return;
			}

			this.cmbTipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idRolTipoPrest, TipoPrestamo From RolTipoPrestamo Where idRolTipoPrest = 9");
			
			this.cmbFechaEntrega.CalendarInfo.MinDate = DateTime.Today.AddDays(-45);
			this.cmbFechaEntrega.CalendarInfo.MaxDate = DateTime.Today;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
		}

		private void cmbTipo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbNombre.Focus();
		}

		private void Consulta (int idRolPrestamo)
		{
			try
			{
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;

				string sSQL = string.Format("Exec ConsultaIndividualFiniquitos {0}", idRolPrestamo);
				SqlDataReader drFiniquitos = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drFiniquitos.Read();
				if (drFiniquitos.HasRows)
				{
					this.txtIdRolPrestamo.Value = idRolPrestamo;
					this.txtIdAsiento.Value = drFiniquitos.GetInt32(1);
					this.cmbTipo.Value = drFiniquitos.GetInt32(2);
					this.txtNumero.Text = drFiniquitos.GetValue(3).ToString();
					this.cmbNombre.Value = drFiniquitos.GetInt32(4);
					this.cmbFechaEntrega.Value = drFiniquitos.GetDateTime(5);
					this.txtValor.Value = drFiniquitos.GetDecimal(6);
					this.txtConcepto.Text = drFiniquitos.GetValue(7).ToString();
					this.txtEstado.Value = drFiniquitos.GetInt32(8);
					this.lblEstado.Text = drFiniquitos.GetValue(9).ToString();
				}
				drFiniquitos.Close();
			
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDesgloseFiniquitos {0}", (int)this.txtIdRolPrestamo.Value));

				this.cmbNombre.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaEmpleadosAnticiposPrestamos '20301231', 9, 0");

				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

				if ((int)this.txtEstado.Value == 0)
				{
					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					if (miAcceso.Anular) this.btnAnular.Enabled = true;
				}

				if ((int)this.txtEstado.Value == 2) if (miAcceso.Retencion) this.btnConsignacion.Enabled = true;

				this.btnCancelar.Enabled = true;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaPrestamos miBusqueda = new frmBuscaPrestamos(0))
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					this.Consulta((int)miBusqueda.ultraGrid1.ActiveRow.Cells["idRolPrestamo"].Value);					
				}
			}		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdRolPrestamo.Value);

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("Finiquito Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstado.Value == 1)
			{
				MessageBox.Show("Finiquito Autorizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("Finiquito Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if (Funcion.ValidaBloqueo((int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF)) return;

			if (miAcceso.CambiarFecha) this.cmbFechaEntrega.Enabled = true;

			this.txtValor.Enabled = true;
			this.txtConcepto.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			
			bNuevo = false;
			bEdicion = true;

			FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, true, (int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF);

			this.cmbTipo.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbNombre, "Seleccione el nombre del empleado")) return;
			
			if (!FuncionesContabilidad.bValidaAnticipoDuplicado(this.cmbTipo, this.cmbNombre, (int)this.txtIdRolPrestamo.Value, cdsSeteoF)) return;

			if (!Validacion.vbFechaEnDocumentos(this.cmbFechaEntrega, "Seleccione una fecha", false, "BANCOS", cdsSeteoF)) return;
			
			string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec FiniquitosValidaFechaSalida '{0}', {1}", Convert.ToDateTime(this.cmbFechaEntrega.Value).ToString("yyyyMMdd"), (int)this.cmbNombre.Value));
			if (sMensaje.Length > 0)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!Validacion.vbCampoDecimalVacio(this.txtValor, "Ingrese el valor", 1.00m, 10000.00m)) return;	
			
			if (!Validacion.vbTexto(this.txtConcepto, 10, 300, "Concepto")) return;
			
			#region Valida Grid
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese el desglose de cuentas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
				return;
			}

			decimal dValorBanco = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["idCuenta"].Value == System.DBNull.Value || dr.Cells["Codigo"].Value == System.DBNull.Value || dr.Cells["Descripcion"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}

				if (Convert.ToDecimal(dr.Cells["Debe"].Value) == 0.00m && Convert.ToDecimal(dr.Cells["Haber"].Value) == 0.00m)
				{
					MessageBox.Show("Ingrese un valor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}

				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CuentaCorriente Where idCuenta = {0}", (int)dr.Cells["idCuenta"].Value)) > 0)
				{
					if (Convert.ToDecimal(dr.Cells["Debe"].Value) > 0.00m)
					{
						MessageBox.Show("El valor en la cuenta bancaria va al Haber", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}

					dValorBanco = Convert.ToDecimal(dr.Cells["Haber"].Value);
				}
			}

			decimal dDebe = 0.00m;
			decimal dHaber = 0.00m;			

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["Debe"].Value != System.DBNull.Value) dDebe += Convert.ToDecimal(dr.Cells["Debe"].Value);
				if (dr.Cells["Haber"].Value != System.DBNull.Value) dHaber += Convert.ToDecimal(dr.Cells["Haber"].Value);				
			}
						
			if (Math.Round(dDebe, 2) != Math.Round(dHaber, 2)) 
			{
				decimal dDiferencia = dDebe - dHaber;
				
				MessageBox.Show("Registro Descuadrado, \n\nDiferencia : " + dDiferencia.ToString("#,##0.00"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

				return;
			}

			if (Math.Round(dValorBanco, 2) != Math.Round(Convert.ToDecimal(txtValor.Value), 2)) 
			{
				MessageBox.Show("El valor del Finiquito debe coincidir al valor de la cuenta bancaria", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtValor.Focus();
				return;
			}
			#endregion Valida Grid

			if (bNuevo) this.txtNumero.Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select dbo.SecuencialPrestamoEmpleados({0})", (int)this.cmbTipo.Value));

			string sSQL = string.Format("Exec GuardaRolPrestamo {0}, {1}, '{2}', 2, {3}, '{4}', 0, 0, {5}, '{6}', 0", 
				(int)this.txtIdRolPrestamo.Value, (int)this.cmbTipo.Value, this.txtNumero.Text.ToString(), (int)this.cmbNombre.Value, Convert.ToDateTime(this.cmbFechaEntrega.Value).ToString("yyyyMMdd"), 
				Convert.ToDecimal(this.txtValor.Value), this.txtConcepto.Text.ToString());
			this.txtIdRolPrestamo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				sSQL = string.Format("Exec GuardaDesgloseFiniquitos {0}, {1}, {2}, {3}, {4}, '{5}'", 
					(int)dr.Cells["idDesglose"].Value, (int)this.txtIdRolPrestamo.Value, (int)dr.Cells["idCuenta"].Value, 
					Convert.ToDecimal(dr.Cells["Debe"].Value), Convert.ToDecimal(dr.Cells["Haber"].Value), dr.Cells["Detalle"].Value.ToString());
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDesgloseFiniquitos {0}", (int)this.txtIdRolPrestamo.Value));
	
			string sSQLAut = string.Format("Exec CreaAutorizacionesDePago {0}, {1}, '{2}', '{3}', {4}, 2, '1 DE 1', '{5}', 0", 
				(int)this.cmbTipo.Value, (int)this.txtIdRolPrestamo.Value, Convert.ToDateTime(this.cmbFechaEntrega.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.cmbFechaEntrega.Value).ToString("yyyyMMdd"), 
				Convert.ToDecimal(this.txtValor.Value), this.txtConcepto.Text.ToString());
			Funcion.EjecutaSQL(cdsSeteoF, sSQLAut, true);

			FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, false, (int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF);

			#region Controles
			this.cmbNombre.Enabled = false;
			this.cmbFechaEntrega.Enabled = false;
			this.txtValor.Enabled = false;
			this.txtConcepto.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);		
			
			bNuevo = false;
			bEdicion = false;

			this.Consulta((int)this.txtIdRolPrestamo.Value);	
			#endregion Controles
		}

		private void cmbNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbFechaEntrega.Focus();
		}

		private void cmbFechaEntrega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
			if (e.KeyChar == 13) this.txtValor.Focus();
		}

		private void txtValor_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
			if (e.KeyChar == 13) this.txtConcepto.Focus();
		}

		private void txtConcepto_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Debe", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Haber", 2);
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			#region F3
			if (e.KeyCode == Keys.F3)
			{
				if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" && bEdicion)
				{
					using (BuscaCuentas miBusqueda = new BuscaCuentas("", 1))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{								
							this.ultraGrid1.ActiveRow.Cells["idCuenta"].Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["idCuenta"].Value;
							this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
							this.ultraGrid1.ActiveRow.Cells["Descripcion"].Value = miBusqueda.ultraGrid1.ActiveRow.Cells["Descripcion"].Value.ToString();
						}
					}
				}										
			}
			#endregion F3

			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if (!bEliminaAlValidar)
				{
					for (int i = 0; i < e.Rows.Length; i++)
					{													
						if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar las lineas seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							if ((int)e.Rows[i].Cells["idDesglose"].Value == 0) e.DisplayPromptMsg = false;
							else
							{
								string sSQL = string.Format("Delete From DesgloseDeCuentas Where idDesglose = {0}", (int)e.Rows[i].Cells["idDesglose"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

								e.DisplayPromptMsg = false;

								this.btnCancelar.Enabled = false;
							}
						}
						else e.Cancel = true;						
					}
				}
				else e.DisplayPromptMsg = false;
			}			
		}

		private void txtValor_Enter(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDesglose"].Value = 0;
				e.Row.Cells["idRolPrestamo"].Value = 0;
				e.Row.Cells["Debe"].Value = 0.00m;
				e.Row.Cells["Haber"].Value = 0.00m;
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdRolPrestamo.Value);

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("Finiquito Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (!miAcceso.Anular)
			{
				MessageBox.Show("No tiene acceso a Anular", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstado.Value == 1)
			{
				MessageBox.Show("Finiquito Autorizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("Finiquito Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if (Funcion.ValidaBloqueo((int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF)) return;

			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Anular este Finiquito", "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{				
				using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
				{
					if (DialogResult.OK == miNota.ShowDialog())
					{
						try
						{
							string sSQL = string.Format("Exec AnulaFiniquitos {0}, {1}, {2}, {3}, '{4}'", 
								(int)this.txtIdRolPrestamo.Value, (int)this.txtIdAsiento.Value, (int)this.cmbTipo.Value, (int)this.txtEstado.Value, miNota.txtNotas.Text.ToString());
							Funcion.EjecutaSQL(cdsSeteoF, sSQL);

							MessageBox.Show("Registro Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

							this.Consulta((int)this.txtIdRolPrestamo.Value);
						}
						catch(Exception Exc)
						{
							MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
			}
		}

		private void cmbNombre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			#region Valida duplicado comentado
			/*
			if (bEdicion)
			{
				if (bNuevo)
				{
					if (this.cmbNombre.ActiveRow != null)
					{						
						int idRP = 0;
						string sMensaje = "";
						bool bEditable = false;

						string sSQL = string.Format("Exec ValidaFiniquitos {0}", (int)this.cmbNombre.Value);
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
						dr.Read();
						if (dr.HasRows)
						{
							idRP = dr.GetInt32(0);
							sMensaje = dr.GetString(1);
							bEditable = dr.GetBoolean(2);
						}
						dr.Close();	

						if (idRP > 0)
						{
							if (bEditable)
							{
								if (DialogResult.Yes == MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
								{
									this.Consulta(idRP);
									bNuevo = false;
									this.cmbNombre.Enabled = false;
								}
								else e.Cancel = true;
							}
							else
							{
								MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
								e.Cancel = true;
							}
						}
					}
				}
			}
			*/
			#endregion Valida duplicado comentado
		}

		private void frmFiniquitosAnticiposViaje_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdRolPrestamo.Value > 0)
				{
					if (e.KeyCode == Keys.E) if ((int)this.txtEstado.Value == 0) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.A) if ((int)this.txtEstado.Value == 0) if (this.btnAnular.Enabled) this.btnAnular_Click(sender, e);					
					if (e.KeyCode == Keys.F5) this.Consulta((int)this.txtIdRolPrestamo.Value);					
				}
				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.ultraGrid1.ActiveRow.Cells["idCuenta"].Value == System.DBNull.Value && (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "DEBE" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "HABER" )) e.Handled = true;

			int iCuentaBancaria = 0;
			if (this.ultraGrid1.ActiveRow.Cells["idCuenta"].Value != System.DBNull.Value)
				iCuentaBancaria = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CuentaCorriente Where idCuenta = {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCuenta"].Value), true);
			
			if (iCuentaBancaria > 0 && this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "DEBE") e.Handled = true;

			if ((decimal)this.ultraGrid1.ActiveRow.Cells["Debe"].Value > 0 && this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "HABER") e.Handled = true;
			if ((decimal)this.ultraGrid1.ActiveRow.Cells["Haber"].Value > 0 && this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "DEBE") e.Handled = true;
		}

		private void cmbFechaEntrega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
		}

		private void txtValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
		}

		private void frmFiniquitosAnticiposViaje_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un Finiquito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void cmbNombre_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnConsignacion_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Consignar este Finiquito", "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update RolPrestamo Set Estado = 4 Where idRolPrestamo = {0}", (int)this.txtIdRolPrestamo.Value));

				this.Consulta((int)this.txtIdRolPrestamo.Value);
			}
		}
	}
}

