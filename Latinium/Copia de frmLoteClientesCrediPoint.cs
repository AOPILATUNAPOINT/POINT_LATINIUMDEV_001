using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmLoteClientesCrediPoint.
	/// </summary>
	public class frmLoteClientesCrediPoint : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblCaducidad;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLotes;
		private System.Windows.Forms.Label lblVoucher;
		private System.Windows.Forms.Label lblCtaCorriente;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.SimpleButton btTXT;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private DevExpress.XtraEditors.SimpleButton btExcel;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLoteAnterior;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLote;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMes;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDia;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmLoteClientesCrediPoint()
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
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoLoteSolidario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapturaLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroLote");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsientoLoteSolidario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapturaLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmLoteClientesCrediPoint));
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			this.uGridLotes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblCaducidad = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btTXT = new DevExpress.XtraEditors.SimpleButton();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.btExcel = new DevExpress.XtraEditors.SimpleButton();
			this.lblVoucher = new System.Windows.Forms.Label();
			this.lblCtaCorriente = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtLoteAnterior = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtLote = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.lblContador = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLoteAnterior)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridLotes
			// 
			this.uGridLotes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridLotes.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLotes.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridLotes.DisplayLayout.Appearance = appearance1;
			this.uGridLotes.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 368;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 44;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 25;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Header.Caption = "Lote";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 100;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "N. Operación";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 120;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 75;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Cédula";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 120;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 379;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridLotes.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.Appearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.BandLabelAppearance = appearance4;
			this.uGridLotes.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridLotes.DisplayLayout.GroupByBox.Prompt = " ";
			appearance5.BackColor = System.Drawing.SystemColors.Control;
			appearance5.BackColor2 = System.Drawing.SystemColors.Control;
			appearance5.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance5.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.PromptAppearance = appearance5;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.uGridLotes.DisplayLayout.Override.CardAreaAppearance = appearance7;
			this.uGridLotes.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLotes.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLotes.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLotes.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridLotes.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridLotes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLotes.Location = new System.Drawing.Point(8, 120);
			this.uGridLotes.Name = "uGridLotes";
			this.uGridLotes.Size = new System.Drawing.Size(840, 304);
			this.uGridLotes.TabIndex = 3;
			this.uGridLotes.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridLotes_BeforeCellUpdate);
			this.uGridLotes.BeforeRowUpdate += new Infragistics.Win.UltraWinGrid.CancelableRowEventHandler(this.uGridLotes_BeforeRowUpdate);
			this.uGridLotes.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridLotes_AfterCellUpdate);
			this.uGridLotes.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridLotes_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn6.DefaultValue = new System.DateTime(((long)(0)));
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// lblCaducidad
			// 
			this.lblCaducidad.AutoSize = true;
			this.lblCaducidad.Location = new System.Drawing.Point(8, 11);
			this.lblCaducidad.Name = "lblCaducidad";
			this.lblCaducidad.Size = new System.Drawing.Size(35, 16);
			this.lblCaducidad.TabIndex = 193;
			this.lblCaducidad.Text = "Desde";
			this.lblCaducidad.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance11;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 192;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 16);
			this.label1.TabIndex = 195;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance12;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(56, 40);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 194;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.ValueChanged += new System.EventHandler(this.dtHasta_ValueChanged);
			// 
			// btTXT
			// 
			this.btTXT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btTXT.Location = new System.Drawing.Point(656, 40);
			this.btTXT.Name = "btTXT";
			this.btTXT.Size = new System.Drawing.Size(104, 22);
			this.btTXT.TabIndex = 197;
			this.btTXT.Text = "&Generar Lote";
			this.btTXT.Click += new System.EventHandler(this.btTXT_Click);
			// 
			// btBuscar
			// 
			this.btBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btBuscar.Location = new System.Drawing.Point(784, 8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(64, 22);
			this.btBuscar.TabIndex = 196;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btExcel
			// 
			this.btExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btExcel.Location = new System.Drawing.Point(768, 40);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(80, 22);
			this.btExcel.TabIndex = 198;
			this.btExcel.Text = "&Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// lblVoucher
			// 
			this.lblVoucher.AutoSize = true;
			this.lblVoucher.BackColor = System.Drawing.Color.Transparent;
			this.lblVoucher.Location = new System.Drawing.Point(184, 11);
			this.lblVoucher.Name = "lblVoucher";
			this.lblVoucher.Size = new System.Drawing.Size(24, 16);
			this.lblVoucher.TabIndex = 203;
			this.lblVoucher.Text = "Mes";
			this.lblVoucher.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCtaCorriente
			// 
			this.lblCtaCorriente.AutoSize = true;
			this.lblCtaCorriente.BackColor = System.Drawing.Color.Transparent;
			this.lblCtaCorriente.Location = new System.Drawing.Point(184, 43);
			this.lblCtaCorriente.Name = "lblCtaCorriente";
			this.lblCtaCorriente.Size = new System.Drawing.Size(20, 16);
			this.lblCtaCorriente.TabIndex = 202;
			this.lblCtaCorriente.Text = "Día";
			this.lblCtaCorriente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(312, 43);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(76, 16);
			this.label2.TabIndex = 207;
			this.label2.Text = "N. Ultimo Lote";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(312, 11);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(40, 16);
			this.label3.TabIndex = 206;
			this.label3.Text = "N. Lote";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// txtLoteAnterior
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLoteAnterior.Appearance = appearance13;
			this.txtLoteAnterior.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtLoteAnterior.Enabled = false;
			this.txtLoteAnterior.FormatString = "";
			this.txtLoteAnterior.Location = new System.Drawing.Point(416, 40);
			this.txtLoteAnterior.Name = "txtLoteAnterior";
			this.txtLoteAnterior.PromptChar = ' ';
			this.txtLoteAnterior.Size = new System.Drawing.Size(96, 22);
			this.txtLoteAnterior.TabIndex = 208;
			// 
			// txtLote
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLote.Appearance = appearance14;
			this.txtLote.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtLote.FormatString = "";
			this.txtLote.Location = new System.Drawing.Point(416, 8);
			this.txtLote.Name = "txtLote";
			this.txtLote.PromptChar = ' ';
			this.txtLote.Size = new System.Drawing.Size(96, 22);
			this.txtLote.TabIndex = 209;
			this.txtLote.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtLote_KeyPress);
			this.txtLote.ValueChanged += new System.EventHandler(this.txtLote_ValueChanged);
			// 
			// txtMes
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMes.Appearance = appearance15;
			this.txtMes.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtMes.FormatString = "";
			this.txtMes.Location = new System.Drawing.Point(224, 8);
			this.txtMes.Name = "txtMes";
			this.txtMes.PromptChar = ' ';
			this.txtMes.Size = new System.Drawing.Size(64, 22);
			this.txtMes.TabIndex = 210;
			// 
			// txtDia
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDia.Appearance = appearance16;
			this.txtDia.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtDia.FormatString = "";
			this.txtDia.Location = new System.Drawing.Point(224, 40);
			this.txtDia.Name = "txtDia";
			this.txtDia.PromptChar = ' ';
			this.txtDia.Size = new System.Drawing.Size(64, 22);
			this.txtDia.TabIndex = 211;
			// 
			// txtBusqueda
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance17;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(592, 8);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(176, 22);
			this.txtBusqueda.TabIndex = 622;
			this.txtBusqueda.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBusqueda_KeyPress);
			this.txtBusqueda.ValueChanged += new System.EventHandler(this.txtBusqueda_ValueChanged);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(536, 11);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(38, 16);
			this.label4.TabIndex = 623;
			this.label4.Text = "Cliente";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(8, 72);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 22);
			this.lblContador.TabIndex = 624;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(592, 72);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 22);
			this.lblEstado.TabIndex = 634;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(856, 8);
			this.groupBox1.TabIndex = 635;
			this.groupBox1.TabStop = false;
			// 
			// frmLoteClientesCrediPoint
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(856, 430);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.txtDia);
			this.Controls.Add(this.txtMes);
			this.Controls.Add(this.txtLote);
			this.Controls.Add(this.txtLoteAnterior);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblVoucher);
			this.Controls.Add(this.lblCtaCorriente);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblCaducidad);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.btTXT);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridLotes);
			this.Name = "frmLoteClientesCrediPoint";
			this.Text = "Lotes Banco Solidario";
			this.Load += new System.EventHandler(this.frmLoteClientesCrediPoint_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLoteAnterior)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmLoteClientesCrediPoint_Load(object sender, System.EventArgs e)
		{	
//			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
//			this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			miAcceso = new Acceso(cdsSeteoF, "08320102");

			if (!Funcion.Permiso("956", cdsSeteoF))
			{		
		
				MessageBox.Show("No tiene Acceso LOTES SOLIDARIO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;

			}
			int iLotAnt = (int)Funcion.iEscalarSQL(cdsSeteoF, "Exec RetornaUltimoLoteSolidario");
			this.txtLoteAnterior.Value = iLotAnt;

			this.txtLote.Value = (int)this.txtLoteAnterior.Value + 1;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			bool bFechas = false;

			DateTime dtDesde = DateTime.Today;
			DateTime dtHasta = DateTime.Today;

			if (this.dtDesde.Value != System.DBNull.Value || this.dtHasta.Value != System.DBNull.Value)
			{
				bFechas = true;

				dtDesde = (DateTime)this.dtDesde.Value;
				dtHasta = (DateTime)this.dtHasta.Value;
			}

			string sSQL = string.Format("Exec ConsultaLotesSolidario '{0}', '{1}', {2}, '{3}', {4}, 2", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int)this.txtLote.Value, this.txtBusqueda.Text.ToString(), bFechas);
			this.uGridLotes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			int idAsiento = 0;			
			idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Declare @idAsiento Int = 0 If Exists(Select ISNULL(idAsiento, 0) From Asiento Where Numero = '{0}' And idTipoAsiento = 1 And Borrar = 0) Select @idAsiento = ISNULL(idAsiento, 0) From Asiento Where Numero = '{0}' And idTipoAsiento = 1 And Borrar = 0 Select @idAsiento", this.txtLote.Value.ToString()));
			if (idAsiento > 0) this.lblEstado.Text = "CONTABILIZADO";	else this.lblEstado.Text = "PENDIENTE";

			ContSel();
		}

		private void uGridLotes_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btTXT_Click(object sender, System.EventArgs e)
		{			
			if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show("No Hay datos para generar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtLote.Focus();				
			}
			else if (this.lblEstado.Text.ToString() == "CONTABILIZADO")
			{
				MessageBox.Show("El Lote ya esta CANCELADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtLote.Focus();				
			}

				//			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Compra Where NumeroLote = {0}", (int)this.txtLote.Value)) > 0)
				//			{
				//				MessageBox.Show(string.Format("El Lote Número {0} ya existe", (int)this.txtLote.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				//				this.txtLote.Focus();
				//				return;
				//			}
			else if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();				
			}
			else if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}
			else
			{
				this.Cursor = Cursors.WaitCursor;
				#region Variables
				char Pad = '0';
				string Fecha = this.txtDia.Value.ToString().PadLeft(2, Pad) + this.txtMes.Value.ToString().PadLeft(2, Pad);

				DateTime dtDesde = (DateTime)this.dtDesde.Value;
				DateTime dtHasta = (DateTime)this.dtHasta.Value;
			
				string sSQLAL = "";
				int iCont = 0;
				#endregion Variables

				#region Archivos de Texto
				/* CUOTAS POINT */
				string sSQL = string.Format("Exec CreaTXTSolidario 'Exec POINT.dbo.ConsultaCuotasSolidario {0}', 'CUOTASPOINT_{1}Lote{0}.txt'", (int)this.txtLote.Value, Fecha);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

				/* CREDITOS POINT */
				sSQL = string.Format("Exec CreaTXTSolidario 'Exec POINT.dbo.ConsultaCreditosSolidario {0}', 'CREDITOSPOINT_{1}Lote{0}.txt'", (int)this.txtLote.Value, Fecha);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

				/* CLIENTES POINT */
				sSQL = string.Format("Exec CreaTXTSolidario 'Exec POINT.dbo.TemporalLotesUnibanco {0}', 'CLIENTESPOINT_{1}Lote{0}.txt'", 
					(int)this.txtLote.Value, Fecha);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				#endregion Archivos de Texto --

				#region Actualiza El Numero de Lote en Facturas
				//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
				//			{
				//				if ((bool)dr.Cells["CapturaLote"].Value)
				//				{	
				//					sSQLAL = string.Format("Update Compra Set NumeroLote = {0}, CapturaLote = 1 Where idCompra = {1}", (int)this.txtLote.Value, (int)dr.Cells["idCompra"].Value);//dr.Cells["NumeroLote"].Value
				//					Funcion.EjecutaSQL(cdsSeteoF, sSQLAL, true);
				//
				//					iCont++;
				//				}
				//			}
				#endregion Actualiza El Numero de Lote en Facturas

				MessageBox.Show(string.Format("Se creo el Lote N. {0} con {1} documentos", (int)this.txtLote.Value, iCont));

				this.btBuscar_Click(sender, e);

				this.Cursor = Cursors.Default;
			}
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{
					string stDir = MenuLatinium.stDirInicio + "\\"; 
					ultraGridExcelExporter1.Export(this.uGridLotes, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void dtHasta_ValueChanged(object sender, System.EventArgs e)
		{
			DateTime dtHasta = (DateTime)this.dtHasta.Value;
			if (dtHasta.Month == DateTime.Today.Month)
			{
				this.txtMes.Value = DateTime.Today.Month;
				this.txtDia.Value = DateTime.Today.Day;
			}
			else
			{
				this.txtMes.Value = dtHasta.Month;
				this.txtDia.Value = DateTime.DaysInMonth(dtHasta.Year, dtHasta.Month);
			}
		}

		private void txtLote_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.btBuscar_Click(sender, e);

				this.txtBusqueda.Focus();
			}
		}

		private void ContSel()
		{
			int iCont = 0;

			this.lblContador.Text = "";

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
			{
				if ((bool)dr.Cells["CapturaLote"].Value) iCont++;
			}

			this.lblContador.Text = this.uGridLotes.Rows.Count + " REGISTROS - " + iCont.ToString() + " SELECCIONADOS";
		}

		private void uGridLotes_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "CapturaLote")
			{	
				if ((bool)(bool)e.Cell.Row.Cells["CapturaLote"].Value) e.Cell.Row.Cells["NumeroLote"].Value = (int)this.txtLote.Value;
				else e.Cell.Row.Cells["NumeroLote"].Value = 0;

				string sSQL = string.Format("Exec ActualizaLoteSoliudarioFacturas {0}, {1}, {2}", (int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["NumeroLote"].Value, (bool)e.Cell.Row.Cells["CapturaLote"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
	
				ContSel();
			}
		}

		private void txtBusqueda_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btBuscar_Click(sender, e);
		}

		private void txtBusqueda_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtLote_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
			
		private void uGridLotes_BeforeRowUpdate(object sender, Infragistics.Win.UltraWinGrid.CancelableRowEventArgs e)
		{
			
		}

		private void uGridLotes_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "CapturaLote")
			{
				if (this.lblEstado.Text.ToString() == "CONTABILIZADO")
				{
					MessageBox.Show("ESTE LOTE SE ENCUENTRA CANCELADO NO PUEDE MODIFICARSE \n\n COMUNIQUESE CON EL ADMINISTRADOR", "POINT TECHNOLOGY", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
			}
		}
	}
}

