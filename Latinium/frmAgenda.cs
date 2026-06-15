using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmAgenda.
	/// </summary>
	public class frmAgenda : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.ComboBoxEdit cmbTipoConsulta;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private DevExpress.XtraEditors.ComboBoxEdit cmbEstados;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArtículo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAgenda()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyectoZoom");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDevolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Devuelto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyectoZoom");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEntrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDevolucion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Devuelto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciudad");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbTipoConsulta = new DevExpress.XtraEditors.ComboBoxEdit();
			this.txtArtículo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbEstados = new DevExpress.XtraEditors.ComboBoxEdit();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoConsulta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArtículo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "_DetCompra";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 60;
			ultraGridColumn2.Header.Caption = "Artículo";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn2.Width = 160;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Código";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 184;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Artículo";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 230;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Número";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 90;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 66;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 110;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Entrega";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 70;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Ent";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 40;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Devolución";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 85;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Dev";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 40;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 42;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 100;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 150;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Ubicación";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 120;
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
			appearance2.FontData.SizeInPoints = 8.5F;
			ultraGridBand1.Override.HeaderAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(4, 112);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1240, 480);
			this.ultraGrid1.TabIndex = 247;
			this.ultraGrid1.DoubleClick += new System.EventHandler(this.ultraGrid1_DoubleClick);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn8.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn10.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn12.DataType = typeof(int);
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
																																 ultraDataColumn14,
																																 ultraDataColumn15});
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(392, 68);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(96, 25);
			this.btBuscar.TabIndex = 248;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// cmbTipoConsulta
			// 
			this.cmbTipoConsulta.Location = new System.Drawing.Point(64, 8);
			this.cmbTipoConsulta.Name = "cmbTipoConsulta";
			// 
			// cmbTipoConsulta.Properties
			// 
			this.cmbTipoConsulta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																						new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbTipoConsulta.Size = new System.Drawing.Size(176, 20);
			this.cmbTipoConsulta.TabIndex = 249;
			this.cmbTipoConsulta.SelectedIndexChanged += new System.EventHandler(this.cmbTipoConsulta_SelectedIndexChanged);
			// 
			// txtArtículo
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArtículo.Appearance = appearance9;
			this.txtArtículo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArtículo.Location = new System.Drawing.Point(64, 69);
			this.txtArtículo.Name = "txtArtículo";
			this.txtArtículo.Size = new System.Drawing.Size(312, 22);
			this.txtArtículo.TabIndex = 250;
			this.txtArtículo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtArtículo_KeyPress);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 72);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(42, 16);
			this.label8.TabIndex = 268;
			this.label8.Text = "Artículo";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 16);
			this.label1.TabIndex = 269;
			this.label1.Text = "Período";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(38, 16);
			this.label2.TabIndex = 271;
			this.label2.Text = "Estado";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEstados
			// 
			this.cmbEstados.Location = new System.Drawing.Point(64, 38);
			this.cmbEstados.Name = "cmbEstados";
			// 
			// cmbEstados.Properties
			// 
			this.cmbEstados.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbEstados.Size = new System.Drawing.Size(176, 20);
			this.cmbEstados.TabIndex = 270;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(432, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(33, 16);
			this.label3.TabIndex = 275;
			this.label3.Text = "Hasta";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(264, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(35, 16);
			this.label4.TabIndex = 274;
			this.label4.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance10;
			this.dtHasta.AutoSelect = false;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(480, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 273;
			this.dtHasta.Value = new System.DateTime(2013, 6, 18, 0, 0, 0, 0);
			// 
			// dtDesde
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance11;
			this.dtDesde.AutoSelect = false;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(312, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 272;
			this.dtDesde.Value = new System.DateTime(2013, 6, 18, 0, 0, 0, 0);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 96);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1256, 8);
			this.groupBox1.TabIndex = 276;
			this.groupBox1.TabStop = false;
			// 
			// frmAgenda
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1248, 598);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.txtArtículo);
			this.Controls.Add(this.cmbTipoConsulta);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "frmAgenda";
			this.Text = "Agenda";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAgenda_KeyDown);
			this.Load += new System.EventHandler(this.frmAgenda_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoConsulta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArtículo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iTipoBusqueda = 0;
		int iEstado = 0;

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			if (this.cmbTipoConsulta.Text.ToString() == "TODO") iTipoBusqueda = 0;
			if (this.cmbTipoConsulta.Text.ToString() == "MES ACTUAL") iTipoBusqueda = 1;
			if (this.cmbTipoConsulta.Text.ToString() == "RANGO DE FECHAS") iTipoBusqueda = 2;

			if (this.cmbEstados.Text.ToString() == "TODO") iEstado = 0;
			if (this.cmbEstados.Text.ToString() == "PENDIENTE") iEstado = 1;
			if (this.cmbEstados.Text.ToString() == "APROBADO") iEstado = 5;
			if (this.cmbEstados.Text.ToString() == "RECHAZADO") iEstado = 4;
			if (this.cmbEstados.Text.ToString() == "FACTURADO") iEstado = 9;
			if (this.cmbEstados.Text.ToString() == "ENTREGADO") iEstado = 15;
			if (this.cmbEstados.Text.ToString() == "DEVUELTO") iEstado = 16;

			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			string sSQL = string.Format("Exec ConsultaAgenda {0}, {1}, '{2}', '{3}', '{4}'", iTipoBusqueda, iEstado, this.txtArtículo.Text.ToString(), dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmAgenda_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0624");

			if (!Funcion.Permiso("989", cdsSeteoF))
			{				
				MessageBox.Show("No tiene Acceso AGENDA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.cmbTipoConsulta.Text = "TODO";
			this.cmbTipoConsulta.Properties.Items.Add("TODO");
			this.cmbTipoConsulta.Properties.Items.Add("MES ACTUAL");
			this.cmbTipoConsulta.Properties.Items.Add("RANGO DE FECHAS");

			this.cmbEstados.Properties.Items.Add("TODO");
			this.cmbEstados.Properties.Items.Add("PENDIENTE");
			this.cmbEstados.Properties.Items.Add("APROBADO");
			this.cmbEstados.Properties.Items.Add("RECHAZADO");
			this.cmbEstados.Properties.Items.Add("FACTURADO");
			this.cmbEstados.Properties.Items.Add("ENTREGADO");
			this.cmbEstados.Properties.Items.Add("DEVUELTO");
				
			this.cmbEstados.Text = "PENDIENTE";
		}

		private void cmbTipoConsulta_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.cmbTipoConsulta.Text.ToString() == "TODO")
			{
				this.dtDesde.Enabled = false;
				this.dtHasta.Enabled = false;
				this.dtDesde.Value = DateTime.Today;
				this.dtHasta.Value = DateTime.Today;
			}
			if (this.cmbTipoConsulta.Text.ToString() == "MES ACTUAL")
			{
				this.dtDesde.Enabled = false;
				this.dtHasta.Enabled = false;
				this.dtDesde.Value = Funcion.dtEscalarSQL(cdsSeteoF, "SELECT DATEADD(s,0,DATEADD(mm, DATEDIFF(m,0,getdate()),0))");
				this.dtHasta.Value = Funcion.dtEscalarSQL(cdsSeteoF, "SELECT DATEADD(s,-1,DATEADD(mm, DATEDIFF(m,0,GETDATE())+1,0))");
			}
			if (this.cmbTipoConsulta.Text.ToString() == "RANGO DE FECHAS")
			{
				this.dtDesde.Enabled = true;
				this.dtHasta.Enabled = true;
				this.dtDesde.Value = DateTime.Today;
				this.dtHasta.Value = DateTime.Today;
			}
		}

		private void txtArtículo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btBuscar_Click(sender, e);
		}

		private void frmAgenda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void ultraGrid1_DoubleClick(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			if (this.ultraGrid1.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value) return;

			frmVistaAgendaGestionCliente miVista = new frmVistaAgendaGestionCliente((int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value);
			miVista.ShowDialog();
			Cursor = Cursors.Default;
		}
	}
}

