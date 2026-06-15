using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmInformesEstadosFinancieros.
	/// </summary>
	public class frmInformesEstadosFinancieros: System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuentas;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optReportes;
		private System.Windows.Forms.Button btnImpresion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmInformesEstadosFinancieros()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmInformesEstadosFinancieros));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ID_CuentaEstadosABs");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipo_reporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescripcionH");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoH");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ID_CuentaEstadosABs");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipo_reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescripcionH");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoH");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btn_calcula", 0);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnVer = new System.Windows.Forms.Button();
			this.optReportes = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label14 = new System.Windows.Forms.Label();
			this.txtCuentas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnImpresion = new System.Windows.Forms.Button();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.button1 = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optReportes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
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
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(200, 12);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(37, 16);
			this.label17.TabIndex = 657;
			this.label17.Text = "Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 12);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(40, 16);
			this.label18.TabIndex = 656;
			this.label18.Text = "Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 10);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 655;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 10);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 654;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
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
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(352, 94);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(88, 24);
			this.btnVer.TabIndex = 659;
			this.btnVer.Text = "Buscar";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// optReportes
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.optReportes.Appearance = appearance3;
			this.optReportes.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optReportes.CheckedIndex = 0;
			this.optReportes.ItemAppearance = appearance4;
			this.optReportes.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Activos";
			valueListItem2.DataValue = 3;
			valueListItem2.DisplayText = "Pasivos";
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Resultados";
			this.optReportes.Items.Add(valueListItem1);
			this.optReportes.Items.Add(valueListItem2);
			this.optReportes.Items.Add(valueListItem3);
			this.optReportes.ItemSpacingVertical = 10;
			this.optReportes.Location = new System.Drawing.Point(8, 56);
			this.optReportes.Name = "optReportes";
			this.optReportes.Size = new System.Drawing.Size(224, 24);
			this.optReportes.TabIndex = 672;
			this.optReportes.Text = "Activos";
			this.optReportes.ValueChanged += new System.EventHandler(this.optReportes_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn7.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.groupBox2);
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(560, 8);
			this.groupBox1.TabIndex = 676;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(10, 34);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(560, 8);
			this.groupBox2.TabIndex = 677;
			this.groupBox2.TabStop = false;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(8, 98);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(127, 16);
			this.label14.TabIndex = 678;
			this.label14.Text = "Buscar Cuenta Contable";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCuentas
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuentas.Appearance = appearance5;
			this.txtCuentas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCuentas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuentas.Location = new System.Drawing.Point(144, 96);
			this.txtCuentas.Name = "txtCuentas";
			this.txtCuentas.Size = new System.Drawing.Size(200, 21);
			this.txtCuentas.TabIndex = 677;
			this.txtCuentas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuentas_KeyDown);
			// 
			// btnImpresion
			// 
			this.btnImpresion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImpresion.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnImpresion.CausesValidation = false;
			this.btnImpresion.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImpresion.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImpresion.Image = ((System.Drawing.Image)(resources.GetObject("btnImpresion.Image")));
			this.btnImpresion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImpresion.Location = new System.Drawing.Point(464, 8);
			this.btnImpresion.Name = "btnImpresion";
			this.btnImpresion.Size = new System.Drawing.Size(88, 24);
			this.btnImpresion.TabIndex = 732;
			this.btnImpresion.Text = "Imprimir";
			this.btnImpresion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImpresion.Click += new System.EventHandler(this.btnImpresion_Click);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn8.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn10.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource3;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 103;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 150;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 100;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 168;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 340;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 461;
			ultraGridColumn7.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance7.Image = ((object)(resources.GetObject("appearance7.Image")));
			appearance7.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance7.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn7.CellAppearance = appearance7;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			appearance8.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance8.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn7.CellButtonAppearance = appearance8;
			ultraGridColumn7.Header.Caption = "...";
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn7.Width = 33;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Hidden = true;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.GroupByRowAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8.5F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 144);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(544, 288);
			this.ultraGrid1.TabIndex = 733;
			this.ultraGrid1.AfterSelectChange += new Infragistics.Win.UltraWinGrid.AfterSelectChangeEventHandler(this.ultraGrid1_AfterSelectChange);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout_1);
			this.ultraGrid1.Click += new System.EventHandler(this.ultraGrid1_Click);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(464, 94);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(88, 24);
			this.button1.TabIndex = 734;
			this.button1.Text = "Nuevo";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(0, 128);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(560, 8);
			this.groupBox3.TabIndex = 735;
			this.groupBox3.TabStop = false;
			// 
			// frmInformesEstadosFinancieros
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(560, 438);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btnImpresion);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.txtCuentas);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.optReportes);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.KeyPreview = true;
			this.Name = "frmInformesEstadosFinancieros";
			this.Text = "Reportes Estados Financieros";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmInformesEstadosFinancieros_KeyDown);
			this.Load += new System.EventHandler(this.frmInformesEstadosFinancieros_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optReportes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmInformesEstadosFinancieros_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200217");

			if (!Funcion.Permiso("2001", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Estados Financieros ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			
			
			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			DateTime dtFecha = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			
			this.dtDesde.CalendarInfo.MaxDate = dtFecha;
			this.dtHasta.CalendarInfo.MaxDate = dtFecha;
			
			this.dtDesde.Value = "01/" + dtFecha.Month.ToString() + "/" + dtFecha.Year.ToString();
			this.dtHasta.Value = dtFecha.AddDays(-1);

			//this.uGridCentrosDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre, CONVERT(Bit, 0) As Sel From Proyecto Where Visible = 1 Order By Nombre");

			this.optReportes.CheckedIndex = 0;
			
			
			ListaCuentas();
		}

		private void ListaCuentas()
		{
			this.ultraGrid1.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec MovimientosListaCuentas {0}", (int)optReportes.Value)); 

			this.ultraGrid1.DisplayLayout.ViewStyleBand =Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.DisplayLayout.Bands[0].SortedColumns.Add("Descripcion", false, true);
		}

		private void txtCuentas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) ListaCuentas();
		}

		private void optGrupos_ValueChanged(object sender, System.EventArgs e)
		{
			ListaCuentas();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			ListaCuentas();
		}

		private void chkCuentas_CheckedChanged(object sender, System.EventArgs e)
		{
			
		}
		string TipoReporte=null;
		string Tipo=null;
		
		private void btnImpresion_Click(object sender, System.EventArgs e)
		{		
			if (!Validacion.vbValidaFechasDesdeHasta(this.dtDesde, this.dtHasta)) return;

			string sReporte = "Informe1PP.rpt";
			if ((int)this.optReportes.Value == 0)
			{
				Tipo="ACTIVOS";
				TipoReporte = "1.01.";						
			}
			if ((int)this.optReportes.Value == 3)
			{
				Tipo="PASIVOS Y PATRIMONIO";
				TipoReporte = "2.01.";		

			}
			if ((int)this.optReportes.Value == 1)
			{
				Tipo="ESTADO DE RESULTADOS ACUMULADO";
				TipoReporte = "3.01.";	
				sReporte = "InformeResultados.rpt";
				
			}

			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramFieldTipo = new ParameterField ();
			ParameterDiscreteValue discreteValTipo = new ParameterDiscreteValue ();
			paramFieldTipo.ParameterFieldName = "@Tipo";
			discreteValTipo.Value = Tipo;
			paramFieldTipo.CurrentValues.Add (discreteValTipo);
			paramFields.Add (paramFieldTipo);

			ParameterField paramFieldReporte = new ParameterField ();
			ParameterDiscreteValue discreteValReporte = new ParameterDiscreteValue ();
			paramFieldReporte.ParameterFieldName = "@reporte";
			discreteValReporte.Value = TipoReporte;
			paramFieldReporte.CurrentValues.Add (discreteValReporte);
			paramFields.Add (paramFieldReporte);

			ParameterField paramFieldestado = new ParameterField ();
			ParameterDiscreteValue discreteValestado = new ParameterDiscreteValue ();
			paramFieldestado.ParameterFieldName = "@estado";
			discreteValestado.Value = 1;
			paramFieldestado.CurrentValues.Add (discreteValestado);
			paramFields.Add (paramFieldestado);

			ParameterField paramFieldHasta = new ParameterField ();
			ParameterDiscreteValue discreteValHasta = new ParameterDiscreteValue ();
			paramFieldHasta.ParameterFieldName = "@Hasta";
			discreteValHasta.Value = Convert.ToDateTime(this.dtHasta.Value);
			paramFieldHasta.CurrentValues.Add (discreteValHasta);
			paramFields.Add (paramFieldHasta);

			ParameterField paramFieldDesde = new ParameterField ();
			ParameterDiscreteValue discreteValDesde  = new ParameterDiscreteValue ();
			paramFieldDesde.ParameterFieldName = "@Desde";
			discreteValDesde.Value = Convert.ToDateTime(this.dtDesde.Value);
			paramFieldDesde.CurrentValues.Add (discreteValDesde);
			paramFields.Add (paramFieldDesde);
			
			
			ParameterField paramFieldEmpresa = new ParameterField ();
			ParameterDiscreteValue discreteValEmpresa = new ParameterDiscreteValue ();
			paramFieldEmpresa.ParameterFieldName = "@idEmpresa";
			discreteValEmpresa.Value = MenuLatinium.stIdDB; 
			paramFieldEmpresa.CurrentValues.Add (discreteValEmpresa);
			paramFields.Add (paramFieldEmpresa);

			//			if ((int)this.optReportes.Value == 0)
			//			{
			//				string sCuentas = "";
			//
			//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)	
			//					if ((bool) dr.Cells["Sel"].Value)
			//						if (sCuentas.Length < 1) sCuentas = dr.Cells["idCuenta"].Value.ToString(); else sCuentas = sCuentas + ", " + dr.Cells["idCuenta"].Value.ToString();
			//				
			//				if (sCuentas.Length == 0)
			//				{
			//					MessageBox.Show("Seleccione al menos una Cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					return;
			//				}
			//
			//				ParameterField paramFieldCuentas = new ParameterField ();
			//				ParameterDiscreteValue discreteValCuentas = new ParameterDiscreteValue ();
			//				paramFieldCuentas.ParameterFieldName = "@Cuentas";
			//				discreteValCuentas.Value = sCuentas;
			//				paramFieldCuentas.CurrentValues.Add (discreteValCuentas);
			//				paramFields.Add (paramFieldCuentas);
			//			}
			//
			//			if ((int)this.optReportes.Value == 2) sReporte = "BalancesDeComprobacion.rpt";
			//			
			//			if ((int)this.optReportes.Value == 3) sReporte = "BalancesGenerales.rpt";
			//			
			//			if ((int)this.optReportes.Value == 4) sReporte = "BalancesDeResultados.rpt";
			//			
			Reporte Reporte = new Reporte(sReporte, "");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
		}

		private void optReportes_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtCuentas.Enabled = true;
			this.txtCuentas.Text = "";
			this.ultraGrid1.Enabled = true;

			if ((int)this.optReportes.Value == 0)
			{
				this.optReportes.Value=0;
				this.ultraGrid1.DataSource = 
					FuncionesProcedimientos.dtGeneral(string.Format("Exec MovimientosListaCuentas {0}", 
					(int)this.optReportes.Value)); 
			}
			if ((int)this.optReportes.Value == 1)
			{
				this.optReportes.Value=1;
				this.ultraGrid1.DataSource = 
					FuncionesProcedimientos.dtGeneral(string.Format("Exec MovimientosListaCuentas {0}", 
					(int)this.optReportes.Value)); 
			}
			if ((int)this.optReportes.Value == 2)
			{
				this.optReportes.Value=2;
				this.ultraGrid1.DataSource = 
					FuncionesProcedimientos.dtGeneral(string.Format("Exec MovimientosListaCuentas {0}", 
					(int)this.optReportes.Value)); 
			}
			if ((int)this.optReportes.Value == 3)
			{
				this.optReportes.Value=3;
				this.ultraGrid1.DataSource = 
					FuncionesProcedimientos.dtGeneral(string.Format("Exec MovimientosListaCuentas {0}", 
					(int)this.optReportes.Value)); 
			}

		}

		private void frmInformesEstadosFinancieros_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
		}

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
	
		}
		private void Lista_Contador(int idCuenta,int tipo,string codigo,DateTime  desde,DateTime hasta,string DescripcionH,string CodigoH)
		{	    		
			if(idCuenta==14)
			{
			
				using (frmNotasEF NEF = new frmNotasEF(idCuenta,tipo,codigo,desde,hasta,DescripcionH,CodigoH))
				{		
					if (DialogResult.OK == NEF.ShowDialog())
					{
					 
					}
				}
			}
			else
			{
				using (frmCuentasEstadoFinanciero CCL = new frmCuentasEstadoFinanciero(idCuenta,tipo,codigo,desde,hasta,DescripcionH,CodigoH))
				{		
					if (DialogResult.OK == CCL.ShowDialog())
					{
					 
					}
				}
			}
		}

		private void ultraGrid1_AfterSelectChange(object sender, Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs e)
		{
		
		}

		private void ultraGrid1_Click(object sender, System.EventArgs e)
		{
			
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btn_calcula")
			{
				Lista_Contador((int)e.Cell.Row.Cells["ID_CuentaEstadosABs"].Value,(int)e.Cell.Row.Cells["tipo_reporte"].Value,e.Cell.Row.Cells["Codigo"].Value.ToString(),(DateTime)this.dtDesde.Value,(DateTime)this.dtHasta.Value,e.Cell.Row.Cells["DescripcionH"].Value.ToString(),e.Cell.Row.Cells["CodigoH"].Value.ToString());
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			if ((int)this.optReportes.Value == 0)
			{
				Tipo="ACTIVOS";
				TipoReporte = "1.01.";					
			}
			if ((int)this.optReportes.Value == 3)
			{
				Tipo="PASIVOS Y PATRIMONIO";
				TipoReporte = "2.01.";		
			}
			if ((int)this.optReportes.Value == 1)
			{
				Tipo="ESTADO DE RESULTADOS ACUMULADO";
				TipoReporte = "3.01.";		
			}

			string sSQLPa = string.Format("Update ReporteEstadosAB set Estado=0 where TipoReporte ='{0}'",TipoReporte);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLPa );

			MessageBox.Show(string.Format("Proceso Finalizado Correctamente Registros 0"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void ultraGrid1_InitializeLayout_1(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
