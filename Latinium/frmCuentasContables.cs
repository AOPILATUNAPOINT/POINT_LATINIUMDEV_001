using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCuentasContables.
	/// </summary>
	public class frmCuentasContables : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Button btnImpresion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuentas;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optReportes;
		private System.Windows.Forms.Button btnVer;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Label lbl_registros;
		/// <summary>
		/// Variable del diseñador requerida.
		bool Cerrado = false;
		private System.Windows.Forms.Button btnCerrar;
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCuentasContables()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ID_CuentaEstadosABs");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipo_reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescripcionH");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoH");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btn_calcula", 0, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCuentasContables));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btn_enlazar", 1);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ID_CuentaEstadosABs");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipo_reporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescripcionH");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoH");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnImpresion = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label14 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtCuentas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.optReportes = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.button2 = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.lbl_registros = new System.Windows.Forms.Label();
			this.btnCerrar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optReportes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
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
			this.ultraGrid1.DataSource = this.ultraDataSource3;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 103;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Width = 222;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 100;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 65;
			ultraGridColumn5.Header.Caption = "Descripción";
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 217;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 92;
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 90;
			ultraGridColumn8.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			appearance2.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance2.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn8.CellAppearance = appearance2;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			appearance3.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance3.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn8.CellButtonAppearance = appearance3;
			ultraGridColumn8.Header.Caption = "...";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn8.Width = 29;
			ultraGridColumn9.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			appearance4.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance4.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn9.CellAppearance = appearance4;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			appearance5.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance5.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn9.CellButtonAppearance = appearance5;
			ultraGridColumn9.Header.Caption = "...";
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn9.Width = 31;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Hidden = true;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.GroupByRowAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8.5F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance10.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance10.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(32, 144);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(520, 272);
			this.ultraGrid1.TabIndex = 745;
			this.ultraGrid1.AfterSelectChange += new Infragistics.Win.UltraWinGrid.AfterSelectChangeEventHandler(this.ultraGrid1_AfterSelectChange);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			this.ultraGrid1.Click += new System.EventHandler(this.ultraGrid1_Click);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn3.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
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
			this.btnImpresion.Location = new System.Drawing.Point(1640, 16);
			this.btnImpresion.Name = "btnImpresion";
			this.btnImpresion.Size = new System.Drawing.Size(88, 24);
			this.btnImpresion.TabIndex = 744;
			this.btnImpresion.Text = "Imprimir";
			this.btnImpresion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn10.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(32, 104);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(127, 16);
			this.label14.TabIndex = 743;
			this.label14.Text = "Buscar Cuenta Contable";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(224, 24);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(37, 16);
			this.label17.TabIndex = 738;
			this.label17.Text = "Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(32, 24);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(40, 16);
			this.label18.TabIndex = 737;
			this.label18.Text = "Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// txtCuentas
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuentas.Appearance = appearance13;
			this.txtCuentas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCuentas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuentas.Location = new System.Drawing.Point(168, 104);
			this.txtCuentas.Name = "txtCuentas";
			this.txtCuentas.Size = new System.Drawing.Size(200, 21);
			this.txtCuentas.TabIndex = 742;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(24, 48);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1716, 8);
			this.groupBox1.TabIndex = 741;
			this.groupBox1.TabStop = false;
			// 
			// optReportes
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.optReportes.Appearance = appearance14;
			this.optReportes.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optReportes.ItemAppearance = appearance15;
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
			this.optReportes.Location = new System.Drawing.Point(32, 64);
			this.optReportes.Name = "optReportes";
			this.optReportes.Size = new System.Drawing.Size(224, 24);
			this.optReportes.TabIndex = 740;
			this.optReportes.ValueChanged += new System.EventHandler(this.optReportes_ValueChanged);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(376, 104);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 739;
			this.btnVer.Text = "Buscar";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
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
			// dtHasta
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance16;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(280, 24);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 736;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance17;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(88, 24);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 735;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button2.CausesValidation = false;
			this.button2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button2.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.Location = new System.Drawing.Point(1336, 16);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(88, 24);
			this.button2.TabIndex = 747;
			this.button2.Text = "Imprimir";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.CausesValidation = false;
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(496, 16);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(88, 24);
			this.button1.TabIndex = 748;
			this.button1.Text = "Imprimir";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click_1);
			// 
			// lbl_registros
			// 
			this.lbl_registros.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lbl_registros.ForeColor = System.Drawing.Color.Red;
			this.lbl_registros.Location = new System.Drawing.Point(280, 64);
			this.lbl_registros.Name = "lbl_registros";
			this.lbl_registros.Size = new System.Drawing.Size(192, 23);
			this.lbl_registros.TabIndex = 749;
			this.lbl_registros.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnCerrar
			// 
			this.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCerrar.Location = new System.Drawing.Point(496, 56);
			this.btnCerrar.Name = "btnCerrar";
			this.btnCerrar.Size = new System.Drawing.Size(88, 23);
			this.btnCerrar.TabIndex = 750;
			this.btnCerrar.Text = "Cerrar Período";
			this.btnCerrar.Click += new System.EventHandler(this.button3_Click_1);
			// 
			// frmCuentasContables
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(608, 470);
			this.Controls.Add(this.btnCerrar);
			this.Controls.Add(this.lbl_registros);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtCuentas);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.optReportes);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btnImpresion);
			this.Name = "frmCuentasContables";
			this.Text = "Reporte Cuentas Contables";
			this.Load += new System.EventHandler(this.frmCuentasContables_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optReportes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void optReportes_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtCuentas.Enabled = true;
			this.txtCuentas.Text = "";
			this.ultraGrid1.Enabled = true;
			this.lbl_registros.Text = "";

			if ((int)this.optReportes.Value == 0)
			{
				this.optReportes.Value=0;
				this.ultraGrid1.DataSource = 
					FuncionesProcedimientos.dtGeneral(string.Format("Exec MovimientosListaCuentas2 {0}", 
					(int)this.optReportes.Value)); 
			}
			if ((int)this.optReportes.Value == 1)
			{
				this.optReportes.Value=1;
				this.ultraGrid1.DataSource = 
					FuncionesProcedimientos.dtGeneral(string.Format("Exec MovimientosListaCuentas2 {0}", 
					(int)this.optReportes.Value)); 
			}
			if ((int)this.optReportes.Value == 2)
			{
				this.optReportes.Value=2;
				this.ultraGrid1.DataSource = 
					FuncionesProcedimientos.dtGeneral(string.Format("Exec MovimientosListaCuentas2 {0}", 
					(int)this.optReportes.Value)); 
			}
			if ((int)this.optReportes.Value == 3)
			{
				this.optReportes.Value=3;
				this.ultraGrid1.DataSource = 
					FuncionesProcedimientos.dtGeneral(string.Format("Exec MovimientosListaCuentas2 {0}", 
					(int)this.optReportes.Value)); 
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
		ListaCuentas();
		}

		private void ListaCuentas()
		{
			bool cerrado = false;
			// Se consulta si el periodo a consular está cerrado.
			string Query = string.Format("Exec CuentasContablesVerificarCerrado '{0}', '{1}', {2} ", 
			Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
			Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
			(int)this.optReportes.Value);
			
			cerrado = Funcion.bEscalarSQL(cdsSeteoF, Query, false);
			Cerrado = cerrado;

			this.lbl_registros.Text = "";
			this.btnCerrar.Visible = true;

			if(Cerrado == true)
			{
				this.lbl_registros.Text = "PERÍODO CERRADO";
				this.btnCerrar.Visible = false;
			}

			this.ultraGrid1.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec MovimientosListaCuentas2 {0}", (int)optReportes.Value)); 
			this.ultraGrid1.DisplayLayout.ViewStyleBand =Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.DisplayLayout.Bands[0].SortedColumns.Add("Descripcion", false, true);
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
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
				Lista_Contador((int)e.Cell.Row.Cells["ID_CuentaEstadosABs"].Value,(int)e.Cell.Row.Cells["tipo_reporte"].Value,e.Cell.Row.Cells["Codigo"].Value.ToString(),(DateTime)this.dtDesde.Value,(DateTime)this.dtHasta.Value,e.Cell.Row.Cells["DescripcionH"].Value.ToString(),e.Cell.Row.Cells["CodigoH"].Value.ToString(),e.Cell.Row.Cells["CodigoCuenta"].Value.ToString());
			}
			if (e.Cell.Column.ToString() == "btn_enlazar")
			{
				Lista_Enlazar((int)e.Cell.Row.Cells["ID_CuentaEstadosABs"].Value,(int)e.Cell.Row.Cells["tipo_reporte"].Value,e.Cell.Row.Cells["Codigo"].Value.ToString(),(DateTime)this.dtDesde.Value,(DateTime)this.dtHasta.Value,e.Cell.Row.Cells["DescripcionH"].Value.ToString(),e.Cell.Row.Cells["CodigoH"].Value.ToString(),e.Cell.Row.Cells["CodigoCuenta"].Value.ToString());
			}
		}

		private void Lista_Contador(int idCuenta,int tipo,string codigo,DateTime  desde,DateTime hasta,string DescripcionH,string CodigoH, string CodigoCuenta)
		{
//			if(idCuenta==14)
//			{
//			
//				using (frmNotasEF NEF = new frmNotasEF(idCuenta,tipo,codigo,desde,hasta,DescripcionH,CodigoH))
//				{		
//					if (DialogResult.OK == NEF.ShowDialog())
//					{
//					 
//					}
//				}
//			}
//			else
//			{

				using (frmCuentasContablesEstadosFinancieros CCL = new frmCuentasContablesEstadosFinancieros(idCuenta,tipo,codigo,desde,hasta,DescripcionH,CodigoH, CodigoCuenta, Cerrado))
				{		
					if (DialogResult.OK == CCL.ShowDialog())
					{
					 
					}
				}
//			}
		}

		private void Lista_Enlazar(int idCuenta,int tipo,string codigo,DateTime  desde,DateTime hasta,string DescripcionH,string CodigoH, string CodigoCuenta)
		{
//			if(idCuenta==14)
//			{
//			
//				using (frmNotasEF NEF = new frmNotasEF(idCuenta,tipo,codigo,desde,hasta,DescripcionH,CodigoH))
//				{		
//					if (DialogResult.OK == NEF.ShowDialog())
//					{
//					 
//					}
//				}
//			}
//			else
//			{

				using (frmCuentasContablesEnlazar CCL = new frmCuentasContablesEnlazar(idCuenta,tipo,codigo,desde,hasta,DescripcionH,CodigoH, CodigoCuenta, Cerrado))
				{		
					if (DialogResult.OK == CCL.ShowDialog())
					{
					 
					}
				}
//			}
		}

		private void frmCuentasContables_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			DateTime dtFecha = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2014");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2014");
			
			this.dtDesde.CalendarInfo.MaxDate = dtFecha;
			this.dtHasta.CalendarInfo.MaxDate = dtFecha;
			
			this.dtDesde.Value = "01/" + dtFecha.Month.ToString() + "/" + dtFecha.Year.ToString();
			this.dtHasta.Value = dtFecha.AddDays(-1);

			//this.uGridCentrosDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre, CONVERT(Bit, 0) As Sel From Proyecto Where Visible = 1 Order By Nombre");

			this.optReportes.CheckedIndex = 0;
			
			ListaCuentas();
		}

		string TipoReporte=null;
		string Tipo=null;

		private void button1_Click_1(object sender, System.EventArgs e)
		{
			RegistrosReporteFinal();
			if (!Validacion.vbValidaFechasDesdeHasta(this.dtDesde, this.dtHasta)) return;
			string sReporte = "InformeCuentasContables2.rpt";
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
				sReporte = "InformeResultados2.rpt";
				
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

//			this.uGridCuentas.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec GeneraFinalEF {0}, '{1}'", (int)this.optReportes.Value), TipoReporte); 
//
//						if ((int)this.optReportes.Value == 0)
//						{
//							string sCuentas = "";
//			
//							foreach(Infragistics.Win.UltraWinGrid.uGridCuentas dr in this.ultraGrid1.Rows.All)	
//							{
//								if ((bool) dr.Cells["Estado"].Value)
//									if (sCuentas.Length < 1) sCuentas = dr.Cells["idCuenta"].Value.ToString(); else sCuentas = sCuentas + ", " + dr.Cells["idCuenta"].Value.ToString();
//							
//								if (sCuentas.Length == 0)
//								{
//									MessageBox.Show("Seleccione al menos una Cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//									return;
//								}
//							}
//			
//							ParameterField paramFieldCuentas = new ParameterField ();
//							ParameterDiscreteValue discreteValCuentas = new ParameterDiscreteValue ();
//							paramFieldCuentas.ParameterFieldName = "@Cuentas";
//							discreteValCuentas.Value = sCuentas;
//							paramFieldCuentas.CurrentValues.Add (discreteValCuentas);
//							paramFields.Add (paramFieldCuentas);
//						}
			
						//if ((int)this.optReportes.Value == 2) sReporte = "BalancesDeComprobacion.rpt";
						
						//if ((int)this.optReportes.Value == 3) sReporte = "BalancesGenerales.rpt";
						
						//if ((int)this.optReportes.Value == 4) sReporte = "BalancesDeResultados.rpt";
						
			Reporte Reporte = new Reporte(sReporte, "");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
		}

		private void RegistrosReporteFinal()
		{
			string sCuentas = "";
			int iCont=0;
			string sSQLVD = "";
			try
			{		

				//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCuentas.Rows.All)	
				//					if ((bool) dr.Cells["Selec"].Value)
				//					{
				//						sCuentas = dr.Cells["idCuenta"].Value.ToString();
				//						#region Duplicados
				//						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT COUNT(*) FROM ReporteEstadosAB WHERE Cuenta = '{0}' AND DescripcionCuenta = '{1}'AND Estado=1",sCuentas,Descrip)) > 0)
				//						{
				//							sSQLVD = string.Format("Exec ValidaDuplicadosEstadosFinancieros '{0}'", 
				//							sCuentas,Descrip);					
				//							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
				//						}
				//						iCont++;
				//						#endregion Duplicados
				//					
				//					}
				//				#region REPETIDOS
				//				if (Funcion.iEscalarSQL(cdsSeteoF, "select count(*) from  ValidaReporteEstadosAB") > 0)
				//				{
				//					frmValidacionEstadosFinancieros VRB = new frmValidacionEstadosFinancieros();
				//					VRB.ShowDialog();
				//
				//					Funcion.EjecutaSQL(cdsSeteoF, "Delete From ValidaReporteEstadosAB");
				//
				//					this.Cursor = Cursors.Default;
				//
				//					return;
				//				}
				//
				//				#endregion REPETIDOS
				iCont=Funcion.iEscalarSQL(cdsSeteoF, string.Format("EXEC CuentasContablesSelecCodigoCuentaTotal {0}", (int)this.optReportes.Value));
				this.lbl_registros.Text = iCont.ToString()+ " GENERADOS";

				string sSQL = string.Format("Exec RegistrosReporteFinal '{0}', '{1}', {2},{3},{4}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
					0,
					0,MenuLatinium.stUsuario);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				MessageBox.Show(string.Format("Proceso Finalizado Correctamente, {0} Generados, se guardó la configuración.", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Demo", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				this.Cursor = Cursors.Default;							
			}
		}


		private void button3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void button3_Click_1(object sender, System.EventArgs e)
		{

			string sCuentas = "";
			int iCont=0;
			string sSQLVD = "";
			try
			{		
				iCont=Funcion.iEscalarSQL(cdsSeteoF, string.Format("EXEC CuentasContablesSelecCodigoCuentaTotal {0}", (int)this.optReportes.Value));
				this.lbl_registros.Text = iCont.ToString()+ " GENERADOS";
				string sSQL = string.Format("Exec CuentasContablesCerrarPeriodo '{0}', '{1}', {2},'{3}'",
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
				(int)this.optReportes.Value,MenuLatinium.stUsuario);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				//this.btnCerrar.Enabled = false;
				ListaCuentas();
				MessageBox.Show(string.Format("Proceso Finalizado Correctamente, Período Cerrado, se guardó la configuración.", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Demo", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				this.Cursor = Cursors.Default;							
			}
		}







	}
}
