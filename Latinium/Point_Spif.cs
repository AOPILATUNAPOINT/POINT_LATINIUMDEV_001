using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Point_Spif.
	/// </summary>
	public class Point_Spif : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaDesde;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsSeteoRepF;
		private DevExpress.XtraEditors.RadioGroup rdgProductoExistencia;
		private DevExpress.XtraEditors.RadioGroup rdgProductoFacturable;
		private C1.Data.C1DataSet cdsSeteoF;
		private C1.Data.C1DataSet cdsArticuloTabla;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private DevExpress.XtraEditors.RadioGroup rdgEntrada;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSPIFF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnSalir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Point_Spif()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Point_Spif));
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			this.dtFechaHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoRepF = new C1.Data.C1DataSet();
			this.dtFechaDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.rdgProductoExistencia = new DevExpress.XtraEditors.RadioGroup();
			this.rdgProductoFacturable = new DevExpress.XtraEditors.RadioGroup();
			this.uGridSPIFF = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cdsArticuloTabla = new C1.Data.C1DataSet();
			this.label3 = new System.Windows.Forms.Label();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.rdgEntrada = new DevExpress.XtraEditors.RadioGroup();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoRepF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoExistencia.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoFacturable.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSPIFF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgEntrada.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// dtFechaHasta
			// 
			this.dtFechaHasta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.dtFechaHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaHasta.DateButtons.Add(dateButton1);
			this.dtFechaHasta.Format = "dd/MM/yyyy";
			this.dtFechaHasta.Location = new System.Drawing.Point(784, 408);
			this.dtFechaHasta.Name = "dtFechaHasta";
			this.dtFechaHasta.NonAutoSizeHeight = 23;
			this.dtFechaHasta.Size = new System.Drawing.Size(112, 21);
			this.dtFechaHasta.SpinButtonsVisible = true;
			this.dtFechaHasta.TabIndex = 2;
			this.dtFechaHasta.Value = ((object)(resources.GetObject("dtFechaHasta.Value")));
			this.dtFechaHasta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFechaHasta_KeyPress);
			this.dtFechaHasta.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaHasta_BeforeDropDown);
			this.dtFechaHasta.ValueChanged += new System.EventHandler(this.dtFechaHasta_ValueChanged);
			// 
			// cdsSeteoRepF
			// 
			this.cdsSeteoRepF.BindingContextCtrl = this;
			this.cdsSeteoRepF.DataLibrary = "LibFacturacion";
			this.cdsSeteoRepF.DataLibraryUrl = "";
			this.cdsSeteoRepF.DataSetDef = "dsSeteoF";
			this.cdsSeteoRepF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoRepF.Name = "cdsSeteoRepF";
			this.cdsSeteoRepF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoRepF.SchemaDef = null;
			this.cdsSeteoRepF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoRepF_BeforeFill);
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.AllowMonthSelection = true;
			this.dtFechaDesde.AllowWeekSelection = true;
			this.dtFechaDesde.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.dtFechaDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaDesde.DateButtons.Add(dateButton2);
			this.dtFechaDesde.Format = "dd/MM/yyyy";
			this.dtFechaDesde.Location = new System.Drawing.Point(576, 408);
			this.dtFechaDesde.Name = "dtFechaDesde";
			this.dtFechaDesde.NonAutoSizeHeight = 23;
			this.dtFechaDesde.Size = new System.Drawing.Size(112, 21);
			this.dtFechaDesde.SpinButtonsVisible = true;
			this.dtFechaDesde.TabIndex = 1;
			this.dtFechaDesde.Value = ((object)(resources.GetObject("dtFechaDesde.Value")));
			this.dtFechaDesde.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFechaDesde_KeyPress);
			this.dtFechaDesde.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaDesde_BeforeDropDown);
			this.dtFechaDesde.ValueChanged += new System.EventHandler(this.dtFechaDesde_ValueChanged);
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(728, 411);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 17);
			this.label2.TabIndex = 81;
			this.label2.Text = "Hasta";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(520, 411);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 17);
			this.label1.TabIndex = 80;
			this.label1.Text = "Desde";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// rdgProductoExistencia
			// 
			this.rdgProductoExistencia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.rdgProductoExistencia.EditValue = 0;
			this.rdgProductoExistencia.Location = new System.Drawing.Point(8, 406);
			this.rdgProductoExistencia.Name = "rdgProductoExistencia";
			// 
			// rdgProductoExistencia.Properties
			// 
			this.rdgProductoExistencia.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgProductoExistencia.Properties.Appearance.Options.UseBackColor = true;
			this.rdgProductoExistencia.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos los productos"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Productos con existencia"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Producto sin existencia"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Solo productos en remates")});
			this.rdgProductoExistencia.Size = new System.Drawing.Size(160, 112);
			this.rdgProductoExistencia.TabIndex = 3;
			this.rdgProductoExistencia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.rdgProductoExistencia_KeyPress);
			this.rdgProductoExistencia.SelectedIndexChanged += new System.EventHandler(this.rdgProductoExistencia_SelectedIndexChanged);
			// 
			// rdgProductoFacturable
			// 
			this.rdgProductoFacturable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.rdgProductoFacturable.EditValue = 0;
			this.rdgProductoFacturable.Location = new System.Drawing.Point(176, 406);
			this.rdgProductoFacturable.Name = "rdgProductoFacturable";
			// 
			// rdgProductoFacturable.Properties
			// 
			this.rdgProductoFacturable.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgProductoFacturable.Properties.Appearance.Options.UseBackColor = true;
			this.rdgProductoFacturable.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos los productos"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Productos Facturables"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Producto No Facturables"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Solo combos")});
			this.rdgProductoFacturable.Size = new System.Drawing.Size(160, 112);
			this.rdgProductoFacturable.TabIndex = 4;
			this.rdgProductoFacturable.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.rdgProductoFacturable_KeyPress);
			this.rdgProductoFacturable.SelectedIndexChanged += new System.EventHandler(this.rdgProductoFacturable_SelectedIndexChanged);
			// 
			// uGridSPIFF
			// 
			this.uGridSPIFF.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridSPIFF.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridSPIFF.DisplayLayout.Appearance = appearance1;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance2;
			this.uGridSPIFF.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridSPIFF.DisplayLayout.ColumnScrollbarSmallChange = 50;
			this.uGridSPIFF.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSPIFF.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSPIFF.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridSPIFF.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSPIFF.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridSPIFF.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSPIFF.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSPIFF.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSPIFF.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridSPIFF.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSPIFF.Location = new System.Drawing.Point(8, 8);
			this.uGridSPIFF.Name = "uGridSPIFF";
			this.uGridSPIFF.Size = new System.Drawing.Size(1186, 264);
			this.uGridSPIFF.TabIndex = 120;
			this.uGridSPIFF.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridSPIFF_BeforeSelectChange);
			this.uGridSPIFF.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSPIFF_InitializeLayout);
			this.uGridSPIFF.Click += new System.EventHandler(this.uGridSPIFF_Click);
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
			// cdsArticuloTabla
			// 
			this.cdsArticuloTabla.BindingContextCtrl = this;
			this.cdsArticuloTabla.CaseSensitive = false;
			this.cdsArticuloTabla.DataLibrary = "LibFacturacion";
			this.cdsArticuloTabla.DataLibraryUrl = "";
			this.cdsArticuloTabla.DataSetDef = "dsArticuloTabla";
			this.cdsArticuloTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticuloTabla.Name = "cdsArticuloTabla";
			this.cdsArticuloTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticuloTabla.SchemaDef = null;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(520, 443);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 17);
			this.label3.TabIndex = 122;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtArticulo
			// 
			this.txtArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance8;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Location = new System.Drawing.Point(576, 440);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(320, 22);
			this.txtArticulo.TabIndex = 0;
			this.txtArticulo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtArticulo_KeyPress);
			this.txtArticulo.ValueChanged += new System.EventHandler(this.txtArticulo_ValueChanged);
			// 
			// rdgEntrada
			// 
			this.rdgEntrada.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.rdgEntrada.EditValue = 2;
			this.rdgEntrada.Location = new System.Drawing.Point(344, 406);
			this.rdgEntrada.Name = "rdgEntrada";
			// 
			// rdgEntrada.Properties
			// 
			this.rdgEntrada.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgEntrada.Properties.Appearance.Options.UseBackColor = true;
			this.rdgEntrada.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			 new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Existencia"),
																																																			 new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Última entrada "),
																																																			 new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Transito"),
																																																			 new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Ventas")});
			this.rdgEntrada.Size = new System.Drawing.Size(160, 85);
			this.rdgEntrada.TabIndex = 5;
			this.rdgEntrada.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.rdgEntrada_KeyPress);
			this.rdgEntrada.SelectedIndexChanged += new System.EventHandler(this.rdgEntrada_SelectedIndexChanged);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance9.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance9;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand2.Header.Appearance = appearance10;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.ColumnScrollbarSmallChange = 50;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8F;
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
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 280);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1186, 100);
			this.ultraGrid1.TabIndex = 201;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 384);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1208, 8);
			this.groupBox1.TabIndex = 202;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(576, 472);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 203;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1080, 440);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 207;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnVer
			// 
			this.btnVer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(912, 440);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 205;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(992, 440);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(80, 23);
			this.btnExportar.TabIndex = 206;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// Point_Spif
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1202, 526);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.rdgEntrada);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.uGridSPIFF);
			this.Controls.Add(this.rdgProductoFacturable);
			this.Controls.Add(this.rdgProductoExistencia);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.dtFechaHasta);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "Point_Spif";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "SPIFF";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Point_Spif_KeyDown);
			this.Load += new System.EventHandler(this.Point_Spif_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoRepF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoExistencia.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoFacturable.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSPIFF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgEntrada.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoRepF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoRepF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		CultureInfo us = new CultureInfo("en-US");

		private void ConsultaSPIFF()
		{
			Cursor = Cursors.WaitCursor;

			DateTime dtDesde = (DateTime) this.dtFechaDesde.Value;
			DateTime dtHasta = (DateTime) this.dtFechaHasta.Value;

			this.uGridSPIFF.DataSource = null;

			string sSQL = string.Format("Exec Consulta_SPIFF '{0}', '{1}', '{2}', {3}, {4}, {5}", 
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtArticulo.Text, (int)rdgProductoExistencia.EditValue, (int)rdgProductoFacturable.EditValue, (int)this.rdgEntrada.EditValue);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridSPIFF);
			
			DiseñoGrid();

			this.lblContador.Text = this.uGridSPIFF.Rows.Count + " ARTICULOS ENCONTRADOS";

			Cursor = Cursors.Default;
		}

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridSPIFF.Rows.Count > 0)
			{
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Codigo"].Width = 160;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Articulo"].Width = 330;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.ActivateOnly;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Discontinuo"].Width = 20;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Discontinuo"].Header.Caption = "D";
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Discontinuo"].CellActivation = Activation.ActivateOnly;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Facturable"].Header.Caption = "F";
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Facturable"].Width = 20;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Facturable"].CellActivation = Activation.ActivateOnly;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["OdC"].Width = 35;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["OdC"].CellActivation = Activation.ActivateOnly;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["OdC"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Compra"].Width = 35;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Compra"].CellActivation = Activation.ActivateOnly;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Compra"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Saldo"].Hidden = true;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Saldo"].Width = 50;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Saldo"].CellActivation = Activation.ActivateOnly;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Saldo"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;				
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Total"].Width = 45;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.ActivateOnly;
				this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridSPIFF.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridSPIFF.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridSPIFF.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridSPIFF.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0}";

				if (this.uGridSPIFF.Rows.Count > 0)
				{
					string nombrCol = "";
					int iRegion = 0 ;
							
					int iColumnas = (int)this.uGridSPIFF.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 9; j < iColumnas; j++)
					{
						nombrCol = this.uGridSPIFF.DisplayLayout.Bands[0].Columns[j].Header.Caption;

						iRegion =  Funcion.iEscalarSQL(cdsSeteoF, string.Format("If  Exists(Select CODIGO From Bodega Where Codigo = '{0}') begin Select region from Bodega where Codigo = '{0}' end Else begin  Select CONVERT(Int , 0) end ",nombrCol));		
						if ( iRegion == 1) 
						{
							this.uGridSPIFF.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LightBlue;
						}
						if ( iRegion == 2) 
						{
							this.uGridSPIFF.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Yellow;
						}
						if (nombrCol == "BI") this.uGridSPIFF.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						if (nombrCol == "BG") this.uGridSPIFF.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;	
						if (nombrCol == "BM") this.uGridSPIFF.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						
						this.uGridSPIFF.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						this.uGridSPIFF.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridSPIFF.DisplayLayout.Bands[0].Columns[nombrCol].Width = 30;

						this.uGridSPIFF.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridSPIFF.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						this.uGridSPIFF.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridSPIFF.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
					}
				}				
			}

			this.uGridSPIFF.DisplayLayout.Bands[0].Header.Caption = this.uGridSPIFF.Rows.Count + " ARTÍCULOS ENCONTRADOS.";
			#endregion Diseño Grid
		}

		private void DiseñoGridTransito()
		{
			#region Diseño Grid
			if (this.ultraGrid1.Rows.Count > 0)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].Width = 290;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Articulo"].Width = 335;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.ActivateOnly;
				
				if (this.ultraGrid1.Rows.Count > 0)
				{
					string nombrCol = "";
					int iRegion = 0 ;
							
					int iColumnas = (int)this.ultraGrid1.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 3; j < iColumnas; j++)
					{
						nombrCol = this.ultraGrid1.DisplayLayout.Bands[0].Columns[j].Header.Caption;

						iRegion =  Funcion.iEscalarSQL(cdsSeteoF, string.Format("If  Exists(Select CODIGO From Bodega Where Codigo = '{0}') begin Select region from Bodega where Codigo = '{0}' end Else begin  Select CONVERT(Int , 0) end ",nombrCol));		
						if ( iRegion == 1) 
						{
							this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LightBlue;
						}
						if ( iRegion == 2) 
						{
							this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Yellow;
						}
						if (nombrCol == "BI") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						if (nombrCol == "BG") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						if (nombrCol == "BM") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						
						this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].Width = 30;						
					}
				}				
			}		
			#endregion Diseño Grid
		}

		private void Spif()
		{
			Cursor = Cursors.WaitCursor;

			DateTime dtDesde = (DateTime) this.dtFechaDesde.Value;
			DateTime dtHasta = (DateTime) this.dtFechaHasta.Value;
 
			int iEntrada = (int)rdgEntrada.EditValue;
			
			this.uGridSPIFF.DataSource = null;

			string stExec = string.Format("Exec Spif '{0}', '{1}', {2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), iEntrada);
			uGridSPIFF.DataSource = Funcion.dvProcedimiento(cdsSeteoRepF, stExec);

			DiseñoGrid();

			Cursor = Cursors.Default;
		}

		private void Point_Spif_Load(object sender, System.EventArgs e)
		{			
			Funcion.Resolucion(this, this.Height, this.Width);

			this.dtFechaDesde.Value = DateTime.Today;
			this.dtFechaHasta.Value = DateTime.Today;
			
			this.ConsultaSPIFF();
			
			this.txtArticulo.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void dtFechaDesde_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void rdgProductoExistencia_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			ConsultaSPIFF();

			this.MercaderiaEnTransito(0);
		}

		private void rdgProductoFacturable_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			ConsultaSPIFF();

			this.MercaderiaEnTransito(0);
		}

		private void txtArticulo_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void txtArticulo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnVer_Click(sender, e);
		}

		private void dtFechaDesde_ValueChanged(object sender, System.EventArgs e)
		{
			this.dtFechaHasta.CalendarInfo.MinDate = (DateTime)this.dtFechaDesde.Value;
		}

		private void dtFechaHasta_ValueChanged(object sender, System.EventArgs e)
		{
			this.dtFechaDesde.CalendarInfo.MaxDate = (DateTime)this.dtFechaHasta.Value;
		}

		private void rdgEntrada_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			ConsultaSPIFF();

			this.MercaderiaEnTransito(0);
		}

		private void Point_Spif_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void dtFechaHasta_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtFechaDesde_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFechaHasta.Focus();
		}

		private void dtFechaHasta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtArticulo.Focus();
		}

		private void rdgProductoExistencia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.rdgProductoFacturable.Focus();
		}

		private void rdgProductoFacturable_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.rdgEntrada.Focus();
		}

		private void rdgEntrada_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFechaDesde.Focus();
		}
		
		private void uGridSPIFF_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void MercaderiaEnTransito (int idArticulo)
		{
			DateTime dtFecha = DateTime.Today;
			string sSQL = string.Format("Exec MercaderiaEnTransitoAPSpiff {0}, '{1}'", idArticulo, dtFecha.ToString("yyyyMMdd"));
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			DiseñoGridTransito();
		}

		private void uGridSPIFF_Click(object sender, System.EventArgs e)
		{
			if (this.uGridSPIFF.Rows.Count > 0)
			{
				if (this.uGridSPIFF.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value) 
				{
					MercaderiaEnTransito((int)this.uGridSPIFF.ActiveRow.Cells["idArticulo"].Value);
				}
			}
		}

		private void uGridSPIFF_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (this.uGridSPIFF.Rows.Count > 0)
			{
				if (this.uGridSPIFF.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value) 
				{
					MercaderiaEnTransito((int)this.uGridSPIFF.ActiveRow.Cells["idArticulo"].Value);
				}
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			ConsultaSPIFF();

			this.MercaderiaEnTransito(0);
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			//FuncionesProcedimientos.ExportaExcel(this.uGridSPIFF);

			if (this.uGridSPIFF.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				
				ultraGridExcelExporter1.Export(this.uGridSPIFF, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}
	}
}
