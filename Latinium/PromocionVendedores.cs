using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Summary description for PromocionVendedores.
	/// </summary>
	public class PromocionVendedores : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdPromocion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtInicio;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFin;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsPromocionVendedores;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		private C1.Data.C1DataRow drPromocionVendedores;
		private C1.Data.C1DataRow drSeteoF;

		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodos;
		private System.Windows.Forms.ErrorProvider errorProvider;
		
		public PromocionVendedores()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_DetPromocionVendedores", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetPromVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			this.label6 = new System.Windows.Forms.Label();
			this.dtInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsPromocionVendedores = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkTodos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtIdPromocion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPromocionVendedores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 48);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(31, 16);
			this.label6.TabIndex = 80;
			this.label6.Text = "Inicio";
			// 
			// dtInicio
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtInicio.Appearance = appearance1;
			this.dtInicio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtInicio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPromocionVendedores, "_PromocionVendedores.Desde"));
			dateButton1.Caption = "Today";
			this.dtInicio.DateButtons.Add(dateButton1);
			this.dtInicio.Format = "dd/MM/yyyy";
			this.dtInicio.Location = new System.Drawing.Point(88, 48);
			this.dtInicio.Name = "dtInicio";
			this.dtInicio.NonAutoSizeHeight = 23;
			this.dtInicio.Size = new System.Drawing.Size(112, 21);
			this.dtInicio.SpinButtonsVisible = true;
			this.dtInicio.TabIndex = 79;
			this.dtInicio.Value = new System.DateTime(2013, 9, 13, 0, 0, 0, 0);
			this.dtInicio.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtInicio_KeyPress);
			this.dtInicio.Validated += new System.EventHandler(this.dtInicio_Validated);
			this.dtInicio.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtInicio_BeforeDropDown);
			// 
			// cdsPromocionVendedores
			// 
			this.cdsPromocionVendedores.BindingContextCtrl = this;
			this.cdsPromocionVendedores.DataLibrary = "LibFacturacion";
			this.cdsPromocionVendedores.DataLibraryUrl = "";
			this.cdsPromocionVendedores.DataSetDef = "dsPromocionVendedores";
			this.cdsPromocionVendedores.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPromocionVendedores.Name = "cdsPromocionVendedores";
			this.cdsPromocionVendedores.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPromocionVendedores.SchemaDef = null;
			this.cdsPromocionVendedores.BeforeFill += new C1.Data.FillEventHandler(this.cdsPromocionVendedores_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 80);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(19, 16);
			this.label1.TabIndex = 82;
			this.label1.Text = "Fin";
			// 
			// dtFin
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFin.Appearance = appearance2;
			this.dtFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFin.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPromocionVendedores, "_PromocionVendedores.Hasta"));
			dateButton2.Caption = "Today";
			this.dtFin.DateButtons.Add(dateButton2);
			this.dtFin.Format = "dd/MM/yyyy";
			this.dtFin.Location = new System.Drawing.Point(88, 80);
			this.dtFin.Name = "dtFin";
			this.dtFin.NonAutoSizeHeight = 23;
			this.dtFin.Size = new System.Drawing.Size(112, 21);
			this.dtFin.SpinButtonsVisible = true;
			this.dtFin.TabIndex = 81;
			this.dtFin.Value = new System.DateTime(2013, 9, 13, 0, 0, 0, 0);
			this.dtFin.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFin_KeyPress);
			this.dtFin.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFin_BeforeDropDown);
			this.dtFin.ValueChanged += new System.EventHandler(this.dtFin_ValueChanged);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance3;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 100;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 113;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 66;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			ultraGridBand1.Header.Caption = "LOCALES";
			ultraGridBand1.HeaderVisible = true;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultraGrid2.Enabled = false;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(712, 152);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(200, 336);
			this.ultraGrid2.TabIndex = 182;
			this.ultraGrid2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid2_InitializeLayout);
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
			// chkTodos
			// 
			this.chkTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTodos.Appearance = appearance10;
			this.chkTodos.Enabled = false;
			this.chkTodos.Location = new System.Drawing.Point(712, 120);
			this.chkTodos.Name = "chkTodos";
			this.chkTodos.Size = new System.Drawing.Size(168, 22);
			this.chkTodos.TabIndex = 186;
			this.chkTodos.Text = "Seleccionar todos los Locales";
			this.chkTodos.CheckedChanged += new System.EventHandler(this.chkTodos_CheckedChanged);
			// 
			// cmbBodega
			// 
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DisplayMember = "Codigo";
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(760, 320);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(144, 64);
			this.cmbBodega.TabIndex = 183;
			this.cmbBodega.Text = "ultraDropDown1";
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_PromocionVendedores.PromocionVendedores - DetPromocionVendedores";
			this.ultraGrid1.DataSource = this.cdsPromocionVendedores;
			appearance11.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance11;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "DetPromocionVendedores";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Width = 451;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance12;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Premio por $";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.MaskInput = "";
			ultraGridColumn9.Width = 116;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance13;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Premio por %";
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.MaskInput = "";
			ultraGridColumn10.MaxValue = 100;
			ultraGridColumn10.MinValue = null;
			ultraGridColumn10.Width = 116;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand2.Header.Caption = "ARTÍCULOS";
			ultraGridBand2.HeaderVisible = true;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance14;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 10F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance16;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.ultraGrid1.Enabled = false;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(5, 152);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(704, 336);
			this.ultraGrid1.TabIndex = 132;
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
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
			// txtIdPromocion
			// 
			this.txtIdPromocion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdPromocion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPromocionVendedores, "_PromocionVendedores.idPromocionVendedor"));
			this.txtIdPromocion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdPromocion.Enabled = false;
			this.txtIdPromocion.Location = new System.Drawing.Point(888, 8);
			this.txtIdPromocion.Name = "txtIdPromocion";
			this.txtIdPromocion.PromptChar = ' ';
			this.txtIdPromocion.Size = new System.Drawing.Size(24, 22);
			this.txtIdPromocion.TabIndex = 133;
			this.txtIdPromocion.Visible = false;
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_PromocionVendedores";
			this.barraDato1.DataNombreId = "idPromocionVendedor";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsPromocionVendedores;
			this.barraDato1.DataTabla = "PromocionVendedores";
			this.barraDato1.DataTablaHija = "DetPromocionVendedores";
			this.barraDato1.Location = new System.Drawing.Point(8, 8);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(256, 26);
			this.barraDato1.TabIndex = 182;
			this.barraDato1.VisibleBorrar = false;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(16, 112);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(61, 16);
			this.lblCliente.TabIndex = 185;
			this.lblCliente.Text = "Descripción";
			// 
			// txtDescripcion
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance20;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsPromocionVendedores, "_PromocionVendedores.Descripcion"));
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(88, 112);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(424, 22);
			this.txtDescripcion.TabIndex = 184;
			this.txtDescripcion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescripcion_KeyPress);
			this.txtDescripcion.ValueChanged += new System.EventHandler(this.txtDescripcion_ValueChanged);
			// 
			// PromocionVendedores
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(914, 496);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtIdPromocion);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFin);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtInicio);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.chkTodos);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "PromocionVendedores";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "SPIF para Vendedores";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.PromocionVendedores_KeyDown);
			this.Load += new System.EventHandler(this.PromocionVendedores_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPromocionVendedores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void PromocionVendedores_Load(object sender, System.EventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];

			barraDato1.HabilitaControles(false);
			barraDato1.PosUltima();
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			
				
				switch(e.KeyValue)
				{
					case (int) Keys.F3:
						using (BuscaArticulos miBusqueda = new BuscaArticulos("", 1, 0)) //(int)this.cmbBodega.Value
						{
							if (DialogResult.OK == miBusqueda.ShowDialog())
							{
								ultraGrid1.ActiveRow.Cells["idArticulo"].Value = miBusqueda.idArticulo;
								ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
								ultraGrid1.ActiveRow.Cells["Articulo"].Value = miBusqueda.sArticulo;
							}
						}
						break;
						#region Page Down Abre lista desplegable
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;
						#endregion Page Down Abre lista desplegable
						#region Flechas
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
						#endregion Flechas
						#region Enter
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
		}

		private void dtInicio_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtInicio_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.dtFin.Focus();
		}

		private void dtFin_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
		}

		private void dtFin_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDescripcion.Focus();
		}

		private void cdsPromocionVendedores_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPromocionVendedores.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drPromocionVendedores = cdsPromocionVendedores.TableViews["PromocionVendedores"].Rows[0];

				string sSQL = string.Format("Exec ListaBodegasSPIFPromocion {0}, 1", (int)drPromocionVendedores["idPromocionVendedor"]);
				this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			}
			catch{};
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			barraDato1.CrearRegistro();

			string sSQL = string.Format("Exec ListaBodegasSPIFPromocion {0}, 0", 0);
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			this.ultraDataSource1.ResetBand();

			drPromocionVendedores["Desde"] = DateTime.Today;
			drPromocionVendedores["Hasta"] = DateTime.Today;
			
			this.chkTodos.Enabled = true;
			this.ultraGrid2.Enabled = true;
			this.dtInicio.Focus();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			barraDato1.EditarRegistro();
			this.chkTodos.Enabled = true;
			this.ultraGrid2.Enabled = true;
			this.dtInicio.Focus();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			barraDato1.DeshacerRegistro();
			this.dtFin.CalendarInfo.MinDate = DateTime.MinValue;
			this.ultraGrid2.Enabled = false;
			this.chkTodos.Enabled = false;
			barraDato1.ToolTipEditar(Funcion.Tiempo("Deshacer"));
		}


		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			int CantBod = 0;
			bool Sel = false;

			if ((int) this.txtDescripcion.TextLength == 0)
			{
				MessageBox.Show("Ingrese la descripción.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtDescripcion.Focus();
				return;
			}

			if (this.dtFin.Value == null) 
			{
				MessageBox.Show("Ingrese la fecha final.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.dtFin.Focus();
				return;
			}

			if ((DateTime)this.dtFin.Value < (DateTime)this.dtInicio.Value) 
			{
				MessageBox.Show("La fecha final no puede ser menor a la fecha inicial.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.dtFin.Focus();
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				Sel = (bool) dr.Cells["Sel"].Value;
				if (Sel) CantBod++;
			}

			if ((int)CantBod == 0)
			{
				MessageBox.Show("Seleccione bodega(s).", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if ((int)this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese el(los) articulo(s).", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["Articulo"].Value == System.DBNull.Value || dr.Cells["Articulo"].Value.ToString().Length == 0)
				{
					MessageBox.Show(string.Format("Ingrese un artículo o elimine la fila.",  "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error));	
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["Articulo"].Value != System.DBNull.Value || dr.Cells["Articulo"].Value.ToString().Length == 0)
				{
					if ((double)dr.Cells["Valor"].Value == 0 && (double)dr.Cells["Porcentaje"].Value == 0)
					{
						MessageBox.Show(string.Format("Ingrese Porcentaje o Valor en el artículo '{0}'.", (string)dr.Cells["Articulo"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}
			}
			
			barraDato1.GrabaRegistro();
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				int Bod = (int)dr.Cells["Bodega"].Value;
				Sel = (bool)dr.Cells["Sel"].Value;
				if (Sel)
				{
					string sSQL = string.Format("Exec GuardaBodegaPromocion {0}, {1}, {2}", Bod, drPromocionVendedores["idPromocionVendedor"], Sel);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				}
			}

			this.chkTodos.Enabled = false;
			this.ultraGrid2.Enabled = false;
		}

		private void BorraErrores()
		{
			errorProvider.SetError(this.dtFin, ""); 
			errorProvider.SetError(this.txtDescripcion, "");
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico

			#region Decimales para el Valor y Porcentaje
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Valor"].Format = stFormato;
			e.Layout.Bands[0].Columns["Porcentaje"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Valor"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Porcentaje"].MaskInput = stInput;
      #endregion Decimales para el Valor y Porcentaje
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void dtInicio_Validated(object sender, System.EventArgs e)
		{
			DateTime Inicio = (DateTime)drPromocionVendedores["Desde"];
			this.dtFin.CalendarInfo.MinDate = Inicio;
			if ((DateTime)drPromocionVendedores["Hasta"] < (DateTime)drPromocionVendedores["Desde"])
			{
				drPromocionVendedores["Hasta"] = (DateTime)drPromocionVendedores["Desde"];
			}
		}

		private void txtDescripcion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtDescripcion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraGrid1.Focus();
		}

		private void chkTodos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if (this.chkTodos.Checked) dr.Cells["Sel"].Value = 1;
				else dr.Cells["Sel"].Value = 0;
			}
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Porcentaje")
			{ 
				double Valor = (double)e.Cell.Row.Cells["Valor"].Value;
				if (Valor > 0)
				{
					e.Cancel = true;
					MessageBox.Show("Si escribe un valor en $ no puede escribir un valor en %.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			
			if (e.Cell.Column.ToString() == "Valor")
			{ 
				double Porcentaje = (double)e.Cell.Row.Cells["Porcentaje"].Value;
				if (Porcentaje > 0)
				{
					e.Cancel = true;
					MessageBox.Show("Si escribe un valor en % no puede escribir un valor en $.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void PromocionVendedores_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) 
			{
				if (this.barraDato1.bEditar)
				{
					this.barraDato1_Deshacer(sender, e);
				}
				else
					Close();
			}
			if (e.KeyCode == Keys.N)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Crear(sender, e);
				}
			}
			if (e.KeyCode == Keys.E)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Editar(sender, e);
				}
			}
			if (e.KeyCode == Keys.C)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Buscar(sender, e);
				}
			}
			if (e.KeyCode == Keys.F2)
			{
				if (this.barraDato1.bEditar)
				{
					this.barraDato1_Grabar(sender, e);
				}
			}
			if (e.KeyCode == Keys.F5)
			{
				if (!this.barraDato1.bEditar)
				{
					barraDato1.IdRegistro = (int)drPromocionVendedores["idPromocionVendedor"];
					barraDato1.ProximoId(5);					
				}
			}
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using (frmBuscaPromocionVendedor miBusqueda = new frmBuscaPromocionVendedor())
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					try
					{
						barraDato1.IdRegistro = miBusqueda.idPromocionVendedor;
						barraDato1.ProximoId(5);
					}
					catch{}
				}
			}
		}

		private void dtFin_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void ultraGrid2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}

