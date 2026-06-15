using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmConteoDeInventario.
	/// </summary>
	public class frmConteoDeInventario : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label55;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridConteo;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnTomaDeInventario;
		private System.Windows.Forms.Button btnCerrarInventario;
		private System.Windows.Forms.Label lblSeguimiento;
		private System.Windows.Forms.ComboBox cboFecha;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConteoDeInventario()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConteoDeInventario));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TFSalida");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TFEntrada");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Existencia");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Conteo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencias");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TFSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TFEntrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Existencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Conteo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencias");
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label7 = new System.Windows.Forms.Label();
			this.label55 = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridConteo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnTomaDeInventario = new System.Windows.Forms.Button();
			this.btnCerrarInventario = new System.Windows.Forms.Button();
			this.lblSeguimiento = new System.Windows.Forms.Label();
			this.cboFecha = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridConteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(968, 8);
			this.groupBox1.TabIndex = 370;
			this.groupBox1.TabStop = false;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 43);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(42, 17);
			this.label7.TabIndex = 369;
			this.label7.Text = "Artículo";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.BackColor = System.Drawing.Color.Transparent;
			this.label55.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label55.Location = new System.Drawing.Point(8, 11);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(41, 17);
			this.label55.TabIndex = 368;
			this.label55.Text = "Bodega";
			this.label55.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label53.Location = new System.Drawing.Point(360, 11);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 17);
			this.label53.TabIndex = 367;
			this.label53.Text = "Fecha";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtArticulo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance1;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtArticulo.Location = new System.Drawing.Point(72, 40);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(264, 22);
			this.txtArticulo.TabIndex = 2;
			this.txtArticulo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArticulo_KeyDown);
			// 
			// btnExcel
			// 
			this.btnExcel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(656, 72);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 4;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(360, 39);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 3;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(72, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(264, 22);
			this.cmbBodega.TabIndex = 0;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// dtFecha
			// 
			appearance2.FontData.Name = "Tahoma";
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance2;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(416, 9);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			// 
			// uGridConteo
			// 
			this.uGridConteo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridConteo.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridConteo.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridConteo.DisplayLayout.Appearance = appearance3;
			this.uGridConteo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 44;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Articulo";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 276;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 85;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Header.Caption = "TF Salida";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 114;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Header.Caption = "TF Entrada";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 114;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance6;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Width = 114;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance7;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 114;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance8;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Width = 114;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.uGridConteo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridConteo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridConteo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridConteo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridConteo.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Tahoma";
			appearance10.FontData.SizeInPoints = 8F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridConteo.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.uGridConteo.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConteo.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConteo.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.uGridConteo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridConteo.Location = new System.Drawing.Point(8, 120);
			this.uGridConteo.Name = "uGridConteo";
			this.uGridConteo.Size = new System.Drawing.Size(952, 248);
			this.uGridConteo.TabIndex = 366;
			this.uGridConteo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridConteo_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
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
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(496, 43);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 371;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnTomaDeInventario
			// 
			this.btnTomaDeInventario.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnTomaDeInventario.Image = ((System.Drawing.Image)(resources.GetObject("btnTomaDeInventario.Image")));
			this.btnTomaDeInventario.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTomaDeInventario.Location = new System.Drawing.Point(737, 72);
			this.btnTomaDeInventario.Name = "btnTomaDeInventario";
			this.btnTomaDeInventario.Size = new System.Drawing.Size(128, 24);
			this.btnTomaDeInventario.TabIndex = 373;
			this.btnTomaDeInventario.Text = "Toma De Inventario";
			this.btnTomaDeInventario.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnTomaDeInventario.Click += new System.EventHandler(this.btnTomaDeInventario_Click);
			// 
			// btnCerrarInventario
			// 
			this.btnCerrarInventario.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCerrarInventario.Image = ((System.Drawing.Image)(resources.GetObject("btnCerrarInventario.Image")));
			this.btnCerrarInventario.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCerrarInventario.Location = new System.Drawing.Point(871, 72);
			this.btnCerrarInventario.Name = "btnCerrarInventario";
			this.btnCerrarInventario.Size = new System.Drawing.Size(88, 24);
			this.btnCerrarInventario.TabIndex = 376;
			this.btnCerrarInventario.Text = "Diferencias";
			this.btnCerrarInventario.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCerrarInventario.Click += new System.EventHandler(this.btnCerrarInventario_Click);
			// 
			// lblSeguimiento
			// 
			this.lblSeguimiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblSeguimiento.AutoSize = true;
			this.lblSeguimiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSeguimiento.Location = new System.Drawing.Point(16, 43);
			this.lblSeguimiento.Name = "lblSeguimiento";
			this.lblSeguimiento.Size = new System.Drawing.Size(0, 17);
			this.lblSeguimiento.TabIndex = 377;
			this.lblSeguimiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cboFecha
			// 
			this.cboFecha.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboFecha.Location = new System.Drawing.Point(848, 9);
			this.cboFecha.MaxDropDownItems = 15;
			this.cboFecha.Name = "cboFecha";
			this.cboFecha.Size = new System.Drawing.Size(112, 21);
			this.cboFecha.TabIndex = 387;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(792, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 386;
			this.label1.Text = "Fecha";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmConteoDeInventario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(968, 374);
			this.Controls.Add(this.cboFecha);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.lblSeguimiento);
			this.Controls.Add(this.btnCerrarInventario);
			this.Controls.Add(this.btnTomaDeInventario);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.uGridConteo);
			this.KeyPreview = true;
			this.Name = "frmConteoDeInventario";
			this.Text = "Conteo De Inventario";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConteoDeInventario_KeyDown);
			this.Load += new System.EventHandler(this.frmConteoDeInventario_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridConteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmConteoDeInventario_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0329");

			if (!Funcion.Permiso("1946", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Conteo de Inventario ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}


			string sSQL = string.Format("Exec ListaBodegas {0}, 47", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			
			int iBodega = 0;

			sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodega > 0)
			{
				this.cmbBodega.Value = iBodega;
				this.cmbBodega.Enabled = false;
			}
			else this.cmbBodega.Value = 1;
									
			this.dtFecha.Value = DateTime.Today;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today.AddDays(2);			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridConteo);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Bodega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
										
			//				string sSQLT = string.Format("Exec ConteoValidaTransferencias {0}, '{1}'", (int)this.cmbBodega.Value, dtFecha.ToString("yyyyMMdd"));
			//				string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLT);
			//
			//				if (sMensaje.Length > 0)
			//				{
			//					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				}			
			//				else
			//				{
			
			int iBodega = 0;				
						
			if (this.cmbBodega.ActiveRow != null) iBodega = (int) this.cmbBodega.Value;
		
			string sSQL = string.Format("Exec ListaConteoDeInventario {0}, '{1}', '{2}'", 
				iBodega, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtArticulo.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridConteo);					

			this.lblContador.Text = this.uGridConteo.Rows.Count + " ARTICULOS ENCONTRADOS";
			//				}			
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFecha.Focus();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtArticulo.Focus();
		}

		private void txtArticulo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void btnTomaDeInventario_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime)this.dtFecha.Value;
						
			if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Bodega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (this.uGridConteo.Rows.Count == 0)
			{
				MessageBox.Show("Cargue Primero el Inventario a la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			Cursor = Cursors.WaitCursor;

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = '{0}' And CONVERT(DATE, Fecha) = '{1}'", (int)this.cmbBodega.Value, dtFecha.ToString("yyyyMMdd"))) == 0)
			{
				if (DialogResult.Yes == MessageBox.Show("No existe Conteo de Inventario de esta Bodega en esta Fecha \n\n\n¿Desea Crearlo?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))	
				{
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = {0} And CONVERT(DATE, Fecha) <> '{1}' And Estado = 1", (int)this.cmbBodega.Value, dtFecha.ToString("yyyyMMdd"))) > 0)
					{
						MessageBox.Show("Existen conteos abiertos de esta bodega, cierre antes de generar un nuevo conteo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					using (frmCIInformeTF informe = new frmCIInformeTF((int)this.cmbBodega.Value))
					{							
						if (DialogResult.OK == informe.ShowDialog())
						{
							frmTomaInventario TomaInventario = new frmTomaInventario((int)this.cmbBodega.Value, dtFecha);
							TomaInventario.ShowDialog();
						}
					}
				}
			}
			else
			{
				frmTomaInventario TomaInventario = new frmTomaInventario((int)this.cmbBodega.Value, dtFecha);
				TomaInventario.ShowDialog();
			}
				
			Cursor = Cursors.Default;
		}

		private void Aceptar()
		{
			/*
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
													
				try
				{		
					this.Cursor = Cursors.WaitCursor;

					oTransaction = oConexion.BeginTransaction();
								
					SqlCommand oCmdIngreso = oConexion.CreateCommand();
					oCmdIngreso.Connection = oConexion;
					oCmdIngreso.CommandType = CommandType.StoredProcedure;
					oCmdIngreso.CommandTimeout = 0;
					oCmdIngreso.Transaction = oTransaction;
					oCmdIngreso.CommandText = "CreaAjusteConteoInventario";

					oCmdIngreso.Parameters.Add("@Bodega", SqlDbType.Int).Value = (int)this.cmbBodega.Value;
					oCmdIngreso.Parameters.Add("@Fecha", SqlDbType.DateTime).Value = (DateTime)this.dtFecha.Value;
					string sNumeroAjuste = oCmdIngreso.ExecuteScalar().ToString();

					this.lblSeguimiento.Text = "AJUSTE TERMINADO";

					SqlCommand oCmdEgreso = oConexion.CreateCommand();
					oCmdEgreso.Connection = oConexion;
					oCmdEgreso.CommandType = CommandType.StoredProcedure;
					oCmdEgreso.CommandTimeout = 0;
					oCmdEgreso.Transaction = oTransaction;
					oCmdEgreso.CommandText = "CreaAjusteEgresoConteoInventario";

					oCmdEgreso.Parameters.Add("@Bodega", SqlDbType.Int).Value = (int)this.cmbBodega.Value;
					oCmdEgreso.Parameters.Add("@Fecha", SqlDbType.DateTime).Value = (DateTime)this.dtFecha.Value;
					string sNumeroAjusteEgreso = oCmdEgreso.ExecuteScalar().ToString();

					//								string sSQLAjusteEgreso = string.Format("Exec CreaAjusteEgresoConteoInventario '{0}', {1}", dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value);					
					//								oCmdEgreso.CommandText = sSQLAjusteEgreso;
								
					this.lblSeguimiento.Text = "EGRESO TERMINADO";

					SqlCommand oCmdTransferencia = oConexion.CreateCommand();
					oCmdTransferencia.Connection = oConexion;
					oCmdTransferencia.CommandType = CommandType.StoredProcedure;
					oCmdTransferencia.CommandTimeout = 0;
					oCmdTransferencia.Transaction = oTransaction;
					oCmdTransferencia.CommandText = "CreaTransferenciaFaltantesConteo";

					oCmdTransferencia.Parameters.Add("@Bodega", SqlDbType.Int).Value = (int)this.cmbBodega.Value;
					oCmdTransferencia.Parameters.Add("@Fecha", SqlDbType.DateTime).Value = (DateTime)this.dtFecha.Value;								
					string sNumeroTransferencia = oCmdTransferencia.ExecuteScalar().ToString();
					
					//								string sSQLTransferencia = string.Format("Exec CreaTransferenciaFaltantesConteo '{0}', {1}", dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value);					
					//								oCmdTransferencia.CommandText = sSQLTransferencia;

					this.lblSeguimiento.Text = "TRANSFERENCIA TERMINADA";

					SqlCommand oCmdSeriales = oConexion.CreateCommand();
					oCmdSeriales.Connection = oConexion;
					oCmdSeriales.CommandType = CommandType.StoredProcedure;
					oCmdSeriales.CommandTimeout = 0;
					oCmdSeriales.Transaction = oTransaction;
					oCmdSeriales.CommandText = "GeneraMovimientoSerialesDesdeConteo";

					oCmdSeriales.Parameters.Add("@Bodega", SqlDbType.Int).Value = (int)this.cmbBodega.Value;
					oCmdSeriales.Parameters.Add("@Fecha", SqlDbType.DateTime).Value = (DateTime)this.dtFecha.Value;
					oCmdSeriales.ExecuteNonQuery();

					//								string sSQLSeriales = string.Format("Exec GeneraMovimientoSerialesDesdeConteo '{0}', {1}", dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value);					
					//								oCmdActualiza.CommandText = sSQLSeriales;
								
					this.lblSeguimiento.Text = "AJUSTE SERIALES TERMINADO";

					SqlCommand oCmdCierre = oConexion.CreateCommand();
					oCmdCierre.Connection = oConexion;
					oCmdCierre.CommandType = CommandType.StoredProcedure;
					oCmdCierre.CommandTimeout = 0;
					oCmdCierre.Transaction = oTransaction;
					oCmdCierre.CommandText = "ActualizaCierreConteoInventario";

					oCmdCierre.Parameters.Add("@Fecha", SqlDbType.DateTime).Value = (DateTime)this.dtFecha.Value;
					oCmdCierre.Parameters.Add("@Bodega", SqlDbType.Int).Value = (int)this.cmbBodega.Value;
					oCmdCierre.Parameters.Add("@idCierre", SqlDbType.Int).Value = (int)informe.txtidCierre.Value;
					oCmdCierre.Parameters.Add("@NumeroTransferencia", SqlDbType.VarChar).Value = sNumeroTransferencia;
					oCmdCierre.Parameters.Add("@NumeroAjuste", SqlDbType.VarChar).Value = sNumeroAjuste;	
					oCmdCierre.Parameters.Add("@NumeroEgreso", SqlDbType.VarChar).Value = sNumeroAjusteEgreso;
					oCmdCierre.ExecuteNonQuery();

					//								string sSQLCierre = string.Format("Exec ActualizaCierreConteoInventario '{0}', {1}, {2}, '{3}', '{4}', '{5}'", 
					//									dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, (int)informe.txtidCierre.Value, sNumeroTransferencia, sNumeroAjuste, sNumeroAjusteEgreso);
					//								oCmdActualiza.CommandText = sSQLCierre;
								
					this.lblSeguimiento.Text = "CIERRE TERMINADO";

					oTransaction.Commit();
						
					this.lblSeguimiento.Text = "GUARDA CAMBIOS EN BDD";
	
					oConexion.Close();

					this.lblSeguimiento.Text = "CONEXION CERRADA - FINALIZO CORRECTAMENTE";

					this.Cursor = Cursors.Default;

					MessageBox.Show("Inventario de la Bodega " + this.cmbBodega.Text.ToString() + " Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				catch (Exception ex)
				{
					try
					{
						oTransaction.Rollback();
									
						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);

						this.Cursor = Cursors.Default;
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							

						this.Cursor = Cursors.Default;
					}
				}
				finally
				{
					oConexion.Close();
				}
			}	
			*/
		}
		
		public string Modulo11(string number)
		{
			int pivote = 2;
			int longitudCadena = number.Length;
			int cantidadTotal = 0;
			int b = 1;

			for (int i = 0; i < longitudCadena; i++)
			{
				if (pivote == 8) pivote = 2;

				int temporal = int.Parse(number.Substring(i, b));

				temporal *= pivote;

				pivote++;

				cantidadTotal += temporal;
			}

			cantidadTotal = 11 - cantidadTotal % 11;

			if (cantidadTotal == 11) cantidadTotal = 0;
			else if (cantidadTotal == 10) cantidadTotal = 1;

			return cantidadTotal.ToString();
		}

		public string invertirCadena(String cadena)
		{
			String cadenaInvertida = "";
			for (int x = cadena.Length - 1; x >= 0; x--)
				cadenaInvertida = cadenaInvertida + cadena.Substring(x, 1);

			return cadenaInvertida;
		}

		private void btnCerrarInventario_Click(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Bodega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = {0} And CONVERT(DATE, Fecha) = '{1}'", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"))) == 0)
			{
				MessageBox.Show("No existe Conteo de esta Bodega en esta Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
						
			using (frmInformeConteo informe = new frmInformeConteo((DateTime)this.dtFecha.Value, (int)this.cmbBodega.Value))
			{				
				if (DialogResult.OK == informe.ShowDialog())
				{
					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 2048;"))
					{
						oConexion.Open();

						SqlTransaction oTransaction = null;
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandType = CommandType.Text;
						oCmdActualiza.CommandTimeout = 240;
						
						try
						{
							oTransaction = oConexion.BeginTransaction();
							oCmdActualiza.Transaction = oTransaction;

							int idIngresoBodega = 0;
							int idEgresoBodega = 0;
							//							int idTransferencia = 0;
							int idFactura = 0;

							#region Crea Ajuste de Ingreso
							oCmdActualiza.CommandText = string.Format("Exec CIValidaCreaAjusteDeIngreso {0}, '{1}'", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
							int iAjusteIngreso = (int)oCmdActualiza.ExecuteScalar();
							
							if (iAjusteIngreso > 0)
							{								
								oCmdActualiza.CommandText = string.Format("Exec CICreaAjusteIngreso {0}, '{1}'", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
								idIngresoBodega = (int)oCmdActualiza.ExecuteScalar();

								/*
								oCmdActualiza.CommandText = string.Format("Exec CIListaAjusteDeIngreso {0}, '{1}'", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
								SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
								DataTable oDTDetalle = new DataTable();
								oDADetalle.Fill(oDTDetalle);
																			
								foreach (DataRow oDRDetalleIngreso in oDTDetalle.Rows)
								{								
									string sSQLDetalleIngreso = string.Format("Insert Into DetCompra (idCompra, idArticulo, Codigo, Articulo, Cantidad, Precio, SubTotal, Bodega)	Values ({0}, {1}, '{2}', '{3}', {4}, {5}, {6}, {7})", 
										idIngresoBodega, (int)oDRDetalleIngreso["idArticulo"], oDRDetalleIngreso["Codigo"].ToString(), oDRDetalleIngreso["Articulo"].ToString(), (int)oDRDetalleIngreso["ConteoExistencia"], (decimal)oDRDetalleIngreso["CostoUltimo"], (decimal)oDRDetalleIngreso["Subtotal"], (int)this.cmbBodega.Value);
									oCmdActualiza.CommandText = sSQLDetalleIngreso;
									oCmdActualiza.ExecuteNonQuery();
								}

								oCmdActualiza.CommandText = string.Format("Exec CITotalizaAjustes {0}", idIngresoBodega);
								oCmdActualiza.ExecuteNonQuery();
								
								oCmdActualiza.CommandText = string.Format("Exec GuardaMovimientosInventario {0}, 9", idIngresoBodega);
								oCmdActualiza.ExecuteNonQuery();
								*/								
								oCmdActualiza.CommandText = string.Format("Exec CreaAsiIngreso {0}", idIngresoBodega);
								oCmdActualiza.ExecuteNonQuery();								
							}
							#endregion Crea Ajuste de Ingreso
				
							#region Crea Ajuste de Egreso
							oCmdActualiza.CommandText = string.Format("Exec CIValidaCreaAjusteDeEgreso {0}, '{1}'", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
							int iAjusteEgreso = (int)oCmdActualiza.ExecuteScalar();

							if (iAjusteEgreso > 0)
							{	
								
								oCmdActualiza.CommandText = string.Format("Exec NumeracionLocales 1, {0}, 1", (int)this.cmbBodega.Value);
								string sNumero = oCmdActualiza.ExecuteScalar().ToString();

								oCmdActualiza.CommandText = string.Format("Exec CICreaAjusteEgreso {0}, '{1}'", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
								idEgresoBodega = (int)oCmdActualiza.ExecuteScalar();
								
								/*
								oCmdActualiza.CommandText = string.Format("Exec CIListaAjusteDeEgreso {0}, '{1}'", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
								SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
								DataTable oDTDetalle = new DataTable();
								oDADetalle.Fill(oDTDetalle);																			
																			
								foreach (DataRow oDRDetalleEgreso in oDTDetalle.Rows)
								{
									string sSQLDetalleEgreso = string.Format("Insert Into DetCompra (idCompra, idArticulo, Codigo, Articulo, Cantidad, Precio, SubTotal, Bodega)	Values ({0}, {1}, '{2}', '{3}', {4}, {5}, {6}, {7})", 
										idEgresoBodega, (int)oDRDetalleEgreso["idArticulo"], oDRDetalleEgreso["Codigo"].ToString(), oDRDetalleEgreso["Articulo"].ToString(), (int)oDRDetalleEgreso["CantidadJustifica"], (decimal)oDRDetalleEgreso["CostoUltimo"], (decimal)oDRDetalleEgreso["Subtotal"], (int)this.cmbBodega.Value);
									oCmdActualiza.CommandText = sSQLDetalleEgreso;
									oCmdActualiza.ExecuteNonQuery();									
								}
								
								oCmdActualiza.CommandText = string.Format("Exec GuardaMovimientosInventario {0}, 8", idEgresoBodega);
								oCmdActualiza.ExecuteNonQuery();
								*/
								oCmdActualiza.CommandText = string.Format("Exec CreaAsiEgreso {0}", idEgresoBodega);
								oCmdActualiza.ExecuteNonQuery();	
							}
							#endregion Crea Ajuste de Egreso
				
							#region Crea Factura - Bloqueado Transferencia
							string sSQL = string.Format("Exec CIValidaTransferenciaFaltantes '{0}', {1}", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value);
							oCmdActualiza.CommandText = sSQL;
							int iTransf = (int)oCmdActualiza.ExecuteScalar();

							if (iTransf > 0)
							{
								oCmdActualiza.CommandText = string.Format("Exec CIValidaBodegaRemates {0}", (int)this.cmbBodega.Value);
								int iLocal = (int)oCmdActualiza.ExecuteScalar();

								string sSQLNumero = string.Format("Exec NumeracionLocales 1, {0}, 1", iLocal);
								oCmdActualiza.CommandText = sSQLNumero;
								string sNumero = oCmdActualiza.ExecuteScalar().ToString();

								oCmdActualiza.CommandText = string.Format("Exec CICreaFacturaFaltantes {0}, '{1}', '{2}', {3}", 
									(int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), sNumero, iLocal);
								idFactura = (int)oCmdActualiza.ExecuteScalar();
							
								if (Funcion.bEscalarSQL(cdsSeteoF, "Select FacturaElectronica From SeteoF", false))
								{
									oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 1", idFactura);;
									string sClaveAcceso = oCmdActualiza.ExecuteScalar().ToString();

									string cadInvertida = invertirCadena(sClaveAcceso.Substring(0, 48));
									string ValM11 = Modulo11(cadInvertida).ToString().Trim();
									string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

									oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, idFactura);
									oCmdActualiza.ExecuteNonQuery();
								}

								//oCmdActualiza.CommandText = string.Format("Exec CreaTransferenciaFaltantesConteo '{0}', {1}", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value);					
								//idTransferencia = (int)oCmdActualiza.ExecuteScalar();
							}
							#endregion Crea Factura - Bloqueado Transferencia

							#region Seriales
							oCmdActualiza.CommandText = string.Format("Exec CIRegistroSerialesContadas '{0}', {1}", 
								Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value);
							oCmdActualiza.ExecuteNonQuery();
							/*
							oCmdActualiza.CommandText = string.Format("Exec CIListaSerialesConteo '{0}', {1}", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value);
							SqlDataAdapter oDASeriales = new SqlDataAdapter(oCmdActualiza);
							DataTable oDTSeriales = new DataTable();
							oDASeriales.Fill(oDTSeriales);	
													
							foreach (DataRow oDRSeriales in oDTSeriales.Rows)
							{
								if ((int)oDRSeriales["idSerial"] > 0)
								{
									#region Serial Existente
									oCmdActualiza.CommandText = string.Format("Update DetalleSeriales Set Disponible = 0 Where idSerial = {0}", (int)oDRSeriales["idSerial"]);
									oCmdActualiza.ExecuteNonQuery();

									if (!(bool)oDRSeriales["CambioArticulo"])
									{
										oCmdActualiza.CommandText = string.Format("Update Seriales Set Bodega = {1}, Estado = 1 Where idSerial = {0}", (int)oDRSeriales["idSerial"], (int)this.cmbBodega.Value);
										oCmdActualiza.ExecuteNonQuery();

										oCmdActualiza.CommandText = string.Format("Exec GuardaDetalleSeriales {0}, 0, 0, {1}, '', 1, 'MOVIMIENTO REGISTRADO DESDE CONTEO DE INVENTARIO', 26", (int)oDRSeriales["idSerial"], (int)this.cmbBodega.Value);
										oCmdActualiza.ExecuteNonQuery();
									}
									else
									{
										#region Cambia estado 24 la Serial y genera un Movimiento de bloqueo
										oCmdActualiza.CommandText = string.Format("Update Seriales Set Bodega = {1}, Estado = 24 Where idSerial = {0}", (int)oDRSeriales["idSerial"], (int)this.cmbBodega.Value);
										oCmdActualiza.ExecuteNonQuery();

										oCmdActualiza.CommandText = string.Format("Exec GuardaDetalleSeriales {0}, 0, 0, {1}, '', 1, 'BLOQUEADO POR CAMBIO DE PRODUCTO REGISTRADO DESDE CONTEO DE INVENTARIO', 24",	(int)oDRSeriales["idSerial"], (int)this.cmbBodega.Value);
										oCmdActualiza.ExecuteNonQuery();
										#endregion Cambia estado 24 la Serial y genera un Movimiento de bloqueo

										#region Registra Serial Nueva
										oCmdActualiza.CommandText = string.Format("Exec GuardaSeriales 0, {0}, '{1}', 1, {2}", (int)oDRSeriales["idArticulo"], oDRSeriales["Serial"].ToString(), (int)this.cmbBodega.Value);
										int idSerialNueva = (int)oCmdActualiza.ExecuteScalar();

										oCmdActualiza.CommandText = string.Format("Exec GuardaDetalleSeriales {0}, 0, 0, {1}, '', 1, 'CAMBIO DE PRODUCTO REGISTRADO DESDE CONTEO DE INVENTARIO', 100", idSerialNueva, (int)this.cmbBodega.Value);
										oCmdActualiza.ExecuteNonQuery();
										#endregion Registra Serial Nueva
									}
									#endregion Serial Existente
								}
								else if ((int)oDRSeriales["idSerial"] == 0)
								{									
									oCmdActualiza.CommandText = string.Format("Exec GuardaSeriales 0, {0}, '{1}', 1, {2}", 
										(int)oDRSeriales["idArticulo"], oDRSeriales["Serial"].ToString(), (int)this.cmbBodega.Value);
									int idSerialNueva = (int)oCmdActualiza.ExecuteScalar();
																		
									oCmdActualiza.CommandText = string.Format("Exec GuardaDetalleSeriales {0}, 0, 0, {1}, '', 1, 'SERIAL REGISTRADA DESDE CONTEO DE INVENTARIO', 26", idSerialNueva, (int)this.cmbBodega.Value);
									oCmdActualiza.ExecuteNonQuery();
								}
							}
							*/
							#endregion Seriales

							#region Bloqueo Seriales
							oCmdActualiza.CommandText = string.Format("Exec CIListaSeriesNoContadas '{0}', {1}", 
								Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value);
							oCmdActualiza.ExecuteNonQuery();
							/*
							SqlDataAdapter oDASerialesNC = new SqlDataAdapter(oCmdActualiza);
							DataTable oDTSerialesNC = new DataTable();
							oDASerialesNC.Fill(oDTSerialesNC);	
													
							foreach (DataRow oDRSerialesNC in oDTSerialesNC.Rows)
							{								
								oCmdActualiza.CommandText = string.Format("Update DetalleSeriales Set Disponible = 0 Where idSerial = {0}", (int)oDRSerialesNC["idSerial"]);
								oCmdActualiza.ExecuteNonQuery();
																	
								oCmdActualiza.CommandText = string.Format("Exec GuardaDetalleSeriales {0}, 0, {1}, {2}, '', 1, 'SERIAL BLOQUEADA DESDE CONTEO DE INVENTARIO', 25", 
									(int)oDRSerialesNC["idSerial"], (int)oDRSerialesNC["idDetalleSerial"], (int)this.cmbBodega.Value);
								oCmdActualiza.ExecuteNonQuery();

								oCmdActualiza.CommandText = string.Format("Update Seriales Set Estado = 25, Bodega = {1} Where idSerial = {0}", (int)oDRSerialesNC["idSerial"], (int)this.cmbBodega.Value);
								oCmdActualiza.ExecuteNonQuery();
							}
							*/
							#endregion Bloqueo Seriales

							string sSQLCierre = string.Format("Exec ActualizaCierreConteoInventario '{0}', {1}, {2}, {3}, {4}, {5}, '{6}'", 
								Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value, (int)informe.txtidCierre.Value, idIngresoBodega, idEgresoBodega, idFactura, informe.txtNotas.Text.ToString());
							oCmdActualiza.CommandText = sSQLCierre;
							oCmdActualiza.ExecuteNonQuery();
							
							oTransaction.Commit();	
					
							oConexion.Close();

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
		}

		private void frmConteoDeInventario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null) return;

			string sSQL = string.Format("Select Fecha From CierreConteoInventario Where Bodega = {0} And Estado = 2 Order By Fecha", (int)this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
			this.cboFecha.Items.Clear();

			DateTime Fecha = DateTime.Today;
			
			while (dr.Read())
			{
				Fecha = dr.GetDateTime(0);
				this.cboFecha.Items.Add(Fecha.ToString("dd/MM/yyyy"));
			}
			dr.Close();

			this.cboFecha.Text = Fecha.ToString("dd/MM/yyyy");
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridConteo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}

