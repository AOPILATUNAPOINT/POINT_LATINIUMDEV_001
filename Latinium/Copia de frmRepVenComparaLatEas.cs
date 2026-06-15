using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRepVenComparaLatEas.
	/// </summary>
	public class frmRepVenComparaLatEas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button cmdBuscar;
		private System.Windows.Forms.Button cmbExcel;
		private System.Windows.Forms.DateTimePicker txtFechaInicia;
		private System.Windows.Forms.DateTimePicker txtFechaFinaliza;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdComisiones;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.GroupBox groupBox3;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdLocales;
		private DevExpress.XtraEditors.CheckEdit chkLocls;
		private System.ComponentModel.IContainer components;

		public frmRepVenComparaLatEas()
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRepVenComparaLatEas));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CODDEPPREF");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nDia");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanFAC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NC_Easy");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anul_Easy");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTot");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Totl");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FLat");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NCLat");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anu_LAt");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtLat");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TivaLat");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotLat");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CODDEPPREF");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("nDia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanFAC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NC_Easy");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anul_Easy");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTot");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TIva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Totl");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FLat");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NCLat");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anu_LAt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtLat");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TivaLat");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotLat");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Selct");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtFechaFinaliza = new System.Windows.Forms.DateTimePicker();
			this.label2 = new System.Windows.Forms.Label();
			this.txtFechaInicia = new System.Windows.Forms.DateTimePicker();
			this.cmdBuscar = new System.Windows.Forms.Button();
			this.cmbExcel = new System.Windows.Forms.Button();
			this.grdComisiones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.grdLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.chkLocls = new DevExpress.XtraEditors.CheckEdit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdComisiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkLocls.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(24, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Inicia:";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtFechaFinaliza);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.txtFechaInicia);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(176, 80);
			this.groupBox1.TabIndex = 1;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Rango de FECHAS";
			// 
			// txtFechaFinaliza
			// 
			this.txtFechaFinaliza.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFechaFinaliza.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.txtFechaFinaliza.Location = new System.Drawing.Point(72, 46);
			this.txtFechaFinaliza.Name = "txtFechaFinaliza";
			this.txtFechaFinaliza.Size = new System.Drawing.Size(88, 20);
			this.txtFechaFinaliza.TabIndex = 6;
			this.txtFechaFinaliza.Value = new System.DateTime(2014, 6, 23, 0, 0, 0, 0);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(11, 48);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(47, 16);
			this.label2.TabIndex = 2;
			this.label2.Text = "Finaliza:";
			// 
			// txtFechaInicia
			// 
			this.txtFechaInicia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFechaInicia.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.txtFechaInicia.Location = new System.Drawing.Point(72, 22);
			this.txtFechaInicia.Name = "txtFechaInicia";
			this.txtFechaInicia.Size = new System.Drawing.Size(88, 20);
			this.txtFechaInicia.TabIndex = 5;
			this.txtFechaInicia.Value = new System.DateTime(2014, 6, 23, 0, 0, 0, 0);
			// 
			// cmdBuscar
			// 
			this.cmdBuscar.Image = ((System.Drawing.Image)(resources.GetObject("cmdBuscar.Image")));
			this.cmdBuscar.Location = new System.Drawing.Point(48, 104);
			this.cmdBuscar.Name = "cmdBuscar";
			this.cmdBuscar.Size = new System.Drawing.Size(48, 40);
			this.cmdBuscar.TabIndex = 2;
			this.toolTip1.SetToolTip(this.cmdBuscar, "Buscar INFORMACION");
			this.cmdBuscar.Click += new System.EventHandler(this.cmdBuscar_Click);
			// 
			// cmbExcel
			// 
			this.cmbExcel.Image = ((System.Drawing.Image)(resources.GetObject("cmbExcel.Image")));
			this.cmbExcel.Location = new System.Drawing.Point(96, 104);
			this.cmbExcel.Name = "cmbExcel";
			this.cmbExcel.Size = new System.Drawing.Size(48, 40);
			this.cmbExcel.TabIndex = 3;
			this.toolTip1.SetToolTip(this.cmbExcel, "Exportar EXCEL");
			this.cmbExcel.Click += new System.EventHandler(this.cmbExcel_Click);
			// 
			// grdComisiones
			// 
			this.grdComisiones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdComisiones.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdComisiones.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdComisiones.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.AutoEdit = false;
			appearance2.BackColor = System.Drawing.Color.Yellow;
			ultraGridColumn1.CellAppearance = appearance2;
			ultraGridColumn1.Header.Caption = "BODEGA";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 50;
			ultraGridColumn2.AutoEdit = false;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(192)), ((System.Byte)(0)));
			ultraGridColumn2.CellAppearance = appearance3;
			ultraGridColumn2.Header.Caption = "DIA";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 35;
			ultraGridColumn3.AutoEdit = false;
			ultraGridColumn3.Header.Caption = "# FC Easy";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 60;
			ultraGridColumn4.AutoEdit = false;
			ultraGridColumn4.Header.Caption = "# NC Easy";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 60;
			ultraGridColumn5.AutoEdit = false;
			ultraGridColumn5.Header.Caption = "# AN Easy";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 60;
			ultraGridColumn6.AutoEdit = false;
			ultraGridColumn6.Header.Caption = "St Easy";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.AutoEdit = false;
			ultraGridColumn7.Header.Caption = "Iva Easy";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 80;
			ultraGridColumn8.AutoEdit = false;
			ultraGridColumn8.Header.Caption = "Tot Easy";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 80;
			ultraGridColumn9.AutoEdit = false;
			ultraGridColumn9.Header.Caption = "# FC Lati";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 60;
			ultraGridColumn10.AutoEdit = false;
			ultraGridColumn10.Header.Caption = "# NC Lati";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 60;
			ultraGridColumn11.AutoEdit = false;
			ultraGridColumn11.Header.Caption = "# AN Lati";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 60;
			ultraGridColumn12.AutoEdit = false;
			ultraGridColumn12.Header.Caption = "St Lati";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.AutoEdit = false;
			ultraGridColumn13.Header.Caption = "Iva Lati";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 80;
			ultraGridColumn14.AutoEdit = false;
			ultraGridColumn14.Header.Caption = "Tot Lati";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 80;
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
																										 ultraGridColumn14});
			this.grdComisiones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdComisiones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.grdComisiones.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdComisiones.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdComisiones.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdComisiones.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.grdComisiones.Location = new System.Drawing.Point(8, 160);
			this.grdComisiones.Name = "grdComisiones";
			this.grdComisiones.Size = new System.Drawing.Size(1000, 336);
			this.grdComisiones.TabIndex = 11;
			this.grdComisiones.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.grdComisiones_InitializeRow);
			this.grdComisiones.DoubleClick += new System.EventHandler(this.grdComisiones_DoubleClick);
			this.grdComisiones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdComisiones_InitializeLayout);
			// 
			// ultraDataSource1
			// 
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
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.grdLocales);
			this.groupBox3.Controls.Add(this.chkLocls);
			this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox3.Location = new System.Drawing.Point(192, 8);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(296, 144);
			this.groupBox3.TabIndex = 170;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "  LOCALES COMERCIALES  ";
			// 
			// grdLocales
			// 
			this.grdLocales.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdLocales.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.Caption = "-";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn16.Width = 30;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Width = 227;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.grdLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.grdLocales.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdLocales.Location = new System.Drawing.Point(8, 25);
			this.grdLocales.Name = "grdLocales";
			this.grdLocales.Size = new System.Drawing.Size(280, 114);
			this.grdLocales.TabIndex = 170;
			// 
			// chkLocls
			// 
			this.chkLocls.Location = new System.Drawing.Point(216, 7);
			this.chkLocls.Name = "chkLocls";
			// 
			// chkLocls.Properties
			// 
			this.chkLocls.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.chkLocls.Properties.Appearance.Options.UseFont = true;
			this.chkLocls.Properties.Caption = "Todos";
			this.chkLocls.Size = new System.Drawing.Size(64, 18);
			this.chkLocls.TabIndex = 169;
			this.chkLocls.CheckedChanged += new System.EventHandler(this.chkLocls_CheckedChanged);
			// 
			// frmRepVenComparaLatEas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1016, 501);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.grdComisiones);
			this.Controls.Add(this.cmbExcel);
			this.Controls.Add(this.cmdBuscar);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRepVenComparaLatEas";
			this.Text = "Compara Facturas de EasyMax Vs. Latinium (frmRepVenComparaLatEas)";
			this.Load += new System.EventHandler(this.frmRepVenComparaLatEas_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdComisiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkLocls.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		C1.Data.C1DataRow drSeteoF;

		private void frmRepVenComparaLatEas_Load(object sender, System.EventArgs e)
		{
			this.txtFechaInicia.Value = DateTime.Today;
			this.txtFechaFinaliza.Value = DateTime.Today;

			string stExecL = string.Format("Exec Bodega_Reporte 0");
			grdLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExecL);

			// Por default todo seleccionado
			this.chkLocls.Checked = true;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr2 in grdLocales.Rows.All)
			{
				if (this.chkLocls.Checked) dr2.Cells["Selct"].Value = 1;
				else dr2.Cells["Selct"].Value = 0;
			}

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}

		private void cmdBuscar_Click(object sender, System.EventArgs e)
		{
			//MsgAyuda.Wait("Por favor, espere un momento...");
			Cursor.Current = Cursors.WaitCursor;
			
			DateTime dtDesde = (DateTime)this.txtFechaInicia.Value;
			DateTime dtHasta = (DateTime)this.txtFechaFinaliza.Value;

			string lc = "";
			int totColumna = 0;

			#region grlc
			//Buscar registros seleccionado en locales
			if (chkLocls.Checked == true)
				lc = "Todos";
			else
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdLocales.Rows.All)
				{
					if ((bool)dr.Cells["Selct"].Value)
					{
						lc = lc + dr.Cells["Codigo"].Value + "-";
						totColumna = 1;
					}
				}
				if (totColumna==1) 
				{
					lc = lc.Substring(0,lc.Length-1);
					totColumna = 0; 
				}
				else lc = "";
			}
			#endregion grlc
			
			this.grdComisiones.DataSource = ultraDataSource1;

			string stExec = string.Format("Exec Compara_EasyMax_Latinium '{0}','{1}','{2}'", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"),lc);
			grdComisiones.DisplayLayout.Bands[0].Reset();
			grdComisiones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);	
			//MsgAyuda.Wait();
			Cursor.Current = Cursors.Default;

		}

		private void cmbExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(this.grdComisiones, "ComparaEasyMaxLatinium.xls");
			/*if (DialogResult.Yes == MessageBox.Show("¿Desea Ver el Archivo Generado?", "Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))*/
		  System.Diagnostics.Process.Start(stDir + "ComparaEasyMaxLatinium.xls");
			return;

		}

		private void grdComisiones_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			/* Cambiar el color de una línea*/
			if ((((int)e.Row.Cells["CanFAC"].Value+(int)e.Row.Cells["Anul_Easy"].Value) != ((int)e.Row.Cells["FLat"].Value+(int)e.Row.Cells["Anu_Lat"].Value)) ||
				  ((int)e.Row.Cells["NC_Easy"].Value != (int)e.Row.Cells["NCLat"].Value) || ((int)e.Row.Cells["Anul_Easy"].Value != (int)e.Row.Cells["Anu_Lat"].Value))
			{
				e.Row.Appearance.BackColor = Color.Red;
				//e.Row.Appearance.BackColor2 = Color.White;
				e.Row.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			}

		}

		private void grdComisiones_DoubleClick(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			
			string CodBodega = this.grdComisiones.ActiveRow.Cells["CodDepPref"].Value.ToString();
			int diaFecha = (int)this.grdComisiones.ActiveRow.Cells["nDia"].Value;
			int mesFecha = (int)this.txtFechaInicia.Value.Month;
			int anoFecha = (int)this.txtFechaInicia.Value.Year;

			DateTime fechaBusk = DateTime.Parse(diaFecha.ToString()+"/"+mesFecha.ToString()+"/"+anoFecha.ToString());

			DiferenciasEasyMaxLatinium DifEL = new DiferenciasEasyMaxLatinium(CodBodega,fechaBusk);
			DifEL.ShowDialog();
			Cursor = Cursors.Default;
		}

		private void chkLocls_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdLocales.Rows.All)
			{
				if (this.chkLocls.Checked) dr.Cells["Selct"].Value = 1;
				else dr.Cells["Selct"].Value = 0;
			}

		}

		private void grdComisiones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
