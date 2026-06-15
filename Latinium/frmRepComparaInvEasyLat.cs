using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRepComparaInvEasyLat.
	/// </summary>
	public class frmRepComparaInvEasyLat : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox3;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdLocales;
		private DevExpress.XtraEditors.CheckEdit chkLocls;
		private System.Windows.Forms.Button cmbExcel;
		private System.Windows.Forms.Button cmdBuscar;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdDatLE;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.DateTimePicker txtFecha;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtpHora;
		private System.Windows.Forms.Label lblHora;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private DevExpress.XtraEditors.CheckEdit checkEdit2;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.ComponentModel.IContainer components;

		public frmRepComparaInvEasyLat()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Selct");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRepComparaInvEasyLat));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Selcn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodEasyMax");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantEasyMax");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantLat");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dif_E_L");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TomaInv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TS");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dif_Inv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobrante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Faltante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dif_Valorada");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Selcn");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodEasyMax");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantEasyMax");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantLat");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dif_E_L");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TomaInv");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dif_Inv");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobrante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Faltante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dif_Valorada");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.grdLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.chkLocls = new DevExpress.XtraEditors.CheckEdit();
			this.cmbExcel = new System.Windows.Forms.Button();
			this.cmdBuscar = new System.Windows.Forms.Button();
			this.grdDatLE = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.button1 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.txtFecha = new System.Windows.Forms.DateTimePicker();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.dtpHora = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.lblHora = new System.Windows.Forms.Label();
			this.checkEdit2 = new DevExpress.XtraEditors.CheckEdit();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkLocls.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdDatLE)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHora)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit2.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.grdLocales);
			this.groupBox3.Controls.Add(this.chkLocls);
			this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox3.Location = new System.Drawing.Point(8, 8);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(264, 432);
			this.groupBox3.TabIndex = 174;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "  LOCALES COMERCIALES  ";
			// 
			// grdLocales
			// 
			this.grdLocales.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdLocales.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "-";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn2.Width = 30;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 227;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.grdLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdLocales.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdLocales.Location = new System.Drawing.Point(8, 40);
			this.grdLocales.Name = "grdLocales";
			this.grdLocales.Size = new System.Drawing.Size(248, 384);
			this.grdLocales.TabIndex = 170;
			// 
			// chkLocls
			// 
			this.chkLocls.Location = new System.Drawing.Point(176, 16);
			this.chkLocls.Name = "chkLocls";
			// 
			// chkLocls.Properties
			// 
			this.chkLocls.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.chkLocls.Properties.Appearance.Options.UseFont = true;
			this.chkLocls.Properties.Caption = "Todos";
			this.chkLocls.Size = new System.Drawing.Size(72, 24);
			this.chkLocls.TabIndex = 169;
			this.chkLocls.CheckedChanged += new System.EventHandler(this.chkLocls_CheckedChanged);
			// 
			// cmbExcel
			// 
			this.cmbExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbExcel.Image = ((System.Drawing.Image)(resources.GetObject("cmbExcel.Image")));
			this.cmbExcel.Location = new System.Drawing.Point(64, 465);
			this.cmbExcel.Name = "cmbExcel";
			this.cmbExcel.Size = new System.Drawing.Size(48, 40);
			this.cmbExcel.TabIndex = 173;
			this.toolTip1.SetToolTip(this.cmbExcel, "Exporta información a EXCEL");
			this.cmbExcel.Click += new System.EventHandler(this.cmbExcel_Click);
			// 
			// cmdBuscar
			// 
			this.cmdBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmdBuscar.Image = ((System.Drawing.Image)(resources.GetObject("cmdBuscar.Image")));
			this.cmdBuscar.Location = new System.Drawing.Point(16, 465);
			this.cmdBuscar.Name = "cmdBuscar";
			this.cmdBuscar.Size = new System.Drawing.Size(48, 40);
			this.cmdBuscar.TabIndex = 172;
			this.toolTip1.SetToolTip(this.cmdBuscar, "Buscar información ");
			this.cmdBuscar.Click += new System.EventHandler(this.cmdBuscar_Click);
			// 
			// grdDatLE
			// 
			this.grdDatLE.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdDatLE.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdDatLE.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdDatLE.DisplayLayout.Appearance = appearance1;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn17.Header.VisiblePosition = 13;
			ultraGridColumn18.Header.VisiblePosition = 14;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.grdDatLE.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.grdDatLE.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdDatLE.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdDatLE.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdDatLE.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdDatLE.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdDatLE.Location = new System.Drawing.Point(280, 8);
			this.grdDatLE.Name = "grdDatLE";
			this.grdDatLE.Size = new System.Drawing.Size(920, 432);
			this.grdDatLE.TabIndex = 175;
			this.grdDatLE.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdDatLE_AfterCellUpdate);
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
																																 ultraDataColumn14,
																																 ultraDataColumn15});
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.Location = new System.Drawing.Point(904, 465);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(48, 40);
			this.button1.TabIndex = 176;
			this.toolTip1.SetToolTip(this.button1, "Realizar conteo de INVENTARIO");
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.Location = new System.Drawing.Point(952, 465);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(48, 40);
			this.button2.TabIndex = 194;
			this.toolTip1.SetToolTip(this.button2, "Cerrar Inventario");
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button3.Enabled = false;
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.Location = new System.Drawing.Point(600, 465);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(48, 40);
			this.button3.TabIndex = 240;
			this.toolTip1.SetToolTip(this.button3, "Realizar CRUCE de Productos");
			this.button3.Click += new System.EventHandler(this.button3_Click);
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
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(128, 475);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(120, 16);
			this.label1.TabIndex = 177;
			this.label1.Text = "Fecha Toma Inventario";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFecha
			// 
			this.txtFecha.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtFecha.CustomFormat = "";
			this.txtFecha.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFecha.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.txtFecha.Location = new System.Drawing.Point(272, 472);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtFecha.Size = new System.Drawing.Size(117, 26);
			this.txtFecha.TabIndex = 193;
			this.txtFecha.Value = new System.DateTime(2014, 6, 23, 0, 0, 0, 0);
			// 
			// dtpHora
			// 
			this.dtpHora.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpHora.Appearance = appearance6;
			this.dtpHora.DateTime = new System.DateTime(1753, 1, 1, 0, 0, 0, 0);
			this.dtpHora.DropDownButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Never;
			this.dtpHora.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dtpHora.FormatString = "HH:mm";
			this.dtpHora.Location = new System.Drawing.Point(504, 472);
			this.dtpHora.MaskInput = "{LOC}hh:mm";
			this.dtpHora.Name = "dtpHora";
			this.dtpHora.Size = new System.Drawing.Size(80, 28);
			this.dtpHora.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.dtpHora.TabIndex = 238;
			this.dtpHora.Value = null;
			// 
			// lblHora
			// 
			this.lblHora.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblHora.AutoSize = true;
			this.lblHora.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblHora.Location = new System.Drawing.Point(432, 475);
			this.lblHora.Name = "lblHora";
			this.lblHora.Size = new System.Drawing.Size(43, 22);
			this.lblHora.TabIndex = 237;
			this.lblHora.Text = "Hora";
			this.lblHora.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// checkEdit2
			// 
			this.checkEdit2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.checkEdit2.Location = new System.Drawing.Point(1016, 473);
			this.checkEdit2.Name = "checkEdit2";
			// 
			// checkEdit2.Properties
			// 
			this.checkEdit2.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.checkEdit2.Properties.Appearance.Options.UseFont = true;
			this.checkEdit2.Properties.Caption = "Todo / Diferencias";
			this.checkEdit2.Size = new System.Drawing.Size(176, 24);
			this.checkEdit2.TabIndex = 239;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(8, 448);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1192, 8);
			this.groupBox1.TabIndex = 241;
			this.groupBox1.TabStop = false;
			// 
			// frmRepComparaInvEasyLat
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1208, 510);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.checkEdit2);
			this.Controls.Add(this.dtpHora);
			this.Controls.Add(this.lblHora);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.grdDatLE);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.cmbExcel);
			this.Controls.Add(this.cmdBuscar);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRepComparaInvEasyLat";
			this.Text = "Conteo de Inventario";
			this.Load += new System.EventHandler(this.frmRepComparaInvEasyLat_Load);
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkLocls.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdDatLE)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHora)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit2.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		C1.Data.C1DataRow drSeteoF;
		int Cuenta = 0;
		int TotlRegistros = 0;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmRepComparaInvEasyLat_Load(object sender, System.EventArgs e)
		{
			/*int fAnio = DateTime.Today.Year;
			int fMes = DateTime.Today.Month;
			int fDia = DateTime.Today.Day;
			int fHora = DateTime.Now.Hour;
      int fMinutos = DateTime.Now.Minute;
			int fSegundos = DateTime.Now.Second;

			DateTime date1 = new DateTime(fAnio,fMes,fDia,fHora,fMinutos,fSegundos);
			this.txtFecha.Value = date1;*/

			this.txtFecha.Value = DateTime.Today;
			this.dtpHora.Value = DateTime.Now;

			string stExecL = string.Format("Exec Bodega_Reporte 0");
			grdLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExecL);

			// Por default todo seleccionado
			/*this.chkLocls.Checked = false;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr2 in grdLocales.Rows.All)
			{
				if (this.chkLocls.Checked) dr2.Cells["Selct"].Value = 1;
				else dr2.Cells["Selct"].Value = 0;
			}*/

		}

		private void cmdBuscar_Click(object sender, System.EventArgs e)
		{
			Cursor.Current = Cursors.WaitCursor;
			
			string lc = "";
			int totColumna = 0;
			int fBusk = 0;
			DateTime FeBusk = (DateTime)this.txtFecha.Value;

			#region grlc
			//Buscar registros seleccionado en locales
			if (chkLocls.Checked == true)
			{
				lc = "Todos";
				fBusk = 1;
			}
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
					fBusk = 0;
				}
				else 
				{
					lc = "";
					fBusk = 1;
				}
			}
			#endregion grlc
			
			//this.grdDatLE.DataSource = ultraDataSource1;

			if(lc.ToString().Equals(""))
				MessageBox.Show("Seleccione una BODEGA \n para realizar proceso de compraciones.?", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			else
			{
				string stExec = string.Format("Exec Compara_Inventarios_Easy_Lat {0},'{1}','{2}'", fBusk, lc, FeBusk.ToString("yyyyMMdd"));
				grdDatLE.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);	
			}
			//MsgAyuda.Wait();
			TotlRegistros = this.grdDatLE.Rows.Count;
			Cursor.Current = Cursors.Default;
		}

		private void chkLocls_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdLocales.Rows.All)
			{
				if (this.chkLocls.Checked) dr.Cells["Selct"].Value = 1;
				else dr.Cells["Selct"].Value = 0;
			}

		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			
			string laBodg = "";
			DateTime laFecha = (DateTime)this.txtFecha.Value;
			DateTime laHora = (DateTime)this.dtpHora.Value;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdLocales.Rows.All)
			{
				if ((bool)dr.Cells["Selct"].Value) laBodg = laBodg + dr.Cells["Codigo"].Value;
			}

			if (laBodg.ToString().Trim().Equals(""))
				MessageBox.Show("Seleccione una BODEGA \npara realizar el CONTEO de INVENTARIO.?", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			else
			{
				/* Verifica si inventario se encuentra cerrado o no */
				string TipoINV = "";
				DateTime FeBusk = (DateTime)this.txtFecha.Value;
				TipoINV = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select dbo.Estado_Toma_Inventario('{0}','{1}')",laBodg,FeBusk.ToString("yyyyMMdd")));
				if (TipoINV.Trim().Equals("ABIERTO") || TipoINV.Trim().Equals("null"))
				{
//					frmTomaInventario TomaInv = new frmTomaInventario(laBodg,laFecha,laHora,0);
//					TomaInv.ShowDialog();
				}
				else
				{
					if (DialogResult.No == MessageBox.Show("Fecha de CONTEO CERRADA\nDesea visualizar información de CONTEO.?", "Aviso...!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
						return;
					else
					{
//						frmTomaInventario TomaInv = new frmTomaInventario(laBodg,laFecha,laHora,1);
//						TomaInv.ShowDialog();
					}
				}
			}
			Cursor = Cursors.Default;
		}

		private void cmbExcel_Click(object sender, System.EventArgs e)
		{
			/* Exporta información a EXECL */
			/*string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(this.grdDatLE, "Difern_TomaInventario.xls");
			if (DialogResult.Yes == MessageBox.Show("¿Desea Ver Archivo Generado?", "Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			   System.Diagnostics.Process.Start(stDir + "TomaInventario.xls");*/
			if (this.grdDatLE.Rows.Count == 0)			{				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				return;			}
			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 			{				string stDir = MenuLatinium.stDirInicio + "\\"; 				ultraGridExcelExporter1.Export(this.grdDatLE, saveFileDialog1.FileName);				System.Diagnostics.Process.Start(saveFileDialog1.FileName);			}
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			/* Cerrar toma de inventario */
			Cursor.Current = Cursors.WaitCursor;
			
			string lc = "";
			int totColumna = 0;
			int fBusk = 0;

			DateTime FeBusk = (DateTime)this.txtFecha.Value;

			#region grlc
			//Buscar registros seleccionado en locales
			if (chkLocls.Checked == true)
			{
				lc = "Todos";
				fBusk = 1;
			}
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
					fBusk = 0;
				}
				else 
				{
					lc = "";
					fBusk = 1;
				}
			}
			#endregion grlc
			
			if(lc.ToString().Equals("")) MessageBox.Show("Seleccione una BODEGA \n Para CERRAR Toma Inventario.?", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			else
			{
				if (DialogResult.No == MessageBox.Show("ESTA SEGURO DE REALIZAR PROCESO DE CIERRE.?", "Aviso...!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) return;
				else
				{
					string stExec = string.Format("Exec Cerrar_Toma_Inventario '{0}','{1}'", lc, FeBusk.ToString("yyyyMMdd"));
					Funcion.EjecutaSQL(cdsSeteoF, stExec);
					MessageBox.Show("Proceso terminado...", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				}
			}
			
			Cursor.Current = Cursors.Default;
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			/* Procesar cruce de inventarios */
			int Sleccionados = 0;
			int idArt1 = 0;
			int idArt2 = 0;
			int i = 1;
			string laBodg = "";
			DateTime laFecha = (DateTime)this.txtFecha.Value;
			DateTime laHora = (DateTime)this.dtpHora.Value;
			string TipoINV = "";
			DateTime CrFeBusk = (DateTime)this.txtFecha.Value;

			/* Verifica si inventario se encuentra cerrado o no */
      foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdLocales.Rows.All)
			{
				if ((bool)dr.Cells["Selct"].Value) laBodg = laBodg + dr.Cells["Codigo"].Value;
			}

			if (laBodg.ToString().Trim().Equals(""))
			{
				MessageBox.Show("Seleccione una BODEGA \npara realizar este proceso...!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			else
			{
				DateTime FeBusk = (DateTime)this.txtFecha.Value;
				TipoINV = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select dbo.Estado_Toma_Inventario('{0}','{1}')",laBodg,FeBusk.ToString("yyyyMMdd")));
				if (TipoINV.Trim().Equals("CERRADO"))
				{
					MessageBox.Show("Fecha de CONTEO se encuentra CERRADA.! \nPor favor registrar fecha correcta para realizar proceso...!!! ", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdDatLE.Rows.All)
			{
				if ((bool)dr.Cells["Selcn"].Value) Sleccionados = Sleccionados+1;
			}

			if (Sleccionados>2 || Sleccionados<2)
			{
				MessageBox.Show("Seleccione 2 PRODUCTOS\nPara realizar proceso de cruce de mercadería.?", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			} 
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdDatLE.Rows.All)
			{
				if ((bool)dr.Cells["Selcn"].Value) 
				{
					if (i==1)	
						idArt1 = (int)dr.Cells["idArticulo"].Value;
					if (i==2)
						idArt2 = (int)dr.Cells["idArticulo"].Value;
					i = i+1;
				}
			}

			using (frmCruceMercaderia CrucInv = new frmCruceMercaderia(laBodg,laFecha,laHora,idArt1,idArt2))
			{				
				if (DialogResult.OK == CrucInv.ShowDialog())
				{
					this.cmdBuscar_Click(sender, e);
				}
			}

		}

		private void grdDatLE_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			//Cuenta = 0;			int NoCuenta = 0;			if(e.Cell.Column.ToString() == "Selcn")
			{
				Cuenta = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdDatLE.Rows.All)
				{
					if ((bool)dr.Cells["Selcn"].Value) Cuenta = Cuenta+1;
					else NoCuenta = NoCuenta+1;
				}

				if ((bool)e.Cell.Row.Cells["Selcn"].Value)
				{
					if(Cuenta>2)
					{
						MessageBox.Show("Solo se puede seleccionar 2 PRODUCTOS", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);						e.Cell.Row.Cells["Selcn"].Value = false;						if (Cuenta>0)	Cuenta = Cuenta-1;						return;			
					}
					if(TotlRegistros==NoCuenta) Cuenta=0;
				}
			}
		}

	}
}
