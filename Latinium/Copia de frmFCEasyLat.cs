using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFCEasyLat.
	/// </summary>
	public class frmFCEasyLat : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button cmdBuscar;
		private System.Windows.Forms.Button cmbExcel;
		private System.Windows.Forms.DateTimePicker txtFechaInicia;
		private System.Windows.Forms.DateTimePicker txtFechaFinaliza;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox txtNDocFC;
		private System.Windows.Forms.TextBox txtCedRuc;
		private System.Windows.Forms.CheckBox ckClteProvdr;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdFactras;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugrFcLatinium;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Button btnCopiar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.ComponentModel.IContainer components;

		public frmFCEasyLat()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFCEasyLat));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtFechaFinaliza = new System.Windows.Forms.DateTimePicker();
			this.label2 = new System.Windows.Forms.Label();
			this.txtFechaInicia = new System.Windows.Forms.DateTimePicker();
			this.cmdBuscar = new System.Windows.Forms.Button();
			this.cmbExcel = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btnCopiar = new System.Windows.Forms.Button();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtNDocFC = new System.Windows.Forms.TextBox();
			this.txtCedRuc = new System.Windows.Forms.TextBox();
			this.ckClteProvdr = new System.Windows.Forms.CheckBox();
			this.grdFactras = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ugrFcLatinium = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label5 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdFactras)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugrFcLatinium)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
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
			this.cmdBuscar.Location = new System.Drawing.Point(816, 24);
			this.cmdBuscar.Name = "cmdBuscar";
			this.cmdBuscar.Size = new System.Drawing.Size(48, 40);
			this.cmdBuscar.TabIndex = 2;
			this.toolTip1.SetToolTip(this.cmdBuscar, "Buscar INFORMACION");
			this.cmdBuscar.Click += new System.EventHandler(this.cmdBuscar_Click);
			// 
			// cmbExcel
			// 
			this.cmbExcel.Image = ((System.Drawing.Image)(resources.GetObject("cmbExcel.Image")));
			this.cmbExcel.Location = new System.Drawing.Point(864, 24);
			this.cmbExcel.Name = "cmbExcel";
			this.cmbExcel.Size = new System.Drawing.Size(48, 40);
			this.cmbExcel.TabIndex = 3;
			this.toolTip1.SetToolTip(this.cmbExcel, "Exportar EXCEL");
			this.cmbExcel.Click += new System.EventHandler(this.cmbExcel_Click);
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
			// btnCopiar
			// 
			this.btnCopiar.Enabled = false;
			this.btnCopiar.Image = ((System.Drawing.Image)(resources.GetObject("btnCopiar.Image")));
			this.btnCopiar.Location = new System.Drawing.Point(976, 24);
			this.btnCopiar.Name = "btnCopiar";
			this.btnCopiar.Size = new System.Drawing.Size(48, 40);
			this.btnCopiar.TabIndex = 179;
			this.toolTip1.SetToolTip(this.btnCopiar, "Registrar factura en Latinium");
			this.btnCopiar.Click += new System.EventHandler(this.btnCopiar_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(200, 16);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(73, 16);
			this.label3.TabIndex = 7;
			this.label3.Text = "NDocumento:";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(376, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(83, 16);
			this.label4.TabIndex = 162;
			this.label4.Text = "# Cedula / RUC";
			// 
			// txtNDocFC
			// 
			this.txtNDocFC.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNDocFC.Location = new System.Drawing.Point(200, 32);
			this.txtNDocFC.Name = "txtNDocFC";
			this.txtNDocFC.Size = new System.Drawing.Size(152, 20);
			this.txtNDocFC.TabIndex = 165;
			this.txtNDocFC.Text = "";
			// 
			// txtCedRuc
			// 
			this.txtCedRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCedRuc.Location = new System.Drawing.Point(376, 32);
			this.txtCedRuc.Name = "txtCedRuc";
			this.txtCedRuc.Size = new System.Drawing.Size(112, 20);
			this.txtCedRuc.TabIndex = 166;
			this.txtCedRuc.Text = "";
			// 
			// ckClteProvdr
			// 
			this.ckClteProvdr.Location = new System.Drawing.Point(496, 32);
			this.ckClteProvdr.Name = "ckClteProvdr";
			this.ckClteProvdr.Size = new System.Drawing.Size(80, 32);
			this.ckClteProvdr.TabIndex = 167;
			this.ckClteProvdr.Text = "Cliente / Proveedor";
			// 
			// grdFactras
			// 
			this.grdFactras.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdFactras.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdFactras.DisplayLayout.Appearance = appearance1;
			this.grdFactras.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdFactras.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdFactras.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdFactras.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdFactras.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdFactras.Location = new System.Drawing.Point(8, 96);
			this.grdFactras.Name = "grdFactras";
			this.grdFactras.Size = new System.Drawing.Size(1288, 224);
			this.grdFactras.TabIndex = 176;
			this.grdFactras.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.grdFactras_BeforeSelectChange);
			this.grdFactras.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdFactras_InitializeLayout_1);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(bool);
			ultraDataColumn4.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// ugrFcLatinium
			// 
			this.ugrFcLatinium.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugrFcLatinium.Cursor = System.Windows.Forms.Cursors.Default;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ugrFcLatinium.DisplayLayout.Appearance = appearance6;
			this.ugrFcLatinium.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ugrFcLatinium.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugrFcLatinium.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugrFcLatinium.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugrFcLatinium.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ugrFcLatinium.Location = new System.Drawing.Point(8, 368);
			this.ugrFcLatinium.Name = "ugrFcLatinium";
			this.ugrFcLatinium.Size = new System.Drawing.Size(1288, 144);
			this.ugrFcLatinium.TabIndex = 177;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(16, 336);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(212, 26);
			this.label5.TabIndex = 178;
			this.label5.Text = "L   a   t   i   n   i   u   m";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance11.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance11;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 457;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 10F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.ultraGrid1.Location = new System.Drawing.Point(1152, 216);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(112, 56);
			this.ultraGrid1.TabIndex = 180;
			this.ultraGrid1.Visible = false;
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// frmFCEasyLat
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1304, 533);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btnCopiar);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtCedRuc);
			this.Controls.Add(this.txtNDocFC);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.ugrFcLatinium);
			this.Controls.Add(this.grdFactras);
			this.Controls.Add(this.ckClteProvdr);
			this.Controls.Add(this.cmbExcel);
			this.Controls.Add(this.cmdBuscar);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmFCEasyLat";
			this.Text = "Compara Facturas de EasyMax Vs. Latinium (frmFCEasyLat.cs)";
			this.Load += new System.EventHandler(this.frmFCEasyLat_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdFactras)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugrFcLatinium)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		C1.Data.C1DataRow drSeteoF;

		private void frmFCEasyLat_Load(object sender, System.EventArgs e)
		{
			this.txtFechaInicia.Value = DateTime.Today;
			this.txtFechaFinaliza.Value = DateTime.Today;

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idCliente, Ruc, Nombre, Sel From Cliente Where Proveedor = 1 Order By Nombre"));
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}

		private void cmdBuscar_Click(object sender, System.EventArgs e)
		{
			//MsgAyuda.Wait("Por favor, espere un momento...");
			
			this.btnCopiar.Enabled = false;

			DateTime dtDesde = (DateTime)this.txtFechaInicia.Value;
			DateTime dtHasta = (DateTime)this.txtFechaFinaliza.Value;

			string NDocFC = this.txtNDocFC.Text;
			string nCedRuc = this.txtCedRuc.Text;
			int tpCuenta = 0;
			if(this.ckClteProvdr.Checked == true) tpCuenta = 1;
			else tpCuenta = 0;

			//this.grdFactras.DataSource = ultraDataSource1;

			string stExec = string.Format("Exec FC_EasyMAX '{0}','{1}','{2}','{3}',{4}", 
				     dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"),NDocFC,nCedRuc,tpCuenta);
			grdFactras.DisplayLayout.Bands[0].Reset();
			grdFactras.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);	

		}

		private void cmbExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(this.grdFactras, "ComparaEasyMaxLatinium.xls");
			/*if (DialogResult.Yes == MessageBox.Show("¿Desea Ver el Archivo Generado?", "Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))*/
		  System.Diagnostics.Process.Start(stDir + "ComparaEasyMaxLatinium.xls");
			return;

		}

		private void grdFactras_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
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

		private void grdFactras_DoubleClick(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			
			string CodBodega = this.grdFactras.ActiveRow.Cells["CodDepPref"].Value.ToString();
			int diaFecha = (int)this.grdFactras.ActiveRow.Cells["nDia"].Value;
			int mesFecha = (int)this.txtFechaInicia.Value.Month;
			int anoFecha = (int)this.txtFechaInicia.Value.Year;

			DateTime fechaBusk = DateTime.Parse(diaFecha.ToString()+"/"+mesFecha.ToString()+"/"+anoFecha.ToString());

			DiferenciasEasyMaxLatinium DifEL = new DiferenciasEasyMaxLatinium(CodBodega,fechaBusk);
			DifEL.ShowDialog();
			Cursor = Cursors.Default;
		}

		private void grdFactras_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void grdFactras_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			string nFcLat = this.txtNDocFC.Text;
			int tpDoc = 0;
			int nReg = 0;

			if(this.ckClteProvdr.Checked == true) tpDoc = 1;
			else tpDoc = 0;

			if (grdFactras.ActiveRow == null) return;
			if (this.grdFactras.ActiveRow.Cells["Numero"].Value == DBNull.Value) return;

			string sSQL = string.Format("Exec Busca_FC_Latinium '{0}',{1}", nFcLat,tpDoc);
			this.ugrFcLatinium.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ugrFcLatinium.Rows.All)
			{
					nReg = nReg+1;
			}

			if(nReg==0)
				this.btnCopiar.Enabled = true;

		}

		private void btnCopiar_Click(object sender, System.EventArgs e)
		{
			string nFcEas = this.txtNDocFC.Text;
			int tpCP = 0;

			if(this.ckClteProvdr.Checked == true) tpCP = 1;
			else tpCP = 0;

			if (DialogResult.No == MessageBox.Show("ESTA SEGURO DE COPIAR.?", "Aviso...!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				return;
			else
			{
				string stExec = string.Format("Exec CopiaFC_EasyMAX_Latinium '{0}',{1}", nFcEas, tpCP);
				Funcion.EjecutaSQL(cdsSeteoF, stExec);
				MessageBox.Show("Proceso terminado...", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}

			this.btnCopiar.Enabled = false;
		}

		private void grdFactras_InitializeLayout_1(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idCliente"].Value != System.DBNull.Value)
			{
				if (e.Cell.Column.ToString() == "Sel")
				{
					this.ultraGrid1.UpdateData();
					
					string sSQL = string.Format("Exec ClienteActualizaSel {0}, {1}", (int)e.Cell.Row.Cells["idCliente"].Value, (bool)e.Cell.Row.Cells["Sel"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);		
				}
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
