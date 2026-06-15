using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmActualizacionDePrecios.
	/// </summary>
	public class frmActualizacionDePrecios : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.RadioGroup rdgProductoFacturable;
		private DevExpress.XtraEditors.RadioGroup rdgProductoExistencia;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnActualizar;
		private System.Windows.Forms.Button btnProcesar;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		bool bDescuento = false;
		bool bDescontinuado = false;
		decimal dPorcTarjeta = 0.00m;
		private System.Windows.Forms.Label lblContador;
		decimal dPorcCredito = 0.00m;
		bool bActiva = true;
		bool bActualiza = true;

		public frmActualizacionDePrecios()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubgrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubgrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmActualizacionDePrecios));
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.rdgProductoFacturable = new DevExpress.XtraEditors.RadioGroup();
			this.rdgProductoExistencia = new DevExpress.XtraEditors.RadioGroup();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label12 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.btnProcesar = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoFacturable.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoExistencia.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
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
			// txtBuscar
			// 
			this.txtBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance1;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(544, 424);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(250, 22);
			this.txtBuscar.TabIndex = 0;
			this.txtBuscar.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscar_KeyPress);
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(480, 427);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 158;
			this.label1.Text = "Buscar";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// rdgProductoFacturable
			// 
			this.rdgProductoFacturable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.rdgProductoFacturable.EditValue = 1;
			this.rdgProductoFacturable.Location = new System.Drawing.Point(8, 423);
			this.rdgProductoFacturable.Name = "rdgProductoFacturable";
			// 
			// rdgProductoFacturable.Properties
			// 
			this.rdgProductoFacturable.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgProductoFacturable.Properties.Appearance.Options.UseBackColor = true;
			this.rdgProductoFacturable.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos los productos"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "No Descontinuados"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Descontinuados"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Combos")});
			this.rdgProductoFacturable.Size = new System.Drawing.Size(464, 24);
			this.rdgProductoFacturable.TabIndex = 3;
			this.rdgProductoFacturable.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.rdgProductoFacturable_KeyPress);
			this.rdgProductoFacturable.SelectedIndexChanged += new System.EventHandler(this.rdgProductoFacturable_SelectedIndexChanged);
			// 
			// rdgProductoExistencia
			// 
			this.rdgProductoExistencia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.rdgProductoExistencia.EditValue = 1;
			this.rdgProductoExistencia.Location = new System.Drawing.Point(8, 392);
			this.rdgProductoExistencia.Name = "rdgProductoExistencia";
			// 
			// rdgProductoExistencia.Properties
			// 
			this.rdgProductoExistencia.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgProductoExistencia.Properties.Appearance.Options.UseBackColor = true;
			this.rdgProductoExistencia.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos los productos"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Con existencia"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Sin existencia"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "En remates")});
			this.rdgProductoExistencia.Size = new System.Drawing.Size(464, 24);
			this.rdgProductoExistencia.TabIndex = 2;
			this.rdgProductoExistencia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.rdgProductoExistencia_KeyPress);
			this.rdgProductoExistencia.SelectedIndexChanged += new System.EventHandler(this.rdgProductoExistencia_SelectedIndexChanged);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance3;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.RowAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(0, 0);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1258, 322);
			this.ultraGrid1.TabIndex = 23;
			this.ultraGrid1.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellListCloseUp += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellListCloseUp);
			this.ultraGrid1.InitializeTemplateAddRow += new Infragistics.Win.UltraWinGrid.InitializeTemplateAddRowEventHandler(this.ultraGrid1_InitializeTemplateAddRow);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// cmbSubGrupo
			// 
			this.cmbSubGrupo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance10;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Subgrupo";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSubGrupo.DisplayMember = "Subgrupo";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(544, 400);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(250, 22);
			this.cmbSubGrupo.TabIndex = 4;
			this.cmbSubGrupo.ValueMember = "idSubgrupo";
			this.cmbSubGrupo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbSubGrupo_KeyPress);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(480, 402);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(53, 16);
			this.label12.TabIndex = 189;
			this.label12.Text = "SubGrupo";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 384);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1258, 8);
			this.groupBox2.TabIndex = 193;
			this.groupBox2.TabStop = false;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			appearance11.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance11;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand3.Header.Appearance = appearance12;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance13;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance15;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.ultraGrid2.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(0, 328);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(1258, 58);
			this.ultraGrid2.TabIndex = 199;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1082, 424);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(80, 23);
			this.btnExportar.TabIndex = 202;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnVer
			// 
			this.btnVer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(824, 424);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 1;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1170, 424);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 203;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnActualizar
			// 
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnActualizar.Location = new System.Drawing.Point(994, 424);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(80, 23);
			this.btnActualizar.TabIndex = 629;
			this.btnActualizar.Text = "&Actualizar";
			this.btnActualizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// btnProcesar
			// 
			this.btnProcesar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnProcesar.CausesValidation = false;
			this.btnProcesar.Image = ((System.Drawing.Image)(resources.GetObject("btnProcesar.Image")));
			this.btnProcesar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnProcesar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnProcesar.Location = new System.Drawing.Point(906, 424);
			this.btnProcesar.Name = "btnProcesar";
			this.btnProcesar.Size = new System.Drawing.Size(80, 23);
			this.btnProcesar.TabIndex = 630;
			this.btnProcesar.Text = "&Procesar";
			this.btnProcesar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(826, 400);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 631;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmActualizacionDePrecios
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1258, 456);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnProcesar);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.rdgProductoFacturable);
			this.Controls.Add(this.rdgProductoExistencia);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.ultraGrid1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmActualizacionDePrecios";
			this.Text = "Actualización de precios";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmActualizacionDePrecios_KeyDown);
			this.Load += new System.EventHandler(this.frmActualizacionDePrecios_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoFacturable.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoExistencia.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ListaArticulosPrecios()
		{
			this.Cursor = Cursors.WaitCursor;
			
			int idSubGrupo = 0;		 
			if (this.cmbSubGrupo.ActiveRow != null) idSubGrupo = (int)this.cmbSubGrupo.Value;
			
			this.ultraGrid1.DataSource = null;
			this.ultraGrid2.DataSource = null;

			string sSQL = string.Format("Exec ConsultaActualizacionPrecios '{0}', {1}, {2}, {3}, {4}", 
				this.txtBuscar.Text.ToString().Trim(), idSubGrupo, (int)rdgProductoExistencia.EditValue, (int)rdgProductoFacturable.EditValue, false);
			this.ultraGrid1.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
			
			DiseñoGrid();

			this.lblContador.Text = this.ultraGrid1.Rows.Count + " ARTICULOS ENCONTRADOS";

			this.Cursor = Cursors.Default;
		}
		
		private void ActualizaMovimientos()
		{
			DateTime dtDesde = DateTime.Now.AddDays(-15);
			DateTime dtHasta = DateTime.Now;

			string sSQL = string.Format("Exec CreaListaArticulosActualizacionDePrecios '{0}', '{1}'", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			
			bDescuento = true;
			bDescontinuado = true;			
		}

		private void DiseñoGrid()
		{
			if (this.ultraGrid1.Rows.Count > 0)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idActualizacion"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Sel"].Width = 20;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Discontinuo"].Width = 20;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Catalogo"].Width = 20;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Obsequio"].Width = 20;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["ProductoWeb"].Width = 20;
				
				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Grupo"].Width = 170;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["SubGrupo"].Width = 170;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].Width = 170;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Marca"].Width = 170;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Modelo"].Width = 170;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Articulo"].Width = 220;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Reemplazo"].Width = 170;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idSubGrupo"].Hidden = true;					

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["O_C"].Width = 50;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Payjoy"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPJ"].Width = 60;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].Width = 70;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CostoUltimo"].Width = 75;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Distribuidor"].Width = 75;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Tarjeta"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Credito"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["WEB"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["TCSI"].Width = 60;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC3"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC6"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC9"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC12"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC15"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC18"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC21"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC24"].Width = 60;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCred"].Width = 40;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargDist"].Width = 40;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCont"].Width = 40;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargTarj"].Width = 40;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargWEB"].Width = 40;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MARGTCSI"].Width = 40;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC3"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC6"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC9"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC12"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC15"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC18"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC21"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC24"].Width = 50;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPVP"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["INCRTCSI"].Width = 50;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC3"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC6"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC9"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC12"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC15"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC18"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC21"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC24"].Width = 50;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Sel"].Header.Caption = "...";				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Discontinuo"].Header.Caption = "D";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Catalogo"].Header.Caption = "C";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Obsequio"].Header.Caption = "Obs";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["ProductoWeb"].Header.Caption = "Web";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CostoUltimo"].Header.Caption = "Costo";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargDist"].Header.Caption = "%";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCont"].Header.Caption = "%";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargTarj"].Header.Caption = "%";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargWEB"].Header.Caption = "%";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["INCRTCSI"].Header.Caption = "INCRTCSI";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["TCSI"].Header.Caption = " TCSI";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MARGTCSI"].Header.Caption = "%";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCred"].Header.Caption = "%";

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Payjoy"].Header.Caption = "PayJoy";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPJ"].Header.Caption = "% IncrPJ";


				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Grupo"].CellActivation = Activation.ActivateOnly;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["SubGrupo"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Marca"].CellActivation = Activation.ActivateOnly;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Modelo"].CellActivation = Activation.ActivateOnly;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.ActivateOnly;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CostoUltimo"].CellActivation = Activation.AllowEdit;	
				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Reemplazo"].CellActivation = Activation.ActivateOnly;				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargDist"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCont"].CellActivation = Activation.Disabled;				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargTarj"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Credito"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCred"].CellActivation = Activation.Disabled;
			//	this.ultraGrid1.DisplayLayout.Bands[0].Columns["WEB"].CellActivation = Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargWEB"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["INCRTCSI"].CellActivation = Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["O_C"].CellActivation = Activation.Disabled;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Payjoy"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPJ"].CellActivation = Activation.AllowEdit;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MARGTCSI"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["TCSI"].CellActivation = Activation.Disabled;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC3"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC6"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC9"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC12"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC15"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC18"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC21"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC24"].CellActivation = Activation.Disabled;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC3"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC6"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC9"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC12"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC15"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC18"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC21"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC24"].CellActivation = Activation.Disabled;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CostoUltimo"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Distribuidor"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargDist"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Contado"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCont"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Tarjeta"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPVP"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargTarj"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Credito"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCred"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				//this.ultraGrid1.DisplayLayout.Bands[0].Columns["WEB"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["TCSI"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["INCRTCSI"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargWEB"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["O_C"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Payjoy"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPJ"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				
					
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MARGTCSI"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC3"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC6"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC9"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC12"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC15"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC18"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC21"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC24"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC3"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC6"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC9"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC12"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC15"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC18"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC21"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["PrecioC24"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC3"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC6"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC9"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC12"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC15"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC18"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC21"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargPrecioC24"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Distribuidor"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargDist"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Contado"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCont"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["WEB"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargWEB"].Hidden = true;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CostoUltimo"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Distribuidor"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargDist"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Contado"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCont"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Tarjeta"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPVP"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargTarj"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Credito"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCred"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["WEB"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargWEB"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["INCRTCSI"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["TCSI"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MARGTCSI"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Payjoy"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPJ"].Format = "#,##0.00";

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC3"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC6"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC9"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC12"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC15"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC18"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC21"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPrecioC24"].Format = "#,##0.00";

				int iContColumnas = (int)this.ultraGrid1.DisplayLayout.Bands[0].Columns.Count;

				int iContador = 27 + 24 + 6;

				string nombrCol = "";

				int iRegion = 0 ; 

				for (int j = iContador; j < iContColumnas; j++)
				{
					nombrCol = this.ultraGrid1.DisplayLayout.Bands[0].Columns[j].Header.Caption;
					iRegion =  Funcion.iEscalarSQL(cdsSeteoF, string.Format("If  Exists(Select CODIGO From Bodega Where Codigo = '{0}') begin Select region from Bodega where Codigo = '{0}' end Else begin  Select CONVERT(Int , 0) end ",nombrCol));		
					
					if (iRegion == 1) 
					{
						this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LightBlue;
					}
					if (iRegion == 2) 
					{
						this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Yellow;
					}

					if (nombrCol == "BI") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
					if (nombrCol == "BG") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
					if (nombrCol == "BM") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;

					this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.Disabled;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].Width = 30;					

					this.ultraGrid1.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.ultraGrid1.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.ultraGrid1.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
				}
			}
		}

		private void DiseñoGridTransito()
		{
			if (this.ultraGrid2.Rows.Count > 0)
			{
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["Codigo"].Width = 210;
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["Articulo"].Width = 770;
				
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.ActivateOnly;
				
				int iContColumnas = (int)this.ultraGrid2.DisplayLayout.Bands[0].Columns.Count;

				int iContador = 3;
				string nombrCol = "";
				int iRegion = 0 ; 

				for (int j = iContador; j < iContColumnas; j++)
				{	
	
					nombrCol = this.ultraGrid2.DisplayLayout.Bands[0].Columns[j].Header.Caption;

					iRegion =  Funcion.iEscalarSQL(cdsSeteoF, string.Format("If  Exists(Select CODIGO From Bodega Where Codigo = '{0}') begin Select region from Bodega where Codigo = '{0}' end Else begin  Select CONVERT(Int , 0) end ",nombrCol));		
					if ( iRegion == 1) 
					{
						this.ultraGrid2.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LightBlue;
					}
					if ( iRegion == 2) 
					{
						this.ultraGrid2.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Yellow;
					}
					if (nombrCol == "BI") this.ultraGrid2.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
					if (nombrCol == "BG") this.ultraGrid2.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;		
					if (nombrCol == "BM") this.ultraGrid2.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;	

					this.ultraGrid2.DisplayLayout.Bands[0].Columns[j].Width = 30;
					this.ultraGrid2.DisplayLayout.Bands[0].Columns[j].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.ultraGrid2.DisplayLayout.Bands[0].Columns[j].CellActivation = Activation.Disabled;
				}
			}
		}	

		private void frmActualizacionDePrecios_Load(object sender, System.EventArgs e)
		{	
			Funcion.Resolucion(this, this.Height, this.Width);
	  			
			this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSubgrupo, SubGrupo From ArticuloSubGrupo Order By SubGrupo");
			
			ActualizaMovimientos();
			ListaArticulosPrecios();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
		
		private void txtBuscar_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnVer_Click(sender, e);
		}
		
		private void rdgProductoExistencia_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			ListaArticulosPrecios();
		}

		private void rdgProductoFacturable_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			ListaArticulosPrecios();
		}
		
		private void frmActualizacionDePrecios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();			
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}
	
		private void Actualiza(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sSQLSel = string.Format("Exec ActualizaActPrecios {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9},{10}, {11}", 
				(bool) e.Cell.Row.Cells["Sel"].Value, (int) e.Cell.Row.Cells["idActualizacion"].Value, 
				(decimal) e.Cell.Row.Cells["Distribuidor"].Value, (decimal) e.Cell.Row.Cells["MargDist"].Value,
				(decimal) e.Cell.Row.Cells["Contado"].Value, (decimal) e.Cell.Row.Cells["MargCont"].Value,
				(decimal) e.Cell.Row.Cells["Tarjeta"].Value, (decimal) e.Cell.Row.Cells["MargTarj"].Value,
				(decimal) e.Cell.Row.Cells["Credito"].Value, (decimal) e.Cell.Row.Cells["MargCred"].Value,
				(decimal) e.Cell.Row.Cells["WEB"].Value, (decimal) e.Cell.Row.Cells["MargWEB"].Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLSel, true);				
		}
		
		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{	
			bActualiza = true;
			if (e.Cell.Column.ToString() == "Sel")
			{	
				if (bActualiza)	Actualiza(e);				
			}
			if (e.Cell.Column.ToString() == "Catalogo")
			{	
				if (bActualiza)	Actualiza(e);				
			}
			if (e.Cell.Column.ToString() == "Discontinuo")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDsctoDsc {0}, {1}, 3", (bool)e.Cell.Row.Cells["Discontinuo"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
			}
			if (e.Cell.Column.ToString() == "Obsequio")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDsctoDsc {0}, {1}, 5", (bool)e.Cell.Row.Cells["Obsequio"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
			}
			if (e.Cell.Column.ToString() == "ProductoWeb")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDsctoDsc {0}, {1}, 6", (bool)e.Cell.Row.Cells["ProductoWeb"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
			}

			if (e.Cell.Column.ToString() == "DSCTO")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDsctoDsc {0}, {1}, 1", (bool)e.Cell.Row.Cells["DSCTO"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
			}

			if (e.Cell.Column.ToString() == "Dsc")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDsctoDsc {0}, {1}, 2", (bool)e.Cell.Row.Cells["Dsc"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
			}

			if (e.Cell.Column.ToString() == "Catalogo")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDsctoDsc {0}, {1}, 4", (bool)e.Cell.Row.Cells["Catalogo"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
			}

			if (e.Cell.Column.ToString() == "INCRTCSI")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionDePrecios Set INCRTCSI = {0} Where idActualizacion = {1} Update Articulo Set Incr6 = {0} Where idArticulo = {2}", 
					(decimal)e.Cell.Row.Cells["INCRTCSI"].Value, (int)e.Cell.Row.Cells["idActualizacion"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["TCSI"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["INCRTCSI"].Value), 2);

				e.Cell.Row.Cells["MargTarj"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value)) * 100.00m);
				e.Cell.Row.Cells["MargCred"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value)) * 100.00m);
				//e.Cell.Row.Cells["MargWEB"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["WEB"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["WEB"].Value)) * 100.00m);
				e.Cell.Row.Cells["MARGTCSI"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value)) * 100.00m);
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionPrecios Set Sel = 1, Tarjeta = {0}, Credito = {1}, MargTarj = {3}, MargCred = {4}, WEB= {5}, MargWEB = {6}, MargTCSI = {7}, TCSI = {8} Where idArticulo = {2}", 
					(decimal)e.Cell.Row.Cells["Tarjeta"].Value, (decimal)e.Cell.Row.Cells["Credito"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value, 
					(decimal)e.Cell.Row.Cells["MargTarj"].Value, (decimal)e.Cell.Row.Cells["MargCred"].Value, (decimal)e.Cell.Row.Cells["WEB"].Value, (decimal)e.Cell.Row.Cells["MargWEB"].Value, (decimal)e.Cell.Row.Cells["MARGTCSI"].Value, (decimal)e.Cell.Row.Cells["TCSI"].Value  ));
				e.Cell.Row.Cells["Sel"].Value = true;
			}

			if (e.Cell.Column.ToString() == "IncrPVP")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionDePrecios Set IncrPVP = {0} Where idActualizacion = {1} Update Articulo Set Incr4 = {0} Where idArticulo = {2}", 
					(decimal)e.Cell.Row.Cells["IncrPVP"].Value, (int)e.Cell.Row.Cells["idActualizacion"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["Credito"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["IncrPVP"].Value), 2);
				e.Cell.Row.Cells["TCSI"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["INCRTCSI"].Value), 2);

				e.Cell.Row.Cells["MargTarj"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value)) * 100.00m);
				e.Cell.Row.Cells["MargCred"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value)) * 100.00m);
				//e.Cell.Row.Cells["MargWEB"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["WEB"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["WEB"].Value)) * 100.00m);
				e.Cell.Row.Cells["MARGTCSI"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value)) * 100.00m);
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionPrecios Set Sel = 1, Tarjeta = {0}, Credito = {1}, MargTarj = {3}, MargCred = {4}, WEB= {5}, MargWEB = {6}, MargTCSI = {7}, TCSI = {8} Where idArticulo = {2}", 
					(decimal)e.Cell.Row.Cells["Tarjeta"].Value, (decimal)e.Cell.Row.Cells["Credito"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value, 
					(decimal)e.Cell.Row.Cells["MargTarj"].Value, (decimal)e.Cell.Row.Cells["MargCred"].Value, (decimal)e.Cell.Row.Cells["WEB"].Value, (decimal)e.Cell.Row.Cells["MargWEB"].Value, (decimal)e.Cell.Row.Cells["MARGTCSI"].Value, (decimal)e.Cell.Row.Cells["TCSI"].Value));
				e.Cell.Row.Cells["Sel"].Value = true;
			}

			if (e.Cell.Column.ToString() == "IncrPrecioC3")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionDePrecios Set IncrPrecioC3 = {0} Where idActualizacion = {1} Update Articulo Set IncrC3 = {0} Where idArticulo = {2}", 
					(decimal)e.Cell.Row.Cells["IncrPrecioC3"].Value, (int)e.Cell.Row.Cells["idActualizacion"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["PrecioC3"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["IncrPrecioC3"].Value), 2);

				e.Cell.Row.Cells["MargPrecioC3"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["PrecioC3"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["PrecioC3"].Value)) * 100.00m);
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionPrecios Set Sel = 1, PrecioC3 = {0}, MargPrecioC3 = {1} Where idArticulo = {2}", 
					(decimal)e.Cell.Row.Cells["PrecioC3"].Value, (decimal)e.Cell.Row.Cells["MargPrecioC3"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
				e.Cell.Row.Cells["Sel"].Value = true;
			}			

			if (e.Cell.Column.ToString() == "CostoUltimo")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Articulo Set CostoUltimo = {0} Where idArticulo = {1}", (decimal)e.Cell.Row.Cells["CostoUltimo"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				//				e.Cell.Row.Cells["Sel"].Value = true;
			}

			if (e.Cell.Column.ToString() == "Distribuidor")
			{							
//				e.Cell.Row.Cells["Sel"].Value = false;

				e.Cell.Row.Cells["MargDist"].Value = (((Convert.ToDecimal(e.Cell.Row.Cells["Distribuidor"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Distribuidor"].Value)) * 100.00m);	

//				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionPrecios Set DisTribuidor = {0} Where idArticulo = {1}", Convert.ToDecimal(e.Cell.Row.Cells["Distribuidor"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["Sel"].Value = true;
			}

			if (e.Cell.Column.ToString() == "Contado")
			{
				//e.Cell.Row.Cells["Sel"].Value = false;

				e.Cell.Row.Cells["MargCont"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Contado"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Contado"].Value)) * 100.00m);

				e.Cell.Row.Cells["Tarjeta"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Contado"].Value) * dPorcTarjeta, 2);// + (Convert.ToDecimal(e.Cell.Row.Cells["Contado"].Value * dPorcTarjeta)/100;
				e.Cell.Row.Cells["MargTarj"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value)) * 100.00m);

				if ((int)e.Cell.Row.Cells["idSubGrupo"].Value == 131 || (int)e.Cell.Row.Cells["idSubGrupo"].Value == 66) dPorcCredito = 1.27m; else dPorcCredito = 1.37m;
				
				e.Cell.Row.Cells["Credito"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * dPorcCredito, 2);
				e.Cell.Row.Cells["MargCred"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value)) * 100.00m);

//				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionPrecios Set Contado = {0}, Tarjeta = {1}, Credito = {2} Where idArticulo = {3}", 
//					(decimal)e.Cell.Row.Cells["Contado"].Value, (decimal)e.Cell.Row.Cells["Tarjeta"].Value, (decimal)e.Cell.Row.Cells["Credito"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["Sel"].Value = true;
			}

			if (e.Cell.Column.ToString() == "Tarjeta")
			{
				e.Cell.Row.Cells["Credito"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["IncrPVP"].Value), 2);
        e.Cell.Row.Cells["TCSI"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["INCRTCSI"].Value), 2);

				e.Cell.Row.Cells["MargTarj"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value)) * 100.00m);
				e.Cell.Row.Cells["MargCred"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value)) * 100.00m);
				e.Cell.Row.Cells["MARGTCSI"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value)) * 100.00m);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionDePrecios Set Sel = 1, Tarjeta = {0}, Credito = {1}, MargTarj = {3}, MargCred = {4}, MARGTCSI = {5} Where idArticulo = {2}", 
					(decimal)e.Cell.Row.Cells["Tarjeta"].Value, (decimal)e.Cell.Row.Cells["Credito"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value, (decimal)e.Cell.Row.Cells["MargTarj"].Value, (decimal)e.Cell.Row.Cells["MargCred"].Value, (decimal)e.Cell.Row.Cells["MARGTCSI"].Value));

				e.Cell.Row.Cells["Sel"].Value = true;
			}
			if (e.Cell.Column.ToString() == "WEB")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Articulo Set Precio5 = {0} Where idArticulo = {1}", (decimal)e.Cell.Row.Cells["WEB"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
				e.Cell.Row.Cells["MargWEB"].Value = (((Convert.ToDecimal(e.Cell.Row.Cells["WEB"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["WEB"].Value)) * 100.00m);	
				//Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionDePrecios Set Sel = 1, MargWEB = {0} Where idArticulo = {1}", (decimal)e.Cell.Row.Cells["MargWEB"].Value,(int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["Sel"].Value = true;

			}

			if (e.Cell.Column.ToString() == "TCSI")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Articulo Set Precio6 = {0} Where idArticulo = {1}", (decimal)e.Cell.Row.Cells["TCSI"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["MARGTCSI"].Value = (((Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value)) * 100.00m);	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionDePrecios Set Sel = 1, MARGTCSI = {0} Where idArticulo = {1}", 
					(decimal)e.Cell.Row.Cells["MARGTCSI"].Value,(int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["Sel"].Value = true;
			}

			if (e.Cell.Column.ToString() == "IncrPJ")
			{
				decimal incrPJ = Convert.ToDecimal(e.Cell.Row.Cells["IncrPJ"].Value);
				decimal costoUltimo = Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value);

				// Nuevo cálculo: divide en lugar de multiplicar
				decimal costoConIncremento = costoUltimo / (1 - (incrPJ / 100));

				Funcion.EjecutaSQL(cdsSeteoF, string.Format( "UPDATE Articulo SET IncrPrecioC25 = {0}, Precio7 = {2} WHERE idArticulo = {1}",
					incrPJ, (int)e.Cell.Row.Cells["idArticulo"].Value, costoConIncremento));

				Funcion.EjecutaSQL(cdsSeteoF, string.Format(
					"UPDATE ActualizacionDePrecios SET Sel = 1, PrecioC25 = {0}, IncrPrecioC25 = {1} WHERE idArticulo = {2}",
					costoConIncremento, incrPJ, (int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["Payjoy"].Value = costoConIncremento;


				//e.Cell.Row.Cells["Sel"].Value = true;


			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			#region Decimales
			FuncionesProcedimientos.FormatoGrid(e, "CostoUltimo", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Distribuidor", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MargDist", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Contado", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MargCont", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Tarjeta", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MargTarj", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Credito", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MargCred", 2);	
			FuncionesProcedimientos.FormatoGrid(e, "WEB", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MargWEB", 2);	
			FuncionesProcedimientos.FormatoGrid(e, "MargCred", 2);	
			FuncionesProcedimientos.FormatoGrid(e, "TCSI", 2);
			FuncionesProcedimientos.FormatoGrid(e, "INCRTCSI", 2);	
			FuncionesProcedimientos.FormatoGrid(e, "MARGTCSI", 2);	
			FuncionesProcedimientos.FormatoGrid(e, "Payjoy", 2);	
			FuncionesProcedimientos.FormatoGrid(e, "IncrPJ", 2);
			#endregion Decimales
		}
				
		private void rdgProductoExistencia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.rdgProductoFacturable.Focus();
		}

		private void rdgProductoFacturable_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtBuscar.Focus();
		}

		private void cmbSubGrupo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtBuscar.Focus();
		}
				
		private void MercaderiaEnTransito (int idArticulo)
		{
			DateTime dtFecha = DateTime.Today;
			string sSQL = string.Format("Exec MercaderiaEnTransitoAPSpiff {0}, '{1}'", idArticulo, dtFecha.ToString("yyyyMMdd"));
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			DiseñoGridTransito();
		}

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (this.ultraGrid1.Rows.Count > 0)
				if (this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value) 
					MercaderiaEnTransito((int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value);				
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			ListaArticulosPrecios();
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			ActualizaMovimientos();	

			ListaArticulosPrecios();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnProcesar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No existen artículos para procesar la actualización de precios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}
			
			int iSel = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				if ((bool)dr.Cells["Sel"].Value) iSel++;			

			if (iSel == 0)
			{
				MessageBox.Show("No existen articulos seleccionados para procesar la actualización de precios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}
			#endregion Validacion

			#region Actualiza Precios
			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de actualizar los precios de los articulos seleccionados?\n\nEste proceso no puede ser revertido", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				string sSQL = "";						
				
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					if ((bool)dr.Cells["Sel"].Value)
					{
						sSQL = string.Format("Exec ActualizaListaDePrecios {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}", 
							(int)dr.Cells["idArticulo"].Value, (decimal)dr.Cells["Distribuidor"].Value, (decimal)dr.Cells["Contado"].Value, (decimal)dr.Cells["Tarjeta"].Value, (decimal)dr.Cells["Credito"].Value, (decimal)dr.Cells["CostoUltimo"].Value, (decimal)dr.Cells["WEB"].Value, (decimal)dr.Cells["TCSI"].Value,
							(decimal)dr.Cells["PrecioC3"].Value, (decimal)dr.Cells["PrecioC6"].Value, (decimal)dr.Cells["PrecioC9"].Value, (decimal)dr.Cells["PrecioC12"].Value, 
							(decimal)dr.Cells["PrecioC15"].Value, (decimal)dr.Cells["PrecioC18"].Value, (decimal)dr.Cells["PrecioC21"].Value, (decimal)dr.Cells["PrecioC24"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionPrecios Set Sel = 0 Where idActualizacion = {0}", (int)dr.Cells["idActualizacion"].Value));

						Funcion.EjecutaSQL(cdsSeteoF, "Exec WEB_ActualizaPreciosBaratazosOfertas");
					}
				}
			
				MessageBox.Show("Lista de precios actualizada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				
				ActualizaMovimientos();			
			}
			#endregion Actualiza Precios
		}

		private void ultraGrid1_InitializeTemplateAddRow(object sender, Infragistics.Win.UltraWinGrid.InitializeTemplateAddRowEventArgs e)
		{
		
		}

		private void ultraGrid1_AfterCellListCloseUp(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		
		}		
	}
}

