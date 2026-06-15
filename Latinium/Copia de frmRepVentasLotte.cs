using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRepVentasLotte.
	/// </summary>
	public class frmRepVentasLotte : System.Windows.Forms.Form
	{
		private System.Windows.Forms.TextBox txtProducto;
		private System.Windows.Forms.Label label4;
		public DevExpress.XtraEditors.SimpleButton btVer;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaFin;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIni;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugResVentas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public DevExpress.XtraEditors.SimpleButton btExportar;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private System.ComponentModel.IContainer components;

		public frmRepVentasLotte()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRepVentasLotte));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Categoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanVenta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValVenta");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Categoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValVenta");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Categoria");
			this.txtProducto = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.btVer = new DevExpress.XtraEditors.SimpleButton();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFechaFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label10 = new System.Windows.Forms.Label();
			this.dtFechaIni = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ugResVentas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btExportar = new DevExpress.XtraEditors.SimpleButton();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugResVentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			this.SuspendLayout();
			// 
			// txtProducto
			// 
			this.txtProducto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtProducto.Location = new System.Drawing.Point(232, 32);
			this.txtProducto.Name = "txtProducto";
			this.txtProducto.Size = new System.Drawing.Size(296, 20);
			this.txtProducto.TabIndex = 287;
			this.txtProducto.Text = "";
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(232, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(64, 16);
			this.label4.TabIndex = 286;
			this.label4.Text = "Producto";
			// 
			// btVer
			// 
			this.btVer.Appearance.Options.UseTextOptions = true;
			this.btVer.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btVer.Image = ((System.Drawing.Image)(resources.GetObject("btVer.Image")));
			this.btVer.ImageAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.btVer.ImageIndex = 4;
			this.btVer.Location = new System.Drawing.Point(800, 16);
			this.btVer.Name = "btVer";
			this.btVer.Size = new System.Drawing.Size(48, 40);
			this.btVer.TabIndex = 285;
			this.btVer.Click += new System.EventHandler(this.btVer_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(120, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(57, 16);
			this.label1.TabIndex = 284;
			this.label1.Text = "Fecha FIN";
			// 
			// dtFechaFin
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaFin.Appearance = appearance1;
			this.dtFechaFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaFin.DateButtons.Add(dateButton1);
			this.dtFechaFin.Format = "dd/MM/yyyy";
			this.dtFechaFin.Location = new System.Drawing.Point(120, 32);
			this.dtFechaFin.Name = "dtFechaFin";
			this.dtFechaFin.NonAutoSizeHeight = 23;
			this.dtFechaFin.Size = new System.Drawing.Size(104, 21);
			this.dtFechaFin.SpinButtonsVisible = true;
			this.dtFechaFin.TabIndex = 283;
			this.dtFechaFin.Value = ((object)(resources.GetObject("dtFechaFin.Value")));
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(8, 16);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(53, 16);
			this.label10.TabIndex = 282;
			this.label10.Text = "Fecha INI";
			// 
			// dtFechaIni
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaIni.Appearance = appearance2;
			this.dtFechaIni.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaIni.DateButtons.Add(dateButton2);
			this.dtFechaIni.Format = "dd/MM/yyyy";
			this.dtFechaIni.Location = new System.Drawing.Point(8, 32);
			this.dtFechaIni.Name = "dtFechaIni";
			this.dtFechaIni.NonAutoSizeHeight = 23;
			this.dtFechaIni.Size = new System.Drawing.Size(104, 21);
			this.dtFechaIni.SpinButtonsVisible = true;
			this.dtFechaIni.TabIndex = 281;
			this.dtFechaIni.Value = ((object)(resources.GetObject("dtFechaIni.Value")));
			// 
			// ugResVentas
			// 
			this.ugResVentas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugResVentas.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugResVentas.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ugResVentas.DisplayLayout.Appearance = appearance3;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 220;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 350;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 150;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance4;
			ultraGridColumn5.Format = "#,##0";
			ultraGridColumn5.Header.Caption = "Cantidad";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance5;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance6;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "VENTA";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand1.SummaryFooterCaption = "TOTAL GENERAL";
			this.ugResVentas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ugResVentas.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugResVentas.DisplayLayout.Override.ActiveRowAppearance = appearance7;
			this.ugResVentas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugResVentas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugResVentas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.ugResVentas.DisplayLayout.Override.CardAreaAppearance = appearance8;
			this.ugResVentas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugResVentas.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.ugResVentas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugResVentas.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugResVentas.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.ugResVentas.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ugResVentas.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ugResVentas.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ugResVentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugResVentas.Location = new System.Drawing.Point(8, 64);
			this.ugResVentas.Name = "ugResVentas";
			this.ugResVentas.Size = new System.Drawing.Size(952, 392);
			this.ugResVentas.TabIndex = 288;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btExportar
			// 
			this.btExportar.Appearance.Options.UseTextOptions = true;
			this.btExportar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btExportar.Image = ((System.Drawing.Image)(resources.GetObject("btExportar.Image")));
			this.btExportar.ImageAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.btExportar.ImageIndex = 2;
			this.btExportar.Location = new System.Drawing.Point(848, 16);
			this.btExportar.Name = "btExportar";
			this.btExportar.Size = new System.Drawing.Size(48, 40);
			this.btExportar.TabIndex = 289;
			this.btExportar.Click += new System.EventHandler(this.btExportar_Click);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBodega.Location = new System.Drawing.Point(536, 16);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(32, 16);
			this.lblBodega.TabIndex = 291;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 40;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 150;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(536, 32);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(112, 21);
			this.cmbBodega.TabIndex = 290;
			this.cmbBodega.ValueMember = "idBodega";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(656, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(55, 16);
			this.label2.TabIndex = 293;
			this.label2.Text = "Categoría";
			this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 150;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbGrupo.DisplayMember = "Categoria";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(656, 32);
			this.cmbGrupo.MaxDropDownItems = 30;
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(112, 21);
			this.cmbGrupo.TabIndex = 292;
			this.cmbGrupo.ValueMember = "idCategoria";
			// 
			// frmRepVentasLotte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(968, 461);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btExportar);
			this.Controls.Add(this.ugResVentas);
			this.Controls.Add(this.txtProducto);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.btVer);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFechaFin);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.dtFechaIni);
			this.Name = "frmRepVentasLotte";
			this.Text = "Reporte de Ventas por Producto, Local y Categoría";
			this.Load += new System.EventHandler(this.frmRepVentasLotte_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugResVentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRepVentasLotte_Load(object sender, System.EventArgs e)
		{
			this.dtFechaIni.Value = DateTime.Today;
			this.dtFechaFin.Value = DateTime.Today;

			string stExecBodega = string.Format("Select idCategoria, Categoria From ArticuloCategoria Where idCategoria in (214,215)");
			this.cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExecBodega);

			string stExecGrupo = string.Format("Select idBodega, Bodega, Codigo, Nombre From Bodega Where Activo=1 and Almacen=1");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExecGrupo);
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btVer_Click(object sender, System.EventArgs e)
		{
			string Prodct = "";
			int nBod = 0;
			int nCate = 0;

			if (this.dtFechaIni.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione Fecha Inicial para REPORTE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaIni.Focus();
			}
			else if (this.dtFechaFin.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione Fecha Inicial para REPORTE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaFin.Focus();
			}
			else
			{
				DateTime dtFeIni = (DateTime) this.dtFechaIni.Value;
				DateTime dtFeFin = (DateTime) this.dtFechaFin.Value;
				if(this.cmbBodega.ActiveRow != null) nBod = (int)this.cmbBodega.Value;
				if(this.cmbGrupo.ActiveRow != null) nCate = (int)this.cmbGrupo.Value;

				if(this.txtProducto.Text.ToString() != "") Prodct = this.txtProducto.Text.ToString();
				string sSQL = string.Format("Exec VentasxProducto '{0}','{1}','{2}',{3},{4}", 
					dtFeIni.ToString("yyyyMMdd"), 
					dtFeFin.ToString("yyyyMMdd"), 
					Prodct,
					nBod,
					nCate);
				this.ugResVentas.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
			}
		}

		private void btExportar_Click(object sender, System.EventArgs e)
		{
			if (this.ugResVentas.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.ugResVentas, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}
	}
}
