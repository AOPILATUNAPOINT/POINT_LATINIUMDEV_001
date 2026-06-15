using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRepVentasArtLote.
	/// </summary>
	public class frmRepVentasArtLote : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAño;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgReporte;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label lblHabitacion;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTurno;
		private System.ComponentModel.IContainer components;

		public frmRepVentasArtLote()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Categoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValVenta");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRepVentasArtLote));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtAño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.rdgReporte = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label18 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label17 = new System.Windows.Forms.Label();
			this.lblHabitacion = new System.Windows.Forms.Label();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.optTurno = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTurno)).BeginInit();
			this.SuspendLayout();
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
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBodega.Location = new System.Drawing.Point(16, 56);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 17);
			this.lblBodega.TabIndex = 291;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 40;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 150;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(56, 53);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 290;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// txtAño
			// 
			this.txtAño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAño.Enabled = false;
			this.txtAño.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtAño.Location = new System.Drawing.Point(208, 13);
			this.txtAño.MinValue = 2014;
			this.txtAño.Name = "txtAño";
			this.txtAño.PromptChar = ' ';
			this.txtAño.Size = new System.Drawing.Size(88, 22);
			this.txtAño.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAño.TabIndex = 735;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(176, 16);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(23, 17);
			this.label7.TabIndex = 736;
			this.label7.Text = "Año";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 96);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1048, 440);
			this.uGridInformacion.TabIndex = 737;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-24, 80);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1112, 8);
			this.groupBox1.TabIndex = 738;
			this.groupBox1.TabStop = false;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Location = new System.Drawing.Point(984, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.TabIndex = 740;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Location = new System.Drawing.Point(528, 53);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 739;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(272, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(52, 17);
			this.label1.TabIndex = 741;
			this.label1.Text = "Categoria";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataSource = this.ultraDataSource2;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 180;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbGrupo.DisplayMember = "SubGrupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(328, 53);
			this.cmbGrupo.MaxDropDownItems = 30;
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(180, 22);
			this.cmbGrupo.TabIndex = 744;
			this.cmbGrupo.ValueMember = "idSubGrupo";
			// 
			// rdgReporte
			// 
			appearance7.FontData.Name = "Tahoma";
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgReporte.Appearance = appearance7;
			this.rdgReporte.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rdgReporte.CheckedIndex = 0;
			this.rdgReporte.ItemAppearance = appearance8;
			this.rdgReporte.ItemOrigin = new System.Drawing.Point(10, 3);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Diario";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Anual";
			this.rdgReporte.Items.Add(valueListItem1);
			this.rdgReporte.Items.Add(valueListItem2);
			this.rdgReporte.ItemSpacingVertical = 10;
			this.rdgReporte.Location = new System.Drawing.Point(16, 8);
			this.rdgReporte.Name = "rdgReporte";
			this.rdgReporte.Size = new System.Drawing.Size(128, 33);
			this.rdgReporte.TabIndex = 745;
			this.rdgReporte.Text = "Diario";
			this.rdgReporte.ValueChanged += new System.EventHandler(this.rdgReporte_ValueChanged);
			// 
			// dtDesde
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance9;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(368, 14);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 746;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(328, 16);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(39, 17);
			this.label18.TabIndex = 747;
			this.label18.Text = "Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance10;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(536, 14);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 748;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(496, 16);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(37, 17);
			this.label17.TabIndex = 749;
			this.label17.Text = "Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblHabitacion
			// 
			this.lblHabitacion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblHabitacion.Font = new System.Drawing.Font("Tahoma", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblHabitacion.ForeColor = System.Drawing.Color.Firebrick;
			this.lblHabitacion.Location = new System.Drawing.Point(168, 8);
			this.lblHabitacion.Name = "lblHabitacion";
			this.lblHabitacion.Size = new System.Drawing.Size(488, 33);
			this.lblHabitacion.TabIndex = 750;
			this.lblHabitacion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel3.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel3.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel3.Location = new System.Drawing.Point(664, 16);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(34, 15);
			this.ultraLabel3.TabIndex = 1003;
			this.ultraLabel3.Text = "Turno";
			// 
			// optTurno
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTurno.Appearance = appearance11;
			this.optTurno.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTurno.CheckedIndex = 0;
			this.optTurno.ItemAppearance = appearance12;
			this.optTurno.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = 0;
			valueListItem3.DisplayText = "Todos";
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "Turno 1";
			valueListItem5.DataValue = 2;
			valueListItem5.DisplayText = "Turno 2";
			this.optTurno.Items.Add(valueListItem3);
			this.optTurno.Items.Add(valueListItem4);
			this.optTurno.Items.Add(valueListItem5);
			this.optTurno.ItemSpacingVertical = 10;
			this.optTurno.Location = new System.Drawing.Point(712, 8);
			this.optTurno.Name = "optTurno";
			this.optTurno.Size = new System.Drawing.Size(240, 26);
			this.optTurno.TabIndex = 1002;
			this.optTurno.Text = "Todos";
			// 
			// frmRepVentasArtLote
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1064, 542);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.optTurno);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.txtAño);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.rdgReporte);
			this.Controls.Add(this.lblHabitacion);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "frmRepVentasArtLote";
			this.Text = "Reporte de Ventas por Producto, Local y Categoría";
			this.Load += new System.EventHandler(this.frmRepVentasArtLote_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTurno)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRepVentasArtLote_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;

			this.txtAño.MaxValue = DateTime.Today.Year;
			this.txtAño.Value = DateTime.Today.Year;

			string sSQL = string.Format("SELECT idSubGrupo, SubGrupo FROM ArticuloSubGrupo Where idSubGrupo In (207, 214, 208, 209, 221)");
			this.cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);

			this.cmbGrupo.Value = 207;

			sSQL = string.Format("Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 And Almacen = 1");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{			
			int iBodega = 0;
			
			if(this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;
						
			this.uGridInformacion.DataSource = null;

			if ((int)this.rdgReporte.Value == 1)
			{
				#region Validacion
				if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la Fecha Inicial", false, "", cdsSeteo)) return;

				if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la Fecha Final", false, "", cdsSeteo)) return;

				if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;
				#endregion Validacion

				DateTime dtDesde = (DateTime) this.dtDesde.Value;
				DateTime dtHasta = (DateTime) this.dtHasta.Value;

				string sSQL = string.Format("Exec ReporteVentasxDia '{0}', '{1}', {2}, {3}, {4}", 
					dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), iBodega, (int)this.cmbGrupo.Value, (int)this.optTurno.Value);
				this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);

				DiseñoGrid();
			}

			if ((int)this.rdgReporte.Value == 2)
			{
				string sSQL = string.Format("Exec ReporteVentasxMes {0}, {1}, {2}", 
					Convert.ToInt32(this.txtAño.Value), iBodega, (int)this.cmbGrupo.Value);
				this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);

				DiseñoGrid();
			}
		}

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridInformacion.Rows.Count > 0)
			{
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Bodega"].Width = 140;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Bodega"].CellActivation = Activation.Disabled;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Articulo"].Width = 250;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.ActivateOnly;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Precio"].Width = 80;				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Precio"].CellActivation = Activation.Disabled;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Precio"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Precio"].Format = "#,##0.00";
				
				if (this.uGridInformacion.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 3; j < iColumnas; j++)
					{
						nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.Disabled;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 60;

						this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
					}
				}				
			}			
			#endregion Diseño Grid
		}

		private void rdgReporte_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int)this.rdgReporte.Value == 1)
			{
				this.dtDesde.Enabled = true;
				this.dtHasta.Enabled = true;
				this.optTurno.Enabled = true;
				this.txtAño.Enabled = false;
			}
			
			if ((int)this.rdgReporte.Value == 2)
			{			
				this.txtAño.Enabled = true;		
				this.dtDesde.Enabled = false;
				this.dtHasta.Enabled = false;
				this.optTurno.Enabled = false;
			}
		}
	}
}
