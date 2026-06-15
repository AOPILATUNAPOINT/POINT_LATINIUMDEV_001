using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteVentaDescuento.
	/// </summary>
	public class frmReporteVentaDescuento : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.Misc.UltraLabel lblLocal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblHasta;
		private System.Windows.Forms.Label lblDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsReporteDescuentosFacturas;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsFacturacionListaBodegas;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdReporteDescuentosFacturas;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteVentaDescuento()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteVentaDescuento));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Número");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento_Original");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento_Facturado");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion_Modifica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario_Modifica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnFactura", 1);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Descuento_Original", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Descuento_Original", 9, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Descuento_Facturado", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Descuento_Facturado", 10, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Precio", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Precio", 8, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CostoUltimo", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CostoUltimo", 11, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Número");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento_Original");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento_Facturado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoUltimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estacion_Modifica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario_Modifica");
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.lblLocal = new Infragistics.Win.Misc.UltraLabel();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsFacturacionListaBodegas = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblContador = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.lblHasta = new System.Windows.Forms.Label();
			this.lblDesde = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnVer = new System.Windows.Forms.Button();
			this.ugdReporteDescuentosFacturas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsReporteDescuentosFacturas = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExportar = new System.Windows.Forms.Button();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsFacturacionListaBodegas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdReporteDescuentosFacturas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsReporteDescuentosFacturas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1096, 19);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 1003;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.lblLocal);
			this.groupBox2.Controls.Add(this.cmbBodega);
			this.groupBox2.Controls.Add(this.lblContador);
			this.groupBox2.Controls.Add(this.txtBuscar);
			this.groupBox2.Controls.Add(this.label2);
			this.groupBox2.Controls.Add(this.lblHasta);
			this.groupBox2.Controls.Add(this.lblDesde);
			this.groupBox2.Controls.Add(this.dtHasta);
			this.groupBox2.Controls.Add(this.dtDesde);
			this.groupBox2.Controls.Add(this.btnVer);
			this.groupBox2.Location = new System.Drawing.Point(8, 11);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1072, 120);
			this.groupBox2.TabIndex = 1001;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// lblLocal
			// 
			this.lblLocal.AutoSize = true;
			this.lblLocal.BackColor = System.Drawing.Color.Transparent;
			this.lblLocal.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblLocal.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.lblLocal.Location = new System.Drawing.Point(8, 56);
			this.lblLocal.Name = "lblLocal";
			this.lblLocal.Size = new System.Drawing.Size(34, 15);
			this.lblLocal.TabIndex = 716;
			this.lblLocal.Text = "Local:";
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.udsFacturacionListaBodegas;
			ultraGridColumn1.Header.Caption = "Código";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(112, 56);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(272, 21);
			this.cmbBodega.TabIndex = 715;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// udsFacturacionListaBodegas
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.udsFacturacionListaBodegas.Band.Columns.AddRange(new object[] {
																																					 ultraDataColumn1,
																																					 ultraDataColumn2,
																																					 ultraDataColumn3});
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(488, 88);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 714;
			// 
			// txtBuscar
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance1;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(112, 88);
			this.txtBuscar.MaxLength = 300;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(272, 21);
			this.txtBuscar.TabIndex = 706;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 88);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(61, 16);
			this.label2.TabIndex = 713;
			this.label2.Text = "N° Factura:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblHasta
			// 
			this.lblHasta.AutoSize = true;
			this.lblHasta.Location = new System.Drawing.Point(400, 24);
			this.lblHasta.Name = "lblHasta";
			this.lblHasta.Size = new System.Drawing.Size(37, 16);
			this.lblHasta.TabIndex = 712;
			this.lblHasta.Text = "Hasta:";
			// 
			// lblDesde
			// 
			this.lblDesde.AutoSize = true;
			this.lblDesde.Location = new System.Drawing.Point(8, 24);
			this.lblDesde.Name = "lblDesde";
			this.lblDesde.Size = new System.Drawing.Size(40, 16);
			this.lblDesde.TabIndex = 711;
			this.lblDesde.Text = "Desde:";
			// 
			// dtHasta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance2;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(448, 24);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(280, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 709;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance3;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(112, 24);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(272, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 708;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(400, 88);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(80, 23);
			this.btnVer.TabIndex = 707;
			this.btnVer.Text = "&Consultar";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// ugdReporteDescuentosFacturas
			// 
			this.ugdReporteDescuentosFacturas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdReporteDescuentosFacturas.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdReporteDescuentosFacturas.DataSource = this.udsReporteDescuentosFacturas;
			appearance4.BackColor = System.Drawing.Color.White;
			this.ugdReporteDescuentosFacturas.DisplayLayout.Appearance = appearance4;
			appearance5.ImageHAlign = Infragistics.Win.HAlign.Right;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance5;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Almacén";
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn6.AutoEdit = false;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Format = "dd/MM/yyyy";
			ultraGridColumn6.Header.VisiblePosition = 9;
			ultraGridColumn6.Width = 79;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Forma pago";
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Width = 207;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Código";
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance6;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellButtonAppearance = appearance7;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Precio ($)";
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance8;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellButtonAppearance = appearance9;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Descuento original ($)";
			ultraGridColumn13.Header.VisiblePosition = 11;
			ultraGridColumn13.Width = 151;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance10;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellButtonAppearance = appearance11;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Descuento facturado ($)";
			ultraGridColumn14.Header.VisiblePosition = 12;
			ultraGridColumn14.Width = 168;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance12;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Ultimo Costo";
			ultraGridColumn15.Header.VisiblePosition = 13;
			ultraGridColumn15.Width = 110;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.Caption = "Estación modifica";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 134;
			ultraGridColumn17.Header.Caption = "Usuario modifica";
			ultraGridColumn17.Header.VisiblePosition = 14;
			ultraGridColumn17.Width = 130;
			ultraGridColumn18.AutoEdit = false;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn18.Header.Caption = "N°";
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Width = 32;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.Image = ((object)(resources.GetObject("appearance13.Image")));
			appearance13.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn19.CellAppearance = appearance13;
			appearance14.Image = ((object)(resources.GetObject("appearance14.Image")));
			appearance14.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn19.CellButtonAppearance = appearance14;
			ultraGridColumn19.Header.Caption = "Fac.";
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 39;
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
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance15;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance16;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance17;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance18;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4});
			this.ugdReporteDescuentosFacturas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdReporteDescuentosFacturas.DisplayLayout.Override.ActiveRowAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.Transparent;
			this.ugdReporteDescuentosFacturas.DisplayLayout.Override.CardAreaAppearance = appearance20;
			appearance21.ForeColor = System.Drawing.Color.Black;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdReporteDescuentosFacturas.DisplayLayout.Override.CellAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance22.FontData.BoldAsString = "True";
			appearance22.FontData.Name = "Arial";
			appearance22.FontData.SizeInPoints = 10F;
			appearance22.ForeColor = System.Drawing.Color.White;
			appearance22.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdReporteDescuentosFacturas.DisplayLayout.Override.HeaderAppearance = appearance22;
			appearance23.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance23.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdReporteDescuentosFacturas.DisplayLayout.Override.RowAlternateAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdReporteDescuentosFacturas.DisplayLayout.Override.RowSelectorAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdReporteDescuentosFacturas.DisplayLayout.Override.SelectedRowAppearance = appearance25;
			this.ugdReporteDescuentosFacturas.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ugdReporteDescuentosFacturas.Location = new System.Drawing.Point(8, 163);
			this.ugdReporteDescuentosFacturas.Name = "ugdReporteDescuentosFacturas";
			this.ugdReporteDescuentosFacturas.Size = new System.Drawing.Size(1176, 397);
			this.ugdReporteDescuentosFacturas.TabIndex = 999;
			this.ugdReporteDescuentosFacturas.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdReporteDescuentosFacturas_ClickCellButton);
			this.ugdReporteDescuentosFacturas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdReporteDescuentosFacturas_InitializeLayout);
			// 
			// udsReporteDescuentosFacturas
			// 
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			this.udsReporteDescuentosFacturas.Band.Columns.AddRange(new object[] {
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
																																						 ultraDataColumn15,
																																						 ultraDataColumn16,
																																						 ultraDataColumn17});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-56, 139);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1417, 8);
			this.groupBox1.TabIndex = 1000;
			this.groupBox1.TabStop = false;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnExportar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1096, 59);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(80, 23);
			this.btnExportar.TabIndex = 1002;
			this.btnExportar.Text = "&Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
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
			// frmReporteVentaDescuento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1192, 566);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.ugdReporteDescuentosFacturas);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExportar);
			this.Name = "frmReporteVentaDescuento";
			this.Text = "Reporte Descuentos";
			this.Load += new System.EventHandler(this.frmReporteVentaDescuento_Load);
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsFacturacionListaBodegas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdReporteDescuentosFacturas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsReporteDescuentosFacturas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;

		private void frmReporteVentaDescuento_Load(object sender, System.EventArgs e)
		{
			SeguridadAcceso();
			AccesoExcel();
			FechaDefault();
			BodegaVinculacion();
		}

		private void AccesoExcel()
		{
			miAcceso = new Acceso(cdsSeteoF, "20054401");
			if (miAcceso.BExportar) this.btnExportar.Enabled =  true;
		}

		private void SeguridadAcceso()
		{
			string ambiente = "PRO";

			if(ambiente == "PRO")
			{
				if (!Funcion.Permiso("2085", cdsSeteoF))
				{
					MessageBox.Show("No puede ingresar a " + this.Text, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					BeginInvoke(new MethodInvoker(UnloadMe));
					return;
				}
			}
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void FechaDefault()
		{
			//Ubicar por default las fechas al día de hoy para no estesar la base
			#region Fechas default
			DateTime dtInicio = Convert.ToDateTime("01/01/2023");
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse(dtInicio.Day.ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse(dtInicio.Day.ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
			DateTime dtFechaActual = DateTime.Today;
			this.dtDesde.Value = DateTime.Parse("01" + "/" + dtFechaActual.Month.ToString() + "/" + dtFechaActual.Year.ToString());
			this.dtHasta.Value = DateTime.Today;
			DateTime dtFin = DateTime.Today;
			this.dtDesde.CalendarInfo.MaxDate = DateTime.Parse(dtFin.Day.ToString() + "/" + dtFin.Month.ToString() + "/" + dtFin.Year.ToString());
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Parse(dtFin.Day.ToString() + "/" + dtFin.Month.ToString() + "/" + dtFin.Year.ToString());
			#endregion Fechas default
		}

		private void BodegaVinculacion()
		{
			#region Bodega
			string squery = "";
			try
			{
				squery = String.Format("Exec FacturacionListaBodegas {0}, {1}, '{2}'",
					MenuLatinium.IdUsuario,
					1,//LISTA DE ALMACENES ACTIVOS A LA FECHA HASTA DEL RePORTE
					DateTime.Now.ToString("yyyyMMdd"));
				this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, squery);
			}
			catch(Exception ex)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
			}
			#endregion Bodega
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida campos  
			if (!Validacion.vbValidaFechasDesdeHasta(this.dtDesde, this.dtHasta)) return;
			#endregion Valida campos 
			ReporteDescuentosFacturas();
		}

		private void ReporteDescuentosFacturas()
		{
			string squery = "";
			try
			{
				int getcmbBodega = 0;

				try{getcmbBodega = (int)this.cmbBodega.Value;}
				catch(Exception ex){string sms = ex.Message;}

				squery = String.Format("Exec ReporteDescuentosFacturas '{0}', '{1}', {2}, '{3}'",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
					getcmbBodega,
					this.txtBuscar.Text
					);
				FuncionesProcedimientos.dsGeneralModulo(squery, this.ugdReporteDescuentosFacturas);

				this.lblContador.Text = this.ugdReporteDescuentosFacturas.Rows.Count + " REGISTROS ENCONTRADOS";
				this.ugdReporteDescuentosFacturas = Funcion.ContadorFilas(this.ugdReporteDescuentosFacturas ,"Row");
			}
			catch(Exception ex)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.ugdReporteDescuentosFacturas.Rows.Count == 0)
			{
				MessageBox.Show("No hay filas para exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{
					ultraGridExcelExporter1.Export(this.ugdReporteDescuentosFacturas, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void ugdReporteDescuentosFacturas_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int idCompra = 0;
			int idPromocion = 0;
			string FormaPago = "";
			try{FormaPago = (string)e.Cell.Row.Cells["FormaPago"].Value;}catch(Exception ex){string sms = ex.Message;}
			
			if (e.Cell.Column.ToString() == "btnFactura")	
			{
				//FACTURA
				if(FormaPago != "CREDI POINT")
				{
					try
					{
						idCompra = (int)e.Cell.Row.Cells["idCompra"].Value;
						using (frmVenta frmS = new frmVenta(idCompra, 1))
						{		
							if (DialogResult.OK == frmS.ShowDialog())
							{
							}
						}
					}
					catch(Exception ex)
					{
						string sms = ex.Message;
						Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "", Funcion.slogtipoerror);
						return;
					}
				}

				if(FormaPago == "CREDI POINT")
				{
					//PREFACTURA
					try
					{
						idCompra = (int)e.Cell.Row.Cells["idCompra"].Value;
						using (frmPrefactura frmS = new frmPrefactura(idCompra, true))
						{		
							if (DialogResult.OK == frmS.ShowDialog())
							{
							}
						}
					}
					catch(Exception ex)
					{
						string sms = ex.Message;
						Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "", Funcion.slogtipoerror);
						return;
					}
				}
			}
		}

		private void ugdReporteDescuentosFacturas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
