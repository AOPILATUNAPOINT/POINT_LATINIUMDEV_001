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
	/// Summary description for Inventarios.
	/// </summary>
	public class Inventarios : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnCaracteristicas;
		private Acceso miAcceso;

		public Inventarios()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoReemplazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Distribuidor");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contado");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EF_Combo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TC_Combo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CR_Combo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pedido");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OrdenProduccion");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Discontinuo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoReemplazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoUltimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Distribuidor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EF_Combo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TC_Combo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CR_Combo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pedido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Promocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OrdenProduccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Discontinuo");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TF");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transito");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 1, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transito", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transito", 3, true);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TF", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TF", 2, true);
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TF");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transito");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Inventarios));
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblContador = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.optFiltro = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnCaracteristicas = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
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
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 38);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 21);
			this.cmbBodega.TabIndex = 160;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(632, 39);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 18);
			this.lblContador.TabIndex = 138;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 40);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 137;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// optFiltro
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro.Appearance = appearance1;
			this.optFiltro.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro.CheckedIndex = 0;
			this.optFiltro.ItemAppearance = appearance2;
			this.optFiltro.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "(F7) Con Existencia";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "(F8) Sin Existencia";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "(F9) Todos";
			this.optFiltro.Items.Add(valueListItem1);
			this.optFiltro.Items.Add(valueListItem2);
			this.optFiltro.Items.Add(valueListItem3);
			this.optFiltro.ItemSpacingVertical = 10;
			this.optFiltro.Location = new System.Drawing.Point(283, 6);
			this.optFiltro.Name = "optFiltro";
			this.optFiltro.Size = new System.Drawing.Size(336, 24);
			this.optFiltro.TabIndex = 136;
			this.optFiltro.Text = "(F7) Con Existencia";
			this.optFiltro.ValueChanged += new System.EventHandler(this.optFiltro_ValueChanged);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance3;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 139;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 178;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Reemplazo";
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 139;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 72;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance5;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "PI";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 55;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance6;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "PA";
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 55;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.BackColor = System.Drawing.Color.LightGreen;
			appearance7.BackColor2 = System.Drawing.Color.LightGreen;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance7;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "PA Com";
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 55;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance8;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "TC";
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Width = 54;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.BackColor = System.Drawing.Color.LightGreen;
			appearance9.BackColor2 = System.Drawing.Color.LightGreen;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance9;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "TC Com";
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Width = 54;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance10;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "PVP";
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Width = 55;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.BackColor = System.Drawing.Color.LightGreen;
			appearance11.BackColor2 = System.Drawing.Color.LightGreen;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance11;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "PVP Com";
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Width = 55;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance12;
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn16.Width = 50;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance13;
			ultraGridColumn17.Header.VisiblePosition = 13;
			ultraGridColumn17.Width = 50;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.VisiblePosition = 14;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 60;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance14;
			ultraGridColumn19.Header.Caption = "TF";
			ultraGridColumn19.Header.VisiblePosition = 15;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance15;
			ultraGridColumn20.Header.Caption = "ODP";
			ultraGridColumn20.Header.VisiblePosition = 16;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.Caption = "D";
			ultraGridColumn21.Header.VisiblePosition = 17;
			ultraGridColumn21.Width = 20;
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
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			ultraGridBand2.Header.Caption = "Inventario";
			ultraGridBand2.HeaderVisible = true;
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand2.Override.ActiveRowAppearance = appearance16;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 7.25F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance19;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(5, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(921, 222);
			this.ultraGrid1.TabIndex = 134;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			this.ultraGrid1.Click += new System.EventHandler(this.ultraGrid1_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.Double);
			ultraDataColumn5.DefaultValue = 0;
			ultraDataColumn6.DataType = typeof(System.Double);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(System.Double);
			ultraDataColumn7.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(System.Double);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn9.DataType = typeof(System.Double);
			ultraDataColumn9.DefaultValue = 0;
			ultraDataColumn10.DataType = typeof(System.Double);
			ultraDataColumn10.DefaultValue = 0;
			ultraDataColumn11.DataType = typeof(System.Double);
			ultraDataColumn11.DefaultValue = 0;
			ultraDataColumn12.DataType = typeof(System.Double);
			ultraDataColumn12.DefaultValue = 0;
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(System.Double);
			ultraDataColumn14.DefaultValue = 0;
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(bool);
			ultraDataColumn18.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// txtArticulo
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance22;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Location = new System.Drawing.Point(336, 37);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(200, 22);
			this.txtArticulo.TabIndex = 132;
			this.txtArticulo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtArticulo_KeyPress);
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
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource1;
			appearance23.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance23;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "_DetCompra";
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn22.CellAppearance = appearance24;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Width = 30;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance25;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 40;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance26;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Width = 40;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance27;
			ultraGridColumn25.Header.VisiblePosition = 3;
			ultraGridColumn25.Width = 50;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			ultraGridBand3.Header.Caption = "Existencias";
			ultraGridBand3.HeaderVisible = true;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance28;
			summarySettings1.DisplayFormat = "{0: #,##0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance29;
			summarySettings2.DisplayFormat = "{0: #,##0}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance30;
			summarySettings3.DisplayFormat = "{0: #,##0}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand3.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance31.ForeColor = System.Drawing.Color.Black;
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance31;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance32.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance32;
			appearance33.ForeColor = System.Drawing.Color.Black;
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.CellAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance34.FontData.BoldAsString = "True";
			appearance34.FontData.Name = "Arial";
			appearance34.FontData.SizeInPoints = 7.5F;
			appearance34.ForeColor = System.Drawing.Color.White;
			appearance34.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance34;
			appearance35.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance35.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance37;
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(936, 88);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(181, 222);
			this.ultraGrid2.TabIndex = 142;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn20.DefaultValue = 0;
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn21.DefaultValue = 0;
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn22.DefaultValue = 0;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label1.Location = new System.Drawing.Point(8, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 17);
			this.label1.TabIndex = 166;
			this.label1.Text = "Fecha";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance38;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 165;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1120, 8);
			this.groupBox1.TabIndex = 167;
			this.groupBox1.TabStop = false;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label2.Location = new System.Drawing.Point(280, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 17);
			this.label2.TabIndex = 168;
			this.label2.Text = "Buscar";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1024, 6);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(88, 24);
			this.btnExcel.TabIndex = 359;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(544, 36);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 358;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnCaracteristicas
			// 
			this.btnCaracteristicas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCaracteristicas.CausesValidation = false;
			this.btnCaracteristicas.Image = ((System.Drawing.Image)(resources.GetObject("btnCaracteristicas.Image")));
			this.btnCaracteristicas.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCaracteristicas.Location = new System.Drawing.Point(1024, 37);
			this.btnCaracteristicas.Name = "btnCaracteristicas";
			this.btnCaracteristicas.Size = new System.Drawing.Size(88, 23);
			this.btnCaracteristicas.TabIndex = 360;
			this.btnCaracteristicas.Text = "&Información";
			this.btnCaracteristicas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCaracteristicas.Click += new System.EventHandler(this.btnCaracteristicas_Click);
			// 
			// Inventarios
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1120, 316);
			this.Controls.Add(this.btnCaracteristicas);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.optFiltro);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtArticulo);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "Inventarios";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Inventarios";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Inventarios_KeyDown);
			this.Load += new System.EventHandler(this.Inventarios_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iBodegaPredefinida = 0;
		bool bMostrar = false;
		bool bReemplazo = false;
		bool bCombo = false;
		bool bPDistribuidor = false;
		bool bPEfectivo = false;
		bool bPTarjeta = false;
		bool bPCredito = false;
		bool bPedido = false;
		bool bExistencias = false;
		bool bTransito = false;

		private void Configuracion()
		{
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select InvCombos, InvPrecioDistribuidor, InvPrecioEfectivo, InvPrecioTarjeta, InvPrecioCredito, InvVerPedido, InvVerExSucursales, InvVerMercTransito, InvReemplazo From SeteoF", true);
			dr.Read();
			if (dr.HasRows)
			{
				bCombo = dr.GetBoolean(0);
				bPDistribuidor = dr.GetBoolean(1);
				bPEfectivo = dr.GetBoolean(2);
				bPTarjeta = dr.GetBoolean(3);
				bPCredito = dr.GetBoolean(4);
				bPedido = dr.GetBoolean(5);
				bExistencias = dr.GetBoolean(6);
				bTransito = dr.GetBoolean(7);
				bReemplazo = dr.GetBoolean(8);
			}
			dr.Close();

			if (!bReemplazo)
			{				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CodigoReemplazo"].Hidden = true;
			}

			if (!bCombo)
			{				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["EF_Combo"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["TC_Combo"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CR_Combo"].Hidden = true;
			}

			if (!bPDistribuidor)this.ultraGrid1.DisplayLayout.Bands[0].Columns["Distribuidor"].Hidden = true;

			if (!bPEfectivo)
			{				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Contado"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["EF_Combo"].Hidden = true;
			}

			if (!bPTarjeta)
			{				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Tarjeta"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["TC_Combo"].Hidden = true;
			}

			if (!bPCredito)
			{				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Credito"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CR_Combo"].Hidden = true;
			}

			if (!bPedido)
			{				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Pedido"].Hidden = true;
			}

			if (!bExistencias) this.ultraGrid2.Visible = false;					

			if (!bExistencias && !bTransito)
			{
				this.Size = new Size(1001, 632);
			}
			else if (bExistencias && !bTransito)
			{
				this.Size = new Size(1113, 632);
			}			
		}

		private void Inventarios_Load(object sender, System.EventArgs e)
		{
			Funcion.Resolucion(this, this.Height, this.Width);
			
			miAcceso = new Acceso(cdsSeteoF, "0303");			
			
			if (!Funcion.Permiso("37",cdsSeteoF))
			{					
				MessageBox.Show("No tiene acceso al inventario.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}	
			
			//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);//			this.dtFecha.Value = DateTime.Today;			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			
			FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 47, (DateTime)this.dtFecha.Value, false, cdsSeteoF);

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			iBodegaPredefinida = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			
			if (iBodegaPredefinida > 0) this.cmbBodega.Value = iBodegaPredefinida; else this.cmbBodega.Enabled = true;
			
			this.optFiltro.Value = 1;
			
			if (miAcceso.BExportarExcel) this.btnExcel.Enabled = true;		
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
						
			bMostrar = true;
		}		

		private void UnloadMe()
		{
			this.Close();
		}

		private void txtArticulo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnVer_Click(sender, e);
		}

		private void ultraGrid1_Click(object sender, System.EventArgs e)
		{
			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null)
			{
				if ((int)this.cmbBodega.Value > 0) this.Text = "INVENTARIO BODEGA: " + this.cmbBodega.Text.ToString();
				if ((int)this.cmbBodega.Value == 1) this.ultraGrid1.DisplayLayout.Bands[0].Columns["OrdenProduccion"].Hidden = false; else this.ultraGrid1.DisplayLayout.Bands[0].Columns["OrdenProduccion"].Hidden = true;
			}
			else 
			{
				this.Text = "INVENTARIO GENERAL";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["OrdenProduccion"].Hidden = false;
			}			
		}

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (ultraGrid1.ActiveRow != null)			
			{				
				string sSQL = string.Format("Exec ExistArticuloPorBodega {0}, '{1}', {2}", 
					(int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), miAcceso.Buscar);
				this.ultraGrid2.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
			}
		}

		private void Inventarios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();
			if (e.KeyCode == Keys.F7) this.optFiltro.Value = 1;
			if (e.KeyCode == Keys.F8)  this.optFiltro.Value = 2;
			if (e.KeyCode == Keys.F9)  this.optFiltro.Value = 3;
		}

		private void optFiltro_ValueChanged(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DisplayLayout.Bands[0].Header.Caption = "Inventario";
	
//			if (!bMostrar) return;
//			if (this.cmbBodega.Text.ToString().Length > 0) iBodega = (int)this.cmbBodega.Value;
//			string sSQL = string.Format("Exec ConsultaInventario 0, '{0}', {1}, {2}", this.txtArticulo.Text.ToString(), iBodega, (int)this.optFiltro.Value);
//			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
//			if (iBodega == 1) this.ultraGrid1.DisplayLayout.Bands[0].Columns["Pedido"].Hidden = false;
//			else this.ultraGrid1.DisplayLayout.Bands[0].Columns["Pedido"].Hidden = true;
//
////			this.lblContador.Text =  this.ultraGrid1.Rows.Count.ToString() + " Artículo(s) encontrados.";
//			this.ultraGrid1.DisplayLayout.Bands[0].Header.Caption = this.ultraGrid1.DisplayLayout.Bands[0].Header.Caption + " - " + this.ultraGrid1.Rows.Count.ToString() + " Artículos encontrados.";
		}		

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (Convert.ToDecimal(e.Row.Cells["EF_Combo"].Value) > 0.00m)
				e.Row.Appearance.BackColor = Color.LightGreen;
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtArticulo.Focus();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;

			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format(string.Format("Exec ListaInventarios '{0}', {1}, {2}, {3}, '{4}'", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), iBodega, 47, (int)this.optFiltro.Value, this.txtArticulo.Text.ToString()));								
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);

			this.lblContador.Text = this.ultraGrid1.Rows.Count + " ARTICULO(S) ENCONTRADOS";

			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select '' As Bodega, 0 As Saldo, 0 As TF, 0 As Transito");// From Bodega Where Activo = 1 And Inventario = 1 And Almacen = 1 Order By Bodega");			

			this.Cursor = Cursors.Default;
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.txtArticulo.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}
		
		private void btnCaracteristicas_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbGridVacio(this.ultraGrid1, "No hay articulos en la lista de inventarios", "")) return;

			if (this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un articulo de la lista", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Articulos_Caracteristicas Where idArticulo = {0}", (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value)) == 0)
			{
				MessageBox.Show("No existe información registrada del articulo seleccionado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			frmArticulosVistaCaracteristicas AVC = new frmArticulosVistaCaracteristicas((int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value, this.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString());
			AVC.ShowDialog();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}

