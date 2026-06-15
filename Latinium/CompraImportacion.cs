using System;
using System.Globalization;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraImportacion.
	/// </summary>
	public class CompraImportacion : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ToolTip toolTip1;
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataSet cdsCuenta;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label29;
		private C1.Data.C1DataSet cdsImportacion;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalFob;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDetalle;
		private Infragistics.Win.Misc.UltraButton btGenerar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private Infragistics.Win.Misc.UltraButton btAsiento;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataView cdvCuenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsImportacion;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.Misc.UltraButton btPlanCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdImportacion;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoImportacion;
		private C1.Data.C1DataSet cdsTipoImportacion;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIngreso;
		private Infragistics.Win.Misc.UltraButton btArancel;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFactor;
		private Infragistics.Win.Misc.UltraButton btIngreso;
		private Infragistics.Win.Misc.UltraButton btCostea;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFob;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtArancel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.Misc.UltraButton btGenArticulo;
		private Infragistics.Win.Misc.UltraButton btDauA;
		private Infragistics.Win.Misc.UltraButton btDauB;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFormulario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private System.ComponentModel.IContainer components;

		public CompraImportacion()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Iva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdSucursal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CentroCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyecto");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_CompraDetImp", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraDetImp");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoImportacion");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Iva", 9, true, "_CompraDetImp", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Iva", 9, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 8, true, "_CompraDetImp", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 8, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ImportacionTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Orden");
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsImportacion = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.btGenArticulo = new Infragistics.Win.Misc.UltraButton();
			this.btGenerar = new Infragistics.Win.Misc.UltraButton();
			this.label8 = new System.Windows.Forms.Label();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cmbFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.label29 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDetalle = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFormulario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btDauB = new Infragistics.Win.Misc.UltraButton();
			this.btDauA = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.txtArancel = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFob = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngreso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btArancel = new Infragistics.Win.Misc.UltraButton();
			this.txtFactor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btIngreso = new Infragistics.Win.Misc.UltraButton();
			this.btCostea = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btPlanCompra = new Infragistics.Win.Misc.UltraButton();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtTotalFob = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.btAsiento = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.udsImportacion = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtIdImportacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsTipoImportacion = new C1.Data.C1DataSet();
			this.cmbTipoImportacion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFormulario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArancel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFob)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFob)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.ultraButton2);
			this.ultraTabPageControl1.Controls.Add(this.txtArticulo);
			this.ultraTabPageControl1.Controls.Add(this.label2);
			this.ultraTabPageControl1.Controls.Add(this.btGenArticulo);
			this.ultraTabPageControl1.Controls.Add(this.btGenerar);
			this.ultraTabPageControl1.Controls.Add(this.label8);
			this.ultraTabPageControl1.Controls.Add(this.cmbCuenta);
			this.ultraTabPageControl1.Controls.Add(this.cmbFecha);
			this.ultraTabPageControl1.Controls.Add(this.label29);
			this.ultraTabPageControl1.Controls.Add(this.label1);
			this.ultraTabPageControl1.Controls.Add(this.txtDetalle);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(652, 102);
			// 
			// ultraButton2
			// 
			this.ultraButton2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton2.Location = new System.Drawing.Point(528, 72);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(64, 23);
			this.ultraButton2.TabIndex = 72;
			this.ultraButton2.Text = "Excel";
			this.ultraButton2.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// txtArticulo
			// 
			this.txtArticulo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "CompraImportacion.Numero"));
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Location = new System.Drawing.Point(88, 72);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(100, 21);
			this.txtArticulo.TabIndex = 71;
			// 
			// cdsImportacion
			// 
			this.cdsImportacion.BindingContextCtrl = this;
			this.cdsImportacion.DataLibrary = "LibFacturacion";
			this.cdsImportacion.DataLibraryUrl = "";
			this.cdsImportacion.DataSetDef = "dsImportacion";
			this.cdsImportacion.FillOnRequest = false;
			this.cdsImportacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsImportacion.Name = "cdsImportacion";
			this.cdsImportacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsImportacion.SchemaDef = null;
			this.cdsImportacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsImportacion_BeforeFill_1);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(24, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(42, 16);
			this.label2.TabIndex = 70;
			this.label2.Text = "Artículo";
			// 
			// btGenArticulo
			// 
			this.btGenArticulo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGenArticulo.Location = new System.Drawing.Point(200, 72);
			this.btGenArticulo.Name = "btGenArticulo";
			this.btGenArticulo.Size = new System.Drawing.Size(112, 23);
			this.btGenArticulo.TabIndex = 69;
			this.btGenArticulo.Text = "Generar Artículo";
			this.btGenArticulo.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// btGenerar
			// 
			this.btGenerar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGenerar.Location = new System.Drawing.Point(376, 72);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(104, 23);
			this.btGenerar.TabIndex = 67;
			this.btGenerar.Text = "Generar Asiento";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(24, 48);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(40, 16);
			this.label8.TabIndex = 59;
			this.label8.Text = "Detalle";
			// 
			// cmbCuenta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance1;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_CompraImportacion.idCuenta"));
			this.cmbCuenta.DataSource = this.cdvCuenta;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 121;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 234;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuenta.DisplayMember = "CodRapido";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(88, 24);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(208, 21);
			this.cmbCuenta.TabIndex = 65;
			this.cmbCuenta.ValueMember = "idCuenta";
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.Sort = "Codigo";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// cmbFecha
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFecha.Appearance = appearance2;
			this.cmbFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_CompraImportacion.Fecha"));
			this.cmbFecha.DateTime = new System.DateTime(2006, 12, 27, 0, 0, 0, 0);
			this.cmbFecha.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFecha.FormatString = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(376, 24);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbFecha.TabIndex = 63;
			this.cmbFecha.Value = new System.DateTime(2006, 12, 27, 0, 0, 0, 0);
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(320, 24);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(36, 16);
			this.label29.TabIndex = 60;
			this.label29.Text = "Fecha";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(24, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(41, 16);
			this.label1.TabIndex = 6;
			this.label1.Text = "Cuenta";
			// 
			// txtDetalle
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDetalle.Appearance = appearance3;
			this.txtDetalle.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_CompraImportacion.Detalle"));
			this.txtDetalle.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDetalle.Location = new System.Drawing.Point(88, 48);
			this.txtDetalle.Name = "txtDetalle";
			this.txtDetalle.Size = new System.Drawing.Size(208, 21);
			this.txtDetalle.TabIndex = 66;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.ultraLabel1);
			this.ultraTabPageControl2.Controls.Add(this.txtTotal);
			this.ultraTabPageControl2.Controls.Add(this.txtFormulario);
			this.ultraTabPageControl2.Controls.Add(this.btDauB);
			this.ultraTabPageControl2.Controls.Add(this.btDauA);
			this.ultraTabPageControl2.Controls.Add(this.ultraLabel2);
			this.ultraTabPageControl2.Controls.Add(this.txtArancel);
			this.ultraTabPageControl2.Controls.Add(this.txtFob);
			this.ultraTabPageControl2.Controls.Add(this.txtIngreso);
			this.ultraTabPageControl2.Controls.Add(this.btArancel);
			this.ultraTabPageControl2.Controls.Add(this.txtFactor);
			this.ultraTabPageControl2.Controls.Add(this.btIngreso);
			this.ultraTabPageControl2.Controls.Add(this.btCostea);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(652, 102);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Location = new System.Drawing.Point(24, 80);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(77, 14);
			this.ultraLabel1.TabIndex = 92;
			this.ultraLabel1.Text = "FOB + Ar. Tot.";
			this.toolTip1.SetToolTip(this.ultraLabel1, "Fob + Arancel Total");
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance4;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.FormatString = "#,##0.000000";
			this.txtTotal.Location = new System.Drawing.Point(104, 80);
			this.txtTotal.MaskInput = "{LOC}nnnnnn.nnnnnn";
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(88, 21);
			this.txtTotal.TabIndex = 91;
			// 
			// txtFormulario
			// 
			this.txtFormulario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFormulario.Appearance = appearance5;
			this.txtFormulario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFormulario.Location = new System.Drawing.Point(512, 16);
			this.txtFormulario.Name = "txtFormulario";
			this.txtFormulario.Size = new System.Drawing.Size(96, 21);
			this.txtFormulario.TabIndex = 90;
			// 
			// btDauB
			// 
			this.btDauB.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btDauB.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btDauB.Location = new System.Drawing.Point(416, 48);
			this.btDauB.Name = "btDauB";
			this.btDauB.Size = new System.Drawing.Size(56, 23);
			this.btDauB.TabIndex = 89;
			this.btDauB.Text = "DUI B";
			this.toolTip1.SetToolTip(this.btDauB, "Impresion de Arancel de Ingreso ej: \'ING1\', \'ING2\'");
			this.btDauB.Click += new System.EventHandler(this.btDauB_Click);
			// 
			// btDauA
			// 
			this.btDauA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btDauA.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btDauA.Location = new System.Drawing.Point(416, 16);
			this.btDauA.Name = "btDauA";
			this.btDauA.Size = new System.Drawing.Size(56, 23);
			this.btDauA.TabIndex = 88;
			this.btDauA.Text = "DUI A";
			this.toolTip1.SetToolTip(this.btDauA, "Impresion de Aranacel de Ingreso  ej: \'ING1\', \'ING2\'");
			this.btDauA.Click += new System.EventHandler(this.btDauA_Click);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel2.Location = new System.Drawing.Point(24, 48);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(40, 14);
			this.ultraLabel2.TabIndex = 87;
			this.ultraLabel2.Text = "Gastos";
			this.toolTip1.SetToolTip(this.ultraLabel2, "Gasto / Costo + Arancel total");
			// 
			// txtArancel
			// 
			this.txtArancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArancel.Appearance = appearance6;
			this.txtArancel.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_CompraImportacion.DerechoArancel"));
			this.txtArancel.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArancel.FormatString = "#,##0.00";
			this.txtArancel.Location = new System.Drawing.Point(280, 48);
			this.txtArancel.Name = "txtArancel";
			this.txtArancel.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtArancel.PromptChar = ' ';
			this.txtArancel.Size = new System.Drawing.Size(96, 21);
			this.txtArancel.TabIndex = 85;
			this.txtArancel.Visible = false;
			// 
			// txtFob
			// 
			this.txtFob.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFob.Appearance = appearance7;
			this.txtFob.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_CompraImportacion.Fob"));
			this.txtFob.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFob.FormatString = "#,##0.000000";
			this.txtFob.Location = new System.Drawing.Point(104, 48);
			this.txtFob.MaskInput = "{LOC}nnnnnnnn.nnnnnn";
			this.txtFob.Name = "txtFob";
			this.txtFob.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFob.PromptChar = ' ';
			this.txtFob.Size = new System.Drawing.Size(88, 21);
			this.txtFob.TabIndex = 84;
			// 
			// txtIngreso
			// 
			this.txtIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngreso.Appearance = appearance8;
			this.txtIngreso.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_CompraImportacion.Numero"));
			this.txtIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngreso.Location = new System.Drawing.Point(280, 16);
			this.txtIngreso.Name = "txtIngreso";
			this.txtIngreso.Size = new System.Drawing.Size(96, 21);
			this.txtIngreso.TabIndex = 83;
			// 
			// btArancel
			// 
			this.btArancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btArancel.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btArancel.Location = new System.Drawing.Point(216, 48);
			this.btArancel.Name = "btArancel";
			this.btArancel.Size = new System.Drawing.Size(56, 23);
			this.btArancel.TabIndex = 82;
			this.btArancel.Text = "Arancel";
			this.toolTip1.SetToolTip(this.btArancel, "Impresion de Aranacel de Ingreso");
			this.btArancel.Click += new System.EventHandler(this.btArancel_Click);
			this.btArancel.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btArancel_MouseDown);
			// 
			// txtFactor
			// 
			this.txtFactor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFactor.Appearance = appearance9;
			this.txtFactor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_CompraImportacion.Factor"));
			this.txtFactor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFactor.FormatString = "#,##0.000000";
			this.txtFactor.Location = new System.Drawing.Point(104, 16);
			this.txtFactor.MaskInput = "{LOC}nnn.nnnnnn";
			this.txtFactor.Name = "txtFactor";
			this.txtFactor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFactor.PromptChar = ' ';
			this.txtFactor.Size = new System.Drawing.Size(88, 21);
			this.txtFactor.TabIndex = 81;
			// 
			// btIngreso
			// 
			this.btIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btIngreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btIngreso.Location = new System.Drawing.Point(216, 16);
			this.btIngreso.Name = "btIngreso";
			this.btIngreso.Size = new System.Drawing.Size(56, 23);
			this.btIngreso.TabIndex = 80;
			this.btIngreso.Text = "Ingreso";
			this.btIngreso.Click += new System.EventHandler(this.btTotal_Click);
			// 
			// btCostea
			// 
			this.btCostea.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCostea.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCostea.Location = new System.Drawing.Point(24, 16);
			this.btCostea.Name = "btCostea";
			this.btCostea.Size = new System.Drawing.Size(72, 23);
			this.btCostea.TabIndex = 79;
			this.btCostea.Text = "CIF / FOB";
			this.toolTip1.SetToolTip(this.btCostea, "Calculo de costo por Importacion");
			this.btCostea.Click += new System.EventHandler(this.btCostea_Click);
			// 
			// btPlanCompra
			// 
			this.btPlanCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btPlanCompra.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPlanCompra.Location = new System.Drawing.Point(392, 408);
			this.btPlanCompra.Name = "btPlanCompra";
			this.btPlanCompra.Size = new System.Drawing.Size(32, 23);
			this.btPlanCompra.TabIndex = 75;
			this.btPlanCompra.Text = "PC";
			this.toolTip1.SetToolTip(this.btPlanCompra, "Plan de Compras");
			this.btPlanCompra.Click += new System.EventHandler(this.btPlanCompra_Click);
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill_1);
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDatoSQL1.BarraMovimiento = true;
			this.barraDatoSQL1.DataMember = "_CompraImportacion";
			this.barraDatoSQL1.DataNombreId = "idImportacion";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsImportacion;
			this.barraDatoSQL1.DataTabla = "CompraImportacion";
			this.barraDatoSQL1.DataTablaHija = "CompraDetImp";
			this.barraDatoSQL1.Location = new System.Drawing.Point(24, 408);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(304, 24);
			this.barraDatoSQL1.TabIndex = 61;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = true;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDatoSQL1_Buscar);
			this.barraDatoSQL1.Load += new System.EventHandler(this.barraDatoSQL1_Load);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDatoSQL1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDatoSQL1_Imprimir);
			this.barraDatoSQL1.Restaurar += new System.EventHandler(this.barraDatoSQL1_Restaurar);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDatoSQL1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDatoSQL1.Impresora += new System.EventHandler(this.barraDatoSQL1_Impresora);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDatoSQL1_Grabar);
			// 
			// txtTotalFob
			// 
			this.txtTotalFob.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtTotalFob.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalFob.FormatString = "#,##0.00";
			this.txtTotalFob.Location = new System.Drawing.Point(584, 424);
			this.txtTotalFob.Name = "txtTotalFob";
			this.txtTotalFob.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalFob.PromptChar = ' ';
			this.txtTotalFob.Size = new System.Drawing.Size(100, 21);
			this.txtTotalFob.TabIndex = 62;
			// 
			// btFactura
			// 
			this.btFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btFactura.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFactura.Location = new System.Drawing.Point(328, 408);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(32, 23);
			this.btFactura.TabIndex = 69;
			this.btFactura.Text = "F";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// btAsiento
			// 
			this.btAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAsiento.Location = new System.Drawing.Point(360, 408);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(32, 23);
			this.btAsiento.TabIndex = 70;
			this.btAsiento.Text = "A";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// udsImportacion
			// 
			ultraDataColumn1.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(System.Double);
			ultraDataColumn5.DataType = typeof(System.Double);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			this.udsImportacion.Band.Columns.AddRange(new object[] {
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
																															 ultraDataColumn11});
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "_CompraImportacion.CompraImportacion - CompraDetImp";
			this.ultraGrid2.DataSource = this.cdsImportacion;
			appearance10.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance10;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn6.Header.VisiblePosition = 9;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 15;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 11;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 14;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 12;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Width = 96;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance11;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Hidden = true;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance12;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Width = 77;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance13;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Width = 77;
			ultraGridColumn16.Header.Caption = "Cod. Asiento";
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Width = 96;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 96;
			ultraGridColumn18.Format = "dd/MMM/yyyy";
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Width = 101;
			ultraGridColumn19.Header.VisiblePosition = 7;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 8;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.Caption = "Importacion";
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn21.Width = 92;
			ultraGridBand2.Columns.AddRange(new object[] {
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
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand2.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 10F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.ultraGrid2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(24, 152);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(656, 240);
			this.ultraGrid2.TabIndex = 73;
			this.ultraGrid2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid2_KeyDown);
			// 
			// txtIdImportacion
			// 
			this.txtIdImportacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdImportacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "CompraImportacion.idImportacion"));
			this.txtIdImportacion.Location = new System.Drawing.Point(488, 416);
			this.txtIdImportacion.Name = "txtIdImportacion";
			this.txtIdImportacion.PromptChar = ' ';
			this.txtIdImportacion.Size = new System.Drawing.Size(56, 21);
			this.txtIdImportacion.TabIndex = 76;
			// 
			// cdsTipoImportacion
			// 
			this.cdsTipoImportacion.BindingContextCtrl = this;
			this.cdsTipoImportacion.DataLibrary = "LibFacturacion";
			this.cdsTipoImportacion.DataLibraryUrl = "";
			this.cdsTipoImportacion.DataSetDef = "dsTipoImportacion";
			this.cdsTipoImportacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoImportacion.Name = "cdsTipoImportacion";
			this.cdsTipoImportacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipoImportacion.SchemaDef = null;
			this.cdsTipoImportacion.BeforeFill += new C1.Data.FillEventHandler(this.dsTipoImportacion_BeforeFill);
			// 
			// cmbTipoImportacion
			// 
			this.cmbTipoImportacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoImportacion.DataMember = "ImportacionTipo";
			this.cmbTipoImportacion.DataSource = this.cdsTipoImportacion;
			ultraGridColumn22.Header.VisiblePosition = 2;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Width = 212;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn25.Header.VisiblePosition = 3;
			ultraGridColumn25.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbTipoImportacion.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipoImportacion.DisplayMember = "Nombre";
			this.cmbTipoImportacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoImportacion.Location = new System.Drawing.Point(304, 224);
			this.cmbTipoImportacion.Name = "cmbTipoImportacion";
			this.cmbTipoImportacion.Size = new System.Drawing.Size(312, 88);
			this.cmbTipoImportacion.TabIndex = 77;
			this.cmbTipoImportacion.ValueMember = "idImportacionTipo";
			this.cmbTipoImportacion.Visible = false;
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Location = new System.Drawing.Point(24, 16);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(656, 128);
			this.ultraTabControl1.TabIndex = 79;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Básico";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Cálculos";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(652, 102);
			// 
			// CompraImportacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(696, 446);
			this.Controls.Add(this.ultraTabControl1);
			this.Controls.Add(this.cmbTipoImportacion);
			this.Controls.Add(this.txtIdImportacion);
			this.Controls.Add(this.btPlanCompra);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.txtTotalFob);
			this.Controls.Add(this.barraDatoSQL1);
			this.Name = "CompraImportacion";
			this.Text = "Importación";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.CompraImportacion_Closing);
			this.Load += new System.EventHandler(this.CompraImportacion_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFormulario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArancel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFob)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFob)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDatoSQL1_Crear(object sender, System.EventArgs e)
		{
			this.barraDatoSQL1.CrearRegistro();
			ModificaGrilla(true);
		}

		private void barraDatoSQL1_Grabar(object sender, System.EventArgs e)
		{
			this.barraDatoSQL1.GrabaRegistro();
			ModificaGrilla(false);
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
		//	btGenerar_Click(this, e);
			Cursor = Cursors.Default;
		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			this.barraDatoSQL1.EditarRegistro();
			ModificaGrilla(true);
		}

		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDatoSQL1.DeshacerRegistro();
			ModificaGrilla(false);
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			e.Row.Cells["idArtCod"].Value = e.Row.Cells["IdArticulo"].Value;
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Impresora(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{AsientoDetalle.idCuenta} = " + cmbCuenta.Value.ToString();
			Reporte miRep = new Reporte("FacturaImportacion.rpt", stFiltro);
			miRep.MdiParent = MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;		
		}

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid2.Enabled = true;
			if (bModifica)
			{
				ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void barraDatoSQL1_Imprimir(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{CompraDetImp.idImportacion} = " + txtIdImportacion.Value.ToString();
			Reporte miRep = new Reporte("ImportacionFacturas.rpt", stFiltro);
			miRep.MdiParent = MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void barraDatoSQL1_Restaurar(object sender, System.EventArgs e)
		{
			if (DialogResult.No ==	MessageBox.Show(
				"¿Está seguro que desea borrar Importaciones Repetidas?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;
			string stExec = "Exec ImportacionBorraRep";
			Funcion.EjecutaSQL(cdsCliente, stExec, true);
			this.barraDatoSQL1.ProximoId(5);
		}

		private void CompraImportacion_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			txtTotalFob.Width = 0;
			this.barraDatoSQL1.PosUltima();
			txtIdImportacion.Width = 0;
			barraDatoSQL1.HabilitaControles(false);
			string stAudita = "Exec AuditaCrear 53, 6, 'Guia Remision'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
			ModificaGrilla(false);
		}

		private void barraDatoSQL1_Borrar(object sender, System.EventArgs e)
		{
			if (DialogResult.No ==	MessageBox.Show(
				"¿Está seguro que desea borrar la Importación?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;
			this.barraDatoSQL1.BorraRegistro();
		}

		private void BuscarImportacion(int idCompra)
		{
			if (idCompra > 0) 
			{
				this.barraDatoSQL1.IdRegistro = idCompra;
				this.barraDatoSQL1.ProximoId(5);
			}
		}

		private void barraDatoSQL1_Buscar(object sender, System.EventArgs e)
		{
			using (BuscaImportacion miBusqueda = new BuscaImportacion())
			{																																 
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					if (miBusqueda.idBusca.Text.Length == 0) return;
					int id = int.Parse(miBusqueda.idBusca.Text);
					BuscarImportacion(id);
				}
			}
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

			C1.Data.SchemaObjects.DataSetDef dataSetDef = 
				e.DataSet.Schema.DataSetDefs["dsCliente"];

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"],
				"[Proveedor] <> 0 Or Ambos <> 0"));
		}

		private void btTotal_Click(object sender, System.EventArgs e)
		{
			BorraErrores();
			string stNumIngreso = txtIngreso.Text.Trim();
			if (stNumIngreso.Length == 0)
			{
				errorProvider.SetError(txtIngreso, "Ingreso no puede estar vacio");
				return;
			}
			string stSelect = "Select idCompra From Compra Where Numero = '"
				+ stNumIngreso + "' And idTipoFactura In (9, 14, 26, 29)";
			int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect, false);
			if (IdCompra <= 0) 
			{
				errorProvider.SetError(txtIngreso, "Ingreso no Existe");
				return;
			}
			stSelect = "Select idTipoFactura From Compra Where idCompra = " + IdCompra.ToString();
			int idTipoFactura = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
			Compra miCompra = new Compra(idTipoFactura, IdCompra);
			miCompra.MdiParent = this.MdiParent;
			miCompra.Show();
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsImportacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
//			Funcion miFuncion = new Funcion();
//			this.cdsImportacion.Schema.Connections[0].ConnectionString = 
//				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		CultureInfo us = new CultureInfo("en-US");

		private void cdsCliente_BeforeFill_1(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void BorraErrores()
		{
			errorProvider.SetError(txtIngreso, "");
			errorProvider.SetError(cmbCuenta, "");
			errorProvider.SetError(txtFactor, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (cmbCuenta.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbCuenta, "Ingrese Cuenta");
				bOk = false;
			}
			return bOk;
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			if (barraDatoSQL1.bEditar) 
			{
				MessageBox.Show("Para generar debe estar grabado el registro", "Información");
				return;
			}
			string stExec = "EXEC ImportacionDetalle " + txtIdImportacion.Value.ToString()
				+ ", " + cmbCuenta.Value.ToString();
			Funcion.EjecutaSQL(cdsCompra, stExec, true);
			this.barraDatoSQL1.ProximoId(5);
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			int IdCompra = (int) ultraGrid2.ActiveRow.Cells["idCompra"].Value;
			Compra miCompra = new Compra(4, IdCompra);
			miCompra.MdiParent = this.MdiParent;
			miCompra.Show();
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			int IdAsiento = (int) ultraGrid2.ActiveRow.Cells["idAsiento"].Value;
			Asiento miAsiento = new Asiento(IdAsiento);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();		
		}

		private void btCostea_Click(object sender, System.EventArgs e)
		{
			BorraErrores();
			string stNumIngreso = txtIngreso.Text.Trim();
			if (stNumIngreso.Length == 0)
			{
				errorProvider.SetError(txtIngreso, "Ingreso no puede estar vacio");
				return;
			}
			string stSelect = "Select Count(*) From Compra Where Numero = '"
				+ stNumIngreso + "' And idTipoFactura In (9, 14, 26, 29)";
			int IdCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, false);
			if (IdCuenta > 1) 
			{
				errorProvider.SetError(txtIngreso, "Existe mas de un Ingreso con el codigo " + stNumIngreso);
				return;
			}
			if (IdCuenta < 1) 
			{
				errorProvider.SetError(txtIngreso, "NO Existe un Ingreso con el codigo " + stNumIngreso);
				return;
			}
			stSelect = "Select idCompra From Compra Where Numero = '"
				+ stNumIngreso + "' And idTipoFactura In (9, 14, 26, 29)";
			int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect, false);

			if (txtFactor == null || txtFactor.Value == null || (double) txtFactor.Value <= 0)
			{
				errorProvider.SetError(txtFactor, "Ingrese Factor mayor a 0");
				return;
			}
			double dFactor = (double) txtFactor.Value;
			double dFactor1 = (double) txtFob.Value;
			double dTotal = (double) txtTotal.Value;
			stSelect = "EXEC ImportacionCosto " + IdCompra.ToString()
				+ ", " + dFactor.ToString("0.000000", us) 
				+ ", " + dFactor1.ToString("0.000000", us)
				+ ", " + dTotal.ToString("0.000000", us);
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);
			MessageBox.Show("Costos Generados", "Información");
		}

		private void cdsImportacion_BeforeFill_1(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btArancel_Click(object sender, System.EventArgs e)
		{
			BorraErrores();
			string stNumIngreso = txtIngreso.Text.Trim();
			if (stNumIngreso.Length == 0)
			{
				errorProvider.SetError(txtIngreso, "Ingreso no puede estar vacio");
				return;
			}
			string stSelect = "Select idCompra From Compra Where Numero = ("
				+ stNumIngreso + ") And idTipoFactura In (9, 14, 26, 29)";
//			int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect, false);
//			if (IdCompra <= 0) 
//			{
//				errorProvider.SetError(txtIngreso, "Ingreso no Existe");
//				return;
//			}
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Compra.Numero} In [" 
				+ stNumIngreso + "] And {Compra.idTipoFactura} In [9, 14, 26, 29]";

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();

			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "Factor";
			discreteVal.Value = txtFactor.Value;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "pFormulario";
			discreteVal1.Value = txtFormulario.Text;
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);

			Reporte miRep = new Reporte("ImportacionArancel.Rpt", stFiltro);
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
			this.Cursor = Cursors.Default;
		}

		private void btArancel_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if ((e.Button == MouseButtons.Right))
			{
				this.Cursor = Cursors.WaitCursor;
				string stNumIngreso = txtIngreso.Text.Trim();
//				string stSelect = "Select idCompra From Compra Where Numero = ("
//					+ stNumIngreso + ") And idTipoFactura In (9, 14, 26, 29)";
//				int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect, false);
//				if (IdCompra <= 0) 
//				{
//					errorProvider.SetError(txtIngreso, "Ingreso no Existe");
//					return;
//				}
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramField = new ParameterField ();

				ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
				paramField.ParameterFieldName = "Factor";
				discreteVal.Value = txtFactor.Value;
				paramField.CurrentValues.Add (discreteVal);
				paramFields.Add (paramField);

				ParameterField paramField1 = new ParameterField ();
				ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
				paramField1.ParameterFieldName = "pFormulario";
				discreteVal1.Value = txtFormulario.Text;
				paramField1.CurrentValues.Add (discreteVal1);
				paramFields.Add (paramField1);

				string stFiltro = "{Compra.Numero} In [" 
					+ stNumIngreso + "] And {Compra.idTipoFactura} In [9, 14, 26, 29]";
				//				string text1 = "{AsientoDetalle.idCuenta} = " + this.cmbCuenta.Value.ToString();
				Reporte reporte1 = new Reporte("ImpArancelGlobal.rpt", stFiltro);
				reporte1.MdiParent = this.MdiParent;
				reporte1.crVista.ParameterFieldInfo = paramFields;
				reporte1.Show();
				this.Cursor = Cursors.Default;
			}
		}

		private void btPlanCompra_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			CompraImporImp miCompraImp = new CompraImporImp();
			miCompraImp.MdiParent = this.MdiParent;
			miCompraImp.Show();
			Cursor = Cursors.Default;
		}

		private void dsTipoImportacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTipoImportacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void CompraImportacion_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Guia Remision'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\";
			ultraGridExcelExporter1.Export(ultraGrid2, "Importacion.xls");
			if (DialogResult.Yes == MessageBox.Show("¿Desea Ver Archivo Generado?",
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "Importacion.xls");
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			if (barraDatoSQL1.bEditar) 
			{
				MessageBox.Show("Para generar debe estar grabado el registro", "Información");
				return;
			}

			string stArticulo = txtArticulo.Text;
			errorProvider.SetError(txtArticulo, "");
			if (stArticulo.Length == 0)
			{
				errorProvider.SetError(txtArticulo, "Ingrese Articulo");
				txtArticulo.Select();
				return;
			}
			string stExec = "Select Count(*) From Articulo Where Articulo = '"
				+ stArticulo + "'";
			int iCuenta = Funcion.iEscalarSQL(cdsCompra, stExec, true);
			if (iCuenta == 0)
			{
				errorProvider.SetError(txtArticulo, "No Existe este Articulo");
				txtArticulo.Select();
				return;
			}
			if (iCuenta > 1)
			{
				errorProvider.SetError(txtArticulo, "Existe mas de un Articulo que cumple ese patron");
				txtArticulo.Select();
				return;
			}

			// Existe ese articulo
			stExec = "Select idArticulo From Articulo Where Articulo = '"
				+ stArticulo + "'";
			int idArticulo = Funcion.iEscalarSQL(cdsCompra, stExec, true);

			stExec = "EXEC ImportacionArticulo " + txtIdImportacion.Value.ToString()
				+ ", " + idArticulo.ToString();
			Funcion.EjecutaSQL(cdsCompra, stExec, true);
			this.barraDatoSQL1.ProximoId(5);

		}

		private void btDauA_Click(object sender, System.EventArgs e)
		{
			string stNumIngreso = txtIngreso.Text.Trim();
			if (stNumIngreso.Length == 0)
			{
				errorProvider.SetError(txtIngreso, "Ingreso no puede estar vacio");
				return;
			}
			string stSelect = "Select idCompra From Compra Where Numero In ("
				+ stNumIngreso + ") And idTipoFactura In (9, 14, 26, 29)";
			int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect, false);
			if (IdCompra <= 0) 
			{
				errorProvider.SetError(txtIngreso, "Ingreso no Existe");
				return;
			}

			Cursor = Cursors.WaitCursor;
			stSelect = "Exec ArancelCodigoNum " + IdCompra.ToString();
			int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			string stCodigo = "Exec ArancelCodigo1 " + IdCompra.ToString();
			stCodigo = Funcion.sEscalarSQL(cdsCompra, stCodigo, true);
			if (stCodigo == null) stCodigo = "";
			string stFiltro = "{ArticuloArancel.Codigo} = '" + stCodigo.Trim() + "'"
				+ " And {Compra.IdCompra} = " + IdCompra.ToString();

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "NumAranceles";
			discreteVal.Value = iCuenta;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
			
			
			Reporte miRep = new Reporte("DUI_A.rpt", stFiltro);
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void btDauB_Click(object sender, System.EventArgs e)
		{
			string stNumIngreso = txtIngreso.Text.Trim();
			if (stNumIngreso.Length == 0)
			{
				errorProvider.SetError(txtIngreso, "Ingreso no puede estar vacio");
				return;
			}
			string stSelect = "Select idCompra From Compra Where Numero In ("
				+ stNumIngreso + ") And idTipoFactura In (9, 14, 26, 29)";
			int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect, false);
			if (IdCompra <= 0) 
			{
				errorProvider.SetError(txtIngreso, "Ingreso no Existe");
				return;
			}

			Cursor = Cursors.WaitCursor;
			string stCodigo = "Exec ArancelCodigo1 " + IdCompra.ToString();
			stCodigo = Funcion.sEscalarSQL(cdsCompra, stCodigo, true);
			if (stCodigo == null) stCodigo = "";
			string stFiltro = "{ArticuloArancel.Codigo} <> '" + stCodigo.Trim()
				+ "' And {Compra.Numero} In [" + stNumIngreso + "]";

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "pFormulario";
			discreteVal.Value = txtFormulario.Text;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			Reporte miRep = new Reporte("DUI_B.rpt", stFiltro);
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void ultraGrid2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid2.ActiveCell == null) return;
			try
			{
				if (ultraGrid2.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid2.ActiveCell.DroppedDown == false)
							ultraGrid2.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ultraGrid2.ActiveCell.DroppedDown) return;

						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid2.ActiveCell.DroppedDown) return;
			
						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid2.ActiveCell.DroppedDown == false)
							ultraGrid2.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//				MessageBox.Show(ex.Message, "Tecla");
			}

		}

		private void barraDatoSQL1_Load(object sender, System.EventArgs e)
		{
		
		}


	}
}
