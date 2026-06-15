using System;
using System.Globalization;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Depuracion.
	/// </summary>
	public class Depuracion : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraButton btNoCuadrado;
		private Infragistics.Win.Misc.UltraButton btNoExisteCuenta;
		private C1.Data.C1DataSet cdsAsiento;
		private Infragistics.Win.Misc.UltraButton btAsiento;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.Misc.UltraButton btComparaAsiento;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdCompara;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsCompara;
		private Infragistics.Win.Misc.UltraButton btBuscaAsiento;
		private Infragistics.Win.Misc.UltraButton btBuscaFactura;
		private C1.Data.C1DataSet cdsCompra;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private Infragistics.Win.Misc.UltraButton btAsiFact;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnTolerancia;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optCosto;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private Infragistics.Win.Misc.UltraButton btProyLinea;
		private Infragistics.Win.Misc.UltraButton btCentroLinea;
		private Infragistics.Win.Misc.UltraButton btCCFactura;
		private Infragistics.Win.Misc.UltraButton btProyFactura;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView7;
		private DevExpress.XtraGrid.GridControl ugContabilidad;
		private DevExpress.XtraGrid.Views.Grid.GridView ugvContabilidad;
		private DevExpress.XtraEditors.DateEdit cmbDesde;
		private DevExpress.XtraEditors.DateEdit cmbHasta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Depuracion()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPuntos");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodArt");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescArt", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Errores");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cantidad", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cantidad", 1, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Haber", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Haber", 13, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Asiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Haber");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Errores");
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ugContabilidad = new DevExpress.XtraGrid.GridControl();
			this.ugvContabilidad = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView7 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.btAsiento = new Infragistics.Win.Misc.UltraButton();
			this.btNoExisteCuenta = new Infragistics.Win.Misc.UltraButton();
			this.btNoCuadrado = new Infragistics.Win.Misc.UltraButton();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.optCosto = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.spnTolerancia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btAsiFact = new Infragistics.Win.Misc.UltraButton();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.btBuscaFactura = new Infragistics.Win.Misc.UltraButton();
			this.btBuscaAsiento = new Infragistics.Win.Misc.UltraButton();
			this.grdCompara = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsCompara = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btComparaAsiento = new Infragistics.Win.Misc.UltraButton();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btProyFactura = new Infragistics.Win.Misc.UltraButton();
			this.btCCFactura = new Infragistics.Win.Misc.UltraButton();
			this.btCentroLinea = new Infragistics.Win.Misc.UltraButton();
			this.btProyLinea = new Infragistics.Win.Misc.UltraButton();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cmbDesde = new DevExpress.XtraEditors.DateEdit();
			this.cmbHasta = new DevExpress.XtraEditors.DateEdit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugContabilidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugvContabilidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView7)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTolerancia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCompara)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCompara)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.ugContabilidad);
			this.ultraTabPageControl1.Controls.Add(this.btAsiento);
			this.ultraTabPageControl1.Controls.Add(this.btNoExisteCuenta);
			this.ultraTabPageControl1.Controls.Add(this.btNoCuadrado);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(516, 292);
			// 
			// ugContabilidad
			// 
			this.ugContabilidad.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// ugContabilidad.EmbeddedNavigator
			// 
			this.ugContabilidad.EmbeddedNavigator.Name = "";
			this.ugContabilidad.Location = new System.Drawing.Point(16, 40);
			this.ugContabilidad.MainView = this.ugvContabilidad;
			this.ugContabilidad.Name = "ugContabilidad";
			this.ugContabilidad.Size = new System.Drawing.Size(480, 240);
			this.ugContabilidad.TabIndex = 31;
			this.ugContabilidad.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																	this.ugvContabilidad,
																																																	this.gridView7});
			// 
			// ugvContabilidad
			// 
			this.ugvContabilidad.GridControl = this.ugContabilidad;
			this.ugvContabilidad.Name = "ugvContabilidad";
			this.ugvContabilidad.OptionsBehavior.Editable = false;
			this.ugvContabilidad.OptionsSelection.MultiSelect = true;
			this.ugvContabilidad.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.ugvContabilidad.OptionsView.ShowAutoFilterRow = true;
			this.ugvContabilidad.DoubleClick += new System.EventHandler(this.ugvContabilidad_DoubleClick);
			// 
			// gridView7
			// 
			this.gridView7.GridControl = this.ugContabilidad;
			this.gridView7.Name = "gridView7";
			// 
			// btAsiento
			// 
			this.btAsiento.Location = new System.Drawing.Point(344, 9);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(152, 24);
			this.btAsiento.TabIndex = 3;
			this.btAsiento.Text = "Mostrar Asiento";
			this.btAsiento.Visible = false;
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// btNoExisteCuenta
			// 
			this.btNoExisteCuenta.Location = new System.Drawing.Point(184, 9);
			this.btNoExisteCuenta.Name = "btNoExisteCuenta";
			this.btNoExisteCuenta.Size = new System.Drawing.Size(152, 24);
			this.btNoExisteCuenta.TabIndex = 2;
			this.btNoExisteCuenta.Text = "Cuentas no Asignadas";
			this.btNoExisteCuenta.Click += new System.EventHandler(this.btNoExisteCuenta_Click);
			this.btNoExisteCuenta.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btNoExisteCuenta_MouseDown);
			// 
			// btNoCuadrado
			// 
			this.btNoCuadrado.Location = new System.Drawing.Point(16, 9);
			this.btNoCuadrado.Name = "btNoCuadrado";
			this.btNoCuadrado.Size = new System.Drawing.Size(152, 24);
			this.btNoCuadrado.TabIndex = 1;
			this.btNoCuadrado.Text = "Asientos No Cuadrados";
			this.btNoCuadrado.Click += new System.EventHandler(this.btNoCuadrado_Click);
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.optCosto);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel3);
			this.ultraTabPageControl3.Controls.Add(this.spnTolerancia);
			this.ultraTabPageControl3.Controls.Add(this.btAsiFact);
			this.ultraTabPageControl3.Controls.Add(this.cmbCuenta);
			this.ultraTabPageControl3.Controls.Add(this.btBuscaFactura);
			this.ultraTabPageControl3.Controls.Add(this.btBuscaAsiento);
			this.ultraTabPageControl3.Controls.Add(this.grdCompara);
			this.ultraTabPageControl3.Controls.Add(this.btComparaAsiento);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(516, 292);
			// 
			// optCosto
			// 
			this.optCosto.CheckedIndex = 0;
			this.optCosto.ItemAppearance = appearance1;
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Precio";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Costo P.";
			this.optCosto.Items.Add(valueListItem1);
			this.optCosto.Items.Add(valueListItem2);
			this.optCosto.ItemSpacingVertical = 5;
			this.optCosto.Location = new System.Drawing.Point(24, 9);
			this.optCosto.Name = "optCosto";
			this.optCosto.Size = new System.Drawing.Size(96, 43);
			this.optCosto.TabIndex = 8;
			this.optCosto.Text = "Precio";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel3.Location = new System.Drawing.Point(232, 34);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(56, 15);
			this.ultraLabel3.TabIndex = 7;
			this.ultraLabel3.Text = "Tolerancia";
			// 
			// spnTolerancia
			// 
			this.spnTolerancia.Location = new System.Drawing.Point(296, 34);
			this.spnTolerancia.MinValue = 0.01;
			this.spnTolerancia.Name = "spnTolerancia";
			this.spnTolerancia.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnTolerancia.PromptChar = ' ';
			this.spnTolerancia.Size = new System.Drawing.Size(80, 22);
			this.spnTolerancia.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnTolerancia.TabIndex = 6;
			// 
			// btAsiFact
			// 
			this.btAsiFact.Location = new System.Drawing.Point(160, 9);
			this.btAsiFact.Name = "btAsiFact";
			this.btAsiFact.Size = new System.Drawing.Size(75, 24);
			this.btAsiFact.TabIndex = 5;
			this.btAsiFact.Text = "Asi -> Fac";
			this.btAsiFact.Click += new System.EventHandler(this.btAsiFact_Click);
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataMember = "Cuenta";
			this.cmbCuenta.DataSource = this.cdsAsiento;
			ultraGridColumn1.Header.VisiblePosition = 8;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 12;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 9;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 13;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 14;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 15;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 16;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 17;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 18;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn20.Header.VisiblePosition = 19;
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
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuenta.DisplayMember = "CodRapido";
			this.cmbCuenta.Location = new System.Drawing.Point(240, 9);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(136, 22);
			this.cmbCuenta.TabIndex = 4;
			this.cmbCuenta.ValueMember = "idCuenta";
			// 
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsCuenta";
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			this.cdsAsiento.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsiento_BeforeFill);
			// 
			// btBuscaFactura
			// 
			this.btBuscaFactura.Location = new System.Drawing.Point(472, 9);
			this.btBuscaFactura.Name = "btBuscaFactura";
			this.btBuscaFactura.Size = new System.Drawing.Size(24, 24);
			this.btBuscaFactura.TabIndex = 3;
			this.btBuscaFactura.Text = "F";
			this.btBuscaFactura.Click += new System.EventHandler(this.btBuscaFactura_Click);
			// 
			// btBuscaAsiento
			// 
			this.btBuscaAsiento.Location = new System.Drawing.Point(448, 9);
			this.btBuscaAsiento.Name = "btBuscaAsiento";
			this.btBuscaAsiento.Size = new System.Drawing.Size(24, 24);
			this.btBuscaAsiento.TabIndex = 2;
			this.btBuscaAsiento.Text = "A";
			this.btBuscaAsiento.Click += new System.EventHandler(this.btBuscaAsiento_Click);
			// 
			// grdCompara
			// 
			this.grdCompara.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdCompara.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdCompara.DataSource = this.udsCompara;
			appearance2.BackColor = System.Drawing.Color.White;
			this.grdCompara.DisplayLayout.Appearance = appearance2;
			ultraGridColumn21.Header.VisiblePosition = 9;
			ultraGridColumn21.Hidden = true;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance3;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "Precio - CostoP";
			ultraGridColumn22.Header.VisiblePosition = 0;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance4;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 1;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance5;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 3;
			ultraGridColumn26.Header.VisiblePosition = 4;
			ultraGridColumn27.Header.VisiblePosition = 5;
			ultraGridColumn28.Header.VisiblePosition = 6;
			ultraGridColumn29.Header.VisiblePosition = 10;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 11;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 12;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 13;
			ultraGridColumn32.Hidden = true;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance6;
			ultraGridColumn33.Header.VisiblePosition = 7;
			ultraGridColumn33.Hidden = true;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance7;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.Caption = "Diferencia";
			ultraGridColumn34.Header.VisiblePosition = 8;
			ultraGridColumn35.Header.VisiblePosition = 14;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35});
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand2.SummaryFooterCaption = "";
			this.grdCompara.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.grdCompara.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdCompara.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grdCompara.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.grdCompara.DisplayLayout.Override.CardAreaAppearance = appearance8;
			this.grdCompara.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.grdCompara.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 10F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdCompara.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.grdCompara.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCompara.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCompara.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.grdCompara.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grdCompara.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grdCompara.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.grdCompara.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.grdCompara.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.grdCompara.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.grdCompara.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.grdCompara.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdCompara.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdCompara.Location = new System.Drawing.Point(16, 60);
			this.grdCompara.Name = "grdCompara";
			this.grdCompara.Size = new System.Drawing.Size(487, 207);
			this.grdCompara.TabIndex = 1;
			this.grdCompara.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdCompara_KeyDown);
			// 
			// udsCompara
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.Double);
			ultraDataColumn3.DataType = typeof(System.Double);
			ultraDataColumn4.DataType = typeof(System.Double);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.Double);
			ultraDataColumn14.DataType = typeof(System.Double);
			this.udsCompara.Band.Columns.AddRange(new object[] {
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
			// btComparaAsiento
			// 
			this.btComparaAsiento.Location = new System.Drawing.Point(160, 34);
			this.btComparaAsiento.Name = "btComparaAsiento";
			this.btComparaAsiento.Size = new System.Drawing.Size(72, 25);
			this.btComparaAsiento.TabIndex = 0;
			this.btComparaAsiento.Text = "Fac -> Asi";
			this.btComparaAsiento.Click += new System.EventHandler(this.btComparaAsiento_Click);
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.btProyFactura);
			this.ultraTabPageControl4.Controls.Add(this.btCCFactura);
			this.ultraTabPageControl4.Controls.Add(this.btCentroLinea);
			this.ultraTabPageControl4.Controls.Add(this.btProyLinea);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(516, 292);
			// 
			// btProyFactura
			// 
			this.btProyFactura.Location = new System.Drawing.Point(32, 184);
			this.btProyFactura.Name = "btProyFactura";
			this.btProyFactura.Size = new System.Drawing.Size(208, 25);
			this.btProyFactura.TabIndex = 3;
			this.btProyFactura.Text = "Proyecto a Línea de Factura";
			this.btProyFactura.Click += new System.EventHandler(this.btProyFactura_Click);
			// 
			// btCCFactura
			// 
			this.btCCFactura.Location = new System.Drawing.Point(32, 144);
			this.btCCFactura.Name = "btCCFactura";
			this.btCCFactura.Size = new System.Drawing.Size(208, 25);
			this.btCCFactura.TabIndex = 2;
			this.btCCFactura.Text = "Centro de Costo a Línea de Factura";
			this.btCCFactura.Click += new System.EventHandler(this.btCCFactura_Click);
			// 
			// btCentroLinea
			// 
			this.btCentroLinea.Location = new System.Drawing.Point(32, 17);
			this.btCentroLinea.Name = "btCentroLinea";
			this.btCentroLinea.Size = new System.Drawing.Size(208, 25);
			this.btCentroLinea.TabIndex = 1;
			this.btCentroLinea.Text = "Centro de Costo a Línea de Asiento";
			this.btCentroLinea.Click += new System.EventHandler(this.btCentroLinea_Click);
			// 
			// btProyLinea
			// 
			this.btProyLinea.Location = new System.Drawing.Point(32, 52);
			this.btProyLinea.Name = "btProyLinea";
			this.btProyLinea.Size = new System.Drawing.Size(208, 24);
			this.btProyLinea.TabIndex = 0;
			this.btProyLinea.Text = "Proyecto a Línea de Asiento";
			this.btProyLinea.Click += new System.EventHandler(this.btProyLinea_Click);
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl3);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl4);
			this.ultraTabControl1.Location = new System.Drawing.Point(16, 52);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(520, 318);
			this.ultraTabControl1.TabIndex = 0;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Contabilidad";
			ultraTab2.TabPage = this.ultraTabPageControl3;
			ultraTab2.Text = "Comparación";
			ultraTab2.Visible = false;
			ultraTab3.TabPage = this.ultraTabPageControl4;
			ultraTab3.Text = "Corrección";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2,
																																															 ultraTab3});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(516, 292);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(256, 17);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(33, 15);
			this.ultraLabel1.TabIndex = 3;
			this.ultraLabel1.Text = "Hasta";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(48, 17);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(35, 15);
			this.ultraLabel2.TabIndex = 4;
			this.ultraLabel2.Text = "Desde";
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompraTabla";
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill_1);
			// 
			// cmbDesde
			// 
			this.cmbDesde.EditValue = null;
			this.cmbDesde.Location = new System.Drawing.Point(96, 16);
			this.cmbDesde.Name = "cmbDesde";
			// 
			// cmbDesde.Properties
			// 
			this.cmbDesde.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbDesde.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.cmbDesde.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbDesde.Properties.EditFormat.FormatString = "dd/MMM/yyyy";
			this.cmbDesde.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbDesde.Properties.Mask.EditMask = "dd/MMM/yyyy";
			this.cmbDesde.Properties.NullDate = "";
			this.cmbDesde.Size = new System.Drawing.Size(104, 20);
			this.cmbDesde.TabIndex = 65;
			// 
			// cmbHasta
			// 
			this.cmbHasta.EditValue = null;
			this.cmbHasta.Location = new System.Drawing.Point(304, 16);
			this.cmbHasta.Name = "cmbHasta";
			// 
			// cmbHasta.Properties
			// 
			this.cmbHasta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbHasta.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbHasta.Properties.EditFormat.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbHasta.Properties.Mask.EditMask = "dd/MMM/yyyy";
			this.cmbHasta.Properties.NullDate = "";
			this.cmbHasta.Size = new System.Drawing.Size(104, 20);
			this.cmbHasta.TabIndex = 66;
			// 
			// Depuracion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(552, 385);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.ultraTabControl1);
			this.Name = "Depuracion";
			this.Text = "Depuracion";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Depuracion_Closing);
			this.Load += new System.EventHandler(this.Depuracion_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugContabilidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugvContabilidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView7)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTolerancia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCompara)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCompara)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsAsiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsAsiento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void btNoCuadrado_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.EditValue;
			DateTime dtHasta = (DateTime) cmbHasta.EditValue;
			string stExec = String.Format("Exec DepuraAsiento '{0:yyyyMMdd}', '{1:yyyyMMdd}'",
				dtDesde, dtHasta);
			ugContabilidad.DataSource = Funcion.dvProcedimiento(cdsAsiento, stExec);
			GrillaFormato();
			GrillaFormato();

			Cursor = Cursors.Default;
		}

		public void GrillaFormato()
		{
			DevExpress.XtraGrid.Views.Base.ColumnView View = 
				ugContabilidad.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			foreach (DevExpress.XtraGrid.Columns.GridColumn col in View.Columns)
			{
				if (col.FieldName.StartsWith("id"))
					col.Visible = false;
				if (col.FieldName.StartsWith("Total") || col.FieldName.StartsWith("Diferencia")
					|| col.FieldName.StartsWith("Debe") || col.FieldName.StartsWith("Haber"))
				{
					col.DisplayFormat.FormatString = "n2";
					col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
				}
				if (col.FieldName.StartsWith("Fecha") || col.FieldName.StartsWith("Vencimiento"))
				{
					col.DisplayFormat.FormatString = "dd/MMM/yyyy";
					col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
				}			
			}
		}

		private void AbreAsiento()
		{
			Cursor = Cursors.WaitCursor;
			int iFila = ugvContabilidad.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Asiento", "Información");
				return;
			}
		
			int IdAsi = (int) ugvContabilidad.GetRowCellValue(iFila, "idAsiento");
			Asiento miAsiento = new Asiento(IdAsi);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void btNoExisteCuenta_Click(object sender, System.EventArgs e)
		{
      DateTime dtDesde = (DateTime) cmbDesde.EditValue;
			DateTime dtHasta = (DateTime) cmbHasta.EditValue;

			string stExec = String.Format("Exec DepuraSinCuenta '{0:yyyyMMdd}', '{1:yyyyMMdd}'",
				dtDesde, dtHasta);
			ugContabilidad.DataSource = Funcion.dvProcedimiento(cdsAsiento, stExec);
			GrillaFormato();
		}

		private void btNoExisteCuenta_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (DialogResult.No == MessageBox.Show("¿Desea Borrar detalles de asientos que no tiene asientos?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				return;
			string stDelete = "Delete asientodetalle from AsientoDetalle "
				+ "left outer join asiento on asiento.idasiento = asientodetalle.idasiento "
				+ "where isnull(Asiento.idAsiento,-1) = -1";
			Funcion.EjecutaSQL(cdsAsiento, stDelete);
			MessageBox.Show("Detalles Borrados", "Información");
		}
		CultureInfo us = new CultureInfo("en-US");

		private void btComparaAsiento_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.EditValue;
			DateTime dtHasta = (DateTime) cmbHasta.EditValue;
			dtHasta = dtHasta.AddDays(1);

			string stSelCompra = "Select DetCompra.Cantidad, DetCompra.Precio, "
				+ "Compra.idAsiento, DetCompra.idDetCompra, DetCompra.idArticulo, "
				+ "DetCompra.CostoPromedio "
				+ "from DetCompra inner join Articulo "
				+ "on Articulo.idArticulo = DetCompra.idArticulo "
				+ "inner join Compra "
				+ "on Compra.idCompra = DetCompra.idCompra "
				+ "where Compra.Fecha >= '" + dtDesde.ToString("yyyyMMdd")
				+ "' And Compra.Fecha < '" + dtHasta.ToString("yyyyMMdd") + "' "
				+ "And Compra.idTipoFactura = 4 And Articulo.idTipoGrupo = 1 ";

			cdsCompra.Schema.Connections[0].Open();
			SqlDataReader miReader;
			C1.Data.SchemaObjects.Connection miConeccion = cdsCompra.Schema.Connections[0];
			SqlCommand cmdCant = new SqlCommand (stSelCompra
				, (SqlConnection) miConeccion.DbConnection);
			miReader = cmdCant.ExecuteReader();

			int iNumFilas = 0;
			int iMax =	10000;
			int[] IdAsi = new int[iMax];
			int[] IdArt = new int[iMax];
			int[] IdDetCompra = new int[iMax];
			double[] dCosto = new double[iMax];
			double[] dPromedio = new double[iMax];
			try
			{
				while (miReader.Read())
				{
					double dCantidad = miReader.GetDouble(0);
					double dPrecio = miReader.GetDouble(1);
					dCosto[iNumFilas] = dCantidad * dPrecio;
					IdAsi[iNumFilas] = miReader.GetInt32(2);
					IdDetCompra[iNumFilas] = miReader.GetInt32(3);
					IdArt[iNumFilas] = miReader.GetInt32(4);
					dPromedio[iNumFilas] = miReader.GetDouble(5) * dCantidad;
					iNumFilas ++;
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Llenar Grilla");
			}
			miReader.Close();
			cdsCompra.Schema.Connections[0].Close();

			int iTotalLineas = iNumFilas;
			iNumFilas = -1;
			udsCompara.Rows.Clear();
			for (int i=0; i< iTotalLineas; i++)
			{
				string stError = "";
				bool bError = false;
				string stSelAsiento = "Select Count(*) From AsientoDetalle "
					+ "Where AsientoDetalle.idAsiento = " + IdAsi[i].ToString()
					+ " And ROUND(Debe + Haber, 2) = " + dCosto[i].ToString("0.00", us);
				int iCuenta = Funcion.iEscalarSQL(cdsAsiento, stSelAsiento);
				double dMargen = .1;
				if (iCuenta == 0 || (dCosto[i] - dPromedio[i]) > dMargen) 
				{
					if ((dCosto[i] - dPromedio[i]) > dMargen)
						stError = "Promedio Diferente de Precio";
					else
            stError = "Valor no Existe en Asiento";
					bError = true;
				}
				else
				{
				
					stSelAsiento = "select idCuentas2 from articulo "
						+ " inner join ArticuloGrupo"
						+ " On Articulo.idGrupoarticulo = ArticuloGrupo.idGrupoarticulo "
						+ " Where Articulo.idArticulo = " + IdArt[i].ToString();
					int IdCuentaGrupo = Funcion.iEscalarSQL(cdsCompra, stSelAsiento);

					stSelAsiento = "Select Count(*) From AsientoDetalle "
						+ "Where AsientoDetalle.idAsiento = " + IdAsi[i].ToString()
						+ " And ROUND(Debe + Haber, 2) = " + dCosto[i].ToString("0.00", us)
						+ " And AsientoDetalle.idCuenta = " + IdCuentaGrupo.ToString();
					iCuenta = Funcion.iEscalarSQL(cdsAsiento, stSelAsiento);
					if (iCuenta == 0)
					{
						stError += "Codigo Contable no Existe";
						bError = true;
					}
				}

				if (bError)
				{
					// Identifica IdCompra
					string stSelect = "SELECT IdCompra FROM DetCompra "
						+ "Where idDetCompra = "+ IdDetCompra[i].ToString();
					int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect);
					iNumFilas ++;
					udsCompara.Rows.Add();
					udsCompara.Rows[iNumFilas]["Errores"]=stError;
					udsCompara.Rows[iNumFilas]["idAsiento"]=IdAsi[i];
					udsCompara.Rows[iNumFilas]["idCompra"]=IdCompra;
					udsCompara.Rows[iNumFilas]["Costo"]=dCosto[i];
					stSelect = "SELECT Numero FROM Compra "
						+ "Where idCompra = "+ IdCompra.ToString();
					string stCod = Funcion.sEscalarSQL(cdsCompra, stSelect);
					udsCompara.Rows[iNumFilas]["Numero"]=stCod;
					stSelect = "SELECT Codigo FROM Articulo Where idArticulo = "
						+ IdArt[i].ToString();
					stCod = Funcion.sEscalarSQL(cdsCompra, stSelect);
					udsCompara.Rows[iNumFilas]["CodArt"]=stCod;
					stSelect = "SELECT Articulo FROM Articulo Where idArticulo = "
						+ IdArt[i].ToString();
					stCod = Funcion.sEscalarSQL(cdsCompra, stSelect);
					udsCompara.Rows[iNumFilas]["DescArt"]=stCod;
					udsCompara.Rows[iNumFilas]["Haber"]=dCosto[i];
					udsCompara.Rows[iNumFilas]["Cantidad"]=dCosto[i] - dPromedio[i];

					if (IdAsi[i] > 0)
					{
						stSelect = "SELECT CodAsiento FROM Asiento Where idAsiento = "
							+ IdAsi[i].ToString();
						stCod = Funcion.sEscalarSQL(cdsCompra, stSelect);
						udsCompara.Rows[iNumFilas]["Asiento"]=stCod;
						stSelect = "SELECT Total FROM Asiento Where idAsiento = "
							+ IdAsi[i].ToString();
						double dTotal = Funcion.dEscalarSQL(cdsCompra, stSelect);
//						udsCompara.Rows[iNumFilas]["Debe"]=dTotal;
						stSelect = "SELECT Total FROM Compra Where idCompra = "
							+ IdCompra.ToString();
						double dTotal1 = Funcion.dEscalarSQL(cdsCompra, stSelect);
						udsCompara.Rows[iNumFilas]["Haber"]=dTotal - dTotal1;
//						stSelect = "SELECT Fecha FROM Asiento Where idAsiento = "
//							+ IdAsi[i].ToString();
//						stCod = Funcion.sEscalarSQL(cdsCompra, stSelect);
//						udsCompara.Rows[iNumFilas]["FechaA"]=stCod;
					}
				}
			}
			Cursor = Cursors.Default;
			MessageBox.Show("Errores Generados");
		}


		private void btAsiento_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btBuscaAsiento_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			int IdAsi = 0;
			try
			{
				IdAsi = (int) grdCompara.ActiveRow.Cells["idAsiento"].Value;
			}
			catch
			{
				MessageBox.Show("Escoja Asiento para mostrar", "Información");
				return;
			}
			Asiento miAsiento = new Asiento(IdAsi);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void btBuscaFactura_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			int idCompra = (int) grdCompara.ActiveRow.Cells["idCompra"].Value;
			if (idCompra == 0)
			{
				MessageBox.Show("No existe Factura", "Información");
				return;
			}
			int IdTipo = Funcion.iEscalarSQL(cdsCompra, "Select IdTipoFactura"
				+ " From Compra Where idCompra = " + idCompra.ToString());
//			Compra miCompra = new Compra(IdTipo, idCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
		}

		private void cdsCompra_BeforeFill_1(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void btAsiFact_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.EditValue;
			DateTime dtHasta = (DateTime) cmbHasta.EditValue;
			dtHasta = dtHasta.AddDays(1);

			string stPrecio = "Precio";
			if ((int) optCosto.Value == 1) stPrecio = "CostoPromedio";
			if (cmbCuenta.Value == null)
			{
				MessageBox.Show("Ingrese Cuenta");
				return;
			}
			int IdCuentaContable = (int) cmbCuenta.Value;
			string stSelCompra = "Select Abs(Debe - Haber), "
				+ "idCuenta, Asiento.idAsiento "
				+ "From AsientoDetalle "
				+ "Inner Join Asiento "
				+ "On Asiento.idAsiento = AsientoDetalle.idAsiento "
				+ "Where idCuenta = " + IdCuentaContable.ToString()
				+ " And Asiento.Fecha >= '" + dtDesde.ToString("yyyyMMdd")
				+ "' And Asiento.Fecha < '" + dtHasta.ToString("yyyyMMdd") + "'";
//				+ " And Debe > 0 ";

			cdsAsiento.Schema.Connections[0].Open();
			SqlDataReader miReader;
			C1.Data.SchemaObjects.Connection miConeccion = cdsAsiento.Schema.Connections[0];
			SqlCommand cmdCant = new SqlCommand (stSelCompra
				, (SqlConnection) miConeccion.DbConnection);
			miReader = cmdCant.ExecuteReader();

			int iNumFilas = 0;
			int iMax =	10000;
			int[] IdCuenta = new int[iMax];
			int[] IdAsi = new int[iMax];
			double[] dCosto = new double[iMax];
			try
			{
				while (miReader.Read())
				{
					dCosto[iNumFilas] = miReader.GetDouble(0);
					IdCuenta[iNumFilas] = miReader.GetInt32(1);
					IdAsi[iNumFilas] = miReader.GetInt32(2);
					iNumFilas ++;
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Llenar Grilla");
			}
			miReader.Close();
			cdsAsiento.Schema.Connections[0].Close();

			int iTotalLineas = iNumFilas;
			iNumFilas = -1;
			double dMargen = (double) spnTolerancia.Value;
			udsCompara.Rows.Clear();
			for (int i=0; i< iTotalLineas; i++)
			{
				string stError = "";
				bool bError = false;
				string stSelAsiento = "select count(*) From DetCompra "
					+ "inner Join Compra "
					+ "on Compra.idCompra = DetCompra.idCompra "
					+ "Where Compra.idAsiento = " + IdAsi[i].ToString()
					+ " And Abs(ROUND(Cantidad * " + stPrecio	
					+ ",2) - "+ dCosto[i].ToString("0.00", us)
					+ ") <  " + dMargen.ToString("0.00", us);
				int iCuenta = Funcion.iEscalarSQL(cdsAsiento, stSelAsiento);
				if (iCuenta == 0) // || (dCosto[i] - dPromedio[i]) > dMargen) 
				{
					stError = "Valor no Existe en Compra";
					bError = true;
				}

				if (bError)
				{
					// Identifica IdCompra
					string stSelect = "SELECT IdCompra FROM Compra "
						+ "Where idAsiento = "+ IdAsi[i].ToString();
					int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect, false);
					iNumFilas ++;
					udsCompara.Rows.Add();
					udsCompara.Rows[iNumFilas]["Errores"]=stError;
					udsCompara.Rows[iNumFilas]["idAsiento"]=IdAsi[i];
					udsCompara.Rows[iNumFilas]["idCompra"]=IdCompra;
					udsCompara.Rows[iNumFilas]["Costo"]=dCosto[i];
					stSelect = "SELECT Numero FROM Compra "
						+ "Where idCompra = "+ IdCompra.ToString();
					string stCod = Funcion.sEscalarSQL(cdsCompra, stSelect);
					udsCompara.Rows[iNumFilas]["Numero"]=stCod;
					stSelect = "SELECT CodAsiento FROM Asiento Where idAsiento = "
						+ IdAsi[i].ToString();
					stCod = Funcion.sEscalarSQL(cdsCompra, stSelect);
					udsCompara.Rows[iNumFilas]["Asiento"]=stCod;

					udsCompara.Rows[iNumFilas]["Haber"]=dCosto[i];
//					udsCompara.Rows[iNumFilas]["Cantidad"]=dCosto[i] - dPromedio;

					if (IdCompra > 0)
					{
						stSelect = "SELECT Total FROM Asiento Where idAsiento = "
							+ IdAsi[i].ToString();
						double dTotal = Funcion.dEscalarSQL(cdsCompra, stSelect);
						stSelect = "SELECT Total FROM Compra Where idCompra = "
							+ IdCompra.ToString();
						double dTotal1 = Funcion.dEscalarSQL(cdsCompra, stSelect);
						udsCompara.Rows[iNumFilas]["Haber"]=dTotal - dTotal1;
					}
				}
			}
			Cursor = Cursors.Default;
			MessageBox.Show("Errores Generados");
		}

		private void btCentroLinea_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Reemplazar los Centros de Costo en las líneas con los "
				+ "\nCentros definidos en el asiento cuando la linea esta vacia?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stSelect = "Update AsientoDetalle Set AsientoDetalle.idProyecto = Asiento.idProyecto "
				+ "From AsientoDetalle left Join Asiento On Asiento.idAsiento = AsientoDetalle.idAsiento "
				+ "where Asiento.idProyecto > 0 And IsNull(AsientoDetalle.idProyecto, 0) = 0 ";
			Funcion.EjecutaSQL(cdsAsiento, stSelect, true);
			MessageBox.Show("Reemplazo Realizado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btProyLinea_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Reemplazar los Proyectos en las líneas con los "
				+ "\nProyectos definidos en el asiento cuando la línea esta vacia?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stSelect = "Update AsientoDetalle Set AsientoDetalle.idSubProyecto = Asiento.idSubProyecto "
				+ "From AsientoDetalle left Join Asiento On Asiento.idAsiento = AsientoDetalle.idAsiento "
				+ "Where Asiento.idSubProyecto > 0 And IsNull(AsientoDetalle.idSubProyecto, 0) = 0 ";
			Funcion.EjecutaSQL(cdsAsiento, stSelect, true);
			MessageBox.Show("Reemplazo Realizado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void Depuracion_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Depuracion'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
			cmbHasta.EditValue = DateTime.Today;
			cmbDesde.EditValue = new DateTime(DateTime.Today.Year, 1, 1);
		}

		private void Depuracion_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Depuracion'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
		}

		private void grdCompara_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (grdCompara.ActiveCell == null) return;
			try
			{
				if (grdCompara.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(grdCompara.ActiveCell.DroppedDown == false)
							grdCompara.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (grdCompara.ActiveCell.DroppedDown) return;

						grdCompara.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdCompara.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						grdCompara.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						grdCompara.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdCompara.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						grdCompara.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (grdCompara.ActiveCell.DroppedDown) return;
			
						grdCompara.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdCompara.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						grdCompara.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						grdCompara.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdCompara.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						grdCompara.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(grdCompara.ActiveCell.DroppedDown == false)
							grdCompara.ActiveCell.DroppedDown = true;
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

		private void btCCFactura_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Reemplazar los Centros de Costo en las líneas con los "
				+ "\nCentros definidos en la Factura cuando la línea esta vacia?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stSelect = "Update DetCompra Set DetCompra.idProyecto = Compra.idProyecto "
				+ "	From DetCompra left Join Compra On Compra.idCompra = DetCompra.idCompra "
				+ " where Compra.idProyecto > 0 And IsNull(DetCompra.idProyecto, 0) = 0 ";

			Funcion.EjecutaSQL(cdsAsiento, stSelect, true);
			MessageBox.Show("Reemplazo Realizado", "Informaciónn", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btProyFactura_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Reemplazar los SubProyectos en las líneas con los "
				+ "\nProyectos definidos en la Factura cuando la línea esta vacia?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stSelect = "Update DetCompra Set DetCompra.idSubProyecto = Compra.idSubProyecto "
				+ "	From DetCompra left Join Compra On Compra.idCompra = DetCompra.idCompra "
				+ " where Compra.idSubProyecto > 0 And IsNull(DetCompra.idSubProyecto, 0) = 0 ";
			Funcion.EjecutaSQL(cdsAsiento, stSelect, true);
			MessageBox.Show("Reemplazo Realizado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void ugvContabilidad_DoubleClick(object sender, System.EventArgs e)
		{
			AbreAsiento();
		}
	}
}
