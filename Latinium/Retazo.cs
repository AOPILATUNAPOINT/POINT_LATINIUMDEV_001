using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Retazo.
	/// </summary>
	public class Retazo : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLargo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAncho;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.Misc.UltraLabel Articulo;
		private Infragistics.Win.Misc.UltraButton btIngreso;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFecha;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataView cdvCodArt;
		private C1.Data.C1DataSet cdsRetazo;
		private C1.Data.C1DataView cdvRetazoI;
		private C1.Data.C1DataView cdvRetazoE;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtLote;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticulo1;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbArticulo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnchoMin;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbAncho;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnchoMax;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArtFiltro;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAltoMax;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbAlto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAltoMin;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor CmbFechaMin;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor CmbFechaMax;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbFechaFiltro;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private Infragistics.Win.Misc.UltraButton btReporte;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkArtFiltro;
		private Infragistics.Win.Misc.UltraButton btSalida;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Retazo()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem14 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem15 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Retazo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo", -1, "cmbArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Salida");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.btIngreso = new Infragistics.Win.Misc.UltraButton();
			this.cmbArticulo1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.label4 = new System.Windows.Forms.Label();
			this.Articulo = new Infragistics.Win.Misc.UltraLabel();
			this.txtAncho = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtLargo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtLote = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btSalida = new Infragistics.Win.Misc.UltraButton();
			this.btReporte = new Infragistics.Win.Misc.UltraButton();
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbFechaFiltro = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.CmbFechaMax = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.CmbFechaMin = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.txtAltoMax = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbAlto = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.txtAltoMin = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbArtFiltro = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.txtAnchoMax = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbAncho = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.txtAnchoMin = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkArtFiltro = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbArticulo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdvRetazoI = new C1.Data.C1DataView();
			this.cdsRetazo = new C1.Data.C1DataSet();
			this.cdvCodArt = new C1.Data.C1DataView();
			this.cdvRetazoE = new C1.Data.C1DataView();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAncho)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbFechaMax)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbFechaMin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAltoMax)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAlto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAltoMin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArtFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnchoMax)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAncho)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnchoMin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetazoI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetazo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetazoE)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.label3);
			this.ultraTabPageControl1.Controls.Add(this.cmbFecha);
			this.ultraTabPageControl1.Controls.Add(this.btIngreso);
			this.ultraTabPageControl1.Controls.Add(this.cmbArticulo1);
			this.ultraTabPageControl1.Controls.Add(this.label4);
			this.ultraTabPageControl1.Controls.Add(this.Articulo);
			this.ultraTabPageControl1.Controls.Add(this.txtAncho);
			this.ultraTabPageControl1.Controls.Add(this.label2);
			this.ultraTabPageControl1.Controls.Add(this.txtLargo);
			this.ultraTabPageControl1.Controls.Add(this.label1);
			this.ultraTabPageControl1.Controls.Add(this.txtLote);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(588, 310);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(16, 64);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(36, 16);
			this.label3.TabIndex = 8;
			this.label3.Text = "Fecha";
			// 
			// cmbFecha
			// 
			this.cmbFecha.DateTime = new System.DateTime(2006, 3, 31, 0, 0, 0, 0);
			this.cmbFecha.Location = new System.Drawing.Point(64, 64);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.Size = new System.Drawing.Size(96, 21);
			this.cmbFecha.TabIndex = 7;
			this.cmbFecha.Value = new System.DateTime(2006, 3, 31, 0, 0, 0, 0);
			// 
			// btIngreso
			// 
			this.btIngreso.Location = new System.Drawing.Point(512, 64);
			this.btIngreso.Name = "btIngreso";
			this.btIngreso.Size = new System.Drawing.Size(56, 24);
			this.btIngreso.TabIndex = 6;
			this.btIngreso.Text = "Ingreso";
			this.btIngreso.Click += new System.EventHandler(this.btIngreso_Click);
			// 
			// cmbArticulo1
			// 
			this.cmbArticulo1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticulo1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo1.DataSource = this.cdvArticulo;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Width = 248;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbArticulo1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbArticulo1.DisplayMember = "Articulo";
			this.cmbArticulo1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArticulo1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArticulo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticulo1.Location = new System.Drawing.Point(336, 32);
			this.cmbArticulo1.Name = "cmbArticulo1";
			this.cmbArticulo1.Size = new System.Drawing.Size(232, 21);
			this.cmbArticulo1.TabIndex = 5;
			this.cmbArticulo1.ValueMember = "idArticulo";
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(184, 64);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(26, 16);
			this.label4.TabIndex = 12;
			this.label4.Text = "Lote";
			// 
			// Articulo
			// 
			this.Articulo.AutoSize = true;
			this.Articulo.BackColor = System.Drawing.Color.Transparent;
			this.Articulo.Location = new System.Drawing.Point(280, 32);
			this.Articulo.Name = "Articulo";
			this.Articulo.Size = new System.Drawing.Size(42, 14);
			this.Articulo.TabIndex = 4;
			this.Articulo.Text = "Artículo";
			// 
			// txtAncho
			// 
			this.txtAncho.Location = new System.Drawing.Point(192, 32);
			this.txtAncho.Name = "txtAncho";
			this.txtAncho.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAncho.PromptChar = ' ';
			this.txtAncho.Size = new System.Drawing.Size(80, 21);
			this.txtAncho.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAncho.TabIndex = 3;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(152, 32);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(36, 16);
			this.label2.TabIndex = 2;
			this.label2.Text = "Ancho";
			// 
			// txtLargo
			// 
			this.txtLargo.Location = new System.Drawing.Point(64, 32);
			this.txtLargo.Name = "txtLargo";
			this.txtLargo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtLargo.PromptChar = ' ';
			this.txtLargo.Size = new System.Drawing.Size(80, 21);
			this.txtLargo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtLargo.TabIndex = 1;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 32);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Largo";
			// 
			// txtLote
			// 
			this.txtLote.Location = new System.Drawing.Point(216, 64);
			this.txtLote.Name = "txtLote";
			this.txtLote.Size = new System.Drawing.Size(104, 21);
			this.txtLote.TabIndex = 11;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.btSalida);
			this.ultraTabPageControl2.Controls.Add(this.btReporte);
			this.ultraTabPageControl2.Controls.Add(this.btFiltro);
			this.ultraTabPageControl2.Controls.Add(this.label7);
			this.ultraTabPageControl2.Controls.Add(this.cmbFechaFiltro);
			this.ultraTabPageControl2.Controls.Add(this.CmbFechaMax);
			this.ultraTabPageControl2.Controls.Add(this.CmbFechaMin);
			this.ultraTabPageControl2.Controls.Add(this.label6);
			this.ultraTabPageControl2.Controls.Add(this.txtAltoMax);
			this.ultraTabPageControl2.Controls.Add(this.cmbAlto);
			this.ultraTabPageControl2.Controls.Add(this.txtAltoMin);
			this.ultraTabPageControl2.Controls.Add(this.cmbArtFiltro);
			this.ultraTabPageControl2.Controls.Add(this.label5);
			this.ultraTabPageControl2.Controls.Add(this.txtAnchoMax);
			this.ultraTabPageControl2.Controls.Add(this.cmbAncho);
			this.ultraTabPageControl2.Controls.Add(this.txtAnchoMin);
			this.ultraTabPageControl2.Controls.Add(this.chkArtFiltro);
			this.ultraTabPageControl2.Controls.Add(this.cmbArticulo);
			this.ultraTabPageControl2.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(588, 310);
			// 
			// btSalida
			// 
			this.btSalida.Location = new System.Drawing.Point(488, 72);
			this.btSalida.Name = "btSalida";
			this.btSalida.TabIndex = 30;
			this.btSalida.Text = "Salida";
			this.btSalida.Click += new System.EventHandler(this.btSalida_Click);
			// 
			// btReporte
			// 
			this.btReporte.Location = new System.Drawing.Point(488, 40);
			this.btReporte.Name = "btReporte";
			this.btReporte.TabIndex = 29;
			this.btReporte.Text = "Reporte";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// btFiltro
			// 
			this.btFiltro.Location = new System.Drawing.Point(488, 8);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.TabIndex = 28;
			this.btFiltro.Text = "Filtro";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(24, 80);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(36, 16);
			this.label7.TabIndex = 27;
			this.label7.Text = "Fecha";
			// 
			// cmbFechaFiltro
			// 
			this.cmbFechaFiltro.AutoComplete = true;
			this.cmbFechaFiltro.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Mayor";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Igual";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Menor";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Todo";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Entre";
			this.cmbFechaFiltro.Items.Add(valueListItem1);
			this.cmbFechaFiltro.Items.Add(valueListItem2);
			this.cmbFechaFiltro.Items.Add(valueListItem3);
			this.cmbFechaFiltro.Items.Add(valueListItem4);
			this.cmbFechaFiltro.Items.Add(valueListItem5);
			this.cmbFechaFiltro.Location = new System.Drawing.Point(72, 80);
			this.cmbFechaFiltro.Name = "cmbFechaFiltro";
			this.cmbFechaFiltro.Size = new System.Drawing.Size(96, 21);
			this.cmbFechaFiltro.TabIndex = 26;
			this.cmbFechaFiltro.Text = "Mayor";
			this.cmbFechaFiltro.ValueChanged += new System.EventHandler(this.cmbFechaFiltro_ValueChanged);
			// 
			// CmbFechaMax
			// 
			this.CmbFechaMax.Location = new System.Drawing.Point(304, 80);
			this.CmbFechaMax.Name = "CmbFechaMax";
			this.CmbFechaMax.Size = new System.Drawing.Size(96, 21);
			this.CmbFechaMax.TabIndex = 25;
			this.CmbFechaMax.Visible = false;
			// 
			// CmbFechaMin
			// 
			this.CmbFechaMin.FormatString = "dd/MMM/yyyy";
			this.CmbFechaMin.Location = new System.Drawing.Point(184, 80);
			this.CmbFechaMin.Name = "CmbFechaMin";
			this.CmbFechaMin.Size = new System.Drawing.Size(96, 21);
			this.CmbFechaMin.TabIndex = 24;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(24, 56);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(24, 16);
			this.label6.TabIndex = 23;
			this.label6.Text = "Alto";
			// 
			// txtAltoMax
			// 
			this.txtAltoMax.Location = new System.Drawing.Point(304, 56);
			this.txtAltoMax.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtAltoMax.Name = "txtAltoMax";
			this.txtAltoMax.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAltoMax.PromptChar = ' ';
			this.txtAltoMax.Size = new System.Drawing.Size(96, 21);
			this.txtAltoMax.TabIndex = 22;
			this.txtAltoMax.Visible = false;
			// 
			// cmbAlto
			// 
			this.cmbAlto.AutoComplete = true;
			this.cmbAlto.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem6.DataValue = 0;
			valueListItem6.DisplayText = "Mayor";
			valueListItem7.DataValue = 1;
			valueListItem7.DisplayText = "Igual";
			valueListItem8.DataValue = 2;
			valueListItem8.DisplayText = "Menor";
			valueListItem9.DataValue = 3;
			valueListItem9.DisplayText = "Todo";
			valueListItem10.DataValue = 4;
			valueListItem10.DisplayText = "Entre";
			this.cmbAlto.Items.Add(valueListItem6);
			this.cmbAlto.Items.Add(valueListItem7);
			this.cmbAlto.Items.Add(valueListItem8);
			this.cmbAlto.Items.Add(valueListItem9);
			this.cmbAlto.Items.Add(valueListItem10);
			this.cmbAlto.Location = new System.Drawing.Point(72, 56);
			this.cmbAlto.Name = "cmbAlto";
			this.cmbAlto.Size = new System.Drawing.Size(96, 21);
			this.cmbAlto.TabIndex = 21;
			this.cmbAlto.Text = "Mayor";
			this.cmbAlto.ValueChanged += new System.EventHandler(this.cmbAlto_ValueChanged);
			// 
			// txtAltoMin
			// 
			this.txtAltoMin.Location = new System.Drawing.Point(184, 56);
			this.txtAltoMin.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtAltoMin.Name = "txtAltoMin";
			this.txtAltoMin.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAltoMin.PromptChar = ' ';
			this.txtAltoMin.Size = new System.Drawing.Size(96, 21);
			this.txtAltoMin.TabIndex = 20;
			// 
			// cmbArtFiltro
			// 
			this.cmbArtFiltro.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArtFiltro.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArtFiltro.DataSource = this.cdvArticulo;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 5;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Width = 269;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbArtFiltro.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbArtFiltro.DisplayMember = "Articulo";
			this.cmbArtFiltro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArtFiltro.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArtFiltro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArtFiltro.Location = new System.Drawing.Point(128, 8);
			this.cmbArtFiltro.Name = "cmbArtFiltro";
			this.cmbArtFiltro.Size = new System.Drawing.Size(272, 21);
			this.cmbArtFiltro.TabIndex = 19;
			this.cmbArtFiltro.ValueMember = "idArticulo";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(24, 32);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(36, 16);
			this.label5.TabIndex = 18;
			this.label5.Text = "Ancho";
			// 
			// txtAnchoMax
			// 
			this.txtAnchoMax.Location = new System.Drawing.Point(304, 32);
			this.txtAnchoMax.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtAnchoMax.Name = "txtAnchoMax";
			this.txtAnchoMax.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAnchoMax.PromptChar = ' ';
			this.txtAnchoMax.Size = new System.Drawing.Size(96, 21);
			this.txtAnchoMax.TabIndex = 17;
			this.txtAnchoMax.Visible = false;
			// 
			// cmbAncho
			// 
			this.cmbAncho.AutoComplete = true;
			this.cmbAncho.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem11.DataValue = 0;
			valueListItem11.DisplayText = "Mayor";
			valueListItem12.DataValue = 1;
			valueListItem12.DisplayText = "Igual";
			valueListItem13.DataValue = 2;
			valueListItem13.DisplayText = "Menor";
			valueListItem14.DataValue = 3;
			valueListItem14.DisplayText = "Todo";
			valueListItem15.DataValue = 4;
			valueListItem15.DisplayText = "Entre";
			this.cmbAncho.Items.Add(valueListItem11);
			this.cmbAncho.Items.Add(valueListItem12);
			this.cmbAncho.Items.Add(valueListItem13);
			this.cmbAncho.Items.Add(valueListItem14);
			this.cmbAncho.Items.Add(valueListItem15);
			this.cmbAncho.Location = new System.Drawing.Point(72, 32);
			this.cmbAncho.Name = "cmbAncho";
			this.cmbAncho.Size = new System.Drawing.Size(96, 21);
			this.cmbAncho.TabIndex = 16;
			this.cmbAncho.Text = "Mayor";
			this.cmbAncho.ValueChanged += new System.EventHandler(this.cmbAncho_ValueChanged);
			// 
			// txtAnchoMin
			// 
			this.txtAnchoMin.Location = new System.Drawing.Point(184, 32);
			this.txtAnchoMin.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtAnchoMin.Name = "txtAnchoMin";
			this.txtAnchoMin.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAnchoMin.PromptChar = ' ';
			this.txtAnchoMin.Size = new System.Drawing.Size(96, 21);
			this.txtAnchoMin.TabIndex = 14;
			// 
			// chkArtFiltro
			// 
			this.chkArtFiltro.BackColor = System.Drawing.Color.Transparent;
			this.chkArtFiltro.Location = new System.Drawing.Point(24, 8);
			this.chkArtFiltro.Name = "chkArtFiltro";
			this.chkArtFiltro.Size = new System.Drawing.Size(96, 20);
			this.chkArtFiltro.TabIndex = 12;
			this.chkArtFiltro.Text = "Todo Artículo";
			this.chkArtFiltro.CheckedChanged += new System.EventHandler(this.chkArtFiltro_CheckedChanged);
			// 
			// cmbArticulo
			// 
			this.cmbArticulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo.DataSource = this.cdvArticulo;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 4;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 5;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 6;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Width = 198;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn26.Header.VisiblePosition = 7;
			ultraGridColumn27.Header.VisiblePosition = 8;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbArticulo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbArticulo.DisplayMember = "Articulo";
			this.cmbArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticulo.Location = new System.Drawing.Point(136, 160);
			this.cmbArticulo.Name = "cmbArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(320, 96);
			this.cmbArticulo.TabIndex = 11;
			this.cmbArticulo.ValueMember = "idArticulo";
			this.cmbArticulo.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvRetazoI;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.Caption = "Articulo";
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn29.Width = 214;
			ultraGridColumn30.Header.VisiblePosition = 2;
			ultraGridColumn30.Width = 69;
			ultraGridColumn31.Header.VisiblePosition = 3;
			ultraGridColumn31.Width = 71;
			ultraGridColumn32.Format = "dd/MMM/yyyy";
			ultraGridColumn32.Header.VisiblePosition = 4;
			ultraGridColumn32.Width = 80;
			ultraGridColumn33.Header.VisiblePosition = 5;
			ultraGridColumn34.Header.VisiblePosition = 6;
			ultraGridColumn34.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 120);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(560, 168);
			this.ultraGrid1.TabIndex = 10;
			// 
			// cdvRetazoI
			// 
			this.cdvRetazoI.BindingContextCtrl = this;
			this.cdvRetazoI.DataSet = this.cdsRetazo;
			this.cdvRetazoI.TableName = "";
			this.cdvRetazoI.TableViewName = "Retazo";
			// 
			// cdsRetazo
			// 
			this.cdsRetazo.BindingContextCtrl = this;
			this.cdsRetazo.DataLibrary = "LibFacturacion";
			this.cdsRetazo.DataLibraryUrl = "";
			this.cdsRetazo.DataSetDef = "dsRetazo";
			this.cdsRetazo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRetazo.Name = "cdsRetazo";
			this.cdsRetazo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsRetazo.SchemaDef = null;
			this.cdsRetazo.BeforeFill += new C1.Data.FillEventHandler(this.cdsRetazo_BeforeFill);
			// 
			// cdvCodArt
			// 
			this.cdvCodArt.BindingContextCtrl = this;
			this.cdvCodArt.DataSet = this.cdsArticulo;
			this.cdvCodArt.Sort = "Codigo";
			this.cdvCodArt.TableName = "";
			this.cdvCodArt.TableViewName = "Articulo";
			// 
			// cdvRetazoE
			// 
			this.cdvRetazoE.BindingContextCtrl = this;
			this.cdvRetazoE.DataSet = this.cdsRetazo;
			this.cdvRetazoE.TableName = "";
			this.cdvRetazoE.TableViewName = "";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Location = new System.Drawing.Point(16, 16);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(592, 336);
			this.ultraTabControl1.TabIndex = 13;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Ingreso";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Filtro";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(588, 310);
			// 
			// Retazo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(616, 366);
			this.Controls.Add(this.ultraTabControl1);
			this.Name = "Retazo";
			this.Text = "Retazo";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Retazo_Closing);
			this.Load += new System.EventHandler(this.Retazo_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAncho)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbFechaMax)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbFechaMin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAltoMax)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAlto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAltoMin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArtFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnchoMax)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAncho)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnchoMin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetazoI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetazo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetazoE)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void Retazo_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Retazo'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			cmbFecha.Value = DateTime.Now;
		}

		private void BorraErrores()
		{
			errorProvider.SetError(txtLote, "");
			errorProvider.SetError(cmbArticulo1, "");
			errorProvider.SetError(txtAncho, "");
			errorProvider.SetError(txtLargo, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (txtLote.Value == null)
			{
				errorProvider.SetError(txtLote, "Ingrese Lote");
				bOk = false;
			}
			if (cmbArticulo1.Value == null || (int) cmbArticulo1.Value < 1)
			{
				errorProvider.SetError(cmbArticulo1, "Ingrese Articulo");
				bOk = false;
			}
			if ((double) txtLargo.Value < 1)
			{
				errorProvider.SetError(txtLargo, "Longitud no puede ser menor que 1");
				bOk = false;
			}
			if ((double) txtAncho.Value < 1)
			{
				errorProvider.SetError(txtAncho, "Ancho no puede ser menor que 1");
				bOk = false;
			}
			return bOk;
		}

		private void btIngreso_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			C1.Data.C1DataRow drRetazo = cdsRetazo.TableViews["Retazo"].AddNew();
			drRetazo["idArticulo"] = cmbArticulo1.Value;
			drRetazo["Ancho"] = txtAncho.Value;
			drRetazo["Largo"] = txtLargo.Value;
			drRetazo["Fecha"] = dtFecha;
			drRetazo["Lote"] = txtLote.Value;
			cdsRetazo.Update();
			MessageBox.Show("Retazo Ingresado");
			txtLote.Value = "";
			txtAncho.Value = 0;
			txtLargo.Value = 0;
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsRetazo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsRetazo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cmbAncho_ValueChanged(object sender, System.EventArgs e)
		{
			switch((int) cmbAncho.Value)
			{
				case 0: // Mayor
				case 1: // Menor
				case 2: // Igual
					txtAnchoMin.Enabled = true;
					txtAnchoMax.Visible = false;
					break;
				case 3: // Todo
					txtAnchoMin.Enabled = false;
					txtAnchoMax.Visible = false;
					break;
				case 4: // Entre
					txtAnchoMin.Enabled = true;
					txtAnchoMax.Visible = true;
					break;
			}
		}

		private void cmbAlto_ValueChanged(object sender, System.EventArgs e)
		{
			switch((int) cmbAlto.Value)
			{
				case 0: // Mayor
				case 1: // Menor
				case 2: // Igual
					txtAltoMin.Enabled = true;
					txtAltoMax.Visible = false;
					break;
				case 3: // Todo
					txtAltoMin.Enabled = false;
					txtAltoMax.Visible = false;
					break;
				case 4: // Entre
					txtAltoMin.Enabled = true;
					txtAltoMax.Visible = true;
					break;
			}		
		}

		private void cmbFechaFiltro_ValueChanged(object sender, System.EventArgs e)
		{
			switch((int) cmbFechaFiltro.Value)
			{
				case 0: // Mayor
				case 1: // Menor
				case 2: // Igual
					CmbFechaMin.Enabled = true;
					CmbFechaMax.Visible = false;
					break;
				case 3: // Todo
					CmbFechaMin.Enabled = false;
					CmbFechaMax.Visible = false;
					break;
				case 4: // Entre
					CmbFechaMin.Enabled = true;
					CmbFechaMax.Visible = true;
					break;
			}		
		}

		private void chkArtFiltro_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbArtFiltro.Enabled = !chkArtFiltro.Checked;
		}

		private bool VerificaFiltro()
		{
			bool bOk = true;
			errorProvider.SetError(cmbArtFiltro, "");
			if (!chkArtFiltro.Checked)
				if (cmbArtFiltro.Value == null || (int) cmbArtFiltro.Value < 1)
				{
					errorProvider.SetError(cmbArtFiltro, "Ingrese Articulo");
					bOk = false;
				}
			return bOk;
		}

		private string stOperador(int iTipo)
		{
			string stOp = "";
			if (iTipo == 0) stOp = ">=";
			if (iTipo == 1) stOp = "=";
			if (iTipo == 2) stOp = "<=";
			return stOp;
		}

		private string Filtro(int iTipo)
		{
			DateTime dtFechaMin = (DateTime) CmbFechaMin.Value;
			DateTime dtFechaMax = (DateTime) CmbFechaMax.Value;
			string stFiltro = "";
			if (!chkArtFiltro.Checked)
				stFiltro += " And {Retazo.idArticulo} = " + cmbArtFiltro.Value.ToString();
			if ((int) cmbAncho.Value  < 3)
				stFiltro += " And {Retazo.Ancho} " + stOperador((int) cmbAncho.Value)
					+ txtAnchoMin.Value.ToString();
			if ((int) cmbAncho.Value  > 3)
				stFiltro += " And {Retazo.Ancho} >= " + txtAnchoMin.Value.ToString()
					+ " And {Retazo.Ancho} <= " + txtAnchoMax.Value.ToString();
			if ((int) cmbAlto.Value  < 3)
				stFiltro += " And {Retazo.Largo} " + stOperador((int) cmbAlto.Value)
					+ txtAltoMin.Value.ToString();
			if ((int) cmbAlto.Value  > 3)
				stFiltro += " And {Retazo.Largo} >= " + txtAltoMin.Value.ToString()
					+ " And {Retazo.Largo} <= " + txtAltoMax.Value.ToString();
			if (iTipo == 1)
			{
				if ((int) cmbFechaFiltro.Value  < 3)
					stFiltro += " And {Retazo.Fecha} " + stOperador((int) cmbFechaFiltro.Value) + " '"
						+ dtFechaMin.ToString("yyyyMMdd") + "'";
				if ((int) cmbFechaFiltro.Value  > 3)
					stFiltro += " And {Retazo.Fecha} >= '" + dtFechaMin.ToString("yyyyMMdd")
						+ "' And {Retazo.Fecha} <= '" + dtFechaMax.ToString("yyyyMMdd") + "'";

				stFiltro = stFiltro.Replace("{", "");
				stFiltro = stFiltro.Replace("}", "");
			}
			else
			{
				if ((int) cmbFechaFiltro.Value  < 3)
					stFiltro += " And {Retazo.Fecha} " + stOperador((int) cmbFechaFiltro.Value) 
						+ " #" + dtFechaMin.ToString("MM-dd-yyy") + "#";
				if ((int) cmbFechaFiltro.Value  > 3)
					stFiltro += " And {Retazo.Fecha} >= #" + dtFechaMin.ToString("MM-dd-yyyy")
						+ "# And {Retazo.Fecha} <= #" + dtFechaMax.ToString("MM-dd-yyyy") + "#";
			}

			if (stFiltro.StartsWith(" And")) stFiltro = stFiltro.Substring(5);
			return stFiltro;
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			if (!VerificaFiltro()) return;

			string stFiltro = Filtro(1);
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(cdsRetazo, "Retazo", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			cdsRetazo.Fill(fcTotal, false);

//			MessageBox.Show(stFiltro);
		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			if (!VerificaFiltro()) return;
			Cursor = Cursors.WaitCursor;
			string stFiltro = Filtro(2);
			Reporte miRetazo = new Reporte("Retazo", stFiltro);
			miRetazo.MdiParent = this.MdiParent;
			miRetazo.Show();
			Cursor = Cursors.Default;
		}

		CultureInfo us = new CultureInfo("en-US");
		private void btSalida_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(ultraGrid1, "");
			if (ultraGrid1.ActiveRow == null)
			{
				errorProvider.SetError(ultraGrid1, "Seleccione articulo a Egresar");
				return;
			}
			int IdArt = (int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
			string stArt = "Select Articulo From Articulo Where idArticulo = " 
				+ IdArt.ToString();
			stArt = Funcion.sEscalarSQL(cdsRetazo, stArt, true);
			double dAncho = (double) ultraGrid1.ActiveRow.Cells["Ancho"].Value;
			double dLargo = (double) ultraGrid1.ActiveRow.Cells["Largo"].Value;
			if (DialogResult.No == MessageBox.Show("¿Desea Egresar " 
				+ stArt + " de ancho "
				+ dAncho.ToString("0.00", us) + " y largo " 
				+ dLargo.ToString("0.00", us) + "?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			int IdRetazo = (int) ultraGrid1.ActiveRow.Cells["idRetazo"].Value;
			string stSelect = "Update Retazo Set Salida = 1 Where idRetazo = "
				+ IdRetazo.ToString();
			Funcion.EjecutaSQL(cdsRetazo, stSelect, true);
			MessageBox.Show("Salida de Articulo " + stArt + " Registrada");
		}

		private void Retazo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Retazo'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}
	}
}
