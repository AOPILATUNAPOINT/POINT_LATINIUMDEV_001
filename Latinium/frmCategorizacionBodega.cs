using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCategorizacionBodega.
	/// </summary>
	public class frmCategorizacionBodega : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbUbicacionBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTamañoBodega;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBodegaTamaño;
		bool bActualiza = false;
		//string sLetra='A','B','C','D';
		public Infragistics.Win.UltraWinGrid.UltraGrid grdCategorizacionBodegas;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBodegaUbicacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCategoriaBodega;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optBodegas;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCategorizacionBodega()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdCategorizacionBodegas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdCategorizacionUbicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdCategorizacionTamaño");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NivelVentasGeneral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaTecnologia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaTVS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaLineaBlanca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaMobilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaContadoTC");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCategorizacionBodega));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdUbicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdTamaño");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tamaño");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdTamaño");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tamaño");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdTamaño");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tamaño");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdCategorizacionBodegas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdCategorizacionUbicacion", -1, "cmbBodegaUbicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdCategorizacionTamaño", -1, "cmbBodegaTamaño");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NivelVentasGeneral", -1, "cmbCategoriaBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaTecnologia", -1, "cmbCategoriaBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaTVS", -1, "cmbCategoriaBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaLineaBlanca", -1, "cmbCategoriaBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaMobilidad", -1, "cmbCategoriaBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaCredito", -1, "cmbCategoriaBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaContadoTC", -1, "cmbCategoriaBodega");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdUbicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdUbicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ubicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdCategoriaBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdCategoriaBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbUbicacionBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTamañoBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblContador = new System.Windows.Forms.Label();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodegaTamaño = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.grdCategorizacionBodegas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbBodegaUbicacion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCategoriaBodega = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.optBodegas = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUbicacionBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTamañoBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaTamaño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCategorizacionBodegas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoriaBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optBodegas)).BeginInit();
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
			this.cdsSeteoF.AfterFill += new C1.Data.FillEventHandler(this.cdsSeteoF_AfterFill);
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
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
																																 ultraDataColumn11});
			this.ultraDataSource1.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource1_CellDataRequested);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(928, 24);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 1001;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(328, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(88, 23);
			this.btnVer.TabIndex = 1002;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(88, 16);
			this.label2.TabIndex = 1005;
			this.label2.Text = "Tamaño Bodega";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Click += new System.EventHandler(this.label2_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(96, 16);
			this.label1.TabIndex = 1007;
			this.label1.Text = "Ubicacion Bodega";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbUbicacionBodega
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbUbicacionBodega.Appearance = appearance1;
			this.cmbUbicacionBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbUbicacionBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 200;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbUbicacionBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbUbicacionBodega.DisplayMember = "Ubicacion";
			this.cmbUbicacionBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbUbicacionBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUbicacionBodega.Location = new System.Drawing.Point(112, 40);
			this.cmbUbicacionBodega.Name = "cmbUbicacionBodega";
			this.cmbUbicacionBodega.Size = new System.Drawing.Size(184, 21);
			this.cmbUbicacionBodega.TabIndex = 1010;
			this.cmbUbicacionBodega.ValueMember = "IdUbicacion";
			// 
			// cmbTamañoBodega
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTamañoBodega.Appearance = appearance2;
			this.cmbTamañoBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTamañoBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbTamañoBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTamañoBodega.DisplayMember = "Tamaño";
			this.cmbTamañoBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTamañoBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTamañoBodega.Location = new System.Drawing.Point(112, 16);
			this.cmbTamañoBodega.Name = "cmbTamañoBodega";
			this.cmbTamañoBodega.Size = new System.Drawing.Size(184, 21);
			this.cmbTamañoBodega.TabIndex = 1011;
			this.cmbTamañoBodega.ValueMember = "IdTamaño";
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(696, 11);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 1012;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource10
			// 
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn12,
																																	ultraDataColumn13});
			// 
			// cmbBodegaTamaño
			// 
			this.cmbBodegaTamaño.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodegaTamaño.DataSource = this.ultraDataSource10;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbBodegaTamaño.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodegaTamaño.DisplayMember = "Tamaño";
			this.cmbBodegaTamaño.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaTamaño.Location = new System.Drawing.Point(336, 208);
			this.cmbBodegaTamaño.Name = "cmbBodegaTamaño";
			this.cmbBodegaTamaño.Size = new System.Drawing.Size(240, 96);
			this.cmbBodegaTamaño.TabIndex = 1041;
			this.cmbBodegaTamaño.ValueMember = "idTamaño";
			this.cmbBodegaTamaño.Visible = false;
			this.cmbBodegaTamaño.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodegaTamaño_InitializeLayout);
			// 
			// grdCategorizacionBodegas
			// 
			this.grdCategorizacionBodegas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdCategorizacionBodegas.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdCategorizacionBodegas.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdCategorizacionBodegas.DisplayLayout.Appearance = appearance3;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 124;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 180;
			ultraGridColumn9.Header.Caption = "Ubicacion";
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 108;
			ultraGridColumn10.Header.Caption = "Tamaño";
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 76;
			ultraGridColumn11.Header.Caption = "Nivel de Ventas General";
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.Header.Caption = "Venta de Tecnologia";
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn12.Width = 86;
			ultraGridColumn13.Header.Caption = "Venta de TVS";
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn13.Width = 78;
			ultraGridColumn14.Header.Caption = "Venta de Linea Blanca";
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn14.Width = 80;
			ultraGridColumn15.Header.Caption = "Venta de Mobilidad";
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn15.Width = 87;
			ultraGridColumn16.Header.Caption = "Venta Credito";
			ultraGridColumn16.Header.VisiblePosition = 9;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn16.Width = 96;
			ultraGridColumn17.Header.Caption = "Venta de Contado TC";
			ultraGridColumn17.Header.VisiblePosition = 10;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn17.Width = 132;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn17});
			this.grdCategorizacionBodegas.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdCategorizacionBodegas.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.grdCategorizacionBodegas.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 9F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdCategorizacionBodegas.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.grdCategorizacionBodegas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCategorizacionBodegas.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCategorizacionBodegas.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCategorizacionBodegas.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.grdCategorizacionBodegas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdCategorizacionBodegas.Location = new System.Drawing.Point(0, 88);
			this.grdCategorizacionBodegas.Name = "grdCategorizacionBodegas";
			this.grdCategorizacionBodegas.Size = new System.Drawing.Size(1032, 368);
			this.grdCategorizacionBodegas.TabIndex = 1042;
			this.grdCategorizacionBodegas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdCategorizacionBodegas_AfterCellUpdate_1);
			// 
			// cmbBodegaUbicacion
			// 
			this.cmbBodegaUbicacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodegaUbicacion.DataSource = this.ultraDataSource2;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 100;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 81;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.cmbBodegaUbicacion.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbBodegaUbicacion.DisplayMember = "Ubicacion";
			this.cmbBodegaUbicacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaUbicacion.Location = new System.Drawing.Point(592, 200);
			this.cmbBodegaUbicacion.Name = "cmbBodegaUbicacion";
			this.cmbBodegaUbicacion.Size = new System.Drawing.Size(240, 96);
			this.cmbBodegaUbicacion.TabIndex = 1043;
			this.cmbBodegaUbicacion.ValueMember = "IdUbicacion";
			this.cmbBodegaUbicacion.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn14.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn14,
																																 ultraDataColumn15});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn16.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// cmbCategoriaBodega
			// 
			this.cmbCategoriaBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCategoriaBodega.DataSource = this.ultraDataSource3;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbCategoriaBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCategoriaBodega.DisplayMember = "Tipo";
			this.cmbCategoriaBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCategoriaBodega.Location = new System.Drawing.Point(80, 200);
			this.cmbCategoriaBodega.Name = "cmbCategoriaBodega";
			this.cmbCategoriaBodega.Size = new System.Drawing.Size(240, 96);
			this.cmbCategoriaBodega.TabIndex = 1044;
			this.cmbCategoriaBodega.ValueMember = "Tipo";
			this.cmbCategoriaBodega.Visible = false;
			// 
			// optBodegas
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.optBodegas.Appearance = appearance10;
			this.optBodegas.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optBodegas.CheckedIndex = 0;
			this.optBodegas.ItemAppearance = appearance11;
			this.optBodegas.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Activas";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Inactivas";
			this.optBodegas.Items.Add(valueListItem1);
			this.optBodegas.Items.Add(valueListItem2);
			this.optBodegas.Items.Add(valueListItem3);
			this.optBodegas.ItemSpacingVertical = 10;
			this.optBodegas.Location = new System.Drawing.Point(440, 8);
			this.optBodegas.Name = "optBodegas";
			this.optBodegas.Size = new System.Drawing.Size(216, 24);
			this.optBodegas.TabIndex = 1045;
			this.optBodegas.Text = "Todos";
			// 
			// frmCategorizacionBodega
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1032, 478);
			this.Controls.Add(this.optBodegas);
			this.Controls.Add(this.cmbCategoriaBodega);
			this.Controls.Add(this.cmbBodegaUbicacion);
			this.Controls.Add(this.cmbBodegaTamaño);
			this.Controls.Add(this.grdCategorizacionBodegas);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.cmbTamañoBodega);
			this.Controls.Add(this.cmbUbicacionBodega);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnExportar);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCategorizacionBodega";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Categorizacion Bodega";
			this.Load += new System.EventHandler(this.frmCategorizacionBodega_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUbicacionBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTamañoBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaTamaño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCategorizacionBodegas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoriaBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optBodegas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_AfterFill(object sender, C1.Data.FillEventArgs e)
		{
			
		}

		private void ultraDataSource1_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.grdCategorizacionBodegas.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			FuncionesProcedimientos.ExportaExcel(this.grdCategorizacionBodegas);
		}

		private void label2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void frmCategorizacionBodega_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200113");
			
			if (!Funcion.Permiso("1963", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar al Reporte de Categorizacion de bodegas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportarExcel) this.btnExportar.Enabled = true;
			this.btnVer_Click(sender, e);

			this.cmbUbicacionBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select IdUbicacion, Ubicacion From BodegaUbicacion Order by Ubicacion");
			this.cmbTamañoBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select IdTamaño, Tamaño From BodegaTamaño Order by Tamaño");
			this.cmbBodegaTamaño.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select IdTamaño, Tamaño From BodegaTamaño Order by Tamaño");		
			this.cmbBodegaUbicacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select IdUbicacion,UbIcacion as Ubicacion From BodegaUbicacion Order by Ubicacion");		
			this.cmbCategoriaBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select IdCategoriaBodega, Tipo From CategoriaBodega Order by Tipo");	
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Consulta
			int iUbicacion = 0;
			int iTamano=0;
			if (this.cmbUbicacionBodega.ActiveRow != null) iUbicacion = (int)this.cmbUbicacionBodega.Value;	
			if (this.cmbTamañoBodega.ActiveRow != null) iTamano = (int)this.cmbTamañoBodega.Value;

			string sSQL = string.Format("Exec CategorizacionDeBodegas '{0}','{1}',{2}",iUbicacion, iTamano,(int)this.optBodegas.Value);
			this.grdCategorizacionBodegas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.lblContador.Text = this.grdCategorizacionBodegas.Rows.Count + " REGISTROS ENCONTRADOS";

			#endregion Consulta
		}

		private void cmbIndemnizacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void grdCategorizacionBodegas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodegaTamaño_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void grdCategorizacionBodegas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}
		
		private void grdCategorizacionBodegas_AfterCellUpdate_1(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "IdCategorizacionTamaño")
			{	
				bActualiza = true;
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Tamaño de la Bodega?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE CategorizacionBodegas SET IdCategorizacionTamaño = {0} WHERE IdCategorizacionBodegas = {1} ", 
						(int)	this.grdCategorizacionBodegas.ActiveRow.Cells["IdCategorizacionTamaño"].Value, (int)	this.grdCategorizacionBodegas.ActiveRow.Cells["IdCategorizacionBodegas"].Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
				}
				else
				{						
					e.Cell.Row.Cells["IdCategorizacionTamaño"].Value = (int)e.Cell.OriginalValue;
				}					
				bActualiza = false;
			}	


			if (e.Cell.Column.ToString() == "IdCategorizacionUbicacion")
			{	
				bActualiza = true;
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar la Ubicacion de la Bodega?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE CategorizacionBodegas SET IdCategorizacionUbicacion = {0} WHERE IdCategorizacionBodegas = {1} ", 
						(int)	this.grdCategorizacionBodegas.ActiveRow.Cells["IdCategorizacionUbicacion"].Value, (int)	this.grdCategorizacionBodegas.ActiveRow.Cells["IdCategorizacionBodegas"].Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
				}
				else
				{						
					e.Cell.Row.Cells["IdCategorizacionUbicacion"].Value = (int)e.Cell.OriginalValue;
				}					
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "NivelVentasGeneral" )
			{
				bActualiza = true;
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Nivel de Ventas General de la Bodega?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE CategorizacionBodegas SET NivelVentasGeneral = '{0}' WHERE IdCategorizacionBodegas = {1} ", 
						(string)	this.grdCategorizacionBodegas.ActiveRow.Cells["NivelVentasGeneral"].Value, (int)	this.grdCategorizacionBodegas.ActiveRow.Cells["IdCategorizacionBodegas"].Value );
			
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
				}
				else
				{						
					e.Cell.Row.Cells["NivelVentasGeneral"].Value = (string)e.Cell.OriginalValue;
				}					
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "VentaTecnologia" )
			{
				bActualiza = true;
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Nivel de Ventas Tecnologia de la Bodega?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE CategorizacionBodegas SET VentaTecnologia = '{0}' WHERE IdCategorizacionBodegas = {1} ", 
						(string)	this.grdCategorizacionBodegas.ActiveRow.Cells["VentaTecnologia"].Value, (int)	this.grdCategorizacionBodegas.ActiveRow.Cells["IdCategorizacionBodegas"].Value );
			
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
				}
				else
				{						
					e.Cell.Row.Cells["VentaTecnologia"].Value = (string)e.Cell.OriginalValue;
				}					
				bActualiza = false;
			}
			
			if (e.Cell.Column.ToString() == "VentaTVS" )
			{
				bActualiza = true;
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Nivel de Ventas de TVS de la Bodega?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE CategorizacionBodegas SET VentaTVS = '{0}' WHERE IdCategorizacionBodegas = {1} ", 
						(string)	this.grdCategorizacionBodegas.ActiveRow.Cells["VentaTVS"].Value, (int)	this.grdCategorizacionBodegas.ActiveRow.Cells["IdCategorizacionBodegas"].Value );
			
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
				}
				else
				{						
					e.Cell.Row.Cells["VentaTVS"].Value = (string)e.Cell.OriginalValue;
				}					
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "VentaLineaBlanca" )
			{
				bActualiza = true;
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Nivel de Ventas de Linea Blanca de la Bodega?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE CategorizacionBodegas SET VentaLineaBlanca = '{0}' WHERE IdCategorizacionBodegas = {1} ", 
						(string)	this.grdCategorizacionBodegas.ActiveRow.Cells["VentaLineaBlanca"].Value, (int)	this.grdCategorizacionBodegas.ActiveRow.Cells["IdCategorizacionBodegas"].Value );
			
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
				}
				else
				{						
					e.Cell.Row.Cells["VentaLineaBlanca"].Value = (string)e.Cell.OriginalValue;
				}					
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "VentaMobilidad" )
			{
				bActualiza = true;
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Nivel de Ventas de Movilidad de la Bodega?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE CategorizacionBodegas SET VentaMobilidad = '{0}' WHERE IdCategorizacionBodegas = {1} ", 
						(string)	this.grdCategorizacionBodegas.ActiveRow.Cells["VentaMobilidad"].Value, (int)	this.grdCategorizacionBodegas.ActiveRow.Cells["IdCategorizacionBodegas"].Value );
			
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
				}
				else
				{						
					e.Cell.Row.Cells["VentaMobilidad"].Value = (string)e.Cell.OriginalValue;
				}					
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "VentaCredito" )
			{
				bActualiza = true;
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Nivel de Ventas de Credito de la Bodega?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE CategorizacionBodegas SET VentaCredito = '{0}' WHERE IdCategorizacionBodegas = {1} ", 
						(string)	this.grdCategorizacionBodegas.ActiveRow.Cells["VentaCredito"].Value, (int)	this.grdCategorizacionBodegas.ActiveRow.Cells["IdCategorizacionBodegas"].Value );
			
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
				}
				else
				{						
					e.Cell.Row.Cells["VentaCredito"].Value = (string)e.Cell.OriginalValue;
				}					
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "VentaContadoTC" )
			{
				bActualiza = true;
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Nivel de Ventas de Contado de la Bodega?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE CategorizacionBodegas SET VentaContadoTC = '{0}' WHERE IdCategorizacionBodegas = {1} ", 
						(string)	this.grdCategorizacionBodegas.ActiveRow.Cells["VentaContadoTC"].Value, (int)	this.grdCategorizacionBodegas.ActiveRow.Cells["IdCategorizacionBodegas"].Value );
			
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
				}
				else
				{						
					e.Cell.Row.Cells["VentaContadoTC"].Value = (string)e.Cell.OriginalValue;
				}					
				bActualiza = false;
			}
		}
	}
}