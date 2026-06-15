using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteRecepcionDeProductos.
	/// </summary>
	public class frmReporteRecepcionDeProductos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridReporte;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
   
		public frmReporteRecepcionDeProductos()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteRecepcionDeProductos));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DetalleProductos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Accesorios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroFacturacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaFacturacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActaRecibido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("LocalVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaRecepcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioRecepcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotaRecepcion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DetalleProductos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Accesorios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroFacturacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaFacturacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActaRecibido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LocalVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRecepcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioRecepcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotaRecepcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("A", 0);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblBuscar = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.uGridReporte = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(184, 12);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(34, 16);
			this.label3.TabIndex = 854;
			this.label3.Text = "Hasta";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 12);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 853;
			this.label1.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.AutoSelect = false;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(240, 10);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 852;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			this.dtHasta.Validated += new System.EventHandler(this.dtHasta_Validated);
			this.dtHasta.ValueChanged += new System.EventHandler(this.dtHasta_ValueChanged);
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.AutoSelect = false;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 10);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 851;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
			this.dtDesde.Validated += new System.EventHandler(this.dtDesde_Validated);
			// 
			// btnConsultar
			// 
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(304, 80);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 855;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 48);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(43, 16);
			this.label2.TabIndex = 858;
			this.label2.Text = "Bodega";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance3;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 60;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 224;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 46);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(216, 21);
			this.cmbBodega.TabIndex = 857;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			// 
			// lblBuscar
			// 
			this.lblBuscar.AutoSize = true;
			this.lblBuscar.Location = new System.Drawing.Point(8, 83);
			this.lblBuscar.Name = "lblBuscar";
			this.lblBuscar.Size = new System.Drawing.Size(44, 16);
			this.lblBuscar.TabIndex = 860;
			this.lblBuscar.Text = "Número";
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(64, 82);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(216, 19);
			this.txtBuscar.TabIndex = 859;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			this.txtBuscar.ValueChanged += new System.EventHandler(this.txtBuscar_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(System.DateTime);
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
			this.groupBox1.Location = new System.Drawing.Point(0, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(992, 8);
			this.groupBox1.TabIndex = 862;
			this.groupBox1.TabStop = false;
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
			// uGridReporte
			// 
			this.uGridReporte.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridReporte.DataSource = this.ultraDataSource1;
			appearance4.BackColor = System.Drawing.Color.White;
			this.uGridReporte.DisplayLayout.Appearance = appearance4;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 70;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Número";
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 70;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Local";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 140;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Width = 200;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Identificación";
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Width = 85;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Detalle del Producto";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 200;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 200;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 14;
			ultraGridColumn11.Width = 100;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "# Factura";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 90;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "F. Factura";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 75;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Width = 200;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.Caption = "Local Venta";
			ultraGridColumn16.Header.VisiblePosition = 13;
			ultraGridColumn16.Width = 140;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.Caption = "Estado";
			ultraGridColumn17.Header.VisiblePosition = 5;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.Caption = "Fec. Recep.";
			ultraGridColumn18.Header.VisiblePosition = 15;
			ultraGridColumn18.Width = 100;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.Caption = "Usuario Recep.";
			ultraGridColumn19.Header.VisiblePosition = 16;
			ultraGridColumn19.Width = 100;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.Caption = "Notas de Recepción";
			ultraGridColumn20.Header.VisiblePosition = 17;
			ultraGridColumn20.Width = 300;
			ultraGridColumn21.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			appearance5.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn21.CellButtonAppearance = appearance5;
			ultraGridColumn21.Header.Caption = "....";
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn21.Width = 30;
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
			this.uGridReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridReporte.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridReporte.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridReporte.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8.5F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridReporte.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.uGridReporte.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridReporte.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridReporte.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridReporte.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridReporte.Location = new System.Drawing.Point(8, 128);
			this.uGridReporte.Name = "uGridReporte";
			this.uGridReporte.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.uGridReporte.Size = new System.Drawing.Size(976, 328);
			this.uGridReporte.TabIndex = 861;
			this.uGridReporte.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridReporte_ClickCellButton);
			this.uGridReporte.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridReporte_InitializeLayout);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(408, 81);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 863;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(904, 79);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 864;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// optEstados
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance10;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance11;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendiente";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Cancelado";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Anulado";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Verificado";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.Items.Add(valueListItem5);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(400, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(392, 24);
			this.optEstados.TabIndex = 871;
			this.optEstados.Text = "Todos";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// frmReporteRecepcionDeProductos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(992, 466);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.uGridReporte);
			this.Controls.Add(this.lblBuscar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.KeyPreview = true;
			this.Name = "frmReporteRecepcionDeProductos";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Recepción De Productos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteRecepcionDeProductos_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteRecepcionDeProductos_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void frmReporteRecepcionDeProductos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0733");

			if (miAcceso.BExportar) this.btnExcel.Enabled = true;
			if (miAcceso.Nuevo) this.uGridReporte.DisplayLayout.Bands[0].Columns["A"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit; 

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtDesde.Value=DateTime.Today.AddMonths(-1);

      this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.Value=DateTime.Today;

//			string sSQL = string.Format("Exec ListaBodegas {0}, 47", MenuLatinium.IdUsuario);
//			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);			
		}

		private void Consulta()
		{
			#region valida Fechas
			if(!Validacion.vbFechaEnDocumentos(this.dtDesde,"Seleccione la fecha inicial",false,"",cdsSeteoF))return;
			if(!Validacion.vbFechaEnDocumentos(this.dtHasta,"Seleccione la fecha final",false,"",cdsSeteoF))return;
			if(!Validacion.vbComparaFechas((DateTime)this.dtDesde.Value,(DateTime)this.dtHasta.Value,this.dtHasta,"La fecha incial no puede ser posterior a la fecha final",false))return;
			#endregion valida Fechas

			int iBodega=0;
			if(this.cmbBodega.ActiveRow !=null)iBodega=(int)this.cmbBodega.Value;
      
			string sSQL =string.Format("Exec ReporteRecepcionDeProductos '{0}','{1}',{2},{3},'{4}'",
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyMMdd"),Convert.ToDateTime(this.dtHasta.Value).ToString("yyyMMdd"),
				(int)this.optEstados.Value,iBodega,this.txtBuscar.Text.ToString());
			this.uGridReporte.DataSource=Funcion.dvProcedimiento(cdsSeteoF,sSQL);

			this.lblContador.Text = this.uGridReporte.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{						
			Consulta();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void txtBuscar_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
     
		private void uGridReporte_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "A")
			{				      			
				int iEstado = Funcion.iEscalarSQL(cdsSeteoF,string.Format("if Exists(Select idActaRecibido from ActaRecibido where idActaRecibido ={0}) Select Estado from ActaRecibido where idActaRecibido =  {0}  Else Select CONVERT(Int, 0)",
					(int)e.Cell.Row.Cells["idActaRecibido"].Value));

				if (iEstado == 2)
				{
					using (frmNotasRecepcionDeProductos NRDP = new frmNotasRecepcionDeProductos (0))
					{
						if (DialogResult.OK == NRDP.ShowDialog())
						{
							string sSQL = string.Format("Update ActaRecibido set Estado = 4, FechaRecepcion = Getdate() , UsuarioRecepcion = suser_sname(), NotaRecepcion = '{0}' where idActaRecibido = {1}", 
								NRDP.txtNotas.Text.ToString(), (int)e.Cell.Row.Cells["idActaRecibido"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL);

							this.btnConsultar_Click(sender, e);
						}
					}
				}
				else 
				{
					MessageBox.Show("Solo para registros en estado Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);		
					return;
				}				
			}
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnConsultar.Focus();
		}

		private void uGridReporte_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmReporteRecepcionDeProductos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridReporte);
		}

		private void dtDesde_Validated(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value == System.DBNull.Value) this.dtDesde.Value = DateTime.Today;
		}

		private void dtHasta_Validated(object sender, System.EventArgs e)
		{
			if (this.dtHasta.Value == System.DBNull.Value) this.dtHasta.Value = DateTime.Today;
		}

		private void dtHasta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtHasta.Value != System.DBNull.Value)
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 45, (DateTime)this.dtHasta.Value, false, cdsSeteoF);
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			this.Consulta();
		}
	}
}
