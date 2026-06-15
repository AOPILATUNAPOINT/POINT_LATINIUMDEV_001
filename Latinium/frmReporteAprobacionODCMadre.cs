using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using Latinium.Services.OrdenDeCompra.Bodega;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteAprobacionODC.
	/// </summary>
	public class frmReporteAprobacionODCMadre : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optOrigen;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.Misc.UltraLabel lblContador;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstadoDentalis;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridOrdenesDeCompra;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteAprobacionODCMadre()
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
    int idBloqueaTransacciones = 0;
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteAprobacionODCMadre));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdOrdenDeCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proveedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Presupuesto");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem14 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem15 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem16 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdOrdenDeCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Presupuesto");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ir", 0);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aceptar", 1);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.optOrigen = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblContador = new Infragistics.Win.Misc.UltraLabel();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnExcel = new System.Windows.Forms.Button();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optEstadoDentalis = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.uGridOrdenesDeCompra = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.optOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoDentalis)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOrdenesDeCompra)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 96);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1024, 8);
			this.groupBox1.TabIndex = 289;
			this.groupBox1.TabStop = false;
			// 
			// optOrigen
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optOrigen.Appearance = appearance1;
			this.optOrigen.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optOrigen.CheckedIndex = 0;
			this.optOrigen.Enabled = false;
			this.optOrigen.ItemAppearance = appearance2;
			this.optOrigen.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "(F7) Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "(F8) Locales";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "(F9) Importaciones";
			this.optOrigen.Items.Add(valueListItem1);
			this.optOrigen.Items.Add(valueListItem2);
			this.optOrigen.Items.Add(valueListItem3);
			this.optOrigen.ItemSpacingVertical = 10;
			this.optOrigen.Location = new System.Drawing.Point(584, 40);
			this.optOrigen.Name = "optOrigen";
			this.optOrigen.Size = new System.Drawing.Size(312, 26);
			this.optOrigen.TabIndex = 278;
			this.optOrigen.Text = "(F7) Todos";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(192, 10);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 287;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 286;
			this.label6.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance3;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 277;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance4;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 276;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// lblContador
			// 
			appearance5.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Appearance = appearance5;
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.lblContador.Location = new System.Drawing.Point(432, 80);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 0);
			this.lblContador.TabIndex = 285;
			// 
			// btnBuscar
			// 
			this.btnBuscar.Location = new System.Drawing.Point(288, 71);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 280;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel4.Location = new System.Drawing.Point(8, 75);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(37, 15);
			this.ultraLabel4.TabIndex = 284;
			this.ultraLabel4.Text = "Buscar";
			// 
			// txtBusqueda
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance6;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(64, 72);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(200, 21);
			this.txtBusqueda.TabIndex = 279;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn2.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(System.Decimal);
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
			// optEstados
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance7;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 1;
			this.optEstados.ItemAppearance = appearance8;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Todos";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Pendientes";
			valueListItem6.DataValue = 3;
			valueListItem6.DisplayText = "Procesado";
			valueListItem7.DataValue = 9;
			valueListItem7.DisplayText = "Facturado";
			valueListItem8.DataValue = 5;
			valueListItem8.DisplayText = "Aprobado";
			valueListItem9.DataValue = 4;
			valueListItem9.DisplayText = "Rechazado";
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.Items.Add(valueListItem5);
			this.optEstados.Items.Add(valueListItem6);
			this.optEstados.Items.Add(valueListItem7);
			this.optEstados.Items.Add(valueListItem8);
			this.optEstados.Items.Add(valueListItem9);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(8, 40);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(496, 26);
			this.optEstados.TabIndex = 311;
			this.optEstados.Text = "Pendientes";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(944, 6);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 737;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource2;
			this.cmbCuenta.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 144;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 374;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuenta.DisplayMember = "Cuenta";
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(416, 192);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(376, 112);
			this.cmbCuenta.TabIndex = 738;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// optEstadoDentalis
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstadoDentalis.Appearance = appearance9;
			this.optEstadoDentalis.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstadoDentalis.CheckedIndex = 1;
			this.optEstadoDentalis.ItemAppearance = appearance10;
			this.optEstadoDentalis.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem10.DataValue = 0;
			valueListItem10.DisplayText = "Todos";
			valueListItem11.DataValue = 1;
			valueListItem11.DisplayText = "Pendientes";
			valueListItem12.DataValue = 6;
			valueListItem12.DisplayText = "Revisión";
			valueListItem13.DataValue = 3;
			valueListItem13.DisplayText = "Procesado";
			valueListItem14.DataValue = 9;
			valueListItem14.DisplayText = "Facturado";
			valueListItem15.DataValue = 5;
			valueListItem15.DisplayText = "Aprobado";
			valueListItem16.DataValue = 4;
			valueListItem16.DisplayText = "Rechazado";
			this.optEstadoDentalis.Items.Add(valueListItem10);
			this.optEstadoDentalis.Items.Add(valueListItem11);
			this.optEstadoDentalis.Items.Add(valueListItem12);
			this.optEstadoDentalis.Items.Add(valueListItem13);
			this.optEstadoDentalis.Items.Add(valueListItem14);
			this.optEstadoDentalis.Items.Add(valueListItem15);
			this.optEstadoDentalis.Items.Add(valueListItem16);
			this.optEstadoDentalis.ItemSpacingVertical = 10;
			this.optEstadoDentalis.Location = new System.Drawing.Point(8, 40);
			this.optEstadoDentalis.Name = "optEstadoDentalis";
			this.optEstadoDentalis.Size = new System.Drawing.Size(568, 26);
			this.optEstadoDentalis.TabIndex = 739;
			this.optEstadoDentalis.Text = "Pendientes";
			this.optEstadoDentalis.ValueChanged += new System.EventHandler(this.optEstadoDentalis_ValueChanged);
			// 
			// uGridOrdenesDeCompra
			// 
			this.uGridOrdenesDeCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridOrdenesDeCompra.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridOrdenesDeCompra.DataSource = this.ultraDataSource1;
			appearance11.BackColor = System.Drawing.Color.White;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOrdenesDeCompra.DisplayLayout.Appearance = appearance11;
			this.uGridOrdenesDeCompra.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 95;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 59;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Número";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 78;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "RUC";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 92;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.Caption = "Importación";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Width = 191;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 284;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance12;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 89;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 127;
			ultraGridColumn11.Header.Caption = "Cuenta";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn11.Width = 120;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 100;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance13;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 120;
			ultraGridColumn14.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance14.Image = ((object)(resources.GetObject("appearance14.Image")));
			appearance14.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellButtonAppearance = appearance14;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn14.Width = 35;
			ultraGridColumn15.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance15.Image = ((object)(resources.GetObject("appearance15.Image")));
			appearance15.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellButtonAppearance = appearance15;
			ultraGridColumn15.Header.Caption = "Ok";
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn15.Width = 32;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn15});
			ultraGridBand2.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.uGridOrdenesDeCompra.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.HeaderAppearance = appearance18;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance19.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance19.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.RowAlternateAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.uGridOrdenesDeCompra.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridOrdenesDeCompra.Location = new System.Drawing.Point(8, 112);
			this.uGridOrdenesDeCompra.Name = "uGridOrdenesDeCompra";
			this.uGridOrdenesDeCompra.Size = new System.Drawing.Size(1008, 248);
			this.uGridOrdenesDeCompra.TabIndex = 281;
			this.uGridOrdenesDeCompra.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridOrdenesDeCompra_ClickCellButton);
			this.uGridOrdenesDeCompra.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridOrdenesDeCompra_InitializeLayout);
			// 
			// frmReporteAprobacionODCMadre
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1024, 366);
			this.Controls.Add(this.optEstadoDentalis);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.optOrigen);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.uGridOrdenesDeCompra);
			this.KeyPreview = true;
			this.Name = "frmReporteAprobacionODCMadre";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Aprobación Ordenes De Compra Madre";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteAprobacionODC_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteAprobacionODC_Load);
			((System.ComponentModel.ISupportInitialize)(this.optOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoDentalis)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOrdenesDeCompra)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;
     bool ODCDiaCredito = false;
		private void frmReporteAprobacionODC_Load(object sender, System.EventArgs e)
		{

			miAcceso = new Acceso(cdsSeteoF, "0628");

			if (miAcceso.BExportar)
			{
				this.btnExcel.Visible =  true;

			}					
				
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2019");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2019");				
			
			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;

			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select idCuenta, c.Codigo +'  '+ c.Descripcion + ' / ' + g.Grupo as Cuenta  from Cuenta c Inner Join CuentaGrupo g on c.idCuentaAgrupa = g.idCuentaAgrupa where GASTO > 0");

			ODCDiaCredito = Funcion.bEscalarSQL(cdsSeteoF, string.Format("select isnull (ODCDiaCredito, 0) from Seteo"), true);
//			if (ODCDiaCredito)
//			{
//
//				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = false;
//				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Periodo"].Hidden = false;
//				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Presupuesto"].Hidden = false;
//			}

			if (ODCDiaCredito)
			{
				this.optEstadoDentalis.Visible = true;
				this.optEstados.Visible = false;
			}
			else
			{
				this.optEstadoDentalis.Visible = false;
				this.optEstados.Visible = true;
			}
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			int IEstado = 0;
			bool IBorrar = false;

			bool bFechas = false;

			DateTime dtDesde = DateTime.Today;
			DateTime dtHasta = DateTime.Today;

			if (this.dtDesde.Value != System.DBNull.Value || this.dtHasta.Value != System.DBNull.Value)
			{
				bFechas = true;

				dtDesde = (DateTime)this.dtDesde.Value;
				dtHasta = (DateTime)this.dtHasta.Value;
			}

			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;
      int iEstado = 0;

			if (!ODCDiaCredito) 
			{
				iEstado = (int) this.optEstados.Value;
			}
			else
			{
       iEstado = (int) this.optEstadoDentalis.Value;
				if (miAcceso.Imprimir) this.optEstadoDentalis.Value = 1;
			}
			string sSQL = string.Format("Exec ConsultaAprobacionOrdenDeCompraNuevo '{0}', {1}, {2}, {3}, {4}, '{5}', '{6}'", 
				this.txtBusqueda.Text.ToString(), iEstado, IBorrar, (int)this.optOrigen.Value, bFechas, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			this.uGridOrdenesDeCompra.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			if (miAcceso.Editar)
			{
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = false;
			}
			else
			{
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
			}
			if (!ODCDiaCredito)
			{
				Estados ();
			}
			else
			{
				EstadosDentalis();
			}

			if (this.uGridOrdenesDeCompra.Rows.Count > 0) this.uGridOrdenesDeCompra.Focus();
			else this.txtBusqueda.Focus();

			this.lblContador.Text = this.uGridOrdenesDeCompra.Rows.Count + " Registros Encontrados";
		}

		private void frmReporteAprobacionODC_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
			if (e.KeyCode == Keys.F7) this.optOrigen.Value = 0;
			if (e.KeyCode == Keys.F8) this.optOrigen.Value = 1;
			if (e.KeyCode == Keys.F9) this.optOrigen.Value = 2;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			
      Estados ();
			this.btnBuscar_Click (sender,  e);
		}

		private void Estados ()
		{
			if((int) this.optEstados.Value == 0)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = true;
			}

			if((int) this.optEstados.Value == 1)
			{
				if (miAcceso.Editar)
				{
					this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = false;
				}
				else
				{
					this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				}
		
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstados.Value == 3)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			
			if((int) this.optEstados.Value == 9)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstados.Value == 5)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstados.Value == 4)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstados.Value == 6)
			{
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
			}
		}


		private void EstadosDentalis ()
		{
//							this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = false;
//							this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Periodo"].Hidden = false;
//							this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Presupuesto"].Hidden = false;

			if((int) this.optEstadoDentalis.Value == 0)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = true;
			}

			if((int) this.optEstadoDentalis.Value == 1)
			{
				if (miAcceso.Editar)
				{
					this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = false;
				}
				else
				{
					this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				}
		
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstadoDentalis.Value == 3)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			
			if((int) this.optEstadoDentalis.Value == 9)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstadoDentalis.Value == 5)
			{


				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstadoDentalis.Value == 4)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstadoDentalis.Value == 6)
			{
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = false;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Periodo"].Hidden = false;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Presupuesto"].Hidden = false;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = false;
			}
		}

//		private void uGridOrdenesDeCompra_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
//		{
//			if (e.Cell.Column.ToString() == "Aceptar")
//			{	
//				if (Funcion.ValidaBloqueo((int)(int)e.Cell.Row.Cells["IdOrdenDeCompra"].Value, 1, cdsSeteoF)) return;
//
//				#region Dentalis
//				if (ODCDiaCredito)
//				{
//					if ((string)e.Cell.Row.Cells["Estado"].Value =="REVISION")
//					{
//						if (e.Cell.Row.Cells["idCuenta"].Value == System.DBNull.Value)
//						{
//							MessageBox.Show("Seleccione la cuenta del Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//
//
//							return;
//						}	
//						if (e.Cell.Row.Cells["Periodo"].Value == System.DBNull.Value)
//						{
//							MessageBox.Show("Seleccione el Periodo afectado para el Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//
//							return;
//						}	
//						if ((int) e.Cell.Row.Cells["idCuenta"].Value == 0)
//						{
//							MessageBox.Show("Seleccione la cuenta del Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//
//							return;
//						}	
//
//						decimal dPresupuesto = Funcion.decEscalarSQL(cdsSeteoF, string.Format(" Exec [ValidaPResupuestoGasto] {0}, '{1}', {2}",
//							(int) e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Periodo"].Value).ToString("yyyyMMdd") , 1));		
//
//						decimal dValorCuenta = Funcion.decEscalarSQL(cdsSeteoF, string.Format(" Exec [ValidaPResupuestoGasto] {0}, '{1}', {2}",
//							(int) e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Periodo"].Value).ToString("yyyyMMdd") , 2));		
//
//						string sCuenta = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Exec [ValidaPResupuestoGasto] {0}, '{1}', {2}",
//							(int) e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Periodo"].Value).ToString("yyyyMMdd") , 3));	
//
//						decimal dODC = (decimal) e.Cell.Row.Cells["Total"].Value;
//						decimal dDisponible = dPresupuesto - (dValorCuenta + dODC);
//						if (dPresupuesto > 0)  e.Cell.Row.Cells["Presupuesto"].Value = dPresupuesto;
//
//						if	(DialogResult.Yes == MessageBox.Show("Cuenta: "+ sCuenta +"\n"+ "Presupuesto : "+ dPresupuesto.ToString("#,##0.00") + "\n"+" Utilizado Cuenta : "+ dValorCuenta.ToString("#,##0.00") + "\n"+" Monto ODC : "+  dODC.ToString("#,##0.00")+ "\n\n\t Disponible " + dDisponible.ToString("#,##0.00") + "\n\n ¿Esta seguro de Aprobar?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
//						{
//							this.DialogResult = DialogResult.OK;
////							this.Close();
//						}
//						else
//						{
//							return;
//						}
//
//					}
//				}
//				#endregion Dentalis
//
//			
////				idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Exec RegistraBloqueoTransacciones {0} , {1} , {2}, {3} ", 0, 1, Convert.ToInt32( e.Cell.Row.Cells["idCompra"].Value),1));	
////
//
//				int iOptEstado = 0;
//        string sSQLNumero="";
//        string sMensaje ="Aprobación";
//
//				if (ODCDiaCredito)
//				{
//					if ((int) this.optEstadoDentalis.Value == 1)sMensaje ="Aprobación Gerencia";
//					if ((int) this.optEstadoDentalis.Value == 6)sMensaje ="Aprobación Presupuestaria";
//				}
//
//
//				using (AprobarOrdenDeCompra ODC  = new AprobarOrdenDeCompra(sMensaje))
//				{
//					if (DialogResult.OK == ODC.ShowDialog())
//					{		
//						iOptEstado = (int)ODC.optEstados.Value;
//						if(ODCDiaCredito)
//						{
//							if((int) this.optEstadoDentalis.Value == 1 && (int)ODC.optEstados.Value == 5 )
//							{
//								iOptEstado = 6;
//							}
//							if((int) this.optEstadoDentalis.Value == 6 && (int)ODC.optEstados.Value == 5 )
//							{
//								iOptEstado = 5;
//
//								string ssSQL = string.Format("Exec NumeracionLocales 76, 0, 0");
//								sSQLNumero = Funcion.sEscalarSQL(cdsSeteoF, ssSQL);	
//
//
//							}
//						}
//						if (!ODCDiaCredito)
//						{
//							e.Cell.Row.Cells["idCuenta"].Value = 0;
//							e.Cell.Row.Cells["Periodo"].Value = "02/02/2015";
//						}
//
//             int iOrigenPresupuesto = 0;
//              iOrigenPresupuesto = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [PresupuestoValidaOrigenODC] {0}", (int)e.Cell.Row.Cells["IdOrdenDeCompra"].Value));	
//				   
//						if ( iOptEstado == 5)
//						{
//							if( iOrigenPresupuesto > 0 ) 
//							{
//								iOptEstado = 3;
//							}
//						}
//						
//					//		string sSQL = string.Format("update OrdenDeCompra set Notas = Notas + ' '+'{0}', Estado = {1}, idCuenta = {3}, Periodo = '{4}', Numero2='{5}' where  idCompra = {2}" ,ODC.txtNotas.Text, iOptEstado, (int)e.Cell.Row.Cells["IdOrdenDeCompra"].Value, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Periodo"].Value).ToString("yyyyMMdd"), sSQLNumero );
//						string sSQL = string.Format(
//							"UPDATE OrdenDeCompra SET Notas = Notas + ' ' + '{0}', Estado = {1} WHERE IdOrdenDeCompra = {2}",
//							ODC.txtNotas.Text,
//							iOptEstado,
//							(int)e.Cell.Row.Cells["IdOrdenDeCompra"].Value													
//							);
//
//						Funcion.EjecutaSQL(cdsSeteoF, sSQL);
//
//						if ( iOrigenPresupuesto > 0 )
//						{
//							if ( iOptEstado == 3)
//							{
//								string sSQLa = string.Format("update Compra set Notas = Notas + ' '+'{0}', Estado = {1} where  idCompra = {2}" ,ODC.txtNotas.Text, 5, iOrigenPresupuesto  );
//								Funcion.EjecutaSQL(cdsSeteoF, sSQLa);
//							}
//							if ( iOptEstado == 4)
//							{
//								string sSQLb = string.Format("update Compra set Notas = Notas + ' '+'{0}', Estado = {1} where  idCompra = {2}" ,ODC.txtNotas.Text, 6, iOrigenPresupuesto  );
//								Funcion.EjecutaSQL(cdsSeteoF, sSQLb);
//							}
//
//						}
//
//						GenerarOrdenesHijas(
//							(int)e.Cell.Row.Cells["IdOrdenDeCompra"].Value
//							);
//
//						this.btnBuscar_Click(sender, e);
//
//						//						string sSQLBloque = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, Convert.ToInt32( e.Cell.Row.Cells["idCompra"].Value));
//						//						Funcion.EjecutaSQL(cdsSeteoF, sSQLBloque, true);
//
//					}
//					
//				}
//			}
//
//			if (e.Cell.Column.ToString() == "Ir")
//			{	
//				using (frmOrdenDeCompra ORDC  = new frmOrdenDeCompra((int)e.Cell.Row.Cells["IdOrdenDeCompra"].Value))
//				{
//					if (DialogResult.OK == ORDC.ShowDialog())
//					{
//						
//
//					}
//				}
//			}
//		}
//

//		private void uGridOrdenesDeCompra_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
//		{
//			if (e.Cell.Column.ToString() == "Aceptar")
//			{
//				int idOrden = Convert.ToInt32(e.Cell.Row.Cells["IdOrdenDeCompra"].Value);
//
//				if (Funcion.ValidaBloqueo(idOrden, 1, cdsSeteoF))
//					return;
//
//				#region VALIDACIÓN DENTALIS
//				if (ODCDiaCredito)
//				{
//					if (Convert.ToString(e.Cell.Row.Cells["Estado"].Value) == "REVISION")
//					{
//						if (e.Cell.Row.Cells["idCuenta"].Value == System.DBNull.Value ||
//							Convert.ToInt32(e.Cell.Row.Cells["idCuenta"].Value) == 0)
//						{
//							MessageBox.Show("Seleccione la cuenta del Gasto");
//							return;
//						}
//
//						if (e.Cell.Row.Cells["Periodo"].Value == System.DBNull.Value)
//						{
//							MessageBox.Show("Seleccione el Periodo del Gasto");
//							return;
//						}
//
//						string periodo = Convert.ToDateTime(e.Cell.Row.Cells["Periodo"].Value).ToString("yyyyMMdd");
//
//						decimal dPresupuesto = Funcion.decEscalarSQL(
//							cdsSeteoF,
//							string.Format("Exec [ValidaPResupuestoGasto] {0}, '{1}', 1",
//							e.Cell.Row.Cells["idCuenta"].Value, periodo));
//
//						decimal dValorCuenta = Funcion.decEscalarSQL(
//							cdsSeteoF,
//							string.Format("Exec [ValidaPResupuestoGasto] {0}, '{1}', 2",
//							e.Cell.Row.Cells["idCuenta"].Value, periodo));
//
//						string sCuenta = Funcion.sEscalarSQL(
//							cdsSeteoF,
//							string.Format("Exec [ValidaPResupuestoGasto] {0}, '{1}', 3",
//							e.Cell.Row.Cells["idCuenta"].Value, periodo));
//
//						decimal dODC = Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value);
//						decimal dDisponible = dPresupuesto - (dValorCuenta + dODC);
//
//						string msg = "Cuenta: " + sCuenta + "\n" +
//							"Presupuesto: " + dPresupuesto.ToString("#,##0.00") + "\n" +
//							"Utilizado: " + dValorCuenta.ToString("#,##0.00") + "\n" +
//							"ODC: " + dODC.ToString("#,##0.00") + "\n" +
//							"Disponible: " + dDisponible.ToString("#,##0.00") +
//							"\n\n¿Desea aprobar?";
//
//						if (MessageBox.Show(msg, "Confirmación",
//							MessageBoxButtons.YesNo,
//							MessageBoxIcon.Question) != DialogResult.Yes)
//						{
//							return;
//						}
//					}
//				}
//				#endregion
//
//				int iOptEstado = 0;
//				string sSQLNumero = "";
//				string sMensaje = "Aprobación";
//
//				if (ODCDiaCredito)
//				{
//					if ((int)this.optEstadoDentalis.Value == 1)
//						sMensaje = "Aprobación Gerencia";
//
//					if ((int)this.optEstadoDentalis.Value == 6)
//						sMensaje = "Aprobación Presupuestaria";
//				}
//
//				using (AprobarOrdenDeCompra ODC = new AprobarOrdenDeCompra(sMensaje))
//				{
//					if (ODC.ShowDialog() != DialogResult.OK)
//						return;
//
//					iOptEstado = Convert.ToInt32(ODC.optEstados.Value);
//
//					#region AJUSTE ESTADOS DENTALIS
//					if (ODCDiaCredito)
//					{
//						if ((int)this.optEstadoDentalis.Value == 1 && iOptEstado == 5)
//							iOptEstado = 6;
//
//						if ((int)this.optEstadoDentalis.Value == 6 && iOptEstado == 5)
//						{
//							iOptEstado = 5;
//
//							sSQLNumero = Funcion.sEscalarSQL(
//								cdsSeteoF,
//								"Exec NumeracionLocales 76, 0, 0");
//						}
//					}
//					#endregion
//
//					string sSQL = string.Format(
//						"UPDATE OrdenDeCompra SET Notas = Notas + ' {0}', Estado = {1} WHERE IdOrdenDeCompra = {2}",
//						ODC.txtNotas.Text,
//						iOptEstado,
//						idOrden);
//
//					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
//
//					int iOrigen = Funcion.iEscalarSQL(
//						cdsSeteoF,
//						string.Format("Exec [PresupuestoValidaOrigenODC] {0}", idOrden));
//
//					if (iOrigen > 0)
//					{
//						string sqlCompra = "";
//
//						if (iOptEstado == 3)
//						{
//							sqlCompra = string.Format(
//								"UPDATE Compra SET Notas = Notas + ' {0}', Estado = 5 WHERE idCompra = {1}",
//								ODC.txtNotas.Text, iOrigen);
//						}
//						else if (iOptEstado == 4)
//						{
//							sqlCompra = string.Format(
//								"UPDATE Compra SET Notas = Notas + ' {0}', Estado = 6 WHERE idCompra = {1}",
//								ODC.txtNotas.Text, iOrigen);
//						}
//
//						if (sqlCompra != "")
//							Funcion.EjecutaSQL(cdsSeteoF, sqlCompra);
//					}
//
//					#region GENERAR ORDENES HIJAS
//					try
//					{
//						if (iOptEstado == 3 || iOptEstado == 5 || iOptEstado == 6)
//						{
//							GenerarOrdenesHijasDesdeAprobacion(idOrden);
//						}
//					}
//					catch (Exception ex)
//					{
//						MessageBox.Show(
//							"Orden aprobada, pero error generando hijas:\n" + ex.Message);
//					}
//					#endregion
//
//					this.btnBuscar_Click(sender, e);
//				}
//			}
//
//			if (e.Cell.Column.ToString() == "Ir")
//			{
//				frmOrdenDeCompra frm =
//					new frmOrdenDeCompra(
//					Convert.ToInt32(e.Cell.Row.Cells["IdOrdenDeCompra"].Value));
//
//				frm.ShowDialog();
//			}
//		}
//
//		private void GenerarOrdenesHijasDesdeAprobacion(int idOrdenCompra)
//		{
//			OrdenCompraBodegaGeneracionService service =
//				new OrdenCompraBodegaGeneracionService(this.cdsSeteoF);
//
//			SqlConnection oConexion =
//				new SqlConnection(
//				"Password=Bl45o6$9;Persist Security Info=True;User ID=" +
//				MenuLatinium.stUsuario +
//				";Initial Catalog=" +
//				MenuLatinium.stDirFacturacion +
//				";Data Source=" +
//				MenuLatinium.stDirServidor +
//				"; max pool size=1024"
//				);
//
//			oConexion.Open();
//
//			SqlTransaction oTransaction =
//				oConexion.BeginTransaction();
//
//			SqlCommand oCmd =
//				oConexion.CreateCommand();
//
//			oCmd.Connection = oConexion;
//			oCmd.Transaction = oTransaction;
//			oCmd.CommandType = CommandType.Text;
//			oCmd.CommandTimeout = 120;
//
//			try
//			{
//				ArrayList ocGeneradas =
//					new ArrayList();
//
//				Hashtable comprasPorBodega =
//					new Hashtable();
//
//				#region OBTENER CABECERA ODC
//
//				DataTable dtCabecera =
//					FuncionesProcedimientos.dtGeneral(
//					"SELECT * " +
//					"FROM OrdenDeCompra " +
//					"WHERE IdOrdenDeCompra = " + idOrdenCompra
//					);
//
//				if (dtCabecera == null ||
//					dtCabecera.Rows.Count == 0)
//				{
//					throw new Exception(
//						"No se encontró la Orden de Compra."
//						);
//				}
//
//				DataRow drCabecera =
//					dtCabecera.Rows[0];
//
//				int idComprobante = 0;
//
//				if (drCabecera["IdComprobante"] != DBNull.Value)
//				{
//					idComprobante =
//						Convert.ToInt32(
//						drCabecera["IdComprobante"]
//						);
//				}
//
//				int idProveedor = 0;
//
//				if (drCabecera["IdCliente"] != DBNull.Value)
//				{
//					idProveedor =
//						Convert.ToInt32(
//						drCabecera["IdCliente"]
//						);
//				}
//
//				int idFormaPago = 9;
//
//				if (drCabecera.Table.Columns.Contains("idFormaPago") &&
//					drCabecera["idFormaPago"] != DBNull.Value)
//				{
//					idFormaPago =
//						Convert.ToInt32(
//						drCabecera["idFormaPago"]
//						);
//				}
//
//				decimal dPorcentajeIVA =
//					Funcion.decEscalarSQL(
//					cdsSeteoF,
//					string.Format(
//					"SELECT dbo.RetornaImpuestoIVAVigente('{0}',1)",
//					DateTime.Now.ToString("yyyyMMdd")
//					)
//					);
//
//				#endregion
//
//
//				#region RECORRER DETALLES
//
//				DataTable dt =
//					FuncionesProcedimientos.dtGeneral(
//					"SELECT * " +
//					"FROM DetOrdenDeCompra " +
//					"WHERE IdOrdenDeCompra = " + idOrdenCompra
//					);
//
//				foreach (DataRow row in dt.Rows)
//				{
//					int idDetOrdenDeCompra =
//						Convert.ToInt32(
//						row["IdDetOrdenDeCompra"]
//						);
//
//					int idArticulo =
//						Convert.ToInt32(
//						row["IdArticulo"]
//						);
//
//					decimal costo =
//						Convert.ToDecimal(
//						row["Costo"]
//						);
//
//					string notas = "";
//
//					if (row["Notas"] != DBNull.Value)
//					{
//						notas =
//							row["Notas"].ToString();
//					}
//
//					notas =
//						notas.Replace("'", "''");
//
//					#region CONSULTAR BODEGAS
//
//					DataTable dtBodegas =
//						service.Consultar(
//						idDetOrdenDeCompra,
//						oConexion,
//						oTransaction
//						);
//
//					if (dtBodegas == null ||
//						dtBodegas.Rows.Count == 0)
//					{
//						continue;
//					}
//
//					#endregion
//
//
//					#region RECORRER BODEGAS
//
//					Hashtable procesados =
//						new Hashtable();
//
//					foreach (DataRow drBodega in dtBodegas.Rows)
//					{
//						int bodega =
//							Convert.ToInt32(
//							drBodega["Bodega"]
//							);
//
//						string key =
//							idDetOrdenDeCompra.ToString() +
//							"_" +
//							bodega.ToString();
//
//						if (procesados.ContainsKey(key))
//							continue;
//
//						procesados.Add(key, true);
//
//						int cantidadTemporal = 0;
//
//						if (drBodega["CantidadTemporal"] != DBNull.Value)
//						{
//							cantidadTemporal =
//								Convert.ToInt32(
//								drBodega["CantidadTemporal"]
//								);
//						}
//
//						int saldoPendiente = 0;
//
//						if (drBodega["SaldoPendiente"] != DBNull.Value)
//						{
//							saldoPendiente =
//								Convert.ToInt32(
//								drBodega["SaldoPendiente"]
//								);
//						}
//
//						if (cantidadTemporal <= 0)
//							continue;
//
//						if (saldoPendiente <= 0)
//							continue;
//
//						if (cantidadTemporal > saldoPendiente)
//						{
//							cantidadTemporal =
//								saldoPendiente;
//						}
//
//						#endregion
//
//
//						#region CALCULAR TOTALES
//
//						decimal subTotal =
//							(costo * cantidadTemporal);
//
//						if (subTotal <= 0)
//							continue;
//
//						#endregion
//
//
//						#region OBTENER O CREAR OC POR BODEGA
//
//						int idCompraNueva = 0;
//
//						if (comprasPorBodega.ContainsKey(bodega))
//						{
//							idCompraNueva =
//								Convert.ToInt32(
//								comprasPorBodega[bodega]
//								);
//						}
//						else
//						{
//							oCmd.CommandText =
//								"Exec FacturaNumeroEstacion 2";
//
//							string numero =
//								Convert.ToString(
//								oCmd.ExecuteScalar()
//								);
//
//							#region CREAR CABECERA OC
//
//							oCmd.CommandText =
//								string.Format(
//								@"
//						EXEC GrabaMaestroOrdenDeCompra
//							0,
//							2,
//							{0},
//							0,
//							0,
//							{1},
//							{2},
//							'{3}',
//							'{4}',
//							NULL,
//							NULL,
//							0,
//							5,
//							'{5}',
//							0,
//							0,
//							0,
//							0,
//							0,
//							0,
//							0,
//							0,
//							0,
//							0,
//							0,
//							{6}
//						",
//								idComprobante,
//								idProveedor,
//								bodega,
//								numero,
//								DateTime.Now.ToString("yyyyMMdd"),
//								notas,
//								idOrdenCompra
//								);
//
//							idCompraNueva =
//								Convert.ToInt32(
//								oCmd.ExecuteScalar()
//								);
//
//							#endregion
//
//
//							#region COMPLETAR CAMPOS COMPRA
//
//							oCmd.CommandText =
//								string.Format(
//								@"
//						UPDATE Compra
//						SET
//							Usuario = '{1}',
//							FechaIngreso = GETDATE(),
//							FechaVencimiento = Fecha,
//
//							PorcentajeIva = {2},
//
//							Contado = 0,
//							Credito = 0,
//							Saldo = 0,
//
//							SubTotal = 0,
//							SubtotalIva = 0,
//							SubtotalExcento = 0,
//
//							Iva = 0,
//							Total = 0,
//
//							Aprobado = 1,
//							Procesado = 0,
//
//							Nuevo = 1,
//							Editar = 0,
//							Borrar = 0,
//
//							idTipoFactura = 5,
//							idFormaPago = {3},
//
//							ContadoCredito = 2,
//
//							Estado = 5,
//
//							Comprobante = Numero,
//							Numero2 = Numero,
//
//							iODC = 1,
//
//							FechaEmision = Fecha
//
//						WHERE idCompra = {0}
//						",
//								idCompraNueva,
//								MenuLatinium.stUsuario.Replace("'", "''"),
//								dPorcentajeIVA.ToString(
//								System.Globalization.CultureInfo.InvariantCulture),
//								idFormaPago
//								);
//
//							oCmd.ExecuteNonQuery();
//
//							#endregion
//
//
//							comprasPorBodega.Add(
//								bodega,
//								idCompraNueva
//								);
//
//							ocGeneradas.Add(
//								new object[]
//						{
//							idCompraNueva,
//							bodega,
//							numero
//						}
//								);
//						}
//
//						#endregion
//
//
//						#region CREAR DETALLE
//
//						oCmd.CommandText =
//							string.Format(
//							@"
//					EXEC GuardaDetalleOrdenDeCompra
//						0,
//						{0},
//						{1},
//						0,
//						{2},
//						0,
//						{3},
//						15,
//						0,
//						0,
//						{4},
//						19,
//						0,
//						{5},
//						0,
//						0,
//						0,
//						0,
//						0,
//						0,
//						0,
//						0,
//						0,
//						0
//					",
//							idCompraNueva,
//							idArticulo,
//							cantidadTemporal,
//							costo.ToString(
//							System.Globalization.CultureInfo.InvariantCulture),
//							subTotal.ToString(
//							System.Globalization.CultureInfo.InvariantCulture),
//							bodega
//							);
//
//						oCmd.ExecuteNonQuery();
//
//						#endregion
//
//
//						#region ACTUALIZAR GENERADO
//
//						int idOrdenDeCompraBodegaDetalle =
//							Convert.ToInt32(
//							drBodega["idOrdenDeCompraBodegaDetalle"]
//							);
//
//						oCmd.CommandText =
//							string.Format(
//							@"
//					UPDATE OrdenDeCompraBodegaDetalle
//					SET
//						CantidadGenerada =
//							ISNULL(CantidadGenerada,0) + {0},
//
//						CantidadTemporal = 0
//
//					WHERE idOrdenDeCompraBodegaDetalle = {1}
//					",
//							cantidadTemporal,
//							idOrdenDeCompraBodegaDetalle
//							);
//
//						oCmd.ExecuteNonQuery();
//
//						#endregion
//					}
//				}
//
//				#endregion
//
//
//				#region RECALCULAR TOTALES CABECERA
//
//				foreach (DictionaryEntry item in comprasPorBodega)
//				{
//					int idCompra =
//						Convert.ToInt32(
//						item.Value
//						);
//
//					oCmd.CommandText =
//						string.Format(
//						@"
//				DECLARE @Subtotal DECIMAL(18,2)
//				DECLARE @IVA DECIMAL(18,2)
//				DECLARE @TOTAL DECIMAL(18,2)
//
//				SET @Subtotal =
//					ISNULL(
//					(
//						SELECT SUM(Subtotal)
//						FROM DetCompra
//						WHERE idCompra = {0}
//					),0)
//
//				SET @IVA =
//					ROUND((@Subtotal * {1}) / 100.00,2)
//
//				SET @TOTAL =
//					@Subtotal + @IVA
//
//				UPDATE Compra
//				SET
//					SubTotal = @Subtotal,
//					SubtotalIva = @Subtotal,
//					SubtotalExcento = 0,
//
//					Iva = @IVA,
//					Total = @TOTAL,
//
//					Saldo = @TOTAL,
//					Credito = @TOTAL,
//
//					PorcentajeIva = {1},
//
//					FechaIngreso = GETDATE(),
//					FechaVencimiento = Fecha,
//
//					Aprobado = 1,
//					Procesado = 0,
//
//					Estado = 5,
//
//					iODC = 1
//
//				WHERE idCompra = {0}
//				",
//						idCompra,
//						dPorcentajeIVA.ToString(
//						System.Globalization.CultureInfo.InvariantCulture)
//						);
//
//					oCmd.ExecuteNonQuery();
//				}
//
//				#endregion
//
//
//				#region ACTUALIZAR ESTADO ODC PADRE
//
//				oCmd.CommandText =
//					string.Format(
//					@"
//			IF EXISTS
//			(
//				SELECT 1
//				FROM OrdenDeCompraBodegaDetalle odbd
//					INNER JOIN DetOrdenDeCompra d
//						ON d.IdDetOrdenDeCompra = odbd.IdDetOrdenDeCompra
//				WHERE d.IdOrdenDeCompra = {0}
//				AND ISNULL(odbd.SaldoPendiente,0) > 0
//			)
//			BEGIN
//				UPDATE OrdenDeCompra
//				SET Estado = 5
//				WHERE IdOrdenDeCompra = {0}
//			END
//			ELSE
//			BEGIN
//				UPDATE OrdenDeCompra
//				SET Estado = 9
//				WHERE IdOrdenDeCompra = {0}
//			END
//			",
//					idOrdenCompra
//					);
//
//				oCmd.ExecuteNonQuery();
//
//				#endregion
//
//
//				oTransaction.Commit();
//			}
//			catch
//			{
//				oTransaction.Rollback();
//				throw;
//			}
//			finally
//			{
//				oConexion.Close();
//			}
//		}

//		private void uGridOrdenesDeCompra_ClickCellButton(
//			object sender,
//			Infragistics.Win.UltraWinGrid.CellEventArgs e)
//		{
//			try
//			{
//				#region ACEPTAR
//
//				if (e.Cell.Column.ToString() == "Aceptar")
//				{
//					int idOrden =
//						Convert.ToInt32(
//						e.Cell.Row.Cells["IdOrdenDeCompra"].Value);
//
//					if (Funcion.ValidaBloqueo(idOrden, 1, cdsSeteoF))
//						return;
//
//					#region VALIDACIÓN DENTALIS
//
//					if (ODCDiaCredito)
//					{
//						if (Convert.ToString(
//							e.Cell.Row.Cells["Estado"].Value) == "REVISION")
//						{
//							if (e.Cell.Row.Cells["idCuenta"].Value == DBNull.Value ||
//								Convert.ToInt32(
//								e.Cell.Row.Cells["idCuenta"].Value) == 0)
//							{
//								MessageBox.Show(
//									"Seleccione la cuenta del Gasto");
//
//								return;
//							}
//
//							if (e.Cell.Row.Cells["Periodo"].Value == DBNull.Value)
//							{
//								MessageBox.Show(
//									"Seleccione el Periodo del Gasto");
//
//								return;
//							}
//
//							string periodo =
//								Convert.ToDateTime(
//								e.Cell.Row.Cells["Periodo"].Value)
//								.ToString("yyyyMMdd");
//
//							decimal dPresupuesto =
//								Funcion.decEscalarSQL(
//								cdsSeteoF,
//								string.Format(
//								"Exec [ValidaPResupuestoGasto] {0}, '{1}', 1",
//								e.Cell.Row.Cells["idCuenta"].Value,
//								periodo));
//
//							decimal dValorCuenta =
//								Funcion.decEscalarSQL(
//								cdsSeteoF,
//								string.Format(
//								"Exec [ValidaPResupuestoGasto] {0}, '{1}', 2",
//								e.Cell.Row.Cells["idCuenta"].Value,
//								periodo));
//
//							string sCuenta =
//								Funcion.sEscalarSQL(
//								cdsSeteoF,
//								string.Format(
//								"Exec [ValidaPResupuestoGasto] {0}, '{1}', 3",
//								e.Cell.Row.Cells["idCuenta"].Value,
//								periodo));
//
//							decimal dODC =
//								Convert.ToDecimal(
//								e.Cell.Row.Cells["Total"].Value);
//
//							decimal dDisponible =
//								dPresupuesto -
//								(dValorCuenta + dODC);
//
//							string msg =
//								"Cuenta: " + sCuenta + "\n" +
//								"Presupuesto: " +
//								dPresupuesto.ToString("#,##0.00") + "\n" +
//								"Utilizado: " +
//								dValorCuenta.ToString("#,##0.00") + "\n" +
//								"ODC: " +
//								dODC.ToString("#,##0.00") + "\n" +
//								"Disponible: " +
//								dDisponible.ToString("#,##0.00") +
//								"\n\n¿Desea aprobar?";
//
//							if (MessageBox.Show(
//								msg,
//								"Confirmación",
//								MessageBoxButtons.YesNo,
//								MessageBoxIcon.Question)
//								!= DialogResult.Yes)
//							{
//								return;
//							}
//						}
//					}
//
//					#endregion
//
//					int iOptEstado = 0;
//
//					string sMensaje =
//						"Aprobación";
//
//					if (ODCDiaCredito)
//					{
//						if ((int)this.optEstadoDentalis.Value == 1)
//							sMensaje = "Aprobación Gerencia";
//
//						if ((int)this.optEstadoDentalis.Value == 6)
//							sMensaje = "Aprobación Presupuestaria";
//					}
//
//					using (AprobarOrdenDeCompra ODC =
//									 new AprobarOrdenDeCompra(sMensaje))
//					{
//						if (ODC.ShowDialog() != DialogResult.OK)
//							return;
//
//						iOptEstado =
//							Convert.ToInt32(
//							ODC.optEstados.Value);
//
//						#region AJUSTE ESTADOS
//
//						if (ODCDiaCredito)
//						{
//							if ((int)this.optEstadoDentalis.Value == 1 &&
//								iOptEstado == 5)
//							{
//								iOptEstado = 6;
//							}
//
//							if ((int)this.optEstadoDentalis.Value == 6 &&
//								iOptEstado == 5)
//							{
//								iOptEstado = 5;
//							}
//						}
//
//						#endregion
//
//						#region ACTUALIZAR ODC
//
//						string sNotas =
//							ODC.txtNotas.Text
//							.Replace("'", "''");
//
//						string sSQL =
//							string.Format(
//							@"
//					UPDATE OrdenDeCompra
//					SET
//						Notas =
//							ISNULL(Notas,'') +
//							CHAR(13) + CHAR(10) +
//							'{0}',
//
//						Estado = {1}
//
//					WHERE IdOrdenDeCompra = {2}
//					",
//							sNotas,
//							iOptEstado,
//							idOrden);
//
//						Funcion.EjecutaSQL(
//							cdsSeteoF,
//							sSQL);
//
//						#endregion
//
//						#region ACTUALIZAR COMPRA ORIGEN
//
//						int iOrigen =
//							Funcion.iEscalarSQL(
//							cdsSeteoF,
//							string.Format(
//							"Exec PresupuestoValidaOrigenODC {0}",
//							idOrden));
//
//						if (iOrigen > 0)
//						{
//							string sqlCompra =
//								"";
//
//							if (iOptEstado == 3)
//							{
//								sqlCompra =
//									string.Format(
//									@"
//							UPDATE Compra
//							SET
//								Notas =
//									ISNULL(Notas,'') +
//									CHAR(13)+CHAR(10)+
//									'{0}',
//
//								Estado = 5
//
//							WHERE idCompra = {1}
//							",
//									sNotas,
//									iOrigen);
//							}
//							else if (iOptEstado == 4)
//							{
//								sqlCompra =
//									string.Format(
//									@"
//							UPDATE Compra
//							SET
//								Notas =
//									ISNULL(Notas,'') +
//									CHAR(13)+CHAR(10)+
//									'{0}',
//
//								Estado = 6
//
//							WHERE idCompra = {1}
//							",
//									sNotas,
//									iOrigen);
//							}
//
//							if (sqlCompra != "")
//							{
//								Funcion.EjecutaSQL(
//									cdsSeteoF,
//									sqlCompra);
//							}
//						}
//
//						#endregion
//
//						#region GENERAR OCS HIJAS
//
//						try
//						{
//							if (iOptEstado == 3 ||
//								iOptEstado == 5 ||
//								iOptEstado == 6)
//							{
//								GenerarOrdenesHijasDesdeAprobacion(
//									idOrden);
//							}
//						}
//						catch (Exception ex)
//						{
//							MessageBox.Show(
//								"Error generando órdenes hijas:\n\n" +
//								ex.Message,
//								"Error",
//								MessageBoxButtons.OK,
//								MessageBoxIcon.Error);
//
//							return;
//						}
//
//						#endregion
//
//						this.btnBuscar_Click(sender, e);
//					}
//				}
//
//				#endregion
//
//				#region IR
//
//				if (e.Cell.Column.ToString() == "Ir")
//				{
//					frmOrdenDeCompra frm =
//						new frmOrdenDeCompra(
//						Convert.ToInt32(
//						e.Cell.Row.Cells["IdOrdenDeCompra"].Value));
//
//					frm.ShowDialog();
//				}
//
//				#endregion
//			}
//			catch (Exception ex)
//			{
//				MessageBox.Show(
//					ex.Message,
//					"Error",
//					MessageBoxButtons.OK,
//					MessageBoxIcon.Error);
//			}
//		}



		private void uGridOrdenesDeCompra_ClickCellButton(
			object sender,
			Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{

			#region Visualizacion
			if (e.Cell.Column.ToString() == "Ir")
			{	
				using (frmOrdenDeCompra ORDC  = new frmOrdenDeCompra((int)e.Cell.Row.Cells["IdOrdenDeCompra"].Value))
				{
					if (DialogResult.OK == ORDC.ShowDialog())
					{
					}
				}
			}
			#endregion

			#region Aceptar
			if (e.Cell.Column.ToString() == "Aceptar")
			{
				try
				{

					int idOrden =
						Convert.ToInt32(
						e.Cell.Row.Cells["IdOrdenDeCompra"].Value);

					if (Funcion.ValidaBloqueo(idOrden, 1, cdsSeteoF))
						return;

					#region VALIDACIONES

					// TODAS TUS VALIDACIONES ACTUALES

					#endregion

					int iOptEstado = 0;

					using (AprobarOrdenDeCompra ODC =
									 new AprobarOrdenDeCompra("Aprobación"))
					{
						if (ODC.ShowDialog() != DialogResult.OK)
							return;

						iOptEstado =
							Convert.ToInt32(
							ODC.optEstados.Value);

						string sConexion =
							"Password=Bl45o6$9;" +
							"Persist Security Info=True;" +
							"User ID=" + MenuLatinium.stUsuario + ";" +
							"Initial Catalog=" + MenuLatinium.stDirFacturacion + ";" +
							"Data Source=" + MenuLatinium.stDirServidor + ";" +
							"max pool size=1024";

						using (SqlConnection oConexion =
										 new SqlConnection(sConexion))
						{
							oConexion.Open();

							using (SqlTransaction oTransaction =
											 oConexion.BeginTransaction(
											 IsolationLevel.ReadCommitted))
							{
								try
								{
									#region ACTUALIZAR ODC

									string sNotas =
										ODC.txtNotas.Text
										.Replace("'", "''");

									SqlCommand cmdODC =
										new SqlCommand(
										"UPDATE OrdenDeCompra " +
										"SET " +
										"Notas = ISNULL(Notas,'') + CHAR(13)+CHAR(10) + @Notas, " +
										"Estado = @Estado " +
										"WHERE IdOrdenDeCompra = @IdOrden",
										oConexion,
										oTransaction);

									cmdODC.Parameters.Add(
										"@Notas",
										SqlDbType.VarChar).Value =
										sNotas;

									cmdODC.Parameters.Add(
										"@Estado",
										SqlDbType.Int).Value =
										iOptEstado;

									cmdODC.Parameters.Add(
										"@IdOrden",
										SqlDbType.Int).Value =
										idOrden;

									cmdODC.ExecuteNonQuery();

									cmdODC.Dispose();

									#endregion

									#region ACTUALIZAR COMPRA ORIGEN

									int iOrigen = 0;

									SqlCommand cmdOrigen =
										new SqlCommand(
										"Exec PresupuestoValidaOrigenODC @IdOrden",
										oConexion,
										oTransaction);

									cmdOrigen.Parameters.Add(
										"@IdOrden",
										SqlDbType.Int).Value =
										idOrden;

									object objOrigen =
										cmdOrigen.ExecuteScalar();

									if (objOrigen != null &&
										objOrigen != DBNull.Value)
									{
										iOrigen =
											Convert.ToInt32(
											objOrigen);
									}

									cmdOrigen.Dispose();

									if (iOrigen > 0)
									{
										SqlCommand cmdCompra =
											new SqlCommand(
											"UPDATE Compra " +
											"SET Estado = @Estado " +
											"WHERE idCompra = @idCompra",
											oConexion,
											oTransaction);

										if (iOptEstado == 3)
										{
											cmdCompra.Parameters.Add(
												"@Estado",
												SqlDbType.Int).Value = 5;
										}
										else
										{
											cmdCompra.Parameters.Add(
												"@Estado",
												SqlDbType.Int).Value = 6;
										}

										cmdCompra.Parameters.Add(
											"@idCompra",
											SqlDbType.Int).Value =
											iOrigen;

										cmdCompra.ExecuteNonQuery();

										cmdCompra.Dispose();
									}

									#endregion

									#region GENERAR HIJAS

									ArrayList odcGeneradas = null;

									if (iOptEstado == 3 ||
										iOptEstado == 5 ||
										iOptEstado == 6)
									{
										odcGeneradas = GenerarOrdenesHijasDesdeAprobacion(
											idOrden,
											oConexion,
											oTransaction);
									}

									#endregion

									oTransaction.Commit();

									string mensaje = "Proceso realizado correctamente.";

									if (odcGeneradas != null && odcGeneradas.Count > 0)
									{
										mensaje += "\n\nODC generadas:";

										for (int i = 0; i < odcGeneradas.Count; i++)
										{
											mensaje += "\n- " + odcGeneradas[i].ToString();
										}

										mensaje += "\n\nTotal ODC generadas: " + odcGeneradas.Count;
									}
									else
									{
										mensaje += "\nNo se generaron ODC hijas.";
									}

									MessageBox.Show(
										mensaje,
										"OK",
										MessageBoxButtons.OK,
										MessageBoxIcon.Information);
								}
								catch
								{
									try
									{
										oTransaction.Rollback();
									}
									catch
									{
									}

									throw;
								}
							}
						}

						this.btnBuscar_Click(sender, e);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show(
						ex.Message,
						"Error",
						MessageBoxButtons.OK,
						MessageBoxIcon.Error);
				}
			}
			#endregion
		}
//		private void GenerarOrdenesHijasDesdeAprobacion(
//			int idOrdenCompra,
//			SqlConnection oConexion,
//			SqlTransaction oTransaction)
//		{
//			OrdenCompraBodegaGeneracionService service =
//				new OrdenCompraBodegaGeneracionService(
//				this.cdsSeteoF);
//
//			Hashtable comprasPorBodega =
//				new Hashtable();
//
//			#region IVA
//
//			decimal dPorcentajeIVA =
//				Funcion.decEscalarSQL(
//				cdsSeteoF,
//				string.Format(
//				"SELECT dbo.RetornaImpuestoIVAVigente('{0}',1)",
//				DateTime.Now.ToString("yyyyMMdd")
//				));
//
//			#endregion
//
//			#region CABECERA ODC
//
//			DataTable dtCabecera =
//				new DataTable();
//
//			SqlCommand cmdCab =
//				new SqlCommand(
//				"SELECT * " +
//				"FROM OrdenDeCompra " +
//				"WHERE IdOrdenDeCompra = @IdOrden",
//				oConexion,
//				oTransaction);
//
//			cmdCab.Parameters.Add(
//				"@IdOrden",
//				SqlDbType.Int).Value =
//				idOrdenCompra;
//
//			SqlDataAdapter daCab =
//				new SqlDataAdapter(cmdCab);
//
//			daCab.Fill(dtCabecera);
//
//			daCab.Dispose();
//			cmdCab.Dispose();
//
//			if (dtCabecera.Rows.Count == 0)
//			{
//				throw new Exception(
//					"No se encontró la ODC.");
//			}
//
//			DataRow drCabecera =
//				dtCabecera.Rows[0];
//
//			int idProveedor = 0;
//
//			if (drCabecera["IdCliente"] != DBNull.Value)
//			{
//				idProveedor =
//					Convert.ToInt32(
//					drCabecera["IdCliente"]);
//			}
//
//			int idComprobante = 0;
//
//			if (drCabecera["IdComprobante"] != DBNull.Value)
//			{
//				idComprobante =
//					Convert.ToInt32(
//					drCabecera["IdComprobante"]);
//			}
//
//			DateTime fechaOrden =
//				drCabecera["Fecha"] == DBNull.Value
//				? DateTime.Now
//				: Convert.ToDateTime(
//				drCabecera["Fecha"]);
//
//			DateTime fechaTentativaLlegada =
//				DateTime.MinValue;
//
//			bool tieneFechaTentativaLlegada =
//				false;
//
//			if (drCabecera["FechaTentativaLlegada"] != System.DBNull.Value)
//			{
//				fechaTentativaLlegada =
//					Convert.ToDateTime(
//					drCabecera["FechaTentativaLlegada"]);
//
//				tieneFechaTentativaLlegada =
//					true;
//			}
//
//			string notasCabecera =
//				drCabecera["Notas"] == DBNull.Value
//				? ""
//				: drCabecera["Notas"].ToString();
//
//			notasCabecera =
//				notasCabecera.Replace("'", "''");
//
//			#endregion
//
//			#region DETALLES
//
//			DataTable dtDetalles =
//				new DataTable();
//
//			SqlCommand cmdDet =
//				new SqlCommand(
//				"SELECT * " +
//				"FROM DetOrdenDeCompra " +
//				"WHERE IdOrdenDeCompra = @IdOrden",
//				oConexion,
//				oTransaction);
//
//			cmdDet.Parameters.Add(
//				"@IdOrden",
//				SqlDbType.Int).Value =
//				idOrdenCompra;
//
//			SqlDataAdapter daDet =
//				new SqlDataAdapter(cmdDet);
//
//			daDet.Fill(dtDetalles);
//
//			daDet.Dispose();
//			cmdDet.Dispose();
//
//			#endregion
//
//			foreach (DataRow row in dtDetalles.Rows)
//			{
//				int idDetOrdenDeCompra =
//					Convert.ToInt32(
//					row["IdDetOrdenDeCompra"]);
//
//				int idArticulo =
//					Convert.ToInt32(
//					row["IdArticulo"]);
//
//				decimal costo =
//					Convert.ToDecimal(
//					row["Costo"]);
//
//				DataTable dtBodegas =
//					service.Consultar(
//					idDetOrdenDeCompra,
//					oConexion,
//					oTransaction);
//
//				if (dtBodegas == null ||
//					dtBodegas.Rows.Count == 0)
//				{
//					continue;
//				}
//
//				foreach (DataRow drBodega in dtBodegas.Rows)
//				{
//					int bodega =
//						Convert.ToInt32(
//						drBodega["Bodega"]);
//
//					int cantidadTemporal =
//						drBodega["CantidadTemporal"] == DBNull.Value
//						? 0
//						: Convert.ToInt32(
//						drBodega["CantidadTemporal"]);
//
//					int saldoPendiente =
//						drBodega["SaldoPendiente"] == DBNull.Value
//						? 0
//						: Convert.ToInt32(
//						drBodega["SaldoPendiente"]);
//
//					if (cantidadTemporal <= 0)
//						continue;
//
//					if (saldoPendiente <= 0)
//						continue;
//
//					if (cantidadTemporal > saldoPendiente)
//					{
//						cantidadTemporal =
//							saldoPendiente;
//					}
//
//					decimal subTotal =
//						costo * cantidadTemporal;
//
//					int idCompraNueva = 0;
//
//					#region CABECERA COMPRA
//
//					if (comprasPorBodega.ContainsKey(bodega))
//					{
//						idCompraNueva =
//							Convert.ToInt32(
//							comprasPorBodega[bodega]);
//					}
//					else
//					{
//						string numero = "";
//
//						SqlCommand cmdNumero =
//							new SqlCommand(
//							"Exec FacturaNumeroEstacion 2",
//							oConexion,
//							oTransaction);
//
//						object objNumero =
//							cmdNumero.ExecuteScalar();
//
//						if (objNumero != DBNull.Value &&
//							objNumero != null)
//						{
//							numero =
//								Convert.ToString(
//								objNumero);
//						}
//
//						cmdNumero.Dispose();
//
//						SqlCommand cmdCompra =
//							new SqlCommand(
//							"Exec GrabaMaestroOrdenDeCompra " +
//							"0," +                  // @idCompra
//							"2," +                  // @ContadoCredito
//							"@idComprobante," +     // @idComprobante
//							"0," +                  // @idCredTributario
//							"0," +                  // @Consignacion
//							"@idProveedor," +       // @idCliente
//							"@bodega," +           // @Bodega
//							"@numero," +           // @Numero
//							"@fecha," +            // @Fecha
//							"NULL," +              // @FechaTentativaLlegada
//							"NULL," +              // @HoraLlegada
//							"0," +                 // @idImportacionTipo
//							"5," +                 // @Estado
//							"@notas," +            // @Notas
//							"0," +                 // @SubtotalExcento
//							"0," +                 // @Descuento0
//							"0," +                 // @SubtotalIva
//							"0," +                 // @Descuento
//							"0," +                 // @Iva
//							"0," +                 // @Total
//							"0," +                 // @bFechaLlegada
//							"0," +                 // @bHoraLlegada
//							"@idFormaPago," +      // @idFormaPago
//							"0," +                 // @Credito1
//							"0," +                 // @Credito2
//							"0," +                 // @Credito3
//							"0," +                 // @Credito4
//							"@idOrigen",           // @idOrigen
//							oConexion,
//							oTransaction);
//
//						int idFormaPago = 0;
//						cmdCompra.Parameters.Add(
//							"@idFormaPago",
//							SqlDbType.Int).Value =
//							idFormaPago;
//
//						cmdCompra.Parameters.Add(
//							"@idComprobante",
//							SqlDbType.Int).Value =
//							idComprobante;
//
//						cmdCompra.Parameters.Add(
//							"@idProveedor",
//							SqlDbType.Int).Value =
//							idProveedor;
//
//						cmdCompra.Parameters.Add(
//							"@bodega",
//							SqlDbType.Int).Value =
//							bodega;
//
//						cmdCompra.Parameters.Add(
//							"@numero",
//							SqlDbType.VarChar).Value =
//							numero;
//
//						cmdCompra.Parameters.Add(
//							"@fecha",
//							SqlDbType.DateTime).Value =
//							fechaOrden;
//
//						cmdCompra.Parameters.Add(
//							"@notas",
//							SqlDbType.VarChar).Value =
//							notasCabecera;
//
//						cmdCompra.Parameters.Add(
//							"@idOrigen",
//							SqlDbType.Int).Value =
//							idOrdenCompra;
//
//						object objCompra =
//							cmdCompra.ExecuteScalar();
//
//						if (objCompra != DBNull.Value &&
//							objCompra != null)
//						{
//							idCompraNueva =
//								Convert.ToInt32(
//								objCompra);
//						}
//
//						cmdCompra.Dispose();
//
//						SqlCommand cmdUpdate =
//							new SqlCommand(
//							"UPDATE Compra " +
//							"SET " +
//							"Usuario = @Usuario, " +
//							"Fecha = @Fecha, " +
//							"FechaIngreso = GETDATE(), " +
//							"FechaVencimiento = @Fecha, " +
//							"FechaTentativaLlegada = @FechaTentativaLlegada, " +
//							"Notas = @Notas, " +
//							"PorcentajeIva = @IVA, " +
//							"Aprobado = 1, " +
//							"Procesado = 0, " +
//							"iODC = 1, " +
//							"Comprobante = Numero, " +
//							"Numero2 = Numero, " +
//							"Contado = 0, " +
//							"Credito = 0, " +
//							"Saldo = 0, " +
//							"SubtotalIva = 0, " +
//							"SubtotalExcento = 0, " +
//							"P_IdOrigen = @IdOrden, " +
//							"Estado = 5 " +
//							"WHERE idCompra = @idCompra",
//							oConexion,
//							oTransaction);
//
//						cmdUpdate.Parameters.Add(
//							"@Usuario",
//							SqlDbType.VarChar).Value =
//							MenuLatinium.stUsuario;
//
//						cmdUpdate.Parameters.Add(
//							"@IVA",
//							SqlDbType.Decimal).Value =
//							dPorcentajeIVA;
//
//						cmdUpdate.Parameters.Add(
//							"@Fecha",
//							SqlDbType.DateTime).Value =
//							fechaOrden;
//
//						cmdUpdate.Parameters.Add(
//							"@FechaTentativaLlegada",
//							SqlDbType.DateTime).Value =
//							tieneFechaTentativaLlegada
//							? (object)fechaTentativaLlegada
//							: System.DBNull.Value;
//
//						cmdUpdate.Parameters.Add(
//							"@Notas",
//							SqlDbType.VarChar).Value =
//							notasCabecera;
//
//						cmdUpdate.Parameters.Add(
//							"@IdOrden",
//							SqlDbType.Int).Value =
//							idOrdenCompra;
//
//						cmdUpdate.Parameters.Add(
//							"@idCompra",
//							SqlDbType.Int).Value =
//							idCompraNueva;
//
//						cmdUpdate.ExecuteNonQuery();
//
//						cmdUpdate.Dispose();
//
//						comprasPorBodega.Add(
//							bodega,
//							idCompraNueva);
//					}
//
//					#endregion
//
//					#region DETALLE COMPRA
//
//					SqlCommand cmdDetalle =
//						new SqlCommand(
//						"Exec GuardaDetalleOrdenDeCompra " +
//						"0," +						// @idDetCompra
//						"@idCompra," +				// @idCompra
//						"@idArticulo," +			// @idArticulo
//						"0," +						// @idUnidad
//						"@cantidad," +				// @Cantidad
//						"0," +						// @CantAereo
//						"@costo," +				// @Precio
//						"0," +						// @Impuesto
//						"0," +						// @DescuentoArt
//						"0," +						// @DescuentoPorc
//						"@subtotal," +				// @SubTotal
//						"0," +						// @idProyecto
//						"0," +						// @idImportacion
//						"@bodega," +				// @Bodega
//						"0," +						// @Signo
//						"1," +						// @Tipo
//						"0," +						// @EstadoSerial
//						"0," +						// @Courier
//						"0," +						// @Peso
//						"0," +						// @Arancel
//						"0," +						// @idTipoCalculo
//						"0," +						// @DescNC
//						"0," +						// @Valores
//						"0",						// @color
//						oConexion,
//						oTransaction);
//
//					cmdDetalle.Parameters.Add(
//						"@idCompra",
//						SqlDbType.Int).Value =
//						idCompraNueva;
//
//					cmdDetalle.Parameters.Add(
//						"@idArticulo",
//						SqlDbType.Int).Value =
//						idArticulo;
//
//					cmdDetalle.Parameters.Add(
//						"@cantidad",
//						SqlDbType.Decimal).Value =
//						cantidadTemporal;
//
//					cmdDetalle.Parameters.Add(
//						"@costo",
//						SqlDbType.Decimal).Value =
//						costo;
//
//					SqlParameter parCosto =
//						cmdDetalle.Parameters["@costo"];
//
//					parCosto.Precision = 18;
//					parCosto.Scale = 5;
//
//					cmdDetalle.Parameters.Add(
//						"@subtotal",
//						SqlDbType.Decimal).Value =
//						subTotal;
//
//					SqlParameter parSubtotal =
//						cmdDetalle.Parameters["@subtotal"];
//
//					parSubtotal.Precision = 18;
//					parSubtotal.Scale = 4;
//
//					cmdDetalle.Parameters.Add(
//						"@bodega",
//						SqlDbType.Int).Value =
//						bodega;
//
//					cmdDetalle.ExecuteNonQuery();
//
//					cmdDetalle.Dispose();
//
//					#endregion
//
//					#region ACTUALIZAR BODEGA
//
//					SqlCommand cmdUpd =
//						new SqlCommand(
//						"UPDATE OrdenDeCompraBodegaDetalle " +
//						"SET " +
//						"CantidadGenerada = " +
//						"ISNULL(CantidadGenerada,0) + @Cantidad, " +
//						"CantidadTemporal = 0 " +
//						"WHERE " +
//						"idOrdenDeCompraBodegaDetalle = @IdDetalle",
//						oConexion,
//						oTransaction);
//
//					cmdUpd.Parameters.Add(
//						"@Cantidad",
//						SqlDbType.Int).Value =
//						cantidadTemporal;
//
//					cmdUpd.Parameters.Add(
//						"@IdDetalle",
//						SqlDbType.Int).Value =
//						Convert.ToInt32(
//						drBodega["idOrdenDeCompraBodegaDetalle"]);
//
//					cmdUpd.ExecuteNonQuery();
//
//					cmdUpd.Dispose();
//
//					#endregion
//				}
//			}
//
//			#region RECALCULAR
//
//			foreach (DictionaryEntry item in comprasPorBodega)
//			{
//				int idCompra =
//					Convert.ToInt32(
//					item.Value);
//
//				SqlCommand cmdTot =
//					new SqlCommand(
//					"DECLARE @Subtotal DECIMAL(18,2) " +
//					"DECLARE @IVA DECIMAL(18,2) " +
//					"DECLARE @TOTAL DECIMAL(18,2) " +
//
//					"SET @Subtotal = " +
//					"ISNULL( " +
//					"( " +
//					"	SELECT SUM(Subtotal) " +
//					"	FROM DetCompra " +
//					"	WHERE idCompra = @idCompra " +
//					"),0) " +
//
//					"SET @IVA = " +
//					"ROUND((@Subtotal * @PorcIVA) / 100.00,2) " +
//
//					"SET @TOTAL = " +
//					"@Subtotal + @IVA " +
//
//					"UPDATE Compra " +
//					"SET " +
//					"SubTotal = @Subtotal, " +
//					"SubtotalIva = @Subtotal, " +
//					"SubtotalExcento = 0, " +
//					"Iva = @IVA, " +
//					"Total = @TOTAL, " +
//					"Saldo = @TOTAL, " +
//					"Credito = @TOTAL, " +
//					"Contado = 0, " +
//					"PorcentajeIva = @PorcIVA " +
//					"WHERE idCompra = @idCompra",
//					oConexion,
//					oTransaction);
//
//				cmdTot.Parameters.Add(
//					"@idCompra",
//					SqlDbType.Int).Value =
//					idCompra;
//
//				cmdTot.Parameters.Add(
//					"@PorcIVA",
//					SqlDbType.Decimal).Value =
//					dPorcentajeIVA;
//
//				SqlParameter parIVA =
//					cmdTot.Parameters["@PorcIVA"];
//
//				parIVA.Precision = 18;
//				parIVA.Scale = 4;
//
//				cmdTot.ExecuteNonQuery();
//
//				cmdTot.Dispose();
//			}
//
//			#endregion
//
//			#region ESTADO PADRE
//
//			SqlCommand cmdEstado =
//				new SqlCommand(
//				"IF EXISTS " +
//				"( " +
//				"	SELECT 1 " +
//				"	FROM OrdenDeCompraBodegaDetalle odbd " +
//				"		INNER JOIN DetOrdenDeCompra d " +
//				"			ON d.IdDetOrdenDeCompra = " +
//				"				odbd.IdDetOrdenDeCompra " +
//				"	WHERE d.IdOrdenDeCompra = @IdOrden " +
//				"	AND ISNULL(odbd.SaldoPendiente,0) > 0 " +
//				") " +
//				"BEGIN " +
//				"	UPDATE OrdenDeCompra " +
//				"	SET Estado = 5 " +
//				"	WHERE IdOrdenDeCompra = @IdOrden " +
//				"END " +
//				"ELSE " +
//				"BEGIN " +
//				"	UPDATE OrdenDeCompra " +
//				"	SET Estado = 9 " +
//				"	WHERE IdOrdenDeCompra = @IdOrden " +
//				"END",
//				oConexion,
//				oTransaction);
//
//			cmdEstado.Parameters.Add(
//				"@IdOrden",
//				SqlDbType.Int).Value =
//				idOrdenCompra;
//
//			cmdEstado.ExecuteNonQuery();
//
//			cmdEstado.Dispose();
//
//			#endregion
//		}


		private const int ID_TIPO_FACTURA_ODC_HIJA = 44;
		private ArrayList GenerarOrdenesHijasDesdeAprobacion(
			int idOrdenCompra,
			SqlConnection oConexion,
			SqlTransaction oTransaction)
		{
			OrdenCompraBodegaGeneracionService service =
				new OrdenCompraBodegaGeneracionService(this.cdsSeteoF);

			Hashtable comprasPorBodega = new Hashtable();
			ArrayList ordenesGeneradas = new ArrayList();

			decimal dPorcentajeIVA = ObtenerIVA(oConexion, oTransaction);

			int idProveedor;
			int idComprobante;
			DateTime fechaOrden;
			DateTime fechaTentativaLlegada;
			bool tieneFechaTentativaLlegada;
			string notasCabecera;
			bool consignacion;

			DataRow drCabecera = ObtenerCabecera(
				idOrdenCompra,
				oConexion,
				oTransaction,
				out idProveedor,
				out idComprobante,
				out fechaOrden,
				out fechaTentativaLlegada,
				out tieneFechaTentativaLlegada,
				out notasCabecera,
				out consignacion);

			DataTable dtDetalles = ObtenerDetalles(
				idOrdenCompra,
				oConexion,
				oTransaction);

			ProcesarDetalles(
				dtDetalles,
				service,
				comprasPorBodega,
				ordenesGeneradas,
				oConexion,
				oTransaction,
				idProveedor,
				idComprobante,
				fechaOrden,
				fechaTentativaLlegada,
				tieneFechaTentativaLlegada,
				notasCabecera,
				dPorcentajeIVA,
				idOrdenCompra,
				consignacion);

			RecalcularCompras(
				comprasPorBodega,
				oConexion,
				oTransaction,
				dPorcentajeIVA);

			return ordenesGeneradas;
		}


		private decimal ObtenerIVA(SqlConnection conn, SqlTransaction tx)
		{
			string fecha = DateTime.Now.ToString("yyyyMMdd");

			string sql = "SELECT dbo.RetornaImpuestoIVAVigente('" + fecha + "',1)";

			return Funcion.decEscalarSQL(
				cdsSeteoF,
				sql
				);
		}

		
		private DataRow ObtenerCabecera(
			int idOrdenCompra,
			SqlConnection conn,
			SqlTransaction tx,
			out int idProveedor,
			out int idComprobante,
			out DateTime fechaOrden,
			out DateTime fechaTentativaLlegada,
			out bool tieneFechaTentativaLlegada,
			out string notasCabecera,
			out bool consignacion)
		{
			DataTable dt = new DataTable();

			SqlCommand cmd = new SqlCommand(
				"SELECT * FROM OrdenDeCompra WHERE IdOrdenDeCompra = @IdOrden",
				conn,
				tx);

			cmd.Parameters.Add("@IdOrden", SqlDbType.Int).Value = idOrdenCompra;

			SqlDataAdapter da = new SqlDataAdapter(cmd);
			da.Fill(dt);

			if (dt.Rows.Count == 0)
				throw new Exception("No se encontró la ODC.");

			DataRow dr = dt.Rows[0];

			// Proveedor
			if (dr["IdCliente"] == DBNull.Value)
				idProveedor = 0;
			else
				idProveedor = Convert.ToInt32(dr["IdCliente"]);

			// Comprobante
			if (dr["IdComprobante"] == DBNull.Value)
				idComprobante = 0;
			else
				idComprobante = Convert.ToInt32(dr["IdComprobante"]);

			// Fecha orden
			if (dr["Fecha"] == DBNull.Value)
				fechaOrden = DateTime.Now;
			else
				fechaOrden = Convert.ToDateTime(dr["Fecha"]);

			// Fecha tentativa llegada
			tieneFechaTentativaLlegada = (dr["FechaTentativaLlegada"] != DBNull.Value);

			if (tieneFechaTentativaLlegada)
				fechaTentativaLlegada = Convert.ToDateTime(dr["FechaTentativaLlegada"]);
			else
				fechaTentativaLlegada = DateTime.MinValue;

			// Notas
			if (dr["Notas"] == DBNull.Value)
				notasCabecera = "";
			else
				notasCabecera = dr["Notas"].ToString();

			// Consignación
		  consignacion =
				dr["Consignacion"] != DBNull.Value
				&& Convert.ToBoolean(dr["Consignacion"]);

			// Escape SQL clásico (si realmente lo necesitas)
			notasCabecera = notasCabecera.Replace("'", "''");

			return dr;
		}


		private DataTable ObtenerDetalles(
			int idOrdenCompra,
			SqlConnection conn,
			SqlTransaction tx)
		{
			DataTable dt = new DataTable();

			SqlCommand cmd = new SqlCommand(
				"SELECT * FROM DetOrdenDeCompra WHERE IdOrdenDeCompra = @IdOrden",
				conn,
				tx);

			cmd.Parameters.Add("@IdOrden", SqlDbType.Int).Value = idOrdenCompra;

			SqlDataAdapter da = new SqlDataAdapter(cmd);
			da.Fill(dt);

			return dt;
		}


		private void ProcesarDetalles(
			DataTable dtDetalles,
			OrdenCompraBodegaGeneracionService service,
			Hashtable comprasPorBodega,
			ArrayList ordenesGeneradas,
			SqlConnection conn,
			SqlTransaction tx,
			int idProveedor,
			int idComprobante,
			DateTime fechaOrden,
			DateTime fechaTentativaLlegada,
			bool tieneFechaTentativaLlegada,
			string notasCabecera,
			decimal dPorcentajeIVA,
			int idOrdenCompra,
			bool consignacion)
		{
			if (dtDetalles == null || dtDetalles.Rows.Count == 0)
				return;

			foreach (DataRow row in dtDetalles.Rows)
			{
				int idDet = 0;
				int idArticulo = 0;
				decimal costo = 0;

				if (row["IdDetOrdenDeCompra"] != DBNull.Value)
					idDet = Convert.ToInt32(row["IdDetOrdenDeCompra"]);

				if (row["IdArticulo"] != DBNull.Value)
					idArticulo = Convert.ToInt32(row["IdArticulo"]);

				if (row["Costo"] != DBNull.Value)
					costo = Convert.ToDecimal(row["Costo"]);

				DataTable dtBodegas = service.Consultar(idDet, conn, tx);

				if (dtBodegas == null || dtBodegas.Rows.Count == 0)
					continue;

				foreach (DataRow drBodega in dtBodegas.Rows)
				{
					ProcesarBodega(
						drBodega,
						comprasPorBodega,
						ordenesGeneradas,
						conn,
						tx,
						idProveedor,
						idComprobante,
						fechaOrden,
						fechaTentativaLlegada,
						tieneFechaTentativaLlegada,
						notasCabecera,
						dPorcentajeIVA,
						idOrdenCompra,
						idArticulo,
						costo,
						consignacion);
				}
			}
		}


		private void ProcesarBodega(
			DataRow drBodega,
			Hashtable comprasPorBodega,
			ArrayList ordenesGeneradas,
			SqlConnection conn,
			SqlTransaction tx,
			int idProveedor,
			int idComprobante,
			DateTime fechaOrden,
			DateTime fechaTentativaLlegada,
			bool tieneFechaTentativaLlegada,
			string notasCabecera,
			decimal dPorcentajeIVA,
			int idOrdenCompra,
			int idArticulo,
			decimal costo,
			bool consignacion)
		{
			if (drBodega == null)
				return;

			int bodega = 0;
			int cantidadTemporal = 0;
			int saldoPendiente = 0;
			int idProyecto = 0;

			if (drBodega["Bodega"] != DBNull.Value)
				bodega = Convert.ToInt32(drBodega["Bodega"]);

			if (drBodega["idProyecto"] != DBNull.Value)
				idProyecto = Convert.ToInt32(drBodega["idProyecto"]);

			if (drBodega["CantidadTemporal"] != DBNull.Value)
				cantidadTemporal = Convert.ToInt32(drBodega["CantidadTemporal"]);

			if (drBodega["SaldoPendiente"] != DBNull.Value)
				saldoPendiente = Convert.ToInt32(drBodega["SaldoPendiente"]);

			// Validaciones de negocio
			if (cantidadTemporal <= 0 || saldoPendiente <= 0)
				return;

			// Ajuste de saldo
			if (cantidadTemporal > saldoPendiente)
				cantidadTemporal = saldoPendiente;

			decimal subtotal = 0;

			if (cantidadTemporal > 0 && costo > 0)
				subtotal = costo * cantidadTemporal;

			int idCompra = ObtenerOCrearCompra(
				bodega,
				comprasPorBodega,
				ordenesGeneradas,
				conn,
				tx,
				idProveedor,
				idComprobante,
				fechaOrden,
				fechaTentativaLlegada,
				tieneFechaTentativaLlegada,
				notasCabecera,
				dPorcentajeIVA,
				idOrdenCompra,
				consignacion);

			InsertarDetalleCompra(
				conn,
				tx,
				idCompra,
				idArticulo,
				cantidadTemporal,
				costo,
				subtotal,
				bodega,
				idProyecto,
				dPorcentajeIVA);

			ActualizarBodegaDetalle(
				conn,
				tx,
				drBodega,
				cantidadTemporal);
		}


		private const int ESTADO_COMPRA_PENDIENTE = 1;
		private const int ESTADO_COMPRA_APROBADA = 5;

		private int ObtenerOCrearCompra(
			int bodega,
			Hashtable comprasPorBodega,
			ArrayList ordenesGeneradas,
			SqlConnection conn,
			SqlTransaction tx,
			int idProveedor,
			int idComprobante,
			DateTime fechaOrden,
			DateTime fechaTentativaLlegada,
			bool tieneFechaTentativaLlegada,
			string notasCabecera,
			decimal iva,
			int idOrdenCompra,
			bool consignacion)
		{
			if (comprasPorBodega.ContainsKey(bodega))
				return Convert.ToInt32(comprasPorBodega[bodega]);

			int idCompraNueva = 0;

			// 1. Obtener número
			string numero = "";

			SqlCommand cmdNumero = new SqlCommand(
				"Exec FacturaNumeroEstacion 2",
				conn,
				tx);

			object objNumero = cmdNumero.ExecuteScalar();

			if (objNumero != null && objNumero != DBNull.Value)
				numero = objNumero.ToString();

			cmdNumero.Dispose();

			// 2. Crear compra
			SqlCommand cmdCompra =
				new SqlCommand(
				"Exec GrabaMaestroOrdenDeCompra " +
				"0," +                  // @idCompra
				"2," +                  // @ContadoCredito
				"@idComprobante," +     // @idComprobante
				"0," +                  // @idCredTributario
				"@consignacion," +                  // @Consignacion
				"@idProveedor," +       // @idCliente
				"@bodega," +           // @Bodega
				"@numero," +           // @Numero
				"@fecha," +            // @Fecha
				"NULL," +              // @FechaTentativaLlegada
				"NULL," +              // @HoraLlegada
				"0," +                 // @idImportacionTipo
				"@estado," +                 // @Estado
				"@notas," +            // @Notas
				"0," +                 // @SubtotalExcento
				"0," +                 // @Descuento0
				"0," +                 // @SubtotalIva
				"0," +                 // @Descuento
				"0," +                 // @Iva
				"0," +                 // @Total
				"0," +                 // @bFechaLlegada
				"0," +                 // @bHoraLlegada
				"@idFormaPago," +      // @idFormaPago
				"0," +                 // @Credito1
				"0," +                 // @Credito2
				"0," +                 // @Credito3
				"0," +                 // @Credito4
				"@idOrigen",           // @idOrigen
				conn,
				tx);

			int idFormaPago = 6;
			int estadoNuevo = 0;
			estadoNuevo = consignacion ? ESTADO_COMPRA_PENDIENTE : ESTADO_COMPRA_APROBADA;

			cmdCompra.Parameters.Add("@consignacion", SqlDbType.Bit).Value = consignacion;
			cmdCompra.Parameters.Add("@estado", SqlDbType.Int).Value = estadoNuevo;
			cmdCompra.Parameters.Add("@idFormaPago", SqlDbType.Int).Value = idFormaPago;
			cmdCompra.Parameters.Add("@idComprobante", SqlDbType.Int).Value = idComprobante;
			cmdCompra.Parameters.Add("@idProveedor", SqlDbType.Int).Value = idProveedor;
			cmdCompra.Parameters.Add("@bodega", SqlDbType.Int).Value = bodega;
			cmdCompra.Parameters.Add("@numero", SqlDbType.VarChar).Value = numero;
			cmdCompra.Parameters.Add("@fecha", SqlDbType.DateTime).Value = fechaOrden;
			cmdCompra.Parameters.Add("@notas", SqlDbType.VarChar).Value = notasCabecera;
			cmdCompra.Parameters.Add("@idOrigen", SqlDbType.Int).Value = idOrdenCompra;

			object objCompra = cmdCompra.ExecuteScalar();

			if (objCompra != null && objCompra != DBNull.Value)
			{
				idCompraNueva = Convert.ToInt32(objCompra);

				string nombreBodega = "";

				SqlCommand cmdBodega = new SqlCommand(
					"SELECT Nombre FROM Bodega WHERE Bodega = @idBodega",
					conn,
					tx);

				cmdBodega.Parameters.Add(
					"@idBodega",
					SqlDbType.Int).Value = bodega;

				object objBodega = cmdBodega.ExecuteScalar();

				if (objBodega != null && objBodega != DBNull.Value)
				{
					nombreBodega = objBodega.ToString();
				}

				cmdBodega.Dispose();

				string textoODC =
					numero + " | " + nombreBodega;

				if (!ordenesGeneradas.Contains(textoODC))
				{
					ordenesGeneradas.Add(textoODC);
				}
			}

			cmdCompra.Dispose();

			// 3. Actualizar campos adicionales
			SqlCommand cmdUpdate = new SqlCommand(
				"UPDATE Compra SET " +
				"Usuario=@Usuario, Fecha=@Fecha, FechaIngreso=GETDATE(), " +
				"FechaVencimiento=@Fecha, FechaTentativaLlegada=@FechaTentativaLlegada, " +
				"Notas=@Notas, PorcentajeIva=@IVA, Aprobado=1, Procesado=0, iODC=1, " +
				"Estado=@Estado, P_IdOrigen=@IdOrden " +
				"WHERE idCompra=@idCompra",
				conn,
				tx);

			cmdUpdate.Parameters.Add("@Estado", SqlDbType.Int).Value = estadoNuevo;

			cmdUpdate.Parameters.Add("@Usuario", SqlDbType.VarChar).Value = MenuLatinium.stUsuario;
			cmdUpdate.Parameters.Add("@Fecha", SqlDbType.DateTime).Value = fechaOrden;

			cmdUpdate.Parameters.Add("@FechaTentativaLlegada", SqlDbType.DateTime).Value =
				tieneFechaTentativaLlegada ? (object)fechaTentativaLlegada : DBNull.Value;

			cmdUpdate.Parameters.Add("@Notas", SqlDbType.VarChar).Value = notasCabecera;
			cmdUpdate.Parameters.Add("@IVA", SqlDbType.Decimal).Value = iva;
			cmdUpdate.Parameters.Add("@IdOrden", SqlDbType.Int).Value = idOrdenCompra;
			cmdUpdate.Parameters.Add("@idCompra", SqlDbType.Int).Value = idCompraNueva;

			cmdUpdate.ExecuteNonQuery();
			cmdUpdate.Dispose();

			comprasPorBodega.Add(bodega, idCompraNueva);

			return idCompraNueva;
		}


		private void InsertarDetalleCompra(
			SqlConnection conn,
			SqlTransaction tx,
			int idCompra,
			int idArticulo,
			int cantidad,
			decimal costo,
			decimal subtotal,
			int bodega,
			int idProyecto,
			decimal impuesto)
		{
			OrdenDeCompraService service =
				new OrdenDeCompraService(cdsSeteoF);

			// 1. Guardar detalle (SP encapsulado)
			service.GuardarDetalleOrdenCompra(
				conn,
				tx,
				idCompra,
				idArticulo,
				cantidad,
				costo,
				subtotal,
				bodega,
				idProyecto,
				impuesto);

			// 2. Inicializar control de facturación
			service.InicializarControlFacturacion(
				conn,
				tx,
				idCompra,
				idArticulo,
				bodega,
				cantidad);
		}


		private void ActualizarBodegaDetalle(
			SqlConnection conn,
			SqlTransaction tx,
			DataRow drBodega,
			int cantidadTemporal)
		{
			if (drBodega == null)
				return;

			int idDetalle = 0;

			if (drBodega["idOrdenDeCompraBodegaDetalle"] != DBNull.Value)
				idDetalle = Convert.ToInt32(drBodega["idOrdenDeCompraBodegaDetalle"]);

			SqlCommand cmd = new SqlCommand(
				"UPDATE OrdenDeCompraBodegaDetalle SET " +
				"CantidadGenerada = ISNULL(CantidadGenerada,0) + @Cantidad, " +
				"CantidadTemporal = 0 " +
				"WHERE idOrdenDeCompraBodegaDetalle = @Id",
				conn,
				tx);

			cmd.Parameters.Add("@Cantidad", SqlDbType.Int).Value = cantidadTemporal;
			cmd.Parameters.Add("@Id", SqlDbType.Int).Value = idDetalle;

			cmd.ExecuteNonQuery();
			cmd.Dispose();
		}


		private void RecalcularCompras(
			Hashtable comprasPorBodega,
			SqlConnection conn,
			SqlTransaction tx,
			decimal iva)
		{
			if (comprasPorBodega == null || comprasPorBodega.Count == 0)
				return;

			foreach (DictionaryEntry item in comprasPorBodega)
			{
				int idCompra = Convert.ToInt32(item.Value);

				SqlCommand cmdTot = new SqlCommand(
					"DECLARE @Subtotal DECIMAL(18,2) " +
					"DECLARE @IVA DECIMAL(18,2) " +
					"DECLARE @TOTAL DECIMAL(18,2) " +

					"SET @Subtotal = ISNULL( " +
					"(SELECT SUM(Subtotal) FROM DetCompra WHERE idCompra = @idCompra),0) " +

					"SET @IVA = ROUND((@Subtotal * @PorcIVA) / 100.00,2) " +

					"SET @TOTAL = @Subtotal + @IVA " +

					"UPDATE Compra SET " +
					"SubTotal = @Subtotal, " +
					"SubtotalIva = @Subtotal, " +
					"SubtotalExcento = 0, " +
					"Iva = @IVA, " +
					"Total = @TOTAL, " +
					"Saldo = @TOTAL, " +
					"Credito = @TOTAL, " +
					"Contado = 0, " +
					"PorcentajeIva = @PorcIVA " +
					"WHERE idCompra = @idCompra",
					conn,
					tx);

				cmdTot.Parameters.Add("@idCompra", SqlDbType.Int).Value = idCompra;

				cmdTot.Parameters.Add("@PorcIVA", SqlDbType.Decimal).Value = iva;
				SqlParameter parIVA = cmdTot.Parameters["@PorcIVA"];
				parIVA.Precision = 18;
				parIVA.Scale = 4;

				cmdTot.ExecuteNonQuery();
				cmdTot.Dispose();
			}
		}


//		private void ActualizarEstadoPadre(
//			int idOrdenCompra,
//			SqlConnection conn,
//			SqlTransaction tx)
//		{
//			SqlCommand cmd = new SqlCommand(
//				"IF EXISTS " +
//				"( " +
//				"   SELECT 1 " +
//				"   FROM OrdenDeCompraBodegaDetalle odbd " +
//				"       INNER JOIN DetOrdenDeCompra d " +
//				"           ON d.IdDetOrdenDeCompra = odbd.IdDetOrdenDeCompra " +
//				"   WHERE d.IdOrdenDeCompra = @IdOrden " +
//				"   AND ISNULL(odbd.SaldoPendiente,0) > 0 " +
//				") " +
//				"BEGIN " +
//				"   UPDATE OrdenDeCompra " +
//				"   SET Estado = 5 " +
//				"   WHERE IdOrdenDeCompra = @IdOrden " +
//				"END " +
//				"ELSE " +
//				"BEGIN " +
//				"   UPDATE OrdenDeCompra " +
//				"   SET Estado = 9 " +
//				"   WHERE IdOrdenDeCompra = @IdOrden " +
//				"END",
//				conn,
//				tx);
//
//			cmd.Parameters.Add("@IdOrden", SqlDbType.Int).Value = idOrdenCompra;
//
//			cmd.ExecuteNonQuery();
//			cmd.Dispose();
//		}



		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridOrdenesDeCompra);
		}

		private void uGridOrdenesDeCompra_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void optEstadoDentalis_ValueChanged(object sender, System.EventArgs e)
		{
			if (miAcceso.Imprimir) this.optEstadoDentalis.Value = 1;
      EstadosDentalis();
			this.btnBuscar_Click (sender, e);
		}
	}
}
