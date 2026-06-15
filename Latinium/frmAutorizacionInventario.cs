using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions.Shared;
using C1.Data;
using System.Drawing.Printing;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmAutorizacionInventario.
	/// </summary>
	public class frmAutorizacionInventario : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDocumento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridAutorizacion;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAutorizacionInventario()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAutorizacionInventario));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDocumento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSolicitud");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioSolicita");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioAutorizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAutorizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasSolicitud");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasAprobacion");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDocumento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioSolicita");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioAutorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAutorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasAprobacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizar", 0);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anular", 1);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Imprimir", 2);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Documento", 3);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, "", "Total", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 10, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnExcel = new System.Windows.Forms.Button();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbDocumento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.uGridAutorizacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridAutorizacion)).BeginInit();
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
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(192, 10);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 286;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 285;
			this.label6.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(232, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 284;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(72, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 283;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1544, 6);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(88, 24);
			this.btnExcel.TabIndex = 662;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel2.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel2.Location = new System.Drawing.Point(19, 35);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(30, 15);
			this.ultraLabel2.TabIndex = 978;
			this.ultraLabel2.Text = "Local";
			// 
			// cmbBodega
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance3;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 8;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 264;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(72, 32);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(264, 21);
			this.cmbBodega.TabIndex = 979;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(-552, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(2296, 8);
			this.groupBox1.TabIndex = 980;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(432, 81);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 1009;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel1.Location = new System.Drawing.Point(16, 83);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 15);
			this.ultraLabel1.TabIndex = 1008;
			this.ultraLabel1.Text = "Buscar";
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(344, 79);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 1007;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// txtBusqueda
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance4;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(72, 80);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(264, 21);
			this.txtBusqueda.TabIndex = 1006;
			// 
			// optEstados
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance5;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance6;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendientes";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Aprobado";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Anulados";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(384, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(312, 29);
			this.optEstados.TabIndex = 1010;
			this.optEstados.Text = "Todos";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(3, 58);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(62, 16);
			this.label2.TabIndex = 1011;
			this.label2.Text = "Documento";
			// 
			// cmbDocumento
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDocumento.Appearance = appearance7;
			this.cmbDocumento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDocumento.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 264;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbDocumento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbDocumento.DisplayMember = "Nombre";
			this.cmbDocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDocumento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDocumento.Location = new System.Drawing.Point(72, 56);
			this.cmbDocumento.Name = "cmbDocumento";
			this.cmbDocumento.Size = new System.Drawing.Size(264, 21);
			this.cmbDocumento.TabIndex = 1013;
			this.cmbDocumento.ValueMember = "idTipoFactura";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.DateTime);
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
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.AcceptsTab = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance8;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(712, 8);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(824, 24);
			this.txtNotas.TabIndex = 1018;
			this.txtNotas.Visible = false;
			// 
			// uGridAutorizacion
			// 
			this.uGridAutorizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridAutorizacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridAutorizacion.DataSource = this.ultraDataSource1;
			appearance9.BackColor = System.Drawing.Color.White;
			this.uGridAutorizacion.DisplayLayout.Appearance = appearance9;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Tipo Documento";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 125;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 165;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.Caption = "Fecha Solicitud";
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn8.Width = 120;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Usuario Solicita";
			ultraGridColumn9.Header.VisiblePosition = 11;
			ultraGridColumn9.Width = 125;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Usuario Autorización";
			ultraGridColumn10.Header.VisiblePosition = 14;
			ultraGridColumn10.Width = 125;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn11.Header.Caption = "Fecha Autorizacion";
			ultraGridColumn11.Header.VisiblePosition = 13;
			ultraGridColumn11.Width = 120;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn12.Width = 111;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Width = 140;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.Caption = "Notas Solicitud";
			ultraGridColumn14.Header.VisiblePosition = 12;
			ultraGridColumn14.Width = 350;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance10;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 9;
			ultraGridColumn15.Width = 100;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.Caption = "Notas Aprobación";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 350;
			ultraGridColumn17.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance11.ForegroundAlpha = Infragistics.Win.Alpha.Transparent;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			appearance11.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn17.CellButtonAppearance = appearance11;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn17.Width = 40;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance12.ForegroundAlpha = Infragistics.Win.Alpha.Transparent;
			appearance12.Image = ((object)(resources.GetObject("appearance12.Image")));
			appearance12.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn18.CellButtonAppearance = appearance12;
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 45;
			ultraGridColumn19.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance13.ForegroundAlpha = Infragistics.Win.Alpha.Transparent;
			appearance13.Image = ((object)(resources.GetObject("appearance13.Image")));
			appearance13.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn19.CellButtonAppearance = appearance13;
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 55;
			ultraGridColumn20.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance14.ForegroundAlpha = Infragistics.Win.Alpha.Transparent;
			appearance14.Image = ((object)(resources.GetObject("appearance14.Image")));
			appearance14.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn20.CellButtonAppearance = appearance14;
			ultraGridColumn20.Header.Caption = "Doc";
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn20.Width = 45;
			ultraGridBand3.Columns.AddRange(new object[] {
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
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance15;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand3.SummaryFooterCaption = "Total";
			this.uGridAutorizacion.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridAutorizacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridAutorizacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridAutorizacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridAutorizacion.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAutorizacion.DisplayLayout.Override.CellAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8.5F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridAutorizacion.DisplayLayout.Override.HeaderAppearance = appearance18;
			this.uGridAutorizacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAutorizacion.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAutorizacion.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.uGridAutorizacion.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridAutorizacion.Location = new System.Drawing.Point(8, 128);
			this.uGridAutorizacion.Name = "uGridAutorizacion";
			this.uGridAutorizacion.Size = new System.Drawing.Size(1624, 256);
			this.uGridAutorizacion.TabIndex = 1017;
			this.uGridAutorizacion.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridAutorizacion_InitializeRow);
			this.uGridAutorizacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridAutorizacion_ClickCellButton);
			// 
			// frmAutorizacionInventario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1640, 390);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.uGridAutorizacion);
			this.Controls.Add(this.cmbDocumento);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmAutorizacionInventario";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Autorizacion Movimiento Inventario";
			this.Load += new System.EventHandler(this.frmAutorizacionInventario_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridAutorizacion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bEgreso = false;
		bool bNotaEntrega = false;
		bool bActivo = false;

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmAutorizacionInventario_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "112801");

			if (!Funcion.Permiso("2016", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Autorizar Egresos / Notas de Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExcel.Enabled = true;

			#region Fecha
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtDesde.Value = DateTime.Today;

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;
			#endregion Fecha

			#region Combos
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre From Bodega Where Inventario = 1 Order by Nombre");
			this.cmbDocumento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoFactura, Nombre From CompraNumero Where idTipoFactura in (8, 51) Order by Nombre");
			#endregion Combos

			this.btnBuscar_Click (sender, e);
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			#region validaciones
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la fecha inicial", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la fecha final", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;
			#endregion validaciones

			#region Inicializa Valores
			int idLocal = 0;
			int idTipoFactura = 0;
			this.lblContador.Text = "";

			DateTime dtFeIni = (DateTime)this.dtDesde.Value;
			DateTime dtFeFin = (DateTime)this.dtHasta.Value;

			if (this.cmbBodega.ActiveRow != null) idLocal = (int)this.cmbBodega.Value;
			if (this.cmbDocumento.ActiveRow != null) idTipoFactura = (int)this.cmbDocumento.Value;
			#endregion Inicializa Valores

			string sSQL = string.Format("Exec ConsultaEgresosNotasEntregaBoega '{0}', '{1}', {2}, {3}, {4}, '{5}'", dtFeIni.ToString("yyyyMMdd"), dtFeFin.ToString("yyyyMMdd"), (int)this.optEstados.Value,
				idTipoFactura, idLocal, this.txtBusqueda.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridAutorizacion);

			this.lblContador.Text = this.uGridAutorizacion.Rows.Count + " Registros Encontrados";

			this.Controles();
		}

		private void Controles()
		{
			string sSQLCI = string.Format("Exec ConsultaUsuarioAutorizacion {0}", MenuLatinium.IdUsuario);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLCI, true);
			dr.Read();
			if (dr.HasRows)
			{
				bEgreso = dr.GetBoolean(0);
				bNotaEntrega = dr.GetBoolean(1);
				bActivo = dr.GetBoolean(2);
			}
			dr.Close();

			if((int)this.optEstados.Value == 0)
			{
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Autorizar"].Hidden = true;
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Anular"].Hidden = true;
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = true;
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Documento"].Hidden = false;				
			}
			if((int)this.optEstados.Value == 1)
			{
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Autorizar"].Hidden = false;			
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Anular"].Hidden = false;
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = false;
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Documento"].Hidden = false;
			}
			if((int)this.optEstados.Value == 2)
			{
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Autorizar"].Hidden = true;
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Anular"].Hidden = true;
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = false;
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Documento"].Hidden = false;
			}
			if((int)this.optEstados.Value == 3)
			{
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Autorizar"].Hidden = true;
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Anular"].Hidden = true;
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = true;
				this.uGridAutorizacion.DisplayLayout.Bands[0].Columns["Documento"].Hidden = false;
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.uGridAutorizacion.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			FuncionesProcedimientos.ExportaExcel(this.uGridAutorizacion);
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnBuscar_Click (sender, e );
		}

		private void uGridAutorizacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Valores
			int idCompra = 0;
			int idTipoFactura = 0;			

			idCompra = (int)e.Cell.Row.Cells["idCompra"].Value;
			idTipoFactura = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idTipoFactura From Compra Where idcompra =  {0}", idCompra));
			#endregion Valores

			#region Autorizacion
			if (e.Cell.Column.ToString() == "Autorizar")
			{
				#region Seguridad
				if(!bActivo)
				{
					MessageBox.Show("el USUARIO se encuentra INACTIVO, validar información.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if(idTipoFactura == 8)
				{
					if(!bEgreso)
					{
						MessageBox.Show("No tiene acceso a AUTORIZAR EGRESOS.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
				if(idTipoFactura == 51)
				{
					if(!bNotaEntrega)
					{
						MessageBox.Show("No tiene acceso a AUTORIZAR NOTAS DE ENTREGA.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
				#endregion Seguridad

				if (DialogResult.Yes == MessageBox.Show("Esta seguro de AUTORIZAR el documento: " + Convert.ToString(e.Cell.Row.Cells["TipoDocumento"].Value) + ", con Numero: " + Convert.ToString(e.Cell.Row.Cells["Numero"].Value) + ".\n\nEste proceso no puede es reversible.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					using (frmMotivoAutorizacionInventario miNotaA = new frmMotivoAutorizacionInventario())
					{
						if (DialogResult.OK == miNotaA.ShowDialog())
						{
							this.txtNotas.Text = miNotaA.txtNotas.Text;

							using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
							{
								oConexion.Open();

								SqlTransaction oTransaction = null;
								SqlCommand oCmdActualiza = oConexion.CreateCommand();
								oCmdActualiza.Connection = oConexion;
								oCmdActualiza.CommandType = CommandType.Text;
								oCmdActualiza.CommandTimeout = 120;

								try
								{	
									oTransaction = oConexion.BeginTransaction();
									oCmdActualiza.Transaction = oTransaction;

									if(idTipoFactura == 8)
									{
										#region Actualiza Estado Egreso
										string sSQLEgreso = string.Format("Exec ModificaEstadoEgreso {0}, '{1}'", idCompra, this.txtNotas.Text.ToString());
										oCmdActualiza.CommandText = sSQLEgreso;
										oCmdActualiza.ExecuteNonQuery();
										#endregion Actualiza Estado Egreso

										#region Inventario
										oCmdActualiza.CommandText = string.Format("Exec GuardaMovimientosInventario {0}, 8", idCompra);
										oCmdActualiza.ExecuteNonQuery();
										#endregion Inventario
									}
									if(idTipoFactura == 51)
									{
										#region Actualiza Estado Nota de Entrega
										string sSQLNotEntrega = string.Format("Exec ModificaEstadoNotaDeEntrega {0}, '{1}'", idCompra, this.txtNotas.Text.ToString());
										oCmdActualiza.CommandText = sSQLNotEntrega;
										oCmdActualiza.ExecuteNonQuery();
										#endregion Actualiza Estado Nota de Entrega

										#region Inventario
										oCmdActualiza.CommandText = string.Format("Exec GuardaMovimientosInventario {0}, 51", idCompra);
										oCmdActualiza.ExecuteNonQuery();
										#endregion Inventario
									}
							
									oTransaction.Commit();
									//oConexion.Close();
								}						
								catch (Exception ex)
								{
									oTransaction.Rollback();
									oConexion.Close();
									MessageBox.Show("Error", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);												
									return;
								}
								finally
								{
									oConexion.Close();				
								}
							}						
						}
					}
				}
				this.btnBuscar_Click (sender, e);
			}
			#endregion Autorizacion

			#region Anulacion
			if (e.Cell.Column.ToString() == "Anular")
			{	
				#region Seguridad
				if (!miAcceso.Anular) 
				{
					MessageBox.Show("No tiene acceso a ANULAR.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#endregion Seguridad

				if (DialogResult.Yes == MessageBox.Show("Esta seguro de ANULAR el documento: " + Convert.ToString(e.Cell.Row.Cells["TipoDocumento"].Value) + ", con Numero: " + Convert.ToString(e.Cell.Row.Cells["Numero"].Value) + ".\n\nEste proceso no puede es reversible.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					using (frmMotivoAutorizacionInventario miNotaA = new frmMotivoAutorizacionInventario())
					{
						if (DialogResult.OK == miNotaA.ShowDialog())
						{
							this.txtNotas.Text = miNotaA.txtNotas.Text;

							if(idTipoFactura == 8)
							{
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AnulacionCompraVenta {0}, {1}, '{2}'", idCompra, idTipoFactura, this.txtNotas.Text.ToString()));
							}

							if(idTipoFactura == 51)
							{
								string sSQLNE = string.Format("Update Compra set Borrar = 1 Where idCompra = {0}", idCompra);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLNE);
								string sSQLAnH = string.Format("Update NotasEntregaLimpieza set Estado = 1 Where idCompra = {0}", idCompra);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLAnH);
							}
						}						
					}
				}
				this.btnBuscar_Click (sender, e);
			}
			#endregion Anulacion

			#region Imprimir
			if (e.Cell.Column.ToString() == "Imprimir")
			{
				#region Seguridad
				if (!miAcceso.Imprimir) 
				{
					MessageBox.Show("No tiene acceso a IMPRIMIR.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#endregion Seguridad

				if(idTipoFactura == 8)
				{
					#region Imprime Reporte Egreso
					Facturacion.Impresion(idCompra, false, 3, miAcceso.Imprimir, true, false, false, 8, 0.00m, 0, 0, "", DateTime.Today, 0, "", cdsSeteoF);
					#endregion Imprime Reporte Egreso
				}

				if(idTipoFactura == 51)
				{
					#region Imprime Reporte Nota de Entrega
					ParameterFields paramFields = new ParameterFields ();
			
					ParameterField paramFieldId = new ParameterField ();
					ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
					paramFieldId.ParameterFieldName = "@idCpa";
					ValIdId.Value = int.Parse(idCompra.ToString());
					paramFieldId.CurrentValues.Add (ValIdId);
					paramFields.Add (paramFieldId);

					Reporte miRep = new Reporte("FacturaNotaEntrega.rpt", "");
					miRep.MdiParent = this.MdiParent;
					miRep.Titulo("Nota de Entrega");
					miRep.crVista.ParameterFieldInfo = paramFields;
					miRep.Show();
					#endregion Imprime Reporte Nota de Entrega
				}				
			}
			#endregion Imprimir

			#region Ver Documento
			if (e.Cell.Column.ToString() == "Documento")
			{
				#region Seguridad
				if (!miAcceso.BBuscarDocumentos) 
				{
					MessageBox.Show("No tiene acceso a BUSCAR DOCUMENTOS.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#endregion Seguridad

				if (idCompra != 0) 
				{
					if (idTipoFactura == 8)
					{
						frmEgresosDeBodega Egresos = new frmEgresosDeBodega(idCompra);
						Egresos.MdiParent = MdiParent;
						Egresos.Show();
					}

					if (idTipoFactura == 51)
					{
						frmNotaDeEntrega NotaDeEntrega = new frmNotaDeEntrega(idCompra);
						NotaDeEntrega.MdiParent = MdiParent;
						NotaDeEntrega.Show();
					}
				}
			}
			#endregion Ver Documento			
		}

		private void uGridAutorizacion_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			DateTime dtHoy = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);

			int iDias = Funcion.iEscalarSQL(cdsSeteoF, 
				string.Format("Select DATEDIFF(day, '{0}', '{1}') As Dias", 
				Convert.ToDateTime(e.Row.Cells["FechaSolicitud"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(dtHoy).ToString("yyyyMMdd")));

			if((int)this.optEstados.Value == 1)
			{
				if (iDias >= 3)
				{
					e.Row.Appearance.BackColor = Color.Yellow;
					e.Row.Appearance.BackColor2 = Color.Yellow;
				}
				if (iDias >= 5)
				{
					e.Row.Appearance.BackColor = Color.Red;
					e.Row.Appearance.BackColor2 = Color.Red;
				}
			}
		}

		private void ultraLabel3_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
