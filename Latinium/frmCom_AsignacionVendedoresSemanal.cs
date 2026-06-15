using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_AsignacionVendedoresSemanal.
	/// </summary>
	public class frmCom_AsignacionVendedoresSemanal : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbLocales;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridAsignacion;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRangos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Button btnFaltas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnNuevoEmpleado;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbVendedores;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstados;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_AsignacionVendedoresSemanal()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_AsignacionDeVendedoresSemanal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SumaPuerteo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaSemanal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaSemanalContado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaSemanalCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaSemanalGEx");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_AsignacionDeVendedoresSemanal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SumaPuerteo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango", -1, "cmbRangos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Estado", -1, "cmbEstados");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaSemanal");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaSemanalContado");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaSemanalCredito");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaSemanalGEx");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMovimiento", 0);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_AsignacionVendedoresSemanal));
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCalcula", 1);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.label7 = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.cmbLocales = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.uGridAsignacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbRangos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnFaltas = new System.Windows.Forms.Button();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnNuevoEmpleado = new System.Windows.Forms.Button();
			this.cmbVendedores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridAsignacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRangos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.SuspendLayout();
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(288, 43);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(53, 16);
			this.label7.TabIndex = 755;
			this.label7.Text = "Vendedor";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(672, 40);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 748;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbLocales
			// 
			this.cmbLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLocales.DataSource = this.ultraDataSource4;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 160;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbLocales.DisplayMember = "Nombre";
			this.cmbLocales.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbLocales.Location = new System.Drawing.Point(416, 224);
			this.cmbLocales.Name = "cmbLocales";
			this.cmbLocales.Size = new System.Drawing.Size(200, 112);
			this.cmbLocales.TabIndex = 740;
			this.cmbLocales.ValueMember = "Bodega";
			this.cmbLocales.Visible = false;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(bool);
			ultraDataColumn5.DefaultValue = false;
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn7.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn9.DefaultValue = 0;
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn10.DefaultValue = 0;
			ultraDataColumn11.DataType = typeof(System.DateTime);
			ultraDataColumn11.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn12.DataType = typeof(System.DateTime);
			ultraDataColumn12.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn13.DefaultValue = 0;
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn14.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn17.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
			// uGridAsignacion
			// 
			this.uGridAsignacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridAsignacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridAsignacion.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridAsignacion.DisplayLayout.Appearance = appearance1;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.Caption = "Grupo";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 25;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.Caption = "Local";
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn6.Width = 17;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.Caption = "Vendedor";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn7.Width = 220;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.Caption = "Cargo";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 140;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.Caption = "Categoria";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 65;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Format = "";
			ultraGridColumn10.Header.Caption = "Estado";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 75;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Format = "";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 70;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Format = "";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 70;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance2;
			ultraGridColumn13.Header.Caption = "N. Dias";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 44;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance3;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Meta Semanal";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 85;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance4;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Meta Contado";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 85;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance5;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Meta Crédito";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 85;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance6;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Meta GEx";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 85;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance7.Image = ((object)(resources.GetObject("appearance7.Image")));
			appearance7.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance7.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn18.CellButtonAppearance = appearance7;
			ultraGridColumn18.Header.Caption = "...";
			ultraGridColumn18.Header.VisiblePosition = 4;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 30;
			ultraGridColumn19.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			appearance8.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance8.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn19.CellButtonAppearance = appearance8;
			ultraGridColumn19.Header.Caption = "...";
			ultraGridColumn19.Header.VisiblePosition = 5;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 30;
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
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridAsignacion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridAsignacion.DisplayLayout.GroupByBox.Hidden = true;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAsignacion.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.uGridAsignacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridAsignacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridAsignacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridAsignacion.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAsignacion.DisplayLayout.Override.CellAppearance = appearance11;
			this.uGridAsignacion.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridAsignacion.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.uGridAsignacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAsignacion.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAsignacion.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAsignacion.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.uGridAsignacion.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridAsignacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridAsignacion.Location = new System.Drawing.Point(10, 88);
			this.uGridAsignacion.Name = "uGridAsignacion";
			this.uGridAsignacion.Size = new System.Drawing.Size(1100, 288);
			this.uGridAsignacion.TabIndex = 737;
			this.uGridAsignacion.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridAsignacion_BeforeRowsDeleted);
			this.uGridAsignacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridAsignacion_ClickCellButton);
			this.uGridAsignacion.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridAsignacion_AfterCellUpdate);
			this.uGridAsignacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridAsignacion_InitializeLayout);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 43);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 742;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(360, 42);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(208, 19);
			this.txtBuscar.TabIndex = 754;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(584, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 747;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1120, 8);
			this.groupBox1.TabIndex = 745;
			this.groupBox1.TabStop = false;
			// 
			// cmbRangos
			// 
			this.cmbRangos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRangos.DataSource = this.ultraDataSource6;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Width = 160;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbRangos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbRangos.DisplayMember = "Rango";
			this.cmbRangos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbRangos.Location = new System.Drawing.Point(616, 224);
			this.cmbRangos.Name = "cmbRangos";
			this.cmbRangos.Size = new System.Drawing.Size(200, 112);
			this.cmbRangos.TabIndex = 743;
			this.cmbRangos.ValueMember = "idCom_Rango";
			this.cmbRangos.Visible = false;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn18.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// ultraDataSource7
			// 
			ultraDataColumn20.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource5;
			ultraGridColumn22.Header.VisiblePosition = 2;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 180;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 41);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 21);
			this.cmbBodega.TabIndex = 741;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn22.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// btnFaltas
			// 
			this.btnFaltas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnFaltas.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnFaltas.Enabled = false;
			this.btnFaltas.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnFaltas.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnFaltas.Image = ((System.Drawing.Image)(resources.GetObject("btnFaltas.Image")));
			this.btnFaltas.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFaltas.Location = new System.Drawing.Point(960, 7);
			this.btnFaltas.Name = "btnFaltas";
			this.btnFaltas.Size = new System.Drawing.Size(72, 23);
			this.btnFaltas.TabIndex = 752;
			this.btnFaltas.Text = "Faltas";
			this.btnFaltas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnFaltas.Click += new System.EventHandler(this.btnFaltas_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn25.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27});
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1040, 7);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 23);
			this.btnExportar.TabIndex = 751;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnNuevoEmpleado
			// 
			this.btnNuevoEmpleado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNuevoEmpleado.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevoEmpleado.Enabled = false;
			this.btnNuevoEmpleado.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevoEmpleado.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevoEmpleado.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevoEmpleado.Image")));
			this.btnNuevoEmpleado.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevoEmpleado.Location = new System.Drawing.Point(840, 7);
			this.btnNuevoEmpleado.Name = "btnNuevoEmpleado";
			this.btnNuevoEmpleado.Size = new System.Drawing.Size(112, 23);
			this.btnNuevoEmpleado.TabIndex = 749;
			this.btnNuevoEmpleado.Text = "Nuevo Empleado";
			this.btnNuevoEmpleado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevoEmpleado.Click += new System.EventHandler(this.btnNuevoEmpleado_Click);
			// 
			// cmbVendedores
			// 
			this.cmbVendedores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedores.DataSource = this.ultraDataSource2;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 57;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 100;
			ultraGridColumn27.Header.VisiblePosition = 2;
			ultraGridColumn27.Width = 180;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbVendedores.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbVendedores.DisplayMember = "Nombre";
			this.cmbVendedores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbVendedores.Location = new System.Drawing.Point(216, 224);
			this.cmbVendedores.Name = "cmbVendedores";
			this.cmbVendedores.Size = new System.Drawing.Size(200, 112);
			this.cmbVendedores.TabIndex = 739;
			this.cmbVendedores.ValueMember = "idPersonal";
			this.cmbVendedores.Visible = false;
			// 
			// cmbEstados
			// 
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DataSource = this.ultraDataSource7;
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Width = 200;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbEstados.DisplayMember = "Estado";
			this.cmbEstados.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbEstados.Location = new System.Drawing.Point(816, 224);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(200, 112);
			this.cmbEstados.TabIndex = 744;
			this.cmbEstados.ValueMember = "idCom_Estado";
			this.cmbEstados.Visible = false;
			// 
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargos.DataSource = this.ultraDataSource3;
			ultraGridColumn30.Header.VisiblePosition = 0;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Width = 200;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbCargos.DisplayMember = "Cargo";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargos.Location = new System.Drawing.Point(16, 223);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(200, 112);
			this.cmbCargos.TabIndex = 738;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn28.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn28,
																																 ultraDataColumn29});
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(192, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 759;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 10);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(37, 16);
			this.label5.TabIndex = 758;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance16;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 757;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance17;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 756;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.Validating += new System.ComponentModel.CancelEventHandler(this.dtDesde_Validating);
			this.dtDesde.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtDesde_BeforeDropDown);
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// frmCom_AsignacionVendedoresSemanal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1120, 382);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbRangos);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnFaltas);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnNuevoEmpleado);
			this.Controls.Add(this.cmbVendedores);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.cmbCargos);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.cmbLocales);
			this.Controls.Add(this.uGridAsignacion);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnVer);
			this.Name = "frmCom_AsignacionVendedoresSemanal";
			this.Text = "frmr";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCom_AsignacionVendedoresSemanal_KeyDown);
			this.Load += new System.EventHandler(this.frmCom_AsignacionVendedoresSemanal_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridAsignacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRangos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private Acceso miAcceso;
		bool bValidaRango = false;
		bool bValidaMovPer = false;

		private void UnloadMe()
		{
			this.Close();
		}

		private void ListaVendedores()
		{
			this.cmbVendedores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Com_ListaVendedores '{0}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd")));
		}

		private void frmCom_AsignacionVendedoresSemanal_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("02/07/2018");

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Exec Com_RetornaPeriodosSemanales", true);
			dr.Read();
			if (dr.HasRows)
			{				
				this.dtDesde.Value = dr.GetDateTime(0);
				this.dtHasta.Value = dr.GetDateTime(1);
			}
			dr.Close();

			this.dtDesde.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);

			miAcceso = new Acceso(cdsSeteoF, "052105");
			
			if (!Funcion.Permiso("869", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Asignacion de vendedores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
						
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 7, 0");			
			this.cmbLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 7, 0");
			this.cmbCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 59, -1, 0");			
			this.cmbRangos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 60, 0");
			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 61, 0");
			
			if (miAcceso.BExportar) this.btnExportar.Enabled = true;
			//if (miAcceso.Imprimir) this.btnImpresion.Enabled = true;
			if (miAcceso.Nuevo) this.btnNuevoEmpleado.Enabled = true;
			if (miAcceso.ReImprimir) this.btnFaltas.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridAsignacion);
			if (miAcceso.Editar) FuncionesProcedimientos.EstadoGrids(true, this.uGridAsignacion);	
		}

		private void Consulta()
		{			
			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format("Exec Com_ConsultaAsignacionSemanalVendedores '{0}', '{1}', {2}, '{3}'", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), iBodega, this.txtBuscar.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridAsignacion);
			
			this.lblContador.Text = this.uGridAsignacion.Rows.Count + " REGISTROS ENCONTRADOS";

			this.uGridAsignacion.DisplayLayout.Bands[0].SortedColumns.Add("Bodega", false, true);	
			this.uGridAsignacion.Rows.ExpandAll(true);

			//			if (FuncionesProcedimientos.ValidaPeriodoCerrado(this.dtpPeriodo.Text.ToString(), cdsSeteoF)) 
			//			{
			//				FuncionesProcedimientos.EstadoGrids(false, this.uGridAsignacion); 
			//				this.btnNuevoEmpleado.Enabled = false;
			//				this.btnFaltas.Enabled = false;
			//			}
			//			else 
			//			{
			//				if (miAcceso.Editar) FuncionesProcedimientos.EstadoGrids(true, this.uGridAsignacion);
			//				if (miAcceso.Nuevo) this.btnNuevoEmpleado.Enabled = true;
			//				if (miAcceso.ReImprimir) this.btnFaltas.Enabled = true;
			//			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			Consulta();

			ListaVendedores();
		}

		private void Actualiza(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int iBodega = 0;
			int idCargo = 0;
			int idPersonal = 0;
			int idCom_Rango = 0;
			int idCom_Estado = 0;

			if (e.Cell.Row.Cells["Bodega"].Value != System.DBNull.Value) iBodega = (int)e.Cell.Row.Cells["Bodega"].Value;
			if (e.Cell.Row.Cells["idCargo"].Value != System.DBNull.Value) idCargo = (int)e.Cell.Row.Cells["idCargo"].Value;
			if (e.Cell.Row.Cells["idPersonal"].Value != System.DBNull.Value) idPersonal = (int)e.Cell.Row.Cells["idPersonal"].Value;
			if (e.Cell.Row.Cells["idCom_Rango"].Value != System.DBNull.Value) idCom_Rango = (int)e.Cell.Row.Cells["idCom_Rango"].Value;
			if (e.Cell.Row.Cells["idCom_Estado"].Value != System.DBNull.Value) idCom_Estado = (int)e.Cell.Row.Cells["idCom_Estado"].Value;

			string sSQL = string.Format("Exec Com_GuardaAsignacionSemanalVendedores {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', '{9}', {10}, {11}, {12}, {13}, {14}, {15}, '{16}'",
				(int)e.Cell.Row.Cells["idCom_AsignacionDeVendedoresSemanal"].Value, (int)e.Cell.Row.Cells["Numero"].Value, iBodega, (bool)e.Cell.Row.Cells["SumaPuerteo"].Value, idPersonal, idCargo, idCom_Rango, idCom_Estado, 
				Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd"), (int)e.Cell.Row.Cells["NumeroDias"].Value, 
				Convert.ToDecimal(e.Cell.Row.Cells["MetaSemanal"].Value), Convert.ToDecimal(e.Cell.Row.Cells["MetaSemanalContado"].Value), //11
				Convert.ToDecimal(e.Cell.Row.Cells["MetaSemanalCredito"].Value), Convert.ToDecimal(e.Cell.Row.Cells["MetaSemanalGEx"].Value),
				false, "20010101");
			e.Cell.Row.Cells["idCom_AsignacionDeVendedoresSemanal"].Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			sSQL = string.Format("Exec Com_ActualizaDatosNominaAV {0}, {1}, {2}, {3}, {4}", 
				iBodega, idCargo, idCom_Estado, idCom_Rango, idPersonal);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void Metas(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idCom_Estado"].Value != System.DBNull.Value)
			{
				if (e.Cell.Row.Cells["idCargo"].Value != System.DBNull.Value)
				{
					e.Cell.Row.Cells["NumeroDias"].Value = FuncionesProcedimientos.CalculaDias((DateTime)e.Cell.Row.Cells["Desde"].Value, (DateTime)e.Cell.Row.Cells["Hasta"].Value, cdsSeteoF);

//					e.Cell.Row.Cells["MetaSemanal"].Value = FuncionesProcedimientos.CalculaMeta((int)e.Cell.Row.Cells["idCom_Estado"].Value, (int)e.Cell.Row.Cells["Bodega"].Value, (int)e.Cell.Row.Cells["idCargo"].Value, (DateTime)e.Cell.Row.Cells["Desde"].Value, (DateTime)e.Cell.Row.Cells["Hasta"].Value, (int)e.Cell.Row.Cells["NumeroDias"].Value, 1, (int)e.Cell.Row.Cells["Numero"].Value, cdsSeteoF);
//					e.Cell.Row.Cells["MetaSemanalContado"].Value = FuncionesProcedimientos.CalculaMeta((int)e.Cell.Row.Cells["idCom_Estado"].Value, (int)e.Cell.Row.Cells["Bodega"].Value, (int)e.Cell.Row.Cells["idCargo"].Value, (DateTime)e.Cell.Row.Cells["Desde"].Value, (DateTime)e.Cell.Row.Cells["Hasta"].Value, (int)e.Cell.Row.Cells["NumeroDias"].Value, 2, (int)e.Cell.Row.Cells["Numero"].Value, cdsSeteoF);
//					e.Cell.Row.Cells["MetaSemanalCredito"].Value = FuncionesProcedimientos.CalculaMeta((int)e.Cell.Row.Cells["idCom_Estado"].Value, (int)e.Cell.Row.Cells["Bodega"].Value, (int)e.Cell.Row.Cells["idCargo"].Value, (DateTime)e.Cell.Row.Cells["Desde"].Value, (DateTime)e.Cell.Row.Cells["Hasta"].Value, (int)e.Cell.Row.Cells["NumeroDias"].Value, 3, (int)e.Cell.Row.Cells["Numero"].Value, cdsSeteoF);
//					e.Cell.Row.Cells["MetaSemanalGEx"].Value = FuncionesProcedimientos.CalculaMeta((int)e.Cell.Row.Cells["idCom_Estado"].Value, (int)e.Cell.Row.Cells["Bodega"].Value, (int)e.Cell.Row.Cells["idCargo"].Value, (DateTime)e.Cell.Row.Cells["Desde"].Value, (DateTime)e.Cell.Row.Cells["Hasta"].Value, (int)e.Cell.Row.Cells["NumeroDias"].Value, 4, (int)e.Cell.Row.Cells["Numero"].Value, cdsSeteoF);
				}
			}
		}
		
		private void uGridAsignacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnMovimiento")
			{										
				if ((DateTime)e.Cell.Row.Cells["Hasta"].Value < DateTime.Today)
				{
					MessageBox.Show("No puede generar movimientos de periodos pasados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if ((int)e.Cell.Row.Cells["idCom_Estado"].Value == 2)
				{
					MessageBox.Show("No puede generar movimientos de personal inactivo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				#region Valida Movimientos
				string sSQLValida = string.Format("Exec Com_ValidaMovimientos {0}, {1}, '{2}'", 
					(int)e.Cell.Row.Cells["idCom_AsignacionDeVendedoresSemanal"].Value, (int)e.Cell.Row.Cells["idPersonal"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd"));
				if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValida) > 0)
				{
					e.Cell.Value = e.Cell.OriginalValue;
					MessageBox.Show(string.Format("Ya existen movimientos registrados en fechas posteriores a la seleccionada", Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					bValidaMovPer = false;
					return;
				}		
				bValidaMovPer = true;			
				#endregion Valida Movimientos
			
				#region Valida Permisos
				sSQLValida = string.Format("Exec Com_ValidaMovimientoPermisos {0}, '{1}', 0", 
					(int)e.Cell.Row.Cells["idPersonal"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd"));

				if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValida) > 0)
				{
					bValidaMovPer = true;
					e.Cell.Value = e.Cell.OriginalValue;
					MessageBox.Show(string.Format("Existen Permisos registrados en fechas posteriores a la seleccionada", Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					bValidaMovPer = false;
					return;
				}
				#endregion Valida Permisos

				using (frmCom_ValidaCambioFechaAV VC = new frmCom_ValidaCambioFechaAV())
				{				
					if (DialogResult.OK == VC.ShowDialog())
					{	
						#region Movimientos
						if (VC.rdbNuevoMovimiento.Checked) 
						{	
							using (frmCom_CambioVendedor CV = new frmCom_CambioVendedor(
											 (int)e.Cell.Row.Cells["idPersonal"].Value, (int)e.Cell.Row.Cells["idCom_AsignacionDeVendedoresSemanal"].Value,
											 false, (int)e.Cell.Row.Cells["idCom_Estado"].Value, (int)e.Cell.Row.Cells["idCargo"].Value, (int)e.Cell.Row.Cells["idCom_Rango"].Value, 
											 (DateTime)e.Cell.Row.Cells["Desde"].Value, (DateTime)e.Cell.Row.Cells["Hasta"].Value, (DateTime)e.Cell.Row.Cells["Hasta"].Value, (int)e.Cell.Row.Cells["Bodega"].Value, (int)e.Cell.Row.Cells["Numero"].Value))
							{				
								if (DialogResult.OK == CV.ShowDialog())
								{
									if ((int)CV.txtIdCom_AsignacionDeVendedoresN.Value != (int)e.Cell.Row.Cells["idCom_AsignacionDeVendedoresSemanal"].Value)
									{
										Metas(e);
										
										this.Actualiza(e);
									}

									this.btnVer_Click(sender, e);
								}
								else 
								{
									bValidaMovPer = true;
									e.Cell.Value = e.Cell.OriginalValue;
									bValidaMovPer = false;
								}
							}	
						}
						#endregion Movimientos				

						#region Permisos
						if (VC.rdbPFV.Checked) 
						{							
							int idCom_Rango = 0;
							int idCargo = 0;
							if (e.Cell.Row.Cells["idCom_Rango"].Value != System.DBNull.Value) idCom_Rango = (int)e.Cell.Row.Cells["idCom_Rango"].Value;
							if (e.Cell.Row.Cells["idCargo"].Value != System.DBNull.Value) idCargo = (int)e.Cell.Row.Cells["idCargo"].Value;

							using (frmNominaPermisosFaltas NPF = new frmNominaPermisosFaltas((int)e.Cell.Row.Cells["idPersonal"].Value, (int)e.Cell.Row.Cells["idCom_AsignacionDeVendedoresSemanal"].Value, 
											 (int)e.Cell.Row.Cells["Bodega"].Value, "", 0, (DateTime)e.Cell.Row.Cells["Desde"].Value, (DateTime)e.Cell.Row.Cells["Hasta"].Value, (DateTime)e.Cell.Row.Cells["Hasta"].Value, 
											 (int)e.Cell.Row.Cells["idCargo"].Value, idCom_Rango, (int)e.Cell.Row.Cells["idCom_Estado"].Value))
							{				
								if (DialogResult.OK == NPF.ShowDialog())
								{
									//									if ((int)NPF.idCom_AsignacionDeVendedores == 0)
									//									{
									Metas(e);
										
									//this.Actualiza(e);
									//									}
									this.btnVer_Click(sender, e);
								}
								else
								{
									bValidaMovPer = true;
									e.Cell.Value = e.Cell.OriginalValue;
									bValidaMovPer = false;
								}
							}
						}						
						#endregion Permisos
					}
					else 
					{						
						bValidaMovPer = true;
						e.Cell.Value = e.Cell.OriginalValue;
						bValidaMovPer = false;
					}
				}
			}

			if (e.Cell.Column.ToString() == "btnCalcula")
			{
				this.Metas(e);
				this.Actualiza(e);
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();

			ListaVendedores();
		}

		private void uGridAsignacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnNuevoEmpleado_Click(object sender, System.EventArgs e)
		{	
			if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
				return;
			}

			using (frmCom_CambioVendedor CV = new frmCom_CambioVendedor(0, 0, true, 0, 0, 0, Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), Convert.ToDateTime(this.dtHasta.Value), (int)this.cmbBodega.Value, 0))
			{				
				if (DialogResult.OK == CV.ShowDialog())
				{
					this.btnVer_Click(sender, e);
				}
			}
		}

		private void btnFaltas_Click(object sender, System.EventArgs e)
		{						
			if (this.uGridAsignacion.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un vendedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
			try
			{
				using (frmNominaPermisosFaltas NPF = new frmNominaPermisosFaltas((int)this.uGridAsignacion.ActiveRow.Cells["idPersonal"].Value, (int)this.uGridAsignacion.ActiveRow.Cells["idCom_AsignacionDeVendedoresSemanal"].Value, 
								 (int)this.uGridAsignacion.ActiveRow.Cells["Bodega"].Value, "", 1, (DateTime)this.uGridAsignacion.ActiveRow.Cells["Desde"].Value, (DateTime)this.uGridAsignacion.ActiveRow.Cells["Hasta"].Value, (DateTime)this.uGridAsignacion.ActiveRow.Cells["Hasta"].Value, 0, 0, 0))
				{				
					if (DialogResult.OK == NPF.ShowDialog())
					{
						
					}
				}
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void uGridAsignacion_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bValidaRango) return;
			if (bValidaMovPer) return;
			
			if (e.Cell.Column.ToString() == "Bodega")
			{
				Metas(e);

				this.Actualiza(e);	
			}

			if (e.Cell.Column.ToString() == "idCargo")
			{
				Metas(e);

				this.Actualiza(e);	
			}

			if (e.Cell.Column.ToString() == "idCom_Estado")
			{
				this.Actualiza(e);	
			}

			if (e.Cell.Column.ToString() == "idPersonal")
			{
				this.Actualiza(e);	
			}

			if (e.Cell.Column.ToString() == "idCom_Rango")
			{
				this.Actualiza(e);					
			}

			if (e.Cell.Column.ToString() == "Desde")
			{
				Metas(e);
				this.Actualiza(e);											
			}

			if (e.Cell.Column.ToString() == "Hasta")
			{
				Metas(e);
				this.Actualiza(e);											
			}
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value != System.DBNull.Value)
			{
				this.dtHasta.Value = Convert.ToDateTime(this.dtDesde.Value).AddDays(6);

				Consulta();

				ListaVendedores();
			}
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void uGridAsignacion_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.DisplayPromptMsg = false;
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta Seguro de Borrar este registro?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				for (int i=0; i<e.Rows.Length; i++)
				{
					string sSQL = string.Format("Delete From Com_AsignacionDeVendedoresSemanal Where idCom_AsignacionDeVendedoresSemanal = {0}", (int)e.Rows[i].Cells["idCom_AsignacionDeVendedoresSemanal"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				}
			}
			else e.Cancel = true;
		}

		private void dtDesde_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.dtDesde.Value != System.DBNull.Value)
			{
				if (FuncionesProcedimientos.iDiaSemana((DateTime)this.dtDesde.Value, cdsSeteoF) != 2)
				{
					MessageBox.Show("La Fecha Inicial de Consulta De Coincidir en Dia Lunes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridAsignacion);
		}

		private void frmCom_AsignacionVendedoresSemanal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void dtDesde_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}
	}
}
