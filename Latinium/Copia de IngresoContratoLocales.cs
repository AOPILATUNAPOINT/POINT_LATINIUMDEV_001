using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de IngresoContratoLocales.
	/// </summary>
	public class IngresoContratoLocales : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel13;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor6;
		private Infragistics.Win.Misc.UltraLabel ultraLabel15;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaDesde;
		private Infragistics.Win.Misc.UltraLabel ultraLabel8;
		private Infragistics.Win.Misc.UltraButton btnUbicacion;
		private Infragistics.Win.Misc.UltraButton btnTipo;
		private Infragistics.Win.Misc.UltraButton btnCComercial;
		private C1.Data.C1DataView cdvCLocales;
		private C1.Data.C1DataSet cdsCLocales;
		private C1.Data.C1DataSet cdsCc;
		private C1.Data.C1DataView cdvCc;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo2;
		private C1.Data.C1DataView cdvCUbicacion;
		private C1.Data.C1DataSet cdsCUbicacion;
		private C1.Data.C1DataSet cdsCTipos;
		private C1.Data.C1DataView cdvCTipos;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCComercial;
		private Infragistics.Win.Misc.UltraLabel ultraLabel9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbActividad;
		private Infragistics.Win.Misc.UltraButton btnActividad;
		private C1.Data.C1DataSet cdsActividad;
		private C1.Data.C1DataView cdvActividad;
		private Infragistics.Win.Misc.UltraLabel ultraLabel10;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbValorM2;
		private Infragistics.Win.Misc.UltraButton btnValorM2;
		private C1.Data.C1DataView cdvValorM2;
		private C1.Data.C1DataSet cdsValorM2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public IngresoContratoLocales()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContratosCentroC", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroComercial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContratosTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContratosUbicacion", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUbicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContratosActividad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActividad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContratosValorM2", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorM2");
			this.cmbCComercial = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCLocales = new C1.Data.C1DataSet();
			this.cdvCc = new C1.Data.C1DataView();
			this.cdsCc = new C1.Data.C1DataSet();
			this.ultraLabel13 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor6 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel15 = new Infragistics.Win.Misc.UltraLabel();
			this.btnCComercial = new Infragistics.Win.Misc.UltraButton();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraCombo1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCTipos = new C1.Data.C1DataView();
			this.cdsCTipos = new C1.Data.C1DataSet();
			this.ultraTextEditor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.btnTipo = new Infragistics.Win.Misc.UltraButton();
			this.ultraTextEditor4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor5 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.dtFechaDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel8 = new Infragistics.Win.Misc.UltraLabel();
			this.btnUbicacion = new Infragistics.Win.Misc.UltraButton();
			this.cdvCLocales = new C1.Data.C1DataView();
			this.ultraCombo2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCUbicacion = new C1.Data.C1DataView();
			this.cdsCUbicacion = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.ultraLabel9 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbActividad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvActividad = new C1.Data.C1DataView();
			this.cdsActividad = new C1.Data.C1DataSet();
			this.btnActividad = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel10 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbValorM2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvValorM2 = new C1.Data.C1DataView();
			this.cdsValorM2 = new C1.Data.C1DataSet();
			this.btnValorM2 = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbCComercial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCTipos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCTipos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbActividad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvActividad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsActividad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbValorM2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvValorM2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsValorM2)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbCComercial
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCComercial.Appearance = appearance1;
			this.cmbCComercial.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCComercial.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCComercial.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCLocales, "ContratosLocales.idCentroComercial"));
			this.cmbCComercial.DataSource = this.cdvCc;
			ultraGridBand1.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 262;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbCComercial.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCComercial.DisplayMember = "Nombre";
			this.cmbCComercial.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCComercial.Enabled = false;
			this.cmbCComercial.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCComercial.Location = new System.Drawing.Point(432, 16);
			this.cmbCComercial.Name = "cmbCComercial";
			this.cmbCComercial.Size = new System.Drawing.Size(200, 21);
			this.cmbCComercial.TabIndex = 4;
			this.cmbCComercial.ValueMember = "idCentroComercial";
			// 
			// cdsCLocales
			// 
			this.cdsCLocales.BindingContextCtrl = this;
			this.cdsCLocales.DataLibrary = "LibFacturacion";
			this.cdsCLocales.DataLibraryUrl = "";
			this.cdsCLocales.DataSetDef = "dsContrato";
			this.cdsCLocales.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCLocales.Name = "cdsCLocales";
			this.cdsCLocales.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCLocales.SchemaDef = null;
			this.cdsCLocales.BeforeFill += new C1.Data.FillEventHandler(this.cdsCLocales_BeforeFill);
			// 
			// cdvCc
			// 
			this.cdvCc.BindingContextCtrl = this;
			this.cdvCc.DataSet = this.cdsCc;
			this.cdvCc.TableName = "";
			this.cdvCc.TableViewName = "ContratosCentroC";
			// 
			// cdsCc
			// 
			this.cdsCc.BindingContextCtrl = this;
			this.cdsCc.DataLibrary = "LibFacturacion";
			this.cdsCc.DataLibraryUrl = "";
			this.cdsCc.DataSetDef = "dsContrato";
			this.cdsCc.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsCc.Name = "cdsCc";
			this.cdsCc.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCc.SchemaDef = null;
			this.cdsCc.BeforeFill += new C1.Data.FillEventHandler(this.cdsCc_BeforeFill);
			// 
			// ultraLabel13
			// 
			this.ultraLabel13.Location = new System.Drawing.Point(328, 24);
			this.ultraLabel13.Name = "ultraLabel13";
			this.ultraLabel13.Size = new System.Drawing.Size(96, 16);
			this.ultraLabel13.TabIndex = 4;
			this.ultraLabel13.Text = "Centro Comercial:";
			// 
			// ultraTextEditor6
			// 
			this.ultraTextEditor6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCLocales, "ContratosLocales.Codigo"));
			this.ultraTextEditor6.Enabled = false;
			this.ultraTextEditor6.Location = new System.Drawing.Point(72, 16);
			this.ultraTextEditor6.Name = "ultraTextEditor6";
			this.ultraTextEditor6.Size = new System.Drawing.Size(72, 21);
			this.ultraTextEditor6.TabIndex = 128;
			// 
			// ultraLabel15
			// 
			this.ultraLabel15.Location = new System.Drawing.Point(16, 24);
			this.ultraLabel15.Name = "ultraLabel15";
			this.ultraLabel15.Size = new System.Drawing.Size(48, 16);
			this.ultraLabel15.TabIndex = 129;
			this.ultraLabel15.Text = "Código";
			// 
			// btnCComercial
			// 
			this.btnCComercial.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnCComercial.Location = new System.Drawing.Point(632, 16);
			this.btnCComercial.Name = "btnCComercial";
			this.btnCComercial.Size = new System.Drawing.Size(24, 23);
			this.btnCComercial.TabIndex = 131;
			this.btnCComercial.Text = "...";
			this.btnCComercial.Click += new System.EventHandler(this.btnCComercial_Click);
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCLocales, "ContratosLocales.Nombre"));
			this.ultraTextEditor1.Enabled = false;
			this.ultraTextEditor1.Location = new System.Drawing.Point(72, 72);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(208, 21);
			this.ultraTextEditor1.TabIndex = 132;
			this.ultraTextEditor1.ValueChanged += new System.EventHandler(this.ultraTextEditor1_ValueChanged);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.Location = new System.Drawing.Point(16, 80);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(48, 16);
			this.ultraLabel1.TabIndex = 133;
			this.ultraLabel1.Text = "Nombre";
			this.ultraLabel1.Click += new System.EventHandler(this.ultraLabel1_Click);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.Location = new System.Drawing.Point(16, 128);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(64, 16);
			this.ultraLabel2.TabIndex = 135;
			this.ultraLabel2.Text = "Ubicación";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.Location = new System.Drawing.Point(208, 128);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(48, 16);
			this.ultraLabel3.TabIndex = 137;
			this.ultraLabel3.Text = "Tipo";
			// 
			// ultraCombo1
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo1.Appearance = appearance2;
			this.ultraCombo1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCLocales, "ContratosLocales.idTipo"));
			this.ultraCombo1.DataSource = this.cdvCTipos;
			ultraGridBand2.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 262;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.ultraCombo1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraCombo1.DisplayMember = "Nombre";
			this.ultraCombo1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo1.Enabled = false;
			this.ultraCombo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo1.Location = new System.Drawing.Point(248, 120);
			this.ultraCombo1.Name = "ultraCombo1";
			this.ultraCombo1.Size = new System.Drawing.Size(88, 21);
			this.ultraCombo1.TabIndex = 138;
			this.ultraCombo1.ValueMember = "idTipo";
			// 
			// cdvCTipos
			// 
			this.cdvCTipos.BindingContextCtrl = this;
			this.cdvCTipos.DataSet = this.cdsCTipos;
			this.cdvCTipos.TableName = "";
			this.cdvCTipos.TableViewName = "ContratosTipo";
			// 
			// cdsCTipos
			// 
			this.cdsCTipos.BindingContextCtrl = this;
			this.cdsCTipos.DataLibrary = "LibFacturacion";
			this.cdsCTipos.DataLibraryUrl = "";
			this.cdsCTipos.DataSetDef = "dsContrato";
			this.cdsCTipos.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsCTipos.Name = "cdsCTipos";
			this.cdsCTipos.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCTipos.SchemaDef = null;
			this.cdsCTipos.BeforeFill += new C1.Data.FillEventHandler(this.cdsCTipos_BeforeFill);
			// 
			// ultraTextEditor3
			// 
			this.ultraTextEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCLocales, "ContratosLocales.Superficie"));
			this.ultraTextEditor3.Enabled = false;
			this.ultraTextEditor3.Location = new System.Drawing.Point(368, 72);
			this.ultraTextEditor3.Name = "ultraTextEditor3";
			this.ultraTextEditor3.Size = new System.Drawing.Size(72, 21);
			this.ultraTextEditor3.TabIndex = 139;
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.Location = new System.Drawing.Point(304, 72);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(64, 16);
			this.ultraLabel4.TabIndex = 140;
			this.ultraLabel4.Text = "Superficie";
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.Location = new System.Drawing.Point(320, 88);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(24, 16);
			this.ultraLabel5.TabIndex = 141;
			this.ultraLabel5.Text = "m2";
			// 
			// btnTipo
			// 
			this.btnTipo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnTipo.Location = new System.Drawing.Point(336, 120);
			this.btnTipo.Name = "btnTipo";
			this.btnTipo.Size = new System.Drawing.Size(24, 23);
			this.btnTipo.TabIndex = 142;
			this.btnTipo.Text = "...";
			this.btnTipo.Click += new System.EventHandler(this.btnTipo_Click);
			// 
			// ultraTextEditor4
			// 
			this.ultraTextEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCLocales, "ContratosLocales.Contacto"));
			this.ultraTextEditor4.Enabled = false;
			this.ultraTextEditor4.Location = new System.Drawing.Point(72, 176);
			this.ultraTextEditor4.Name = "ultraTextEditor4";
			this.ultraTextEditor4.Size = new System.Drawing.Size(248, 21);
			this.ultraTextEditor4.TabIndex = 143;
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.Location = new System.Drawing.Point(16, 184);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(56, 16);
			this.ultraLabel6.TabIndex = 144;
			this.ultraLabel6.Text = "Contacto";
			// 
			// ultraTextEditor5
			// 
			this.ultraTextEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCLocales, "ContratosLocales.Direccion"));
			this.ultraTextEditor5.Enabled = false;
			this.ultraTextEditor5.Location = new System.Drawing.Point(408, 176);
			this.ultraTextEditor5.Name = "ultraTextEditor5";
			this.ultraTextEditor5.Size = new System.Drawing.Size(248, 21);
			this.ultraTextEditor5.TabIndex = 145;
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.Location = new System.Drawing.Point(352, 184);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(56, 16);
			this.ultraLabel7.TabIndex = 146;
			this.ultraLabel7.Text = "Dirección";
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.AllowMonthSelection = true;
			this.dtFechaDesde.AllowWeekSelection = true;
			this.dtFechaDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaDesde.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCLocales, "ContratosLocales.Fecha"));
			dateButton1.Caption = "Today";
			this.dtFechaDesde.DateButtons.Add(dateButton1);
			this.dtFechaDesde.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.dtFechaDesde.Enabled = false;
			this.dtFechaDesde.Format = "dd/MMM/yyyy";
			this.dtFechaDesde.Location = new System.Drawing.Point(200, 16);
			this.dtFechaDesde.Name = "dtFechaDesde";
			this.dtFechaDesde.NonAutoSizeHeight = 23;
			this.dtFechaDesde.Size = new System.Drawing.Size(112, 21);
			this.dtFechaDesde.SpinButtonsVisible = true;
			this.dtFechaDesde.TabIndex = 148;
			this.dtFechaDesde.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// ultraLabel8
			// 
			this.ultraLabel8.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel8.Location = new System.Drawing.Point(152, 24);
			this.ultraLabel8.Name = "ultraLabel8";
			this.ultraLabel8.Size = new System.Drawing.Size(40, 16);
			this.ultraLabel8.TabIndex = 147;
			this.ultraLabel8.Text = "Fecha";
			// 
			// btnUbicacion
			// 
			this.btnUbicacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnUbicacion.Location = new System.Drawing.Point(176, 120);
			this.btnUbicacion.Name = "btnUbicacion";
			this.btnUbicacion.Size = new System.Drawing.Size(24, 23);
			this.btnUbicacion.TabIndex = 149;
			this.btnUbicacion.Text = "...";
			this.btnUbicacion.Click += new System.EventHandler(this.btnUbicacion_Click);
			// 
			// cdvCLocales
			// 
			this.cdvCLocales.BindingContextCtrl = this;
			this.cdvCLocales.DataSet = this.cdsCLocales;
			this.cdvCLocales.Sort = "idLocales";
			this.cdvCLocales.TableName = "";
			this.cdvCLocales.TableViewName = "ContratosLocales";
			// 
			// ultraCombo2
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo2.Appearance = appearance3;
			this.ultraCombo2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCLocales, "ContratosLocales.idUbicacion"));
			this.ultraCombo2.DataSource = this.cdvCUbicacion;
			ultraGridBand3.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.ultraCombo2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraCombo2.DisplayMember = "Nombre";
			this.ultraCombo2.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo2.Enabled = false;
			this.ultraCombo2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo2.Location = new System.Drawing.Point(72, 120);
			this.ultraCombo2.Name = "ultraCombo2";
			this.ultraCombo2.Size = new System.Drawing.Size(104, 21);
			this.ultraCombo2.TabIndex = 151;
			this.ultraCombo2.ValueMember = "idUbicacion";
			// 
			// cdvCUbicacion
			// 
			this.cdvCUbicacion.BindingContextCtrl = this;
			this.cdvCUbicacion.DataSet = this.cdsCUbicacion;
			this.cdvCUbicacion.TableName = "";
			this.cdvCUbicacion.TableViewName = "ContratosUbicacion";
			// 
			// cdsCUbicacion
			// 
			this.cdsCUbicacion.BindingContextCtrl = this;
			this.cdsCUbicacion.DataLibrary = "LibFacturacion";
			this.cdsCUbicacion.DataLibraryUrl = "";
			this.cdsCUbicacion.DataSetDef = "dsContrato";
			this.cdsCUbicacion.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsCUbicacion.Name = "cdsCUbicacion";
			this.cdsCUbicacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCUbicacion.SchemaDef = null;
			this.cdsCUbicacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsCUbicacion_BeforeFill);
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "ContratosLocales";
			this.barraDato1.DataNombreId = "idLocales";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsCLocales;
			this.barraDato1.DataTabla = "ContratosLocales";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(16, 224);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 152;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// ultraLabel9
			// 
			this.ultraLabel9.Location = new System.Drawing.Point(376, 128);
			this.ultraLabel9.Name = "ultraLabel9";
			this.ultraLabel9.Size = new System.Drawing.Size(56, 16);
			this.ultraLabel9.TabIndex = 154;
			this.ultraLabel9.Text = "Actividad:";
			// 
			// cmbActividad
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbActividad.Appearance = appearance4;
			this.cmbActividad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbActividad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbActividad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCLocales, "ContratosLocales.idActividad"));
			this.cmbActividad.DataSource = this.cdvActividad;
			ultraGridBand4.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 262;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbActividad.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbActividad.DisplayMember = "Nombre";
			this.cmbActividad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbActividad.Enabled = false;
			this.cmbActividad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbActividad.Location = new System.Drawing.Point(440, 120);
			this.cmbActividad.Name = "cmbActividad";
			this.cmbActividad.Size = new System.Drawing.Size(192, 21);
			this.cmbActividad.TabIndex = 153;
			this.cmbActividad.ValueMember = "idActividad";
			// 
			// cdvActividad
			// 
			this.cdvActividad.BindingContextCtrl = this;
			this.cdvActividad.DataSet = this.cdsActividad;
			this.cdvActividad.TableName = "";
			this.cdvActividad.TableViewName = "ContratosActividad";
			// 
			// cdsActividad
			// 
			this.cdsActividad.BindingContextCtrl = this;
			this.cdsActividad.DataLibrary = "LibFacturacion";
			this.cdsActividad.DataLibraryUrl = "";
			this.cdsActividad.DataSetDef = "dsContrato";
			this.cdsActividad.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsActividad.Name = "cdsActividad";
			this.cdsActividad.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsActividad.SchemaDef = null;
			// 
			// btnActividad
			// 
			this.btnActividad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnActividad.Location = new System.Drawing.Point(632, 120);
			this.btnActividad.Name = "btnActividad";
			this.btnActividad.Size = new System.Drawing.Size(24, 23);
			this.btnActividad.TabIndex = 155;
			this.btnActividad.Text = "...";
			this.btnActividad.Click += new System.EventHandler(this.btnActividad_Click);
			// 
			// ultraLabel10
			// 
			this.ultraLabel10.Location = new System.Drawing.Point(464, 72);
			this.ultraLabel10.Name = "ultraLabel10";
			this.ultraLabel10.Size = new System.Drawing.Size(80, 16);
			this.ultraLabel10.TabIndex = 157;
			this.ultraLabel10.Text = "Valor x Metro2";
			// 
			// cmbValorM2
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbValorM2.Appearance = appearance5;
			this.cmbValorM2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbValorM2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbValorM2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCLocales, "ContratosLocales.idActividad"));
			this.cmbValorM2.DataSource = this.cdvValorM2;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbValorM2.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbValorM2.DisplayMember = "ValorM2";
			this.cmbValorM2.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbValorM2.Enabled = false;
			this.cmbValorM2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbValorM2.Location = new System.Drawing.Point(544, 72);
			this.cmbValorM2.Name = "cmbValorM2";
			this.cmbValorM2.Size = new System.Drawing.Size(88, 21);
			this.cmbValorM2.TabIndex = 158;
			this.cmbValorM2.ValueMember = "idValor";
			// 
			// cdvValorM2
			// 
			this.cdvValorM2.BindingContextCtrl = this;
			this.cdvValorM2.DataSet = this.cdsValorM2;
			this.cdvValorM2.TableName = "";
			this.cdvValorM2.TableViewName = "ContratosValorM2";
			// 
			// cdsValorM2
			// 
			this.cdsValorM2.BindingContextCtrl = this;
			this.cdsValorM2.DataLibrary = "LibFacturacion";
			this.cdsValorM2.DataLibraryUrl = "";
			this.cdsValorM2.DataSetDef = "dsContrato";
			this.cdsValorM2.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsValorM2.Name = "cdsValorM2";
			this.cdsValorM2.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsValorM2.SchemaDef = null;
			// 
			// btnValorM2
			// 
			this.btnValorM2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnValorM2.Location = new System.Drawing.Point(632, 72);
			this.btnValorM2.Name = "btnValorM2";
			this.btnValorM2.Size = new System.Drawing.Size(24, 23);
			this.btnValorM2.TabIndex = 159;
			this.btnValorM2.Text = "...";
			this.btnValorM2.Click += new System.EventHandler(this.btnValorM2_Click);
			// 
			// IngresoContratoLocales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(672, 261);
			this.Controls.Add(this.btnValorM2);
			this.Controls.Add(this.cmbValorM2);
			this.Controls.Add(this.ultraLabel10);
			this.Controls.Add(this.ultraLabel9);
			this.Controls.Add(this.cmbActividad);
			this.Controls.Add(this.btnActividad);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.ultraCombo2);
			this.Controls.Add(this.btnUbicacion);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.ultraLabel8);
			this.Controls.Add(this.ultraTextEditor5);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.ultraTextEditor4);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.btnTipo);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.ultraTextEditor3);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.ultraCombo1);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.btnCComercial);
			this.Controls.Add(this.ultraTextEditor6);
			this.Controls.Add(this.ultraLabel15);
			this.Controls.Add(this.ultraLabel13);
			this.Controls.Add(this.cmbCComercial);
			this.MaximizeBox = false;
			this.Name = "IngresoContratoLocales";
			this.Text = "Locales Comerciales";
			this.Load += new System.EventHandler(this.IngresoContratoLocales_Load);
			this.Closed += new System.EventHandler(this.IngresoContratoLocales_Closed);
			((System.ComponentModel.ISupportInitialize)(this.cmbCComercial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCTipos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCTipos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbActividad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvActividad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsActividad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbValorM2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvValorM2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsValorM2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraTextEditor1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraLabel1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void IngresoContratoLocales_Load(object sender, System.EventArgs e)
		{
			barraDato1.ProximoId(4);
			
		}

		private void btnUbicacion_Click(object sender, System.EventArgs e)
		{	
			using (IngresoContratosUbicacion	ingUbi = new IngresoContratosUbicacion())
			{
				ingUbi.ShowDialog();
			}
			this.cdsCUbicacion.Clear();
			this.cdsCUbicacion.Fill();		
		}

		private void btnTipo_Click(object sender, System.EventArgs e)
		{								
			using (IngresoContratosTipo	ingTip = new IngresoContratosTipo())
			{
				ingTip.ShowDialog();
			}
			this.cdsCTipos.Clear();
			this.cdsCTipos.Fill();	
		}

		private void btnCComercial_Click(object sender, System.EventArgs e)
		{
			using (IngresoContratosCentroComercial ingCC = new IngresoContratosCentroComercial())
			{
				ingCC.ShowDialog();
			}
			this.cdsCc.Clear();
			this.cdsCc.Fill();		

		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
		}
		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
		this.barraDato1.GrabaRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
		this.barraDato1.EditarRegistro();	
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
		this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			this.barraDato1.BorraRegistro();
		}

		private void cdsCLocales_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCLocales.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void cdsCTipos_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCTipos.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCc_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCc.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCUbicacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCUbicacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		public int idBusca=0;

		private void IngresoContratoLocales_Closed(object sender, System.EventArgs e)
		{			
			this.DialogResult = DialogResult.OK;
			this.Close();
		}

		private void btnActividad_Click(object sender, System.EventArgs e)
		{
			using (IngresoContratosActividad	ingAct = new IngresoContratosActividad())
			{
				ingAct.ShowDialog();
			}
			this.cdsActividad.Clear();
			this.cdsActividad.Fill();	
		}

		private void btnValorM2_Click(object sender, System.EventArgs e)
		{
			using (IngresoContratosValorM2 ingVm2 = new IngresoContratosValorM2())
			{
				ingVm2.ShowDialog();
			}
			this.cdsValorM2.Clear();
			this.cdsValorM2.Fill();		
		}
	}
}
