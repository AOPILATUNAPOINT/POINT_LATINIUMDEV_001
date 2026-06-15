using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CierreCaja.
	/// </summary>
	public class CierreCaja : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.Misc.UltraLabel ultraLabel8;
		private Infragistics.Win.Misc.UltraLabel ultraLabel10;
		private Infragistics.Win.Misc.UltraLabel ultraLabel11;
		private Infragistics.Win.Misc.UltraLabel ultraLabel13;
		private Infragistics.Win.Misc.UltraLabel ultraLabel14;
		private Infragistics.Win.Misc.UltraLabel ultraLabel9;
		private Infragistics.Win.Misc.UltraLabel ultraLabel12;
		private Infragistics.Win.Misc.UltraLabel ultraLabel15;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel16;
		private Infragistics.Win.Misc.UltraButton btTotal;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaHasta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataView cdvProyecto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsUsuario;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbUsuario;
		private C1.Data.C1DataSet cdsDeposito;
		private C1.Data.C1DataSet cdsBanco;
		private C1.Data.C1DataView cdvBanco;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBanco;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.Misc.UltraLabel ultraLabel17;
		private Infragistics.Win.Misc.UltraLabel ultraLabel18;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCierreCaja;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtM25;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtM1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtM5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtM10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtM50;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtM100;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtB1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtB2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtB5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtB10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtB20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtB50;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtB100;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CierreCaja()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_CierreCajaDeposito", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDeposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCierreCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco", -1, "cmbBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hoja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProcesoOk");
			this.cdsDeposito = new C1.Data.C1DataSet();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel8 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel10 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel11 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel13 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel14 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel9 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFechaDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbFechaHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.ultraLabel12 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel15 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.ultraLabel16 = new Infragistics.Win.Misc.UltraLabel();
			this.btTotal = new Infragistics.Win.Misc.UltraButton();
			this.udsUsuario = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbUsuario = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsBanco = new C1.Data.C1DataSet();
			this.cdvBanco = new C1.Data.C1DataView();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.ultraLabel17 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel18 = new Infragistics.Win.Misc.UltraLabel();
			this.txtIdCierreCaja = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtM25 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtM1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtM5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtM10 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtM50 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtM100 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtB1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtB2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtB5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtB10 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtB20 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtB50 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtB100 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCierreCaja)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtM25)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtM1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtM5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtM10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtM50)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtM100)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB50)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB100)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsDeposito
			// 
			this.cdsDeposito.BindingContextCtrl = this;
			this.cdsDeposito.DataLibrary = "LibFacturacion";
			this.cdsDeposito.DataLibraryUrl = "";
			this.cdsDeposito.DataSetDef = "dsDeposito";
			this.cdsDeposito.EnforceConstraints = false;
			this.cdsDeposito.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsDeposito.Name = "cdsDeposito";
			this.cdsDeposito.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsDeposito.SchemaDef = null;
			this.cdsDeposito.BeforeFill += new C1.Data.FillEventHandler(this.cdsDeposito_BeforeFill);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(18, 34);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(23, 15);
			this.ultraLabel1.TabIndex = 1;
			this.ultraLabel1.Text = "100";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(24, 60);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(17, 15);
			this.ultraLabel2.TabIndex = 4;
			this.ultraLabel2.Text = "50";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(24, 112);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(17, 15);
			this.ultraLabel3.TabIndex = 8;
			this.ultraLabel3.Text = "10";
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(24, 86);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(17, 15);
			this.ultraLabel4.TabIndex = 6;
			this.ultraLabel4.Text = "20";
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.AutoSize = true;
			this.ultraLabel5.Location = new System.Drawing.Point(31, 190);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(10, 15);
			this.ultraLabel5.TabIndex = 14;
			this.ultraLabel5.Text = "1";
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.Location = new System.Drawing.Point(31, 164);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(10, 15);
			this.ultraLabel6.TabIndex = 12;
			this.ultraLabel6.Text = "2";
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.AutoSize = true;
			this.ultraLabel7.Location = new System.Drawing.Point(31, 138);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(10, 15);
			this.ultraLabel7.TabIndex = 10;
			this.ultraLabel7.Text = "5";
			// 
			// ultraLabel8
			// 
			this.ultraLabel8.AutoSize = true;
			this.ultraLabel8.Location = new System.Drawing.Point(167, 164);
			this.ultraLabel8.Name = "ultraLabel8";
			this.ultraLabel8.Size = new System.Drawing.Size(10, 15);
			this.ultraLabel8.TabIndex = 28;
			this.ultraLabel8.Text = "1";
			// 
			// ultraLabel10
			// 
			this.ultraLabel10.AutoSize = true;
			this.ultraLabel10.Location = new System.Drawing.Point(167, 138);
			this.ultraLabel10.Name = "ultraLabel10";
			this.ultraLabel10.Size = new System.Drawing.Size(10, 15);
			this.ultraLabel10.TabIndex = 24;
			this.ultraLabel10.Text = "5";
			// 
			// ultraLabel11
			// 
			this.ultraLabel11.AutoSize = true;
			this.ultraLabel11.Location = new System.Drawing.Point(160, 112);
			this.ultraLabel11.Name = "ultraLabel11";
			this.ultraLabel11.Size = new System.Drawing.Size(17, 15);
			this.ultraLabel11.TabIndex = 22;
			this.ultraLabel11.Text = "10";
			// 
			// ultraLabel13
			// 
			this.ultraLabel13.AutoSize = true;
			this.ultraLabel13.Location = new System.Drawing.Point(160, 60);
			this.ultraLabel13.Name = "ultraLabel13";
			this.ultraLabel13.Size = new System.Drawing.Size(17, 15);
			this.ultraLabel13.TabIndex = 18;
			this.ultraLabel13.Text = "50";
			// 
			// ultraLabel14
			// 
			this.ultraLabel14.AutoSize = true;
			this.ultraLabel14.Location = new System.Drawing.Point(154, 34);
			this.ultraLabel14.Name = "ultraLabel14";
			this.ultraLabel14.Size = new System.Drawing.Size(23, 15);
			this.ultraLabel14.TabIndex = 16;
			this.ultraLabel14.Text = "100";
			// 
			// ultraLabel9
			// 
			this.ultraLabel9.AutoSize = true;
			this.ultraLabel9.Location = new System.Drawing.Point(160, 86);
			this.ultraLabel9.Name = "ultraLabel9";
			this.ultraLabel9.Size = new System.Drawing.Size(17, 15);
			this.ultraLabel9.TabIndex = 30;
			this.ultraLabel9.Text = "25";
			// 
			// cmbFechaDesde
			// 
			this.cmbFechaDesde.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Fecha"));
			dateButton1.Caption = "Today";
			this.cmbFechaDesde.DateButtons.Add(dateButton1);
			this.cmbFechaDesde.Format = "dd/MMM/yyyy";
			this.cmbFechaDesde.Location = new System.Drawing.Point(8, 233);
			this.cmbFechaDesde.Name = "cmbFechaDesde";
			this.cmbFechaDesde.NonAutoSizeHeight = 24;
			this.cmbFechaDesde.Size = new System.Drawing.Size(120, 21);
			this.cmbFechaDesde.SpinButtonsVisible = true;
			this.cmbFechaDesde.TabIndex = 31;
			this.cmbFechaDesde.Value = new System.DateTime(2009, 5, 18, 0, 0, 0, 0);
			// 
			// cmbFechaHasta
			// 
			this.cmbFechaHasta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.FechaFin"));
			dateButton2.Caption = "Today";
			this.cmbFechaHasta.DateButtons.Add(dateButton2);
			this.cmbFechaHasta.Format = "dd/MMM/yyyy";
			this.cmbFechaHasta.Location = new System.Drawing.Point(152, 233);
			this.cmbFechaHasta.Name = "cmbFechaHasta";
			this.cmbFechaHasta.NonAutoSizeHeight = 24;
			this.cmbFechaHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbFechaHasta.SpinButtonsVisible = true;
			this.cmbFechaHasta.TabIndex = 32;
			this.cmbFechaHasta.Value = new System.DateTime(2009, 5, 18, 0, 0, 0, 0);
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.idProyecto"));
			this.cmbProyecto.DataSource = this.cdvProyecto;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
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
																										 ultraGridColumn10});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Location = new System.Drawing.Point(136, 302);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(136, 22);
			this.cmbProyecto.TabIndex = 34;
			this.cmbProyecto.ValueMember = "idProyecto";
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible =1";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			// 
			// ultraLabel12
			// 
			this.ultraLabel12.AutoSize = true;
			this.ultraLabel12.Location = new System.Drawing.Point(32, 267);
			this.ultraLabel12.Name = "ultraLabel12";
			this.ultraLabel12.Size = new System.Drawing.Size(42, 15);
			this.ultraLabel12.TabIndex = 35;
			this.ultraLabel12.Text = "Usuario";
			// 
			// ultraLabel15
			// 
			this.ultraLabel15.AutoSize = true;
			this.ultraLabel15.Location = new System.Drawing.Point(32, 302);
			this.ultraLabel15.Name = "ultraLabel15";
			this.ultraLabel15.Size = new System.Drawing.Size(85, 15);
			this.ultraLabel15.TabIndex = 36;
			this.ultraLabel15.Text = "Centro de Costo";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_CierreCaja.Relation";
			this.ultraGrid1.DataSource = this.cdsDeposito;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 53;
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 58;
			ultraGridColumn13.Header.Caption = "Banco";
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn13.Width = 115;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance2;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 67;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Width = 93;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(296, 17);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(296, 267);
			this.ultraGrid1.TabIndex = 37;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_CierreCaja";
			this.barraDato1.DataNombreId = "idCierreCaja";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsDeposito;
			this.barraDato1.DataTabla = "CierreCaja";
			this.barraDato1.DataTablaHija = "CierreCajaDeposito";
			this.barraDato1.Location = new System.Drawing.Point(296, 302);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 25);
			this.barraDato1.TabIndex = 38;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// ultraLabel16
			// 
			this.ultraLabel16.AutoSize = true;
			this.ultraLabel16.Location = new System.Drawing.Point(136, 233);
			this.ultraLabel16.Name = "ultraLabel16";
			this.ultraLabel16.Size = new System.Drawing.Size(13, 15);
			this.ultraLabel16.TabIndex = 39;
			this.ultraLabel16.Text = "al";
			// 
			// btTotal
			// 
			this.btTotal.Location = new System.Drawing.Point(144, 190);
			this.btTotal.Name = "btTotal";
			this.btTotal.Size = new System.Drawing.Size(40, 25);
			this.btTotal.TabIndex = 41;
			this.btTotal.Text = "Total";
			this.btTotal.Click += new System.EventHandler(this.btTotal_Click);
			// 
			// udsUsuario
			// 
			this.udsUsuario.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn1});
			// 
			// cmbUsuario
			// 
			this.cmbUsuario.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbUsuario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUsuario.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Persona"));
			this.cmbUsuario.DataMember = "Band 0";
			this.cmbUsuario.DataSource = this.udsUsuario;
			this.cmbUsuario.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "Cliente";
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Width = 117;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn16});
			this.cmbUsuario.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbUsuario.DisplayMember = "Usuario";
			this.cmbUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbUsuario.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUsuario.Location = new System.Drawing.Point(136, 267);
			this.cmbUsuario.Name = "cmbUsuario";
			this.cmbUsuario.Size = new System.Drawing.Size(136, 21);
			this.cmbUsuario.TabIndex = 92;
			this.cmbUsuario.ValueMember = "Usuario";
			// 
			// cdsBanco
			// 
			this.cdsBanco.BindingContextCtrl = this;
			this.cdsBanco.DataLibrary = "LibContabilidad";
			this.cdsBanco.DataLibraryUrl = "";
			this.cdsBanco.DataSetDef = "dsCuentaCorriente";
			this.cdsBanco.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsBanco.Name = "cdsBanco";
			this.cdsBanco.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsBanco.SchemaDef = null;
			// 
			// cdvBanco
			// 
			this.cdvBanco.BindingContextCtrl = this;
			this.cdvBanco.DataSet = this.cdsBanco;
			this.cdvBanco.Sort = "Nombre";
			this.cdvBanco.TableName = "";
			this.cdvBanco.TableViewName = "Banco";
			// 
			// cmbBanco
			// 
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataMember = "Banco";
			this.cmbBanco.DataSource = this.cdsBanco;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 4;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 5;
			ultraGridColumn23.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(400, 121);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(136, 94);
			this.cmbBanco.TabIndex = 93;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.Visible = false;
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// ultraLabel17
			// 
			this.ultraLabel17.AutoSize = true;
			this.ultraLabel17.Location = new System.Drawing.Point(96, 9);
			this.ultraLabel17.Name = "ultraLabel17";
			this.ultraLabel17.Size = new System.Drawing.Size(39, 15);
			this.ultraLabel17.TabIndex = 94;
			this.ultraLabel17.Text = "Billetes";
			// 
			// ultraLabel18
			// 
			this.ultraLabel18.AutoSize = true;
			this.ultraLabel18.Location = new System.Drawing.Point(216, 9);
			this.ultraLabel18.Name = "ultraLabel18";
			this.ultraLabel18.Size = new System.Drawing.Size(49, 15);
			this.ultraLabel18.TabIndex = 95;
			this.ultraLabel18.Text = "Monedas";
			// 
			// txtIdCierreCaja
			// 
			this.txtIdCierreCaja.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.idCierreCaja"));
			this.txtIdCierreCaja.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdCierreCaja.FormatString = "#,##0.00";
			this.txtIdCierreCaja.Location = new System.Drawing.Point(8, 0);
			this.txtIdCierreCaja.MinValue = -2147483646;
			this.txtIdCierreCaja.Name = "txtIdCierreCaja";
			this.txtIdCierreCaja.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIdCierreCaja.PromptChar = ' ';
			this.txtIdCierreCaja.Size = new System.Drawing.Size(80, 21);
			this.txtIdCierreCaja.TabIndex = 96;
			// 
			// txtTotal
			// 
			this.txtTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Total"));
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(192, 190);
			this.txtTotal.MinValue = 0;
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(80, 24);
			this.txtTotal.TabIndex = 103;
			// 
			// txtM25
			// 
			this.txtM25.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Moneda25"));
			this.txtM25.Location = new System.Drawing.Point(192, 86);
			this.txtM25.MaskInput = "nnn,nnn";
			this.txtM25.MinValue = 0;
			this.txtM25.Name = "txtM25";
			this.txtM25.PromptChar = ' ';
			this.txtM25.Size = new System.Drawing.Size(80, 22);
			this.txtM25.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtM25.TabIndex = 102;
			// 
			// txtM1
			// 
			this.txtM1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Moneda1"));
			this.txtM1.Location = new System.Drawing.Point(192, 164);
			this.txtM1.MaskInput = "nnn,nnn";
			this.txtM1.MinValue = 0;
			this.txtM1.Name = "txtM1";
			this.txtM1.PromptChar = ' ';
			this.txtM1.Size = new System.Drawing.Size(80, 22);
			this.txtM1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtM1.TabIndex = 101;
			// 
			// txtM5
			// 
			this.txtM5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Moneda5"));
			this.txtM5.Location = new System.Drawing.Point(192, 138);
			this.txtM5.MaskInput = "nnn,nnn";
			this.txtM5.MinValue = 0;
			this.txtM5.Name = "txtM5";
			this.txtM5.PromptChar = ' ';
			this.txtM5.Size = new System.Drawing.Size(80, 22);
			this.txtM5.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtM5.TabIndex = 100;
			// 
			// txtM10
			// 
			this.txtM10.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Moneda10"));
			this.txtM10.Location = new System.Drawing.Point(192, 112);
			this.txtM10.MaskInput = "nnn,nnn";
			this.txtM10.MinValue = 0;
			this.txtM10.Name = "txtM10";
			this.txtM10.PromptChar = ' ';
			this.txtM10.Size = new System.Drawing.Size(80, 22);
			this.txtM10.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtM10.TabIndex = 99;
			// 
			// txtM50
			// 
			this.txtM50.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Moneda50"));
			this.txtM50.Location = new System.Drawing.Point(192, 60);
			this.txtM50.MaskInput = "nnn,nnn";
			this.txtM50.MinValue = 0;
			this.txtM50.Name = "txtM50";
			this.txtM50.PromptChar = ' ';
			this.txtM50.Size = new System.Drawing.Size(80, 22);
			this.txtM50.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtM50.TabIndex = 98;
			// 
			// txtM100
			// 
			this.txtM100.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Moneda100"));
			this.txtM100.Location = new System.Drawing.Point(192, 34);
			this.txtM100.MaskInput = "nnn,nnn";
			this.txtM100.MinValue = 0;
			this.txtM100.Name = "txtM100";
			this.txtM100.PromptChar = ' ';
			this.txtM100.Size = new System.Drawing.Size(80, 22);
			this.txtM100.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtM100.TabIndex = 97;
			// 
			// txtB1
			// 
			this.txtB1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Billete1"));
			this.txtB1.Location = new System.Drawing.Point(56, 190);
			this.txtB1.MaskInput = "nnn,nnn";
			this.txtB1.MinValue = 0;
			this.txtB1.Name = "txtB1";
			this.txtB1.PromptChar = ' ';
			this.txtB1.Size = new System.Drawing.Size(72, 22);
			this.txtB1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtB1.TabIndex = 13;
			// 
			// txtB2
			// 
			this.txtB2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Billete2"));
			this.txtB2.Location = new System.Drawing.Point(56, 164);
			this.txtB2.MaskInput = "nnn,nnn";
			this.txtB2.MinValue = 0;
			this.txtB2.Name = "txtB2";
			this.txtB2.PromptChar = ' ';
			this.txtB2.Size = new System.Drawing.Size(72, 22);
			this.txtB2.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtB2.TabIndex = 11;
			// 
			// txtB5
			// 
			this.txtB5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Billete5"));
			this.txtB5.Location = new System.Drawing.Point(56, 138);
			this.txtB5.MaskInput = "nnn,nnn";
			this.txtB5.MinValue = 0;
			this.txtB5.Name = "txtB5";
			this.txtB5.PromptChar = ' ';
			this.txtB5.Size = new System.Drawing.Size(72, 22);
			this.txtB5.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtB5.TabIndex = 9;
			// 
			// txtB10
			// 
			this.txtB10.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Billete10"));
			this.txtB10.Location = new System.Drawing.Point(56, 112);
			this.txtB10.MaskInput = "nnn,nnn";
			this.txtB10.MinValue = 0;
			this.txtB10.Name = "txtB10";
			this.txtB10.PromptChar = ' ';
			this.txtB10.Size = new System.Drawing.Size(72, 22);
			this.txtB10.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtB10.TabIndex = 7;
			// 
			// txtB20
			// 
			this.txtB20.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Billete20"));
			this.txtB20.Location = new System.Drawing.Point(56, 86);
			this.txtB20.MaskInput = "nnn,nnn";
			this.txtB20.MinValue = 0;
			this.txtB20.Name = "txtB20";
			this.txtB20.PromptChar = ' ';
			this.txtB20.Size = new System.Drawing.Size(72, 22);
			this.txtB20.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtB20.TabIndex = 5;
			// 
			// txtB50
			// 
			this.txtB50.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Billete50"));
			this.txtB50.Location = new System.Drawing.Point(56, 60);
			this.txtB50.MaskInput = "nnn,nnn";
			this.txtB50.MinValue = 0;
			this.txtB50.Name = "txtB50";
			this.txtB50.PromptChar = ' ';
			this.txtB50.Size = new System.Drawing.Size(72, 22);
			this.txtB50.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtB50.TabIndex = 3;
			// 
			// txtB100
			// 
			this.txtB100.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsDeposito, "_CierreCaja.Billete100"));
			this.txtB100.Location = new System.Drawing.Point(56, 34);
			this.txtB100.MaskInput = "nnn,nnn";
			this.txtB100.MinValue = 0;
			this.txtB100.Name = "txtB100";
			this.txtB100.PromptChar = ' ';
			this.txtB100.Size = new System.Drawing.Size(72, 22);
			this.txtB100.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtB100.TabIndex = 0;
			// 
			// CierreCaja
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(616, 342);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtM25);
			this.Controls.Add(this.txtM1);
			this.Controls.Add(this.txtM5);
			this.Controls.Add(this.txtM10);
			this.Controls.Add(this.txtM50);
			this.Controls.Add(this.txtM100);
			this.Controls.Add(this.txtIdCierreCaja);
			this.Controls.Add(this.ultraLabel18);
			this.Controls.Add(this.ultraLabel17);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.cmbUsuario);
			this.Controls.Add(this.btTotal);
			this.Controls.Add(this.ultraLabel16);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.ultraLabel15);
			this.Controls.Add(this.ultraLabel12);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.cmbFechaHasta);
			this.Controls.Add(this.cmbFechaDesde);
			this.Controls.Add(this.ultraLabel9);
			this.Controls.Add(this.ultraLabel8);
			this.Controls.Add(this.ultraLabel10);
			this.Controls.Add(this.ultraLabel11);
			this.Controls.Add(this.ultraLabel13);
			this.Controls.Add(this.ultraLabel14);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.txtB1);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.txtB2);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.txtB5);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.txtB10);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.txtB20);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.txtB50);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtB100);
			this.Name = "CierreCaja";
			this.Text = "Cierre de Caja";
			this.Load += new System.EventHandler(this.CierreCaja_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCierreCaja)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtM25)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtM1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtM5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtM10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtM50)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtM100)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB50)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtB100)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
//			DateTime dtFechaIni = (DateTime) dtFechaDesde.Value;
//			DateTime dtFechaFin = (DateTime) dtFechaHasta.Value;
//			string stSelect = "Exec CreaCierreCaja '" + dtFechaIni.ToString("yyyyMMdd")
//				+ "', '" + dtFechaFin.ToString("yyyyMMdd") + "'";
//			if (chkTipoTrans.Checked) stSelect += ", 0";
//			else stSelect += ", " + cmbTipoFact.EditValue.ToString();
//			Funcion.EjecutaSQL(cdsCliente, stSelect, false);
//			string stFiltro = "";
//			string stTitulo = "";
//			if (this.chkUsuario.Checked == false)
//			{
//				stFiltro = " {TmpCierre.Usuario} = '" 
//					+ this.cmbUsuario.Value.ToString() + "'";
//				stTitulo += "Usuario " + cmbUsuario.Text.Trim() + " ";
//			}
//			Reporte miRep = new Reporte("CierreCaja", stFiltro);
//			miRep.MdiParent = this.MdiParent;
//			miRep.Titulo(stTitulo);
//			miRep.Show();
			Cursor = Cursors.Default;
		}


		private void btTotal_Click(object sender, System.EventArgs e)
		{
			double dTotal = 0;
			dTotal = (int) txtB100.Value * 100.0 + (int)txtB50.Value * 50.0
				+ (int)txtB20.Value * 20.0 + (int)txtB10.Value * 10.0
				+ (int)txtB5.Value * 5.0 + (int)txtB2.Value * 2 +(int) txtB1.Value;
			double dTotalM = (int)txtM100.Value * 100.0 +(int) txtM50.Value * 50.0
				+(int) txtM25.Value * 25.0 + (int)txtM10.Value * 10.0
				+ (int)txtM5.Value * 5.0 + (int)txtM1.Value;
			dTotal += dTotalM/100.0;
			txtTotal.Value = dTotal;
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Borrar " + this.Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea borrar el Registro?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;
			string stSelect = "Delete cierrecajadeposito Where idCierreCaja = " + txtIdCierreCaja.Value.ToString();
			Funcion.EjecutaSQL(cdsDeposito, stSelect, true);
			barraDato1.BorraRegistro();
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene Acceso a Crear " + this.Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDato1.CrearRegistro();
			cmbFechaDesde.Value = DateTime.Today;
			cmbFechaHasta.Value = DateTime.Today;
			cmbUsuario.Value = stUsuario;
			cmbProyecto.Value = IdProyecto;
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No puede Editar este Registro ",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDato1.EditarRegistro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			btTotal_Click(this, e);
			barraDato1.GrabaRegistro();
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir " + this.Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.Cursor = Cursors.WaitCursor;
			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@idCierreCaja";
			discreteVal.Value = int.Parse(txtIdCierreCaja.Value.ToString());
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			Reporte miReporte = new Reporte("CuadreDeposito.Rpt", "");
			miReporte.MdiParent = this.MdiParent;
			miReporte.Titulo("Cuadre deposito");
			//			if (chkResumen.Checked) miReporte.Resumen();
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Show();
			this.Cursor = Cursors.Default;
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir " + this.Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.Cursor = Cursors.WaitCursor;
			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@idCierreCaja";
			discreteVal.Value = int.Parse(txtIdCierreCaja.Value.ToString());
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			Reporte miReporte = new Reporte("CuadreCaja.Rpt", "");
			miReporte.MdiParent = this.MdiParent;
			miReporte.Titulo("Cuadre de Caja");
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Show();
			this.Cursor = Cursors.Default;

		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}
		private Acceso miAcceso;

		int IdProyecto = 0;
		string stUsuario = "";
		private void CierreCaja_Load(object sender, System.EventArgs e)
		{
			txtIdCierreCaja.Width = 0;
			string stExec = "Exec ProyectoPredefinido 1";
			IdProyecto = Funcion.iEscalarSQL(cdsBanco, stExec + ", 1", true);
			stUsuario = Funcion.sEscalarSQL(cdsBanco, "Select suser_sname() ", true);

			miAcceso =new Acceso("MFCC", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			string stUsuarios = "Select Distinct(IsNull(Usuario, '')) From Compra";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsProyecto, stUsuarios, true);
			udsUsuario.Rows.Clear();
			int iCuenta = 0;
			while(miReader.Read())
			{
				udsUsuario.Rows.Add();
				udsUsuario.Rows[iCuenta][0] = miReader.GetString(0);
				iCuenta ++;
			}
			miReader.Close();
			barraDato1.ProximoId(4);
			barraDato1.HabilitaControles(false);
		}

		private void cdsDeposito_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsDeposito.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
