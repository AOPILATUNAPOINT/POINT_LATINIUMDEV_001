using System;
using System.Data;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RepPresupuesto.
	/// </summary>
	public class RepPresupuesto : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvCuenta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCuenta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProyecto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaHasta;
		private C1.Data.C1DataSet cdsSeteoC;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSubProyecto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubProyecto;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsReporte;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataView cdvReporte;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbf;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RepPresupuesto()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Reporte", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Filtro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoFiltro");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.cmbFechaDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbFechaHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.chkCuenta = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkProyecto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.chkSubProyecto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbSubProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvReporte = new C1.Data.C1DataView();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbf = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbf)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbFechaDesde
			// 
			dateButton1.Caption = "Today";
			this.cmbFechaDesde.DateButtons.Add(dateButton1);
			this.cmbFechaDesde.Location = new System.Drawing.Point(80, 26);
			this.cmbFechaDesde.Name = "cmbFechaDesde";
			this.cmbFechaDesde.NonAutoSizeHeight = 23;
			this.cmbFechaDesde.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaDesde.SpinButtonsVisible = true;
			this.cmbFechaDesde.TabIndex = 0;
			this.cmbFechaDesde.Value = new System.DateTime(2011, 3, 15, 0, 0, 0, 0);
			// 
			// cmbFechaHasta
			// 
			dateButton2.Caption = "Today";
			this.cmbFechaHasta.DateButtons.Add(dateButton2);
			this.cmbFechaHasta.Location = new System.Drawing.Point(256, 26);
			this.cmbFechaHasta.Name = "cmbFechaHasta";
			this.cmbFechaHasta.NonAutoSizeHeight = 23;
			this.cmbFechaHasta.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaHasta.SpinButtonsVisible = true;
			this.cmbFechaHasta.TabIndex = 1;
			this.cmbFechaHasta.Value = new System.DateTime(2011, 3, 15, 0, 0, 0, 0);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(32, 26);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 16);
			this.label1.TabIndex = 3;
			this.label1.Text = "Desde";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(208, 26);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 16);
			this.label2.TabIndex = 4;
			this.label2.Text = "Hasta";
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.cdvCuenta;
			ultraGridColumn1.Header.VisiblePosition = 8;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 12;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn4.Header.VisiblePosition = 9;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 2;
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
			ultraGridColumn18.Header.VisiblePosition = 4;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 358;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Hidden = true;
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
			this.cmbCuenta.DisplayMember = "Codigo";
			this.cmbCuenta.Enabled = false;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(144, 103);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(232, 21);
			this.cmbCuenta.TabIndex = 5;
			this.cmbCuenta.ValueMember = "idCuenta";
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.RowFilter = "CodPresup > \'\'";
			this.cdvCuenta.Sort = "Codigo";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			// 
			// chkCuenta
			// 
			this.chkCuenta.Checked = true;
			this.chkCuenta.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCuenta.Location = new System.Drawing.Point(32, 103);
			this.chkCuenta.Name = "chkCuenta";
			this.chkCuenta.Size = new System.Drawing.Size(88, 22);
			this.chkCuenta.TabIndex = 9;
			this.chkCuenta.Text = "Toda Cuenta";
			this.chkCuenta.CheckedChanged += new System.EventHandler(this.chkCuenta_CheckedChanged);
			// 
			// chkProyecto
			// 
			this.chkProyecto.Checked = true;
			this.chkProyecto.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkProyecto.Location = new System.Drawing.Point(32, 138);
			this.chkProyecto.Name = "chkProyecto";
			this.chkProyecto.Size = new System.Drawing.Size(96, 21);
			this.chkProyecto.TabIndex = 11;
			this.chkProyecto.Text = "Toda C. Costo";
			this.chkProyecto.CheckedChanged += new System.EventHandler(this.chkProyecto_CheckedChanged);
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataMember = "Proyecto";
			this.cmbProyecto.DataSource = this.cdsProyecto;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 6;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 7;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 4;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Width = 294;
			ultraGridColumn28.Header.VisiblePosition = 5;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 8;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 9;
			ultraGridColumn30.Hidden = true;
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
																										 ultraGridColumn30});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.Enabled = false;
			this.cmbProyecto.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbProyecto.Location = new System.Drawing.Point(144, 138);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(232, 22);
			this.cmbProyecto.TabIndex = 10;
			this.cmbProyecto.ValueMember = "idProyecto";
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
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			// 
			// chkSubProyecto
			// 
			this.chkSubProyecto.Checked = true;
			this.chkSubProyecto.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkSubProyecto.Location = new System.Drawing.Point(32, 172);
			this.chkSubProyecto.Name = "chkSubProyecto";
			this.chkSubProyecto.Size = new System.Drawing.Size(96, 22);
			this.chkSubProyecto.TabIndex = 13;
			this.chkSubProyecto.Text = "Toda Proyecto";
			this.chkSubProyecto.CheckedChanged += new System.EventHandler(this.chkSubProyecto_CheckedChanged);
			// 
			// cmbSubProyecto
			// 
			this.cmbSubProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubProyecto.DataMember = "SubProyecto";
			this.cmbSubProyecto.DataSource = this.cdsProyecto;
			ultraGridBand3.AddButtonCaption = "Proyecto";
			ultraGridColumn31.Header.VisiblePosition = 3;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn34.Header.VisiblePosition = 4;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 6;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 5;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 1;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 7;
			ultraGridColumn38.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38});
			this.cmbSubProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbSubProyecto.DisplayMember = "Nombre";
			this.cmbSubProyecto.Enabled = false;
			this.cmbSubProyecto.Location = new System.Drawing.Point(144, 172);
			this.cmbSubProyecto.Name = "cmbSubProyecto";
			this.cmbSubProyecto.Size = new System.Drawing.Size(232, 22);
			this.cmbSubProyecto.TabIndex = 12;
			this.cmbSubProyecto.ValueMember = "idSubProyecto";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdsReporte
			// 
			this.cdsReporte.BindingContextCtrl = this;
			this.cdsReporte.DataLibrary = "LibEmpresa";
			this.cdsReporte.DataLibraryUrl = "";
			this.cdsReporte.DataSetDef = "dsReporte";
			this.cdsReporte.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsReporte.Name = "cdsReporte";
			this.cdsReporte.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsReporte.SchemaDef = null;
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DataSource = this.cdvReporte;
			ultraGridColumn39.Header.VisiblePosition = 0;
			ultraGridColumn39.Width = 252;
			ultraGridColumn40.Header.VisiblePosition = 2;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 3;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 1;
			ultraGridColumn42.Width = 220;
			ultraGridColumn43.Header.VisiblePosition = 4;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 5;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 6;
			ultraGridColumn45.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45});
			this.cmbReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbReporte.DisplayMember = "Detalle";
			this.cmbReporte.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(144, 69);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(232, 21);
			this.cmbReporte.TabIndex = 14;
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
			this.cdvReporte.RowFilter = "Tipo = 7";
			this.cdvReporte.TableName = "";
			this.cdvReporte.TableViewName = "Reporte";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(32, 69);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(43, 16);
			this.label3.TabIndex = 15;
			this.label3.Text = "Reporte";
			// 
			// cmbf
			// 
			dateButton3.Caption = "Today";
			this.cmbf.DateButtons.Add(dateButton3);
			this.cmbf.Location = new System.Drawing.Point(272, 241);
			this.cmbf.Name = "cmbf";
			this.cmbf.NonAutoSizeHeight = 23;
			this.cmbf.Size = new System.Drawing.Size(112, 21);
			this.cmbf.SpinButtonsVisible = true;
			this.cmbf.TabIndex = 16;
			this.cmbf.Value = new System.DateTime(2000, 1, 1, 0, 0, 0, 0);
			this.cmbf.Visible = false;
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(72, 216);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 17;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(272, 216);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 18;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// RepPresupuesto
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(400, 264);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbf);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbReporte);
			this.Controls.Add(this.chkSubProyecto);
			this.Controls.Add(this.cmbSubProyecto);
			this.Controls.Add(this.chkProyecto);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.chkCuenta);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.cmbFechaHasta);
			this.Controls.Add(this.cmbFechaDesde);
			this.MaximizeBox = false;
			this.Name = "RepPresupuesto";
			this.Text = "Reporte de Presupuesto";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RepPresupuesto_Closing);
			this.Load += new System.EventHandler(this.RepPresupuesto_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbf)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void chkCuenta_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCuenta.Enabled = !chkCuenta.Checked;
			if (!chkCuenta.Checked) cmbCuenta.Select();
		}

		private void chkProyecto_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbProyecto.Enabled = !chkProyecto.Checked;
			if (!chkProyecto.Checked) cmbProyecto.Select();
		}

		private bool Verifica()
		{
			bool bOk = true;
			errorProvider.SetError(cmbCuenta, "");
			errorProvider.SetError(cmbProyecto, "");
			errorProvider.SetError(cmbSubProyecto, "");
			errorProvider.SetError(cmbReporte, "");

			if (cmbReporte.ActiveRow.Cells["idGrupo"].Value == System.DBNull.Value)
			{
				errorProvider.SetError(cmbReporte, "Ingrese Reporte");
				bOk = false;
				cmbReporte.Select();
			}

			if (!chkCuenta.Checked)
				if (cmbCuenta.Value == null)
				{
					errorProvider.SetError(cmbCuenta, "Ingrese Cuenta");
					bOk = false;
					cmbCuenta.Select();
				}
			if (!chkProyecto.Checked)
				if (cmbProyecto.Value == null)
				{
					errorProvider.SetError(cmbProyecto, "Ingrese Centro de Costo");
					bOk = false;
					cmbProyecto.Select();
				}
			if (!chkSubProyecto.Checked)
				if (cmbSubProyecto.Value == null)
				{
					errorProvider.SetError(cmbSubProyecto, "Ingrese Proyecto");
					bOk = false;
					cmbSubProyecto.Select();
				}

			return bOk;
		}

		CultureInfo us = new CultureInfo("en-US");
		C1.Data.C1DataRow drSeteoC;

		private void SelectPresup() //, string stCentroC)
		{
			DateTime dtFechaIni = (DateTime) cmbFechaDesde.Value;
			DateTime dtFechaFin = (DateTime) cmbFechaHasta.Value;
			dtFechaFin = dtFechaFin.AddDays(1);

			int anio=dtFechaIni.Year; 
			cmbf.Value="01/01/"+ anio.ToString();  
			DateTime dtFechaIni2 = (DateTime) cmbf.Value;

			string stSelect = "Delete TmpPresup";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			// Filtro de Cuenta
//			if (this.chkCuenta.Enabled)
//				if (!this.chkCuenta.Checked)
//				{
//					stFiltro += " AND {Cuenta.idCuenta} = " + 
//						cmbCuenta.Value;
//				}
			// Fin de Filtro de Cuenta

			// Filtro de Proyecto
			string stFiltro = "";
				if (!this.chkProyecto.Checked)
				{
					if ((bool) drSeteoC["ProyectoLinea"])
						stFiltro += " AND PresupuestoDetalle.idProyecto = " + 
							this.cmbProyecto.Value;
					else
						stFiltro += " AND Presupuesto.idProyecto = " + 
							this.cmbProyecto.Value;
				}
			if (!this.chkSubProyecto.Checked)
			{
				if ((bool) drSeteoC["SubProyLinea"])
					stFiltro += " AND PresupuestoDetalle.idSubProyecto = " + 
						this.cmbSubProyecto.Value;
				else
					stFiltro += " AND Presupuesto.idSubProyecto = " + 
						this.cmbSubProyecto.Value;
			}
			// Fin de Filtro de Proyecto

			stFiltro += " And Distribucion = 0 ";

			string stDesde = 	"From PresupuestoDetalle "
				+ "Inner join Presupuesto "
				+ "On Presupuesto.idPresupuesto = "
				+ "PresupuestoDetalle.idPresupuesto ";

			string stFecha = "Where Presupuesto.Fecha < '" 
				+ dtFechaFin.ToString("yyyyMMdd") + "' "
				+  "and  Presupuesto.Fecha >= '" 
				+ dtFechaIni2.ToString("yyyyMMdd") + "' ";

			stSelect = "Insert Into TmpPresup (idCuenta) "
				+ "Select PresupuestoDetalle.idCuenta "
				+ stDesde + stFecha + stFiltro
				+ " Group by idCuenta "
				+ "Order by IdCuenta ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			// Registro de Presupuesto Anterior
			string stAnioIni = "'" + dtFechaIni.ToString("yyyy") + "0101') ";
			stFecha = "Where (Presupuesto.Fecha < '" 
				+ dtFechaIni.ToString("yyyyMMdd") + "' And Presupuesto.Fecha >= " 
				+ stAnioIni + stFiltro
				+ " And TmpPresup.idCuenta = PresupuestoDetalle.idCuenta) ";

			stSelect = "UPDATE TmpPresup SET PresupDA = "
				+ "(SELECT ISNULL(SUM(Debe), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			stSelect = "UPDATE TmpPresup SET PresupHA = "
				+ "(SELECT ISNULL(SUM(Haber), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			// Registro de Presupuesto del Periodo
			stFecha = "Where Presupuesto.Fecha < '" 
				+ dtFechaFin.ToString("yyyyMMdd") 
				+ "' And Presupuesto.Fecha >= '" 
				+ dtFechaIni.ToString("yyyyMMdd") + "' " + stFiltro
				+ " And TmpPresup.idCuenta = PresupuestoDetalle.idCuenta) ";

			stSelect = "UPDATE TmpPresup SET PresupDP = "
				+ "(SELECT ISNULL(SUM(Debe), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			stSelect = "UPDATE TmpPresup SET PresupHP = "
				+ "(SELECT ISNULL(SUM(Haber), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			// Registro de Asiento Anterior
			stDesde = "From AsientoDetalle Inner join Asiento "
				+ "On Asiento.idAsiento = AsientoDetalle.idAsiento ";
			stFiltro = "";
				if (!this.chkProyecto.Checked)
				{
					if ((bool) drSeteoC["ProyectoLinea"])
						stFiltro += " AND AsientoDetalle.idProyecto = " + 
							this.cmbProyecto.Value;
					else
						stFiltro += " AND Asiento.idProyecto = " + 
							this.cmbProyecto.Value;
				}
				if (!this.chkSubProyecto.Checked)
				{
					if ((bool) drSeteoC["SubProyLinea"])
						stFiltro += " AND AsientoDetalle.idSubProyecto = " + 
							this.cmbSubProyecto.Value;
					else
						stFiltro += " AND Asiento.idSubProyecto = " + 
							this.cmbSubProyecto.Value;
				}
			stFiltro += " And Asiento.Borrar = 0 ";

			// Fin de Filtro de Proyecto
			stFecha = "Where Asiento.Fecha < '" 
				+ dtFechaIni.ToString("yyyyMMdd") + "' " + stFiltro
				+ " And TmpPresup.idCuenta = AsientoDetalle.idCuenta) ";
			
			stSelect = "UPDATE TmpPresup SET AsientoDA = "
				+ "(SELECT ISNULL(SUM(Debe), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			stSelect = "UPDATE TmpPresup SET AsientoHA = "
				+ "(SELECT ISNULL(SUM(Haber), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			// Registro de Asiento Periodo
			stFecha = "Where Asiento.Fecha < '" 
				+ dtFechaFin.ToString("yyyyMMdd") 
				+ "' And Asiento.Fecha >= '" 
				+ dtFechaIni.ToString("yyyyMMdd") + "' " + stFiltro
				+ " And TmpPresup.idCuenta = AsientoDetalle.idCuenta) ";
			
			stSelect = "UPDATE TmpPresup SET AsientoDP = "
				+ "(SELECT ISNULL(SUM(Debe), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			stSelect = "UPDATE TmpPresup SET AsientoHP = "
				+ "(SELECT ISNULL(SUM(Haber), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);
		}

		#region Columna Proyecto*CC
		private void SelectPresup1() //, string stCentroC)
		{
			DateTime dtFechaIni = (DateTime) cmbFechaDesde.Value;
			DateTime dtFechaFin = (DateTime) cmbFechaHasta.Value;
			dtFechaFin = dtFechaFin.AddDays(1);

			int anio=dtFechaIni.Year; 
			cmbf.Value="01/01/"+ anio.ToString();  
			DateTime dtFechaIni2 = (DateTime) cmbf.Value;

			string stSelect = "Delete TmpPresup";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			// Filtro de Cuenta
			//			if (this.chkCuenta.Enabled)
			//				if (!this.chkCuenta.Checked)
			//				{
			//					stFiltro += " AND {Cuenta.idCuenta} = " + 
			//						cmbCuenta.Value;
			//				}
			// Fin de Filtro de Cuenta

			// Filtro de Proyecto
			string stFiltro = "";
			if (!this.chkProyecto.Checked)
			{
				if ((bool) drSeteoC["ProyectoLinea"])
					stFiltro += " AND PresupuestoDetalle.idProyecto = " + 
						this.cmbProyecto.Value;
				else
					stFiltro += " AND Proyecto.idProyecto = " + 
						this.cmbProyecto.Value;
				//
				//					stFiltro += " AND Presupuesto.idProyecto = " + 
				//						this.cmbProyecto.Value;
			}
			if (!this.chkSubProyecto.Checked)
			{
				if ((bool) drSeteoC["SubProyLinea"])
					stFiltro += " AND PresupuestoDetalle.idSubProyecto = " + 
						this.cmbSubProyecto.Value;
				else
					stFiltro += " AND subproyectoProyecto.idSubProyecto = " + 
						this.cmbSubProyecto.Value;

				//				stFiltro += " AND Presupuesto.idSubProyecto = " + 
				//					this.cmbSubProyecto.Value;
			}
			// Fin de Filtro de Proyecto

			stFiltro += " And Distribucion = 0 ";

			//			string stDesde = 	"From PresupuestoDetalle "
			//				+ "Inner join Presupuesto "
			//				+ "On Presupuesto.idPresupuesto = "
			//				+ "PresupuestoDetalle.idPresupuesto ";

			string stDesde = 	"From PresupuestoDetalle "
				+ "Inner join Presupuesto "
				+ "On Presupuesto.idPresupuesto = "
				+ "PresupuestoDetalle.idPresupuesto "
				+ "Inner join proyecto "
				+ "On proyecto.nombre = "
				+ "PresupuestoDetalle.centrocostobr "
				+ "Inner join subproyecto "
				+ "On subproyecto.nombre = "
				+ "PresupuestoDetalle.proyectobr ";


			string stFecha = "Where Presupuesto.Fecha < '" 
				+ dtFechaFin.ToString("yyyyMMdd") + "' "
				+  "and  Presupuesto.Fecha >= '" 
				+ dtFechaIni2.ToString("yyyyMMdd") + "' ";

			//			stSelect = "Insert Into TmpPresup (idCuenta) "
			//				+ "Select PresupuestoDetalle.idCuenta"
			//				+ stDesde + stFecha + stFiltro
			//				+ " Group by idCuenta"
			//				+ "Order by IdCuenta ";

			stSelect = "Insert Into TmpPresup (idCuenta,idproyecto,idsubproyecto) "
				+ "Select PresupuestoDetalle.idCuenta,proyecto.idproyecto,subproyecto.idsubproyecto "
				+ stDesde + stFecha + stFiltro
				+ " Group by idCuenta ,proyecto.idproyecto,subproyecto.idsubproyecto "
				+ "Order by IdCuenta ";
			//						stSelect = "Insert Into TmpPresup (idCuenta) "
			//							+ "Select PresupuestoDetalle.idCuenta "
			//							+ stDesde + stFecha + stFiltro
			//							+ " Group by idCuenta "
			//							+ "Order by IdCuenta ";

			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			// Registro de Presupuesto Anterior
			string stAnioIni = "'" + dtFechaIni.ToString("yyyy") + "0101') ";
			stFecha = "Where (Presupuesto.Fecha < '" 
				+ dtFechaIni.ToString("yyyyMMdd") + "' And Presupuesto.Fecha >= " 
				+ stAnioIni + stFiltro
				+ " And TmpPresup.idCuenta = PresupuestoDetalle.idCuenta) ";

			stSelect = "UPDATE TmpPresup SET PresupDA = "
				+ "(SELECT ISNULL(SUM(Debe), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			stSelect = "UPDATE TmpPresup SET PresupHA = "
				+ "(SELECT ISNULL(SUM(Haber), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			// Registro de Presupuesto del Periodo
			stFecha = "Where Presupuesto.Fecha < '" 
				+ dtFechaFin.ToString("yyyyMMdd") 
				+ "' And Presupuesto.Fecha >= '" 
				+ dtFechaIni.ToString("yyyyMMdd") + "' " + stFiltro
				+ " And TmpPresup.idCuenta = PresupuestoDetalle.idCuenta) ";

			stSelect = "UPDATE TmpPresup SET PresupDP = "
				+ "(SELECT ISNULL(SUM(Debe), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			stSelect = "UPDATE TmpPresup SET PresupHP = "
				+ "(SELECT ISNULL(SUM(Haber), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			// Registro de Asiento Anterior
			stDesde = "From AsientoDetalle Inner join Asiento "
				+ "On Asiento.idAsiento = AsientoDetalle.idAsiento ";
			stFiltro = "";
			if (!this.chkProyecto.Checked)
			{
				if ((bool) drSeteoC["ProyectoLinea"])
					stFiltro += " AND AsientoDetalle.idProyecto = " + 
						this.cmbProyecto.Value;
				else
					stFiltro += " AND Asiento.idProyecto = " + 
						this.cmbProyecto.Value;
			}
			if (!this.chkSubProyecto.Checked)
			{
				if ((bool) drSeteoC["SubProyLinea"])
					stFiltro += " AND AsientoDetalle.idSubProyecto = " + 
						this.cmbSubProyecto.Value;
				else
					stFiltro += " AND Asiento.idSubProyecto = " + 
						this.cmbSubProyecto.Value;
			}
			stFiltro += " And Asiento.Borrar = 0 ";

			// Fin de Filtro de Proyecto
			stFecha = "Where Asiento.Fecha < '" 
				+ dtFechaIni.ToString("yyyyMMdd") + "' " + stFiltro
				+ " And TmpPresup.idCuenta = AsientoDetalle.idCuenta) ";
			
			stSelect = "UPDATE TmpPresup SET AsientoDA = "
				+ "(SELECT ISNULL(SUM(Debe), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			stSelect = "UPDATE TmpPresup SET AsientoHA = "
				+ "(SELECT ISNULL(SUM(Haber), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			// Registro de Asiento Periodo
			stFecha = "Where Asiento.Fecha < '" 
				+ dtFechaFin.ToString("yyyyMMdd") 
				+ "' And Asiento.Fecha >= '" 
				+ dtFechaIni.ToString("yyyyMMdd") + "' " + stFiltro
				+ " And TmpPresup.idCuenta = AsientoDetalle.idCuenta) ";
			
			stSelect = "UPDATE TmpPresup SET AsientoDP = "
				+ "(SELECT ISNULL(SUM(Debe), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);

			stSelect = "UPDATE TmpPresup SET AsientoHP = "
				+ "(SELECT ISNULL(SUM(Haber), 0) "
				+ stDesde + stFecha
				+ " FROM TmpPresup, PresupuestoDetalle ";
			Funcion.EjecutaSQL(cdsProyecto, stSelect);
		}

		#endregion Columna Proyecto*CC
		private void CreaPresupMov(int idCuenta)
		{
			DateTime dtFechaIni = (DateTime) cmbFechaDesde.Value;
			DateTime dtFechaFin = (DateTime) cmbFechaHasta.Value;

			string stExec = string.Format("Exec PresupuestoMovimiento {0}, '{1}', '{2}'",
				idCuenta, dtFechaIni, dtFechaFin);
			Funcion.EjecutaSQL(cdsCuenta, stExec);
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!this.Verifica()) return;
			DateTime dtIni = DateTime.Now;

			DateTime dtFechaIni = (DateTime) cmbFechaDesde.Value;
			DateTime dtFechaFin = (DateTime) cmbFechaHasta.Value;

			int anio=dtFechaIni.Year; 
			cmbf.Value="01/01/"+ anio.ToString();  
			DateTime dtFechaIni2 = (DateTime) cmbf.Value;


			string stFiltro = "{Cuenta.CodPresup} > '' ";

			this.Cursor = Cursors.WaitCursor;

			// Definición de titulo
			string stTitulo = cmbReporte.Text.Trim()
				+ " Desde " + dtFechaIni.ToString("dd/MMM/yyyy")
				+ " Hasta " + dtFechaFin.ToString("dd/MMM/yyyy");

			if (!chkProyecto.Checked)
				stTitulo += " Centro de Costo " + cmbProyecto.Text;
			if (!chkSubProyecto.Checked)
				stTitulo += " Proyecto " + cmbSubProyecto.Text;

			string stReporte = "";
			switch( (int) cmbReporte.ActiveRow.Cells["idGrupo"].Value)
			{
				case 0:
					stReporte = "BalancePresupuesto.rpt";
					SelectPresup();
					break;
				case 1:
					stReporte = "BalancePresupuestoC.rpt";
					SelectPresup();
					break;
				case 2:
					stReporte = "BalancePresupMes.rpt";
					stFiltro = " {Presupuesto.Distribucion} = false ";
					break;
				case 3:
					stReporte = "PresupMovimiento.rpt";
					if (cmbCuenta.Value == null) return;
					CreaPresupMov((int) cmbCuenta.Value);
					stFiltro = "";
					break;
				case 4:
					stReporte = "BalancePresupuestoCHOR.rpt";
					string stelect = "Exec P_PresupuestoMeses";
					Funcion.EjecutaSQL(cdsCuenta,stelect);
					break;
				case 5:
					stReporte = "BalancePresupuestoCC.rpt";
					SelectPresup1();
					break;
				default:
					stReporte = cmbReporte.ActiveRow.Cells["Reporte"].Value.ToString();
					stFiltro = "";
					break;
			}

			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Titulo(stTitulo);
			miReporte.Show();
			Cursor = Cursors.Default;
			TimeSpan ts = DateTime.Now.Subtract(dtIni);
			btAceptar.ToolTip = "Tiempo de ultimo reporte : " + ts.TotalSeconds.ToString("n2");
		}

		private void RepPresupuesto_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'RepPresupuesto'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
			drSeteoC = cdsSeteoC.TableViews["SeteoC"].Rows[0];
			cmbReporte.Value = 112;
			cmbFechaDesde.Value = new DateTime (DateTime.Today.Year, 1, 1);
			cmbFechaHasta.Value = new DateTime (DateTime.Today.Year + 1, 1, 1).AddDays(-1);
		}

		private void chkSubProyecto_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbSubProyecto.Enabled = !chkSubProyecto.Checked;
			if (!chkSubProyecto.Checked) cmbSubProyecto.Select();
		}

		private void RepPresupuesto_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'RepPresupuesto'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
		}
	}
}
