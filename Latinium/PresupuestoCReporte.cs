using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de PresupuestoCReporte.
	/// </summary>
	public class PresupuestoCReporte : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaDesde;
		private DevExpress.XtraEditors.CheckEdit chkRubro;
		private DevExpress.XtraEditors.CheckEdit chkSubCapitulo;
		private DevExpress.XtraEditors.CheckEdit chkCapitulo;
		private DevExpress.XtraEditors.CheckEdit chkObra;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private C1.Data.C1DataSet cdsPresupuestoC;
		private C1.Data.C1DataView cdvReporte;
		private C1.Data.C1DataSet cdsReporte;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbObra;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRubro;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubCapitulo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCapitulo;
		private System.Windows.Forms.ErrorProvider errorProvider;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public PresupuestoCReporte()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Reporte", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Filtro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoFiltro");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvReporte = new C1.Data.C1DataView();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.dtFechaHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.chkRubro = new DevExpress.XtraEditors.CheckEdit();
			this.chkSubCapitulo = new DevExpress.XtraEditors.CheckEdit();
			this.chkCapitulo = new DevExpress.XtraEditors.CheckEdit();
			this.chkObra = new DevExpress.XtraEditors.CheckEdit();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.cdsPresupuestoC = new C1.Data.C1DataSet();
			this.cmbObra = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbRubro = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbSubCapitulo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCapitulo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkRubro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubCapitulo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCapitulo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkObra.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPresupuestoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbObra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRubro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubCapitulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCapitulo)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DataSource = this.cdvReporte;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 8;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbReporte.DisplayMember = "Reporte";
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(40, 24);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(184, 21);
			this.cmbReporte.TabIndex = 107;
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
			this.cdvReporte.RowFilter = "tipo=7 and idgrupo=5";
			this.cdvReporte.TableName = "";
			this.cdvReporte.TableViewName = "Reporte";
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
			// dtFechaHasta
			// 
			this.dtFechaHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaHasta.DateButtons.Add(dateButton1);
			this.dtFechaHasta.Format = "dd/MMM/yyyy";
			this.dtFechaHasta.Location = new System.Drawing.Point(104, 96);
			this.dtFechaHasta.Name = "dtFechaHasta";
			this.dtFechaHasta.NonAutoSizeHeight = 23;
			this.dtFechaHasta.Size = new System.Drawing.Size(112, 21);
			this.dtFechaHasta.SpinButtonsVisible = true;
			this.dtFechaHasta.TabIndex = 106;
			this.dtFechaHasta.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.AllowMonthSelection = true;
			this.dtFechaDesde.AllowWeekSelection = true;
			this.dtFechaDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaDesde.DateButtons.Add(dateButton2);
			this.dtFechaDesde.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.dtFechaDesde.Format = "dd/MMM/yyyy";
			this.dtFechaDesde.Location = new System.Drawing.Point(104, 64);
			this.dtFechaDesde.Name = "dtFechaDesde";
			this.dtFechaDesde.NonAutoSizeHeight = 23;
			this.dtFechaDesde.Size = new System.Drawing.Size(112, 21);
			this.dtFechaDesde.SpinButtonsVisible = true;
			this.dtFechaDesde.TabIndex = 105;
			this.dtFechaDesde.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// chkRubro
			// 
			this.chkRubro.EditValue = true;
			this.chkRubro.Location = new System.Drawing.Point(280, 111);
			this.chkRubro.Name = "chkRubro";
			// 
			// chkRubro.Properties
			// 
			this.chkRubro.Properties.Caption = "Todo Rubro";
			this.chkRubro.Size = new System.Drawing.Size(112, 19);
			this.chkRubro.TabIndex = 103;
			this.chkRubro.CheckedChanged += new System.EventHandler(this.chkRubro_CheckedChanged);
			// 
			// chkSubCapitulo
			// 
			this.chkSubCapitulo.EditValue = true;
			this.chkSubCapitulo.Location = new System.Drawing.Point(280, 87);
			this.chkSubCapitulo.Name = "chkSubCapitulo";
			// 
			// chkSubCapitulo.Properties
			// 
			this.chkSubCapitulo.Properties.Caption = "Todo Sub Capitulo";
			this.chkSubCapitulo.Size = new System.Drawing.Size(120, 19);
			this.chkSubCapitulo.TabIndex = 101;
			this.chkSubCapitulo.CheckedChanged += new System.EventHandler(this.chkSubCapitulo_CheckedChanged);
			// 
			// chkCapitulo
			// 
			this.chkCapitulo.EditValue = true;
			this.chkCapitulo.Location = new System.Drawing.Point(280, 63);
			this.chkCapitulo.Name = "chkCapitulo";
			// 
			// chkCapitulo.Properties
			// 
			this.chkCapitulo.Properties.Caption = "Todo Capitulo";
			this.chkCapitulo.Size = new System.Drawing.Size(104, 19);
			this.chkCapitulo.TabIndex = 99;
			this.chkCapitulo.CheckedChanged += new System.EventHandler(this.chkCapitulo_CheckedChanged);
			// 
			// chkObra
			// 
			this.chkObra.EditValue = true;
			this.chkObra.Location = new System.Drawing.Point(280, 39);
			this.chkObra.Name = "chkObra";
			// 
			// chkObra.Properties
			// 
			this.chkObra.Properties.Caption = "Toda Obra";
			this.chkObra.Size = new System.Drawing.Size(96, 19);
			this.chkObra.TabIndex = 97;
			this.chkObra.CheckedChanged += new System.EventHandler(this.chkObra_CheckedChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(40, 96);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 96;
			this.label2.Text = "Hasta:";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(40, 64);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 95;
			this.label1.Text = "Desde:";
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(144, 136);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 94;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(48, 136);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 93;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// cdsPresupuestoC
			// 
			this.cdsPresupuestoC.BindingContextCtrl = this;
			this.cdsPresupuestoC.DataLibrary = "LibFacturacion";
			this.cdsPresupuestoC.DataLibraryUrl = "";
			this.cdsPresupuestoC.DataSetDef = "dsPresupuestoC";
			this.cdsPresupuestoC.EnforceConstraints = false;
			this.cdsPresupuestoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPresupuestoC.Name = "cdsPresupuestoC";
			this.cdsPresupuestoC.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPresupuestoC.SchemaDef = null;
			this.cdsPresupuestoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsPresupuestoC_BeforeFill);
			// 
			// cmbObra
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbObra.Appearance = appearance1;
			this.cmbObra.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbObra.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbObra.DisplayMember = "";
			this.cmbObra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbObra.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbObra.Enabled = false;
			this.cmbObra.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbObra.Location = new System.Drawing.Point(400, 40);
			this.cmbObra.Name = "cmbObra";
			this.cmbObra.Size = new System.Drawing.Size(176, 21);
			this.cmbObra.TabIndex = 108;
			this.cmbObra.ValueMember = "";
			// 
			// cmbRubro
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRubro.Appearance = appearance2;
			this.cmbRubro.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRubro.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRubro.DisplayMember = "";
			this.cmbRubro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRubro.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRubro.Enabled = false;
			this.cmbRubro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRubro.Location = new System.Drawing.Point(400, 112);
			this.cmbRubro.Name = "cmbRubro";
			this.cmbRubro.Size = new System.Drawing.Size(176, 21);
			this.cmbRubro.TabIndex = 109;
			this.cmbRubro.ValueMember = "";
			// 
			// cmbSubCapitulo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubCapitulo.Appearance = appearance3;
			this.cmbSubCapitulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubCapitulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubCapitulo.DisplayMember = "";
			this.cmbSubCapitulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubCapitulo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubCapitulo.Enabled = false;
			this.cmbSubCapitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubCapitulo.Location = new System.Drawing.Point(400, 88);
			this.cmbSubCapitulo.Name = "cmbSubCapitulo";
			this.cmbSubCapitulo.Size = new System.Drawing.Size(176, 21);
			this.cmbSubCapitulo.TabIndex = 110;
			this.cmbSubCapitulo.ValueMember = "";
			// 
			// cmbCapitulo
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCapitulo.Appearance = appearance4;
			this.cmbCapitulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCapitulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCapitulo.DisplayMember = "";
			this.cmbCapitulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCapitulo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCapitulo.Enabled = false;
			this.cmbCapitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCapitulo.Location = new System.Drawing.Point(400, 64);
			this.cmbCapitulo.Name = "cmbCapitulo";
			this.cmbCapitulo.Size = new System.Drawing.Size(176, 21);
			this.cmbCapitulo.TabIndex = 111;
			this.cmbCapitulo.ValueMember = "";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// PresupuestoCReporte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(584, 181);
			this.Controls.Add(this.cmbCapitulo);
			this.Controls.Add(this.cmbSubCapitulo);
			this.Controls.Add(this.cmbRubro);
			this.Controls.Add(this.cmbObra);
			this.Controls.Add(this.cmbReporte);
			this.Controls.Add(this.dtFechaHasta);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.chkRubro);
			this.Controls.Add(this.chkSubCapitulo);
			this.Controls.Add(this.chkCapitulo);
			this.Controls.Add(this.chkObra);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.MaximizeBox = false;
			this.Name = "PresupuestoCReporte";
			this.Text = "Reportes Analisis";
			this.Load += new System.EventHandler(this.PresupuestoCReporte_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkRubro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubCapitulo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCapitulo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkObra.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPresupuestoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbObra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRubro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubCapitulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCapitulo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsPresupuestoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			//carga los datos con la conexion a la base
			Funcion miFuncion = new Funcion();
			cdsPresupuestoC.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
		//llena los combos con informacion
		private void llenaCombos()
		{
			string stSelect="";
			//combo Obra
			stSelect ="Select idPresupuestoC,Obra,NotasObra from P_Presupuesto order by obra";
			cmbObra.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);			
			cmbObra.DisplayMember = "Obra";
			cmbObra.ValueMember = "idPresupuestoC";
			cmbObra.Rows.Band.Columns["idPresupuestoC"].Hidden = true;

			//combo Capitulo
			stSelect ="Select idcapitulo,capitulo,notascapitulo from P_PresupuestoCapitulo order by capitulo";
			cmbCapitulo.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);			
			cmbCapitulo.DisplayMember = "capitulo";
			cmbCapitulo.ValueMember = "idcapitulo";
			cmbCapitulo.Rows.Band.Columns["idcapitulo"].Hidden = true;

			//combo SubCapitulo
			stSelect ="Select idSubcapitulo,Subcapitulo,notas from P_PresupuestoSubCapitulo order by Subcapitulo";
			cmbSubCapitulo.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);			
			cmbSubCapitulo.DisplayMember = "Subcapitulo";
			cmbSubCapitulo.ValueMember = "idSubcapitulo";
			cmbSubCapitulo.Rows.Band.Columns["idSubcapitulo"].Hidden = true;

			//combo Rubro
			stSelect ="Select idPresupuestoCDet,Rubro,Descripcion from P_PresupuestoDetalle order by Rubro";
			cmbRubro.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);			
			cmbRubro.DisplayMember = "Rubro";
			cmbRubro.ValueMember = "idPresupuestoCDet";
			cmbRubro.Rows.Band.Columns["idPresupuestoCDet"].Hidden = true;

		}

		private void PresupuestoCReporte_Load(object sender, System.EventArgs e)
		{
			//inicia las fechas con la fecha actual 
			dtFechaDesde.Value = DateTime.Today.AddMonths(-1);
			dtFechaHasta.Value= DateTime.Today;
			llenaCombos();
		}
		#region Control de la funcionalidad cuando se quita el visto o se pone 
		private void chkObra_CheckedChanged(object sender, System.EventArgs e)
		{
			if(!chkObra.Checked)
				cmbObra.Enabled=true;
			else
				cmbObra.Enabled=false;
		}

		private void chkCapitulo_CheckedChanged(object sender, System.EventArgs e)
		{
			if(!chkCapitulo.Checked)
				cmbCapitulo.Enabled=true;
			else
				cmbCapitulo.Enabled=false;

		}

		private void chkSubCapitulo_CheckedChanged(object sender, System.EventArgs e)
		{
			if(!chkSubCapitulo.Checked)
				cmbSubCapitulo.Enabled=true;
			else
				cmbSubCapitulo.Enabled=false;

		}

		private void chkRubro_CheckedChanged(object sender, System.EventArgs e)
		{
			if(!chkRubro.Checked)
				cmbRubro.Enabled=true;
			else
				cmbRubro.Enabled=false;
		}
	#endregion Control de Checks 
		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		#region Funcion para verificar si se selecciono un combo caso contrario sale errorprovider
		private bool Verifica()
		{
			bool bOk = true;
			errorProvider.SetError(cmbObra, "");
			errorProvider.SetError(cmbCapitulo, "");
			errorProvider.SetError(cmbSubCapitulo, "");
			errorProvider.SetError(cmbRubro, "");
			errorProvider.SetError(cmbReporte, "");

			if (cmbReporte.Text.Length <=0)
			{
				errorProvider.SetError(cmbReporte, "Ingrese Reporte");
				bOk = false;
				cmbReporte.Select();
			}

			if (!chkObra.Checked)
				if (cmbObra.Value == null)
				{
					errorProvider.SetError(cmbObra, "Ingrese Obra");
					bOk = false;
					cmbObra.Select();
				}
			if (!chkCapitulo.Checked)
				if (cmbCapitulo.Value == null)
				{
					errorProvider.SetError(cmbCapitulo, "Ingrese Centro de Costo");
					bOk = false;
					cmbCapitulo.Select();
				}
			if (!chkSubCapitulo.Checked)
				if (cmbSubCapitulo.Value == null)
				{
					errorProvider.SetError(cmbSubCapitulo, "Ingrese Proyecto");
					bOk = false;
					cmbSubCapitulo.Select();
				}

			if (!chkRubro.Checked)
				if (cmbRubro.Value == null)
				{
					errorProvider.SetError(cmbRubro, "Ingrese Proyecto");
					bOk = false;
					cmbRubro.Select();
				}

			return bOk;
		}

		#endregion Funcion para verificar si se selecciono un combo caso contrario sale errorprovider
		
		//variable global
		CultureInfo us = new CultureInfo("en-US");

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!this.Verifica()) return;
			DateTime dtFechaIni = (DateTime) dtFechaDesde.Value;
			DateTime dtFechaFin = (DateTime) dtFechaHasta.Value;

			this.Cursor = Cursors.WaitCursor;
			// Definición de titulo
			string stTitulo = cmbReporte.Text.Trim()
				+ " Desde " + dtFechaIni.ToString("dd/MMM/yyyy")
				+ " Hasta " + dtFechaFin.ToString("dd/MMM/yyyy");
			string stReporte = cmbReporte.Text.Trim()+".rpt";
			
			//si el reporte es P_analisis imprimira desde procedimiento almacenado
			if (stReporte == "P_AnalisisC.rpt")
				repAnalisisCostos(stReporte,stTitulo);
			else
				Imprimir(stReporte,stTitulo);

			this.Cursor = Cursors.Default;
		}
		private void Imprimir(string stReporte ,string stTitulo )
		{		
			DateTime dtDesde = (DateTime) dtFechaDesde.Value;
			DateTime dtHasta = (DateTime) dtFechaHasta.Value;
			
			//capura la posicion del texto si este existe si no pone -1
			int posAvance = stReporte.ToUpper().IndexOf("AVANCE");
		  int pos = stReporte.ToUpper().IndexOf("COMPRA");
			
			string stFiltro = "0=0";
			// si existe la palabra compra en el nombre del reporte filtra con las fechas de la compra
			// caso contrario toma las fechas del avance de obra
			if (pos != -1)
			{
				stFiltro += " and {compra.Fecha} >= #" 
					+ dtDesde.ToString("MM/dd/yyyy", us) + "#";
				stFiltro += " and {compra.Fecha} < #" 
					+ dtHasta.AddDays(1).ToString("MM/dd/yyyy", us) + "#";
			}
			if (posAvance != -1)
			{			
				stFiltro += " and {p_presupuestoavance.Fecha} >= #" 
					+ dtDesde.ToString("MM/dd/yyyy", us) + "#";
				stFiltro += " and {p_presupuestoavance.Fecha} < #" 
					+ dtHasta.AddDays(1).ToString("MM/dd/yyyy", us) + "#";
			}

			if (!chkObra.Checked)
					stFiltro += "And {P_Presupuesto.idPresupuestoC} = "+cmbObra.Value;
			if (!chkCapitulo.Checked)
					stFiltro += "And {P_PresupuestoCapitulo.idCapitulo} = "+cmbCapitulo.Value;
			if (!chkSubCapitulo.Checked)
					stFiltro += "And {P_PresupuestoSubCapitulo.idSubCapitulo} = "+cmbSubCapitulo.Value;
			if (!chkRubro.Checked)
					stFiltro += "And {P_PresupuestoDetalle.idPresupuestoCDet} = "+cmbRubro.Value;

			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Titulo(stTitulo);
			miReporte.Show();
		}

		
		private void repAnalisisCostos(string stReporte ,string stTitulo )
		{
			//esta funcion imprime con procedimiento almacenado 
			DateTime dtDesde = (DateTime) dtFechaDesde.Value;
			DateTime dtHasta = (DateTime) dtFechaHasta.Value;
			
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			#region parametros para filtrar Reportes
				#region parametros para filtrar por obra
					ParameterField paramField5 = new ParameterField ();
					ParameterDiscreteValue discreteVal5 = new ParameterDiscreteValue ();
					paramField5.ParameterFieldName = "@idPresupuestoC";
					if (chkObra.Checked)
						discreteVal5.Value = (int) 0;
					else
						discreteVal5.Value = (int) cmbObra.Value;
					paramField5.CurrentValues.Add (discreteVal5);
					paramFields.Add (paramField5);
					#endregion parametros para filtrar por obra
				#region parametros para filtrar por Capitulo
				ParameterField paramField6 = new ParameterField ();
				ParameterDiscreteValue discreteVal6 = new ParameterDiscreteValue ();
				paramField6.ParameterFieldName = "@idCapitulo";
				if (chkCapitulo.Checked)
					discreteVal6.Value = (int) 0;
				else
					discreteVal6.Value = (int) cmbCapitulo.Value;
				paramField6.CurrentValues.Add (discreteVal6);
				paramFields.Add (paramField6);
				#endregion parametros para filtrar por Capitulo
				#region parametros para filtrar por SubCapitulo
				ParameterField paramField7 = new ParameterField ();
				ParameterDiscreteValue discreteVal7 = new ParameterDiscreteValue ();
				paramField7.ParameterFieldName = "@idSubCapitulo";
				if (chkSubCapitulo.Checked)
					discreteVal7.Value = (int) 0;
				else
					discreteVal7.Value = (int) cmbSubCapitulo.Value;
				paramField7.CurrentValues.Add (discreteVal7);
				paramFields.Add (paramField7);
				#endregion parametros para filtrar por SubCapitulo
				#region parametros para filtrar por Rubros
				ParameterField paramField8 = new ParameterField ();
				ParameterDiscreteValue discreteVal8 = new ParameterDiscreteValue ();
				paramField8.ParameterFieldName = "@idPresupuestoCDet";
				if (chkRubro.Checked)
					discreteVal8.Value = (int) 0;
				else
					discreteVal8.Value = (int) cmbRubro.Value;
				paramField8.CurrentValues.Add (discreteVal8);
				paramFields.Add (paramField8);
				#endregion parametros para filtrar por Rubros
				#region parametros para filtrar por Fechas

			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@FechaDesde";
			discreteVal.Value = dtDesde.ToString("yyyy-MM-dd");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@FechaHasta";
			discreteVal1.Value = dtHasta.ToString("yyyy-MM-dd");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);
			#endregion parametros para filtrar por obra
			#endregion parametros para filtrar Reportes
			Reporte miReporte = new Reporte(stReporte, "");
			miReporte.MdiParent = this.MdiParent;
			miReporte.Titulo(stTitulo);
			miReporte.crVista.ParameterFieldInfo = paramFields;			
			miReporte.Show();			
		}


	}
}
