using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RepContabilidad.
	/// </summary>
	public class RepContabilidad : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.RadioGroup rdgTipo;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private DevExpress.XtraEditors.DateEdit dtFechaDesde;
		private DevExpress.XtraEditors.DateEdit dtFechaHasta;
		private DevExpress.XtraEditors.CheckEdit chkCuenta;
		private DevExpress.XtraEditors.CheckEdit chkSubProyecto;
		private DevExpress.XtraEditors.CheckEdit chkComprobante;
		private C1.Data.C1DataSet cdsSeteoC;
		private DevExpress.XtraEditors.CheckEdit checkEdit1;
		private DevExpress.XtraEditors.CheckEdit checkEdit2;
		private DevExpress.XtraEditors.CheckEdit checkEdit3;
		private DevExpress.XtraEditors.CheckEdit checkEdit4;
		private DevExpress.XtraEditors.CheckEdit checkEdit5;
		private DevExpress.XtraEditors.CheckEdit checkEdit6;
		private DevExpress.XtraEditors.CheckEdit checkEdit7;
		private DevExpress.XtraEditors.CheckEdit checkEdit8;
		private DevExpress.XtraEditors.CheckEdit checkEdit9;
		private DevExpress.XtraEditors.CheckEdit checkEdit10;
		private DevExpress.XtraEditors.CheckEdit checkEdit11;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvCuenta;
		private DevExpress.XtraEditors.LookUpEdit cmbCodHasta;
		private DevExpress.XtraEditors.CheckEdit chkGrupoCuentas;
		private DevExpress.XtraEditors.LookUpEdit cmbComprobante;
		private System.Windows.Forms.GroupBox gbxBalance;
		private System.Windows.Forms.GroupBox gbxMayor;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsAsiento;
		private C1.Data.C1DataView cdvTipoAsiento;
		private DevExpress.XtraEditors.LookUpEdit cmbSubProyecto;
		private DevExpress.XtraEditors.LookUpEdit cmbProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataView cdvSubProyecto;
		private DevExpress.XtraEditors.ComboBoxEdit cmbFechas;
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataSet cdsMoneda;
		private C1.Data.C1DataView cdvMoneda;
		private DevExpress.XtraEditors.LookUpEdit cmbMoneda;
		private C1.Data.C1DataSet cdsReporte;
		private C1.Data.C1DataView cdvReporte;
		private DevExpress.XtraEditors.CheckEdit chkCentroCosto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCentroCosto;
		private DevExpress.XtraEditors.LookUpEdit cmbCodDesde;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private C1.Data.C1DataSet cdsSeguridad;
		private C1.Data.C1DataView cdvCuentaNombre;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRango;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDesde;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtHasta;
		private DevExpress.XtraEditors.CheckEdit chkOrden;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProyecto;
		private DevExpress.XtraEditors.CheckEdit chkSubCentro;
		private DevExpress.XtraEditors.CheckEdit chkUsuario;
		private DevExpress.XtraEditors.CheckEdit chkFechaComp;
		private DevExpress.XtraEditors.CheckEdit chkProyecto;
		private DevExpress.XtraEditors.LookUpEdit cmbUsuario;
		int JuanM =0;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Filtro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idReporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Reporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Filtro");
			this.rdgTipo = new DevExpress.XtraEditors.RadioGroup();
			this.dtFechaDesde = new DevExpress.XtraEditors.DateEdit();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.dtFechaHasta = new DevExpress.XtraEditors.DateEdit();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.chkCuenta = new DevExpress.XtraEditors.CheckEdit();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cmbComprobante = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvTipoAsiento = new C1.Data.C1DataView();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.cmbSubProyecto = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvSubProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cmbProyecto = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cmbCodHasta = new DevExpress.XtraEditors.LookUpEdit();
			this.chkSubProyecto = new DevExpress.XtraEditors.CheckEdit();
			this.chkComprobante = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit1 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit2 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit3 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit4 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit5 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit6 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit7 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit8 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit9 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit10 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit11 = new DevExpress.XtraEditors.CheckEdit();
			this.chkOrden = new DevExpress.XtraEditors.CheckEdit();
			this.gbxBalance = new System.Windows.Forms.GroupBox();
			this.gbxMayor = new System.Windows.Forms.GroupBox();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdvReporte = new C1.Data.C1DataView();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.chkGrupoCuentas = new DevExpress.XtraEditors.CheckEdit();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cmbFechas = new DevExpress.XtraEditors.ComboBoxEdit();
			this.cmbMoneda = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvMoneda = new C1.Data.C1DataView();
			this.cdsMoneda = new C1.Data.C1DataSet();
			this.label3 = new System.Windows.Forms.Label();
			this.chkCentroCosto = new DevExpress.XtraEditors.CheckEdit();
			this.txtCentroCosto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCodDesde = new DevExpress.XtraEditors.LookUpEdit();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.cdvCuentaNombre = new C1.Data.C1DataView();
			this.chkRango = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtDesde = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtHasta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtProyecto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkSubCentro = new DevExpress.XtraEditors.CheckEdit();
			this.chkUsuario = new DevExpress.XtraEditors.CheckEdit();
			this.chkFechaComp = new DevExpress.XtraEditors.CheckEdit();
			this.chkProyecto = new DevExpress.XtraEditors.CheckEdit();
			this.cmbUsuario = new DevExpress.XtraEditors.LookUpEdit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCuenta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodHasta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubProyecto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkComprobante.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit2.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit3.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit4.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit5.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit6.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit7.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit8.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit9.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit10.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit11.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkOrden.Properties)).BeginInit();
			this.gbxBalance.SuspendLayout();
			this.gbxMayor.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupoCuentas.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechas.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMoneda.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMoneda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMoneda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCentroCosto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubCentro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkUsuario.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkFechaComp.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkProyecto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// rdgTipo
			// 
			this.rdgTipo.EditValue = -1;
			this.rdgTipo.Location = new System.Drawing.Point(16, 17);
			this.rdgTipo.Name = "rdgTipo";
			// 
			// rdgTipo.Properties
			// 
			this.rdgTipo.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgTipo.Properties.Appearance.Options.UseBackColor = true;
			this.rdgTipo.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Diario General"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Mayor General"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Balance de Comprobación"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Balance General"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(4, "Balance de Resultados"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(5, "Resumen Asientos")});
			this.rdgTipo.Size = new System.Drawing.Size(168, 216);
			this.rdgTipo.TabIndex = 0;
			this.rdgTipo.SelectedIndexChanged += new System.EventHandler(this.rdgTipo_SelectedIndexChanged);
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.FechaRIni"));
			this.dtFechaDesde.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaDesde.Location = new System.Drawing.Point(80, 319);
			this.dtFechaDesde.Name = "dtFechaDesde";
			// 
			// dtFechaDesde.Properties
			// 
			this.dtFechaDesde.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaDesde.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFechaDesde.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaDesde.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFechaDesde.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaDesde.Size = new System.Drawing.Size(96, 20);
			this.dtFechaDesde.TabIndex = 1;
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
			this.cdsSeteoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoC_BeforeFill);
			// 
			// dtFechaHasta
			// 
			this.dtFechaHasta.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.FechaRFin"));
			this.dtFechaHasta.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaHasta.Location = new System.Drawing.Point(80, 353);
			this.dtFechaHasta.Name = "dtFechaHasta";
			// 
			// dtFechaHasta.Properties
			// 
			this.dtFechaHasta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaHasta.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFechaHasta.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaHasta.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFechaHasta.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaHasta.Size = new System.Drawing.Size(96, 20);
			this.dtFechaHasta.TabIndex = 2;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 319);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 3;
			this.label1.Text = "Desde:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 353);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 4;
			this.label2.Text = "Hasta:";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(520, 310);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 5;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			this.btAceptar.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAceptar_MouseUp);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(616, 310);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 6;
			this.btCancelar.Text = "&Salir";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// chkCuenta
			// 
			this.chkCuenta.EditValue = true;
			this.chkCuenta.Location = new System.Drawing.Point(200, 17);
			this.chkCuenta.Name = "chkCuenta";
			// 
			// chkCuenta.Properties
			// 
			this.chkCuenta.Properties.Caption = "Todas las cuentas";
			this.chkCuenta.Size = new System.Drawing.Size(128, 18);
			this.chkCuenta.TabIndex = 7;
			this.chkCuenta.EnabledChanged += new System.EventHandler(this.chkCuenta_EnabledChanged);
			this.chkCuenta.CheckStateChanged += new System.EventHandler(this.chkCuenta_CheckStateChanged);
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.RowFilter = "Grupo = 0";
			this.cdvCuenta.Sort = "Codigo";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// cmbComprobante
			// 
			this.cmbComprobante.Location = new System.Drawing.Point(336, 155);
			this.cmbComprobante.Name = "cmbComprobante";
			// 
			// cmbComprobante.Properties
			// 
			this.cmbComprobante.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbComprobante.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Tipo", "Tipo", 31, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbComprobante.Properties.DataSource = this.cdvTipoAsiento;
			this.cmbComprobante.Properties.DisplayMember = "Tipo";
			this.cmbComprobante.Properties.NullText = "";
			this.cmbComprobante.Properties.PopupWidth = 300;
			this.cmbComprobante.Properties.ValueMember = "idTipoAsiento";
			this.cmbComprobante.Size = new System.Drawing.Size(160, 20);
			this.cmbComprobante.TabIndex = 10;
			// 
			// cdvTipoAsiento
			// 
			this.cdvTipoAsiento.BindingContextCtrl = this;
			this.cdvTipoAsiento.DataSet = this.cdsAsiento;
			this.cdvTipoAsiento.TableName = "";
			this.cdvTipoAsiento.TableViewName = "AsientoTipo";
			// 
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsAsiento";
			this.cdsAsiento.FillOnRequest = false;
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			this.cdsAsiento.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsiento_BeforeFill);
			// 
			// cmbSubProyecto
			// 
			this.cmbSubProyecto.Location = new System.Drawing.Point(336, 121);
			this.cmbSubProyecto.Name = "cmbSubProyecto";
			// 
			// cmbSubProyecto.Properties
			// 
			this.cmbSubProyecto.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbSubProyecto.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbSubProyecto.Properties.DataSource = this.cdvSubProyecto;
			this.cmbSubProyecto.Properties.DisplayMember = "Nombre";
			this.cmbSubProyecto.Properties.NullText = "";
			this.cmbSubProyecto.Properties.PopupWidth = 300;
			this.cmbSubProyecto.Properties.ValueMember = "idSubProyecto";
			this.cmbSubProyecto.Size = new System.Drawing.Size(160, 20);
			this.cmbSubProyecto.TabIndex = 12;
			// 
			// cdvSubProyecto
			// 
			this.cdvSubProyecto.BindingContextCtrl = this;
			this.cdvSubProyecto.DataSet = this.cdsProyecto;
			this.cdvSubProyecto.Sort = "Nombre";
			this.cdvSubProyecto.TableName = "";
			this.cdvSubProyecto.TableViewName = "SubProyecto";
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
			this.cdsProyecto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProyecto_BeforeFill);
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.Location = new System.Drawing.Point(336, 86);
			this.cmbProyecto.Name = "cmbProyecto";
			// 
			// cmbProyecto.Properties
			// 
			this.cmbProyecto.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbProyecto.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbProyecto.Properties.DataSource = this.cdvProyecto;
			this.cmbProyecto.Properties.DisplayMember = "Nombre";
			this.cmbProyecto.Properties.NullText = "";
			this.cmbProyecto.Properties.PopupWidth = 300;
			this.cmbProyecto.Properties.ValueMember = "idProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(160, 20);
			this.cmbProyecto.TabIndex = 14;
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible = 1";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// cmbCodHasta
			// 
			this.cmbCodHasta.Location = new System.Drawing.Point(336, 52);
			this.cmbCodHasta.Name = "cmbCodHasta";
			// 
			// cmbCodHasta.Properties
			// 
			this.cmbCodHasta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCodHasta.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 66, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbCodHasta.Properties.DataSource = this.cdvCuenta;
			this.cmbCodHasta.Properties.DisplayMember = "CodRapido";
			this.cmbCodHasta.Properties.NullText = "";
			this.cmbCodHasta.Properties.PopupWidth = 300;
			this.cmbCodHasta.Properties.ValueMember = "idCuenta";
			this.cmbCodHasta.Size = new System.Drawing.Size(160, 20);
			this.cmbCodHasta.TabIndex = 16;
			this.cmbCodHasta.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCodHasta_MouseUp);
			// 
			// chkSubProyecto
			// 
			this.chkSubProyecto.EditValue = true;
			this.chkSubProyecto.Location = new System.Drawing.Point(200, 121);
			this.chkSubProyecto.Name = "chkSubProyecto";
			// 
			// chkSubProyecto.Properties
			// 
			this.chkSubProyecto.Properties.Caption = "Todo Proyecto";
			this.chkSubProyecto.Size = new System.Drawing.Size(128, 18);
			this.chkSubProyecto.TabIndex = 23;
			this.chkSubProyecto.CheckedChanged += new System.EventHandler(this.chkSubProyecto_CheckedChanged);
			// 
			// chkComprobante
			// 
			this.chkComprobante.EditValue = true;
			this.chkComprobante.Location = new System.Drawing.Point(200, 155);
			this.chkComprobante.Name = "chkComprobante";
			// 
			// chkComprobante.Properties
			// 
			this.chkComprobante.Properties.Caption = "Todo Tipo de asiento";
			this.chkComprobante.Size = new System.Drawing.Size(128, 18);
			this.chkComprobante.TabIndex = 24;
			this.chkComprobante.CheckedChanged += new System.EventHandler(this.chkComprobante_CheckedChanged);
			// 
			// checkEdit1
			// 
			this.checkEdit1.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.n1"));
			this.checkEdit1.Location = new System.Drawing.Point(16, 17);
			this.checkEdit1.Name = "checkEdit1";
			// 
			// checkEdit1.Properties
			// 
			this.checkEdit1.Properties.Caption = "Nivel 1";
			this.checkEdit1.Size = new System.Drawing.Size(64, 18);
			this.checkEdit1.TabIndex = 25;
			// 
			// checkEdit2
			// 
			this.checkEdit2.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.n2"));
			this.checkEdit2.Location = new System.Drawing.Point(16, 43);
			this.checkEdit2.Name = "checkEdit2";
			// 
			// checkEdit2.Properties
			// 
			this.checkEdit2.Properties.Caption = "Nivel 2";
			this.checkEdit2.Size = new System.Drawing.Size(64, 18);
			this.checkEdit2.TabIndex = 26;
			// 
			// checkEdit3
			// 
			this.checkEdit3.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.n4"));
			this.checkEdit3.Location = new System.Drawing.Point(16, 95);
			this.checkEdit3.Name = "checkEdit3";
			// 
			// checkEdit3.Properties
			// 
			this.checkEdit3.Properties.Caption = "Nivel 4";
			this.checkEdit3.Size = new System.Drawing.Size(64, 18);
			this.checkEdit3.TabIndex = 28;
			// 
			// checkEdit4
			// 
			this.checkEdit4.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.n3"));
			this.checkEdit4.Location = new System.Drawing.Point(16, 69);
			this.checkEdit4.Name = "checkEdit4";
			// 
			// checkEdit4.Properties
			// 
			this.checkEdit4.Properties.Caption = "Nivel 3";
			this.checkEdit4.Size = new System.Drawing.Size(64, 18);
			this.checkEdit4.TabIndex = 27;
			// 
			// checkEdit5
			// 
			this.checkEdit5.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.n5"));
			this.checkEdit5.Location = new System.Drawing.Point(16, 121);
			this.checkEdit5.Name = "checkEdit5";
			// 
			// checkEdit5.Properties
			// 
			this.checkEdit5.Properties.Caption = "Nivel 5";
			this.checkEdit5.Size = new System.Drawing.Size(64, 18);
			this.checkEdit5.TabIndex = 29;
			// 
			// checkEdit6
			// 
			this.checkEdit6.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.t5"));
			this.checkEdit6.Location = new System.Drawing.Point(96, 121);
			this.checkEdit6.Name = "checkEdit6";
			// 
			// checkEdit6.Properties
			// 
			this.checkEdit6.Properties.Caption = "Titulo 5";
			this.checkEdit6.Size = new System.Drawing.Size(64, 18);
			this.checkEdit6.TabIndex = 34;
			// 
			// checkEdit7
			// 
			this.checkEdit7.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.t4"));
			this.checkEdit7.Location = new System.Drawing.Point(96, 95);
			this.checkEdit7.Name = "checkEdit7";
			// 
			// checkEdit7.Properties
			// 
			this.checkEdit7.Properties.Caption = "Titulo 4";
			this.checkEdit7.Size = new System.Drawing.Size(64, 18);
			this.checkEdit7.TabIndex = 33;
			// 
			// checkEdit8
			// 
			this.checkEdit8.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.t3"));
			this.checkEdit8.Location = new System.Drawing.Point(96, 69);
			this.checkEdit8.Name = "checkEdit8";
			// 
			// checkEdit8.Properties
			// 
			this.checkEdit8.Properties.Caption = "Titulo 3";
			this.checkEdit8.Size = new System.Drawing.Size(64, 18);
			this.checkEdit8.TabIndex = 32;
			// 
			// checkEdit9
			// 
			this.checkEdit9.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.t2"));
			this.checkEdit9.Location = new System.Drawing.Point(96, 43);
			this.checkEdit9.Name = "checkEdit9";
			// 
			// checkEdit9.Properties
			// 
			this.checkEdit9.Properties.Caption = "Titulo 2";
			this.checkEdit9.Size = new System.Drawing.Size(64, 18);
			this.checkEdit9.TabIndex = 31;
			// 
			// checkEdit10
			// 
			this.checkEdit10.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.t1"));
			this.checkEdit10.Location = new System.Drawing.Point(96, 17);
			this.checkEdit10.Name = "checkEdit10";
			// 
			// checkEdit10.Properties
			// 
			this.checkEdit10.Properties.Caption = "Titulo 1";
			this.checkEdit10.Size = new System.Drawing.Size(64, 18);
			this.checkEdit10.TabIndex = 30;
			// 
			// checkEdit11
			// 
			this.checkEdit11.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.n6"));
			this.checkEdit11.Location = new System.Drawing.Point(16, 146);
			this.checkEdit11.Name = "checkEdit11";
			// 
			// checkEdit11.Properties
			// 
			this.checkEdit11.Properties.Caption = "Detalles";
			this.checkEdit11.Size = new System.Drawing.Size(72, 18);
			this.checkEdit11.TabIndex = 35;
			// 
			// chkOrden
			// 
			this.chkOrden.Location = new System.Drawing.Point(96, 146);
			this.chkOrden.Name = "chkOrden";
			// 
			// chkOrden.Properties
			// 
			this.chkOrden.Properties.Caption = "Orden";
			this.chkOrden.Size = new System.Drawing.Size(64, 18);
			this.chkOrden.TabIndex = 37;
			// 
			// gbxBalance
			// 
			this.gbxBalance.Controls.Add(this.checkEdit1);
			this.gbxBalance.Controls.Add(this.checkEdit6);
			this.gbxBalance.Controls.Add(this.checkEdit2);
			this.gbxBalance.Controls.Add(this.checkEdit7);
			this.gbxBalance.Controls.Add(this.checkEdit5);
			this.gbxBalance.Controls.Add(this.chkOrden);
			this.gbxBalance.Controls.Add(this.checkEdit11);
			this.gbxBalance.Controls.Add(this.checkEdit10);
			this.gbxBalance.Controls.Add(this.checkEdit9);
			this.gbxBalance.Controls.Add(this.checkEdit3);
			this.gbxBalance.Controls.Add(this.checkEdit4);
			this.gbxBalance.Controls.Add(this.checkEdit8);
			this.gbxBalance.Enabled = false;
			this.gbxBalance.Location = new System.Drawing.Point(520, 9);
			this.gbxBalance.Name = "gbxBalance";
			this.gbxBalance.Size = new System.Drawing.Size(168, 181);
			this.gbxBalance.TabIndex = 44;
			this.gbxBalance.TabStop = false;
			this.gbxBalance.Text = "Balance";
			// 
			// gbxMayor
			// 
			this.gbxMayor.Controls.Add(this.cmbReporte);
			this.gbxMayor.Location = new System.Drawing.Point(520, 224);
			this.gbxMayor.Name = "gbxMayor";
			this.gbxMayor.Size = new System.Drawing.Size(168, 60);
			this.gbxMayor.TabIndex = 45;
			this.gbxMayor.TabStop = false;
			this.gbxMayor.Text = " Reporte ";
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DataSource = this.ultraDataSource1;
			this.cmbReporte.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "Reporte";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 117;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 54;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbReporte.DisplayMember = "Reporte";
			this.cmbReporte.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbReporte.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(16, 26);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(136, 21);
			this.cmbReporte.TabIndex = 43;
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
			this.cdvReporte.RowFilter = "Tipo = 4";
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
			this.cdsReporte.BeforeFill += new C1.Data.FillEventHandler(this.cdsReporte_BeforeFill);
			// 
			// chkGrupoCuentas
			// 
			this.chkGrupoCuentas.EditValue = true;
			this.chkGrupoCuentas.Location = new System.Drawing.Point(200, 52);
			this.chkGrupoCuentas.Name = "chkGrupoCuentas";
			// 
			// chkGrupoCuentas.Properties
			// 
			this.chkGrupoCuentas.Properties.Caption = "Rango de cuentas";
			this.chkGrupoCuentas.Size = new System.Drawing.Size(128, 18);
			this.chkGrupoCuentas.TabIndex = 46;
			this.chkGrupoCuentas.CheckedChanged += new System.EventHandler(this.chkGrupoCuentas_CheckedChanged);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// cmbFechas
			// 
			this.cmbFechas.EditValue = "Hasta Hoy";
			this.cmbFechas.Location = new System.Drawing.Point(16, 284);
			this.cmbFechas.Name = "cmbFechas";
			// 
			// cmbFechas.Properties
			// 
			this.cmbFechas.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbFechas.Properties.Items.AddRange(new object[] {
																															"Hasta Hoy",
																															"Solo Hoy",
																															"Este Mes",
																															"El mes pasado",
																															"Este Año",
																															"El Año pasado",
																															"Periodo Ingresado",
																															"Primer Trimestre",
																															"Segundo Trimestre",
																															"Tercer Trimestre",
																															"Cuarto Trimestre",
																															"Periodo Digitado"});
			this.cmbFechas.Size = new System.Drawing.Size(160, 20);
			this.cmbFechas.TabIndex = 47;
			this.cmbFechas.SelectedIndexChanged += new System.EventHandler(this.cmbFechas_SelectedIndexChanged);
			// 
			// cmbMoneda
			// 
			this.cmbMoneda.Location = new System.Drawing.Point(336, 293);
			this.cmbMoneda.Name = "cmbMoneda";
			// 
			// cmbMoneda.Properties
			// 
			this.cmbMoneda.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbMoneda.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbMoneda.Properties.DataSource = this.cdvMoneda;
			this.cmbMoneda.Properties.DisplayMember = "Nombre";
			this.cmbMoneda.Properties.NullText = "";
			this.cmbMoneda.Properties.PopupWidth = 300;
			this.cmbMoneda.Properties.ValueMember = "idTipoMoneda";
			this.cmbMoneda.Size = new System.Drawing.Size(160, 20);
			this.cmbMoneda.TabIndex = 48;
			this.cmbMoneda.Visible = false;
			// 
			// cdvMoneda
			// 
			this.cdvMoneda.BindingContextCtrl = this;
			this.cdvMoneda.DataSet = this.cdsMoneda;
			this.cdvMoneda.TableName = "";
			this.cdvMoneda.TableViewName = "MonedaTipo";
			// 
			// cdsMoneda
			// 
			this.cdsMoneda.BindingContextCtrl = this;
			this.cdsMoneda.DataLibrary = "LibContabilidad";
			this.cdsMoneda.DataLibraryUrl = "";
			this.cdsMoneda.DataSetDef = "dsMoneda";
			this.cdsMoneda.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsMoneda.Name = "cdsMoneda";
			this.cdsMoneda.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsMoneda.SchemaDef = null;
			this.cdsMoneda.BeforeFill += new C1.Data.FillEventHandler(this.cdsMoneda_BeforeFill);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(200, 293);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 16);
			this.label3.TabIndex = 49;
			this.label3.Text = "Moneda";
			this.label3.Visible = false;
			// 
			// chkCentroCosto
			// 
			this.chkCentroCosto.EditValue = true;
			this.chkCentroCosto.Location = new System.Drawing.Point(200, 190);
			this.chkCentroCosto.Name = "chkCentroCosto";
			// 
			// chkCentroCosto.Properties
			// 
			this.chkCentroCosto.Properties.Caption = "Grupos Centro Costo";
			this.chkCentroCosto.Size = new System.Drawing.Size(128, 18);
			this.chkCentroCosto.TabIndex = 51;
			this.chkCentroCosto.CheckedChanged += new System.EventHandler(this.chkCentroCosto_CheckedChanged);
			// 
			// txtCentroCosto
			// 
			this.txtCentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCentroCosto.Enabled = false;
			this.txtCentroCosto.Location = new System.Drawing.Point(336, 190);
			this.txtCentroCosto.Name = "txtCentroCosto";
			this.txtCentroCosto.Size = new System.Drawing.Size(160, 22);
			this.txtCentroCosto.TabIndex = 52;
			// 
			// cmbCodDesde
			// 
			this.cmbCodDesde.Location = new System.Drawing.Point(336, 17);
			this.cmbCodDesde.Name = "cmbCodDesde";
			// 
			// cmbCodDesde.Properties
			// 
			this.cmbCodDesde.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCodDesde.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 66, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbCodDesde.Properties.DataSource = this.cdvCuenta;
			this.cmbCodDesde.Properties.DisplayMember = "CodRapido";
			this.cmbCodDesde.Properties.NullText = "";
			this.cmbCodDesde.Properties.PopupWidth = 300;
			this.cmbCodDesde.Properties.ValueMember = "idCuenta";
			this.cmbCodDesde.Size = new System.Drawing.Size(160, 20);
			this.cmbCodDesde.TabIndex = 8;
			this.cmbCodDesde.EditValueChanged += new System.EventHandler(this.cmbCodDesde_EditValueChanged);
			this.cmbCodDesde.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCodDesde_MouseUp);
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
			// cdvCuentaNombre
			// 
			this.cdvCuentaNombre.BindingContextCtrl = this;
			this.cdvCuentaNombre.DataSet = this.cdsCuenta;
			this.cdvCuentaNombre.Sort = "Descripcion";
			this.cdvCuentaNombre.TableName = "";
			this.cdvCuentaNombre.TableViewName = "Cuenta";
			// 
			// chkRango
			// 
			this.chkRango.Checked = true;
			this.chkRango.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkRango.Location = new System.Drawing.Point(200, 327);
			this.chkRango.Name = "chkRango";
			this.chkRango.Size = new System.Drawing.Size(64, 22);
			this.chkRango.TabIndex = 53;
			this.chkRango.Text = "Rango";
			this.chkRango.Visible = false;
			// 
			// txtDesde
			// 
			this.txtDesde.Location = new System.Drawing.Point(264, 327);
			this.txtDesde.Name = "txtDesde";
			this.txtDesde.Size = new System.Drawing.Size(80, 22);
			this.txtDesde.TabIndex = 54;
			this.txtDesde.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(360, 327);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(31, 16);
			this.label4.TabIndex = 55;
			this.label4.Text = "hasta";
			this.label4.Visible = false;
			// 
			// txtHasta
			// 
			this.txtHasta.Location = new System.Drawing.Point(416, 327);
			this.txtHasta.Name = "txtHasta";
			this.txtHasta.Size = new System.Drawing.Size(80, 22);
			this.txtHasta.TabIndex = 56;
			this.txtHasta.Visible = false;
			// 
			// txtProyecto
			// 
			this.txtProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyecto.Enabled = false;
			this.txtProyecto.Location = new System.Drawing.Point(336, 224);
			this.txtProyecto.Name = "txtProyecto";
			this.txtProyecto.Size = new System.Drawing.Size(160, 22);
			this.txtProyecto.TabIndex = 59;
			// 
			// chkSubCentro
			// 
			this.chkSubCentro.EditValue = true;
			this.chkSubCentro.Location = new System.Drawing.Point(200, 224);
			this.chkSubCentro.Name = "chkSubCentro";
			// 
			// chkSubCentro.Properties
			// 
			this.chkSubCentro.Properties.Caption = "Grupos de Proyectos";
			this.chkSubCentro.Size = new System.Drawing.Size(128, 18);
			this.chkSubCentro.TabIndex = 58;
			this.chkSubCentro.CheckedChanged += new System.EventHandler(this.chkSubCentro_CheckedChanged);
			// 
			// chkUsuario
			// 
			this.chkUsuario.EditValue = true;
			this.chkUsuario.Location = new System.Drawing.Point(200, 258);
			this.chkUsuario.Name = "chkUsuario";
			// 
			// chkUsuario.Properties
			// 
			this.chkUsuario.Properties.Caption = "Todo Usuario";
			this.chkUsuario.Size = new System.Drawing.Size(112, 18);
			this.chkUsuario.TabIndex = 92;
			this.chkUsuario.CheckedChanged += new System.EventHandler(this.chkUsuario_CheckedChanged);
			// 
			// chkFechaComp
			// 
			this.chkFechaComp.Location = new System.Drawing.Point(32, 250);
			this.chkFechaComp.Name = "chkFechaComp";
			// 
			// chkFechaComp.Properties
			// 
			this.chkFechaComp.Properties.Caption = "Fecha Comprobante";
			this.chkFechaComp.Size = new System.Drawing.Size(128, 18);
			this.chkFechaComp.TabIndex = 94;
			// 
			// chkProyecto
			// 
			this.chkProyecto.EditValue = true;
			this.chkProyecto.Location = new System.Drawing.Point(200, 86);
			this.chkProyecto.Name = "chkProyecto";
			// 
			// chkProyecto.Properties
			// 
			this.chkProyecto.Properties.Caption = "Todo Centro Costo";
			this.chkProyecto.Properties.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Inactive;
			this.chkProyecto.Size = new System.Drawing.Size(128, 18);
			this.chkProyecto.TabIndex = 95;
			this.chkProyecto.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkProyecto_MouseUp);
			this.chkProyecto.CheckedChanged += new System.EventHandler(this.chkProyecto_CheckedChanged);
			// 
			// cmbUsuario
			// 
			this.cmbUsuario.Location = new System.Drawing.Point(336, 258);
			this.cmbUsuario.Name = "cmbUsuario";
			// 
			// cmbUsuario.Properties
			// 
			this.cmbUsuario.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbUsuario.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Usuario", "Usuario", 31, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbUsuario.Properties.DisplayMember = "Usuario";
			this.cmbUsuario.Properties.NullText = "";
			this.cmbUsuario.Properties.PopupWidth = 300;
			this.cmbUsuario.Properties.ValueMember = "Usuario";
			this.cmbUsuario.Size = new System.Drawing.Size(160, 20);
			this.cmbUsuario.TabIndex = 96;
			// 
			// RepContabilidad
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(704, 385);
			this.Controls.Add(this.cmbUsuario);
			this.Controls.Add(this.chkProyecto);
			this.Controls.Add(this.chkFechaComp);
			this.Controls.Add(this.chkUsuario);
			this.Controls.Add(this.txtProyecto);
			this.Controls.Add(this.chkSubCentro);
			this.Controls.Add(this.txtHasta);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtDesde);
			this.Controls.Add(this.chkRango);
			this.Controls.Add(this.txtCentroCosto);
			this.Controls.Add(this.chkCentroCosto);
			this.Controls.Add(this.cmbMoneda);
			this.Controls.Add(this.cmbFechas);
			this.Controls.Add(this.chkGrupoCuentas);
			this.Controls.Add(this.gbxMayor);
			this.Controls.Add(this.gbxBalance);
			this.Controls.Add(this.chkComprobante);
			this.Controls.Add(this.chkSubProyecto);
			this.Controls.Add(this.cmbCodHasta);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.cmbSubProyecto);
			this.Controls.Add(this.cmbComprobante);
			this.Controls.Add(this.cmbCodDesde);
			this.Controls.Add(this.chkCuenta);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.dtFechaHasta);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.rdgTipo);
			this.MaximizeBox = false;
			this.Name = "RepContabilidad";
			this.Text = "Reportes de Contabilidad";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RepContabilidad_Closing);
			this.Load += new System.EventHandler(this.RepContabilidad_Load);
			((System.ComponentModel.ISupportInitialize)(this.rdgTipo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCuenta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodHasta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubProyecto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkComprobante.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit2.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit3.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit4.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit5.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit6.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit7.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit8.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit9.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit10.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit11.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkOrden.Properties)).EndInit();
			this.gbxBalance.ResumeLayout(false);
			this.gbxMayor.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupoCuentas.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechas.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMoneda.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMoneda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMoneda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCentroCosto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubCentro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkUsuario.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkFechaComp.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkProyecto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variable Globales
		C1.Data.C1DataRow drSeteoC;
		double dUtilidad = 0;
		CultureInfo us = new CultureInfo("en-US");

		#endregion Variable Globales
		#region Pantalla Reporte de Contabilidad
		public RepContabilidad()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
			//
		}

		private void RepContabilidad_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'RepContabilidad'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);

			Funcion.EjecutaSQL(cdsAsiento, "Exec AsientoBorraVacio");
			#region Tres Estados en Check
			if (Funcion.bEjecutaSQL(cdsCuenta, "Exec SeteoGExiste 'FTMTE'")) 				
			{
				chkCuenta.Properties.AllowGrayed = true;
				chkGrupoCuentas.Properties.AllowGrayed = true;
				chkProyecto.Properties.AllowGrayed = true;
				chkSubProyecto.Properties.AllowGrayed = true;
				chkComprobante.Properties.AllowGrayed = true;
				chkCentroCosto.Properties.AllowGrayed = true;
				chkSubCentro.Properties.AllowGrayed = true;
				chkUsuario.Properties.AllowGrayed = true;
			}
			#endregion Tres Estados en Check

			#region Asiento
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(cdsAsiento, "Asiento", "idAsiento=0");
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			C1.Data.FilterCondition fcHija = new 
				C1.Data.FilterCondition(cdsAsiento, "AsientoTipo", "idTipoAsiento > 0");
			fcTotal.Add(fcHija);

			C1.Data.FilterCondition fcHija1 = new 
				C1.Data.FilterCondition(cdsAsiento, "AsientoDetalle", "idAsiento = 0");
			fcTotal.Add(fcHija1);
			cdsAsiento.Fill(fcTotal, false);
			#endregion

			string stUsuario = "Select Distinct(IsNull(Usuario, '')) As Usuario From Asiento";
			cmbUsuario.Properties.DataSource = Funcion.dvProcedimiento(cdsAsiento, stUsuario);

			Cursor = Cursors.Default;
			drSeteoC = cdsSeteoC.TableViews["SeteoC"].Rows[0];
			cdvReporte.RowFilter = "Tipo = 4 And idGrupo = 1";
			DataRow[] drRep = cdsReporte.StorageDataSet.Tables["Reporte"].Select(cdvReporte.RowFilter);
			if (drRep.Length > 0) cmbReporte.Value = drRep[0]["idReporte"];
			#region Visibilidad de Combos
			chkCentroCosto_CheckedChanged(this, e);
			chkComprobante_CheckedChanged(this, e);
			chkCuenta_CheckStateChanged(this, e);
			chkGrupoCuentas_CheckedChanged(this, e);
			chkProyecto_CheckedChanged(this, e);
			chkSubCentro_CheckedChanged(this, e);
			chkSubProyecto_CheckedChanged(this, e);
			chkUsuario_CheckedChanged(this, e);
			#endregion Visibilidad combos

			this.rdgTipo.EditValue = 1;
		}

		private void RepContabilidad_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'RepContabilidad'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
		}


		#endregion Pantalla Reporte de Contabilidad
		#region Botones Principales
		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			#region Variables Iniciales
			int iTipoReporte = (int) rdgTipo.EditValue;
			bool bBalance8 = false;
			if (( iTipoReporte == 1 || iTipoReporte == 2 || iTipoReporte == 3 || iTipoReporte == 4 ) && !chkOrden.Checked) 
				bBalance8 = true;
			Reporte miReporte;
			DateTime dtFechaIni = (DateTime) dtFechaDesde.EditValue;
			DateTime dtFechaFin = (DateTime) dtFechaHasta.EditValue;
			dUtilidad = 0;
			string stTitulo = "Desde " + dtFechaIni.ToString("dd/MMM/yyyy")
				+ " Hasta " + dtFechaFin.ToString("dd/MMM/yyyy");
			#endregion Variables Iniciales
			#region Verifica y Graba
			DateTime dtIni = DateTime.Now;
			if (!Verifica()) return;
			Cursor = Cursors.WaitCursor;
			if (Funcion.bEjecutaSQL(cdsCuenta, "Exec SeteoGExiste 'GEUR'"))
			{
				try
				{
					cdsSeteoC.Update();
				}
				catch
				{
				}
			}
			#endregion Verifica y Graba

			#region Filtros
			#region Filtro de Fechas
			string strFiltro = "";
			if (iTipoReporte < 5 && iTipoReporte > 0) // Balances y Mayores
			{
				if (!bBalance8)
				{
					if(chkFechaComp.Checked==true )
					{
						strFiltro = "{AsientoDetalle.FechaComp} < #" 
							+ dtFechaFin.AddDays(1).ToString("MM/dd/yyyy", us) + "#"
							+ " AND {AsientoDetalle.Borrar} = False And {Asiento.Borrar} = False " ;
					
					}
					else
					{
						strFiltro = "{Asiento.Fecha} < #" 
							+ dtFechaFin.AddDays(1).ToString("MM/dd/yyyy", us) + "#"
							+ " AND {AsientoDetalle.Borrar} = False And {Asiento.Borrar} = False " ;
					}
				}
				else // Ejecuta procedimiento en balances
				{
					strFiltro = " ";
				}
			}
			else // Diario y Resumen de Asientos
			{
				if(chkFechaComp.Checked==true )
				{
					strFiltro = "{Asientodetalle.Fechacomp} >= #" 
						+ dtFechaIni.ToString("MM/dd/yyyy", us) + "#"
						+ " AND {Asientodetalle.Fechacomp} < #" 
						+ dtFechaFin.AddDays(1).ToString("MM/dd/yyyy", us) + "#";
					
				}
				else
				{
					strFiltro = "{Asiento.Fecha} >= #" 
						+ dtFechaIni.ToString("MM/dd/yyyy", us) + "#"
						+ " AND {Asiento.Fecha} < #" 
						+ dtFechaFin.AddDays(1).ToString("MM/dd/yyyy", us) + "#";
				}
			}
			#endregion Filtro de Fechas
			#region Filtro de Comprobante
			if (chkComprobante.Enabled)
				if (!chkComprobante.Checked)
				{
					strFiltro += " AND {Asiento.idTipoAsiento} = " + 
						cmbComprobante.EditValue;
					stTitulo += " " + cmbComprobante.Text.Trim();
				}
			#endregion Fin de Filtro de Comprobante
			#region Filtro de Cuenta
			string stFiltroCuenta = "";
			if (chkCuenta.Enabled)
				if (!chkCuenta.Checked)
				{
					if (!chkGrupoCuentas.Checked)
					{
						stFiltroCuenta ="{Cuenta.idCuenta} = " + 
							cmbCodDesde.EditValue;
						strFiltro += " AND {Cuenta.idCuenta} = " + 
							cmbCodDesde.EditValue;

					}
					else
					{
						stFiltroCuenta = "{Cuenta.CodRapido} >= '" + 
							cmbCodDesde.Text.Trim() + "' And {Cuenta.CodRapido} <= '" +
							cmbCodHasta.Text.Trim() + "'";
						if (chkOrden.Checked)
						{
							strFiltro += " AND {Cuenta.CodRapido} >= '" + 
								cmbCodDesde.Text.Trim() + "' And {Cuenta.CodRapido} <= '" +
								cmbCodHasta.Text.Trim() + "'";
						}
						else
						{
							string CodDesde = "Select Codigo FROM Cuenta Where CodRapido = '"
								+ cmbCodDesde.Text.Trim() +"'";
							CodDesde = Funcion.sEscalarSQL(cdsAsiento, CodDesde, true);
							string CodHasta = "Select Codigo FROM Cuenta Where CodRapido = '"
								+ cmbCodHasta.Text.Trim() +"'";
							CodHasta = Funcion.sEscalarSQL(cdsAsiento, CodHasta, true);
						
							stFiltroCuenta = " {TmpBalance.Codigo} >= '" + 
								CodDesde + "' And {TmpBalance.Codigo} <= '" +
								CodHasta + "'";
							////////////////// aqi si es mayor
							///
							if((int) rdgTipo.EditValue == 1)
							{
								stFiltroCuenta = " {TmpMayor.Codigo} >= '" + 
									CodDesde + "' And {TmpMayor.Codigo} <= '" +
									CodHasta + "'";
							}
						}
					}
				}
			#endregion Fin de Filtro de Cuenta
			#region Filtro de Proyecto
			if (chkProyecto.Enabled)
			{
				if (!chkProyecto.Checked)
				{
					if ((bool) drSeteoC["ProyectoLinea"])
						strFiltro += " AND {AsientoDetalle.idProyecto} = " + 
							cmbProyecto.EditValue;
					else
						strFiltro += " AND {Asiento.idProyecto} = " + 
							cmbProyecto.EditValue;
					stTitulo += " C.C. " + cmbProyecto.Text.Trim();
				}
				if (chkProyecto.CheckState == CheckState.Indeterminate)
				{
					if ((bool) drSeteoC["ProyectoLinea"])
						strFiltro += " AND {AsientoDetalle.idProyecto} = 0 ";
					else
						strFiltro += " AND {Asiento.idProyecto} = 0";
					stTitulo += " Sin Centro Costo ";
				}
			}
			#endregion Filtro de Proyecto
			#region Filtro de Usuario
			if (!chkUsuario.Checked)
			{
				strFiltro += " AND {Asiento.Usuario} = '" + 
					cmbUsuario.Text.Trim() + "'";
				stTitulo += " Usuario " + cmbUsuario.Text.ToString().Trim();
			}
			#endregion Filtro de Usuario
			#region Filtro de Centro de Costo
			if (chkCentroCosto.Enabled)
				if (!chkCentroCosto.Checked)
				{
					if ((bool) drSeteoC["ProyectoLinea"])
					{
						strFiltro += " AND {Proyecto.CentroCosto} Like '" + 
							txtCentroCosto.Value.ToString().Trim() + "*'";
					}
					else
					{
						strFiltro += " AND {Proyecto.CentroCosto} Like '" + 
							txtCentroCosto.Value.ToString().Trim() + "*'";
					}

					string stNombreCC = Funcion.sEscalarSQL(cdsAsiento, "select Nombre from proyecto Where CentroCosto = '"
						+ txtCentroCosto.Value.ToString().Trim() + "'");
					stTitulo += " C. Costo " + stNombreCC;
				}
			#endregion Fin de Filtro de Centro de Costo
			#region Filtro de SubCentro de Costo
			if (chkSubCentro.Enabled)
				if (!chkSubCentro.Checked)
				{
					strFiltro += " AND {SubProyecto.CodSubProyecto} Like '" + 
						txtProyecto.Value.ToString().Trim() + "*'";
					stTitulo += " Proyecto " + txtProyecto.Value.ToString().Trim();
				}
			#endregion Fin de Filtro de SubCentro de Costo
			#region Filtro de SubProyecto
			if (chkSubProyecto.Enabled)
				if (!chkSubProyecto.Checked)
				{
					if ((bool) drSeteoC["SubProyLinea"])
						strFiltro += " AND {AsientoDetalle.idSubProyecto} = " + 
							cmbSubProyecto.EditValue;
					else
						strFiltro += " AND {Asiento.idSubProyecto} = " + 
							cmbSubProyecto.EditValue;
					stTitulo += " Proy. " + cmbSubProyecto.Text.Trim();
				}
			#endregion Fin de Filtro de SubProyecto
			string stFiltroB = strFiltro;
			if (bBalance8)
				if ((int) rdgTipo.EditValue == 1)
					strFiltro = stFiltroCuenta;
				else
					strFiltro = stFiltroCuenta;
			#endregion Filtros
			#region Nombre del Reporte
			string strReporte = "";
			switch((int) rdgTipo.EditValue)
			{
				case 0:
					strReporte = "DiarioGeneral.rpt";
					break;
				case 1:
					strReporte = "Mayor.rpt";
					break;
				case 2:
					strReporte = "Balance8.rpt";
					break;
				case 3:
					Contabilidad.TipoCuenta(cdsCuenta, drSeteoC["CuentaActivo"].ToString(), 1);
					Contabilidad.TipoCuenta(cdsCuenta, drSeteoC["CuentaPasivo"].ToString(), 2);
					Contabilidad.TipoCuenta(cdsCuenta, drSeteoC["CuentaCapital"].ToString(), 3);
					CreaAsientoTemp();
					if (bBalance8)
					{
						strFiltro += " {TmpBalance.Tipo} < 4 And {TmpBalance.Tipo} > 0 ";
						strReporte = "Balance8.rpt";
					}
					else
					{
						strFiltro += " AND {Cuenta.Tipo} < 4 And {Cuenta.Tipo} > 0 ";
						strReporte = "Balance.rpt";
					}
					break;
				case 4:
					Contabilidad.TipoCuenta(cdsCuenta, drSeteoC["CuentaIngreso"].ToString(), 4);
					Contabilidad.TipoCuenta(cdsCuenta, drSeteoC["CuentaEgreso"].ToString(), 5);
					CreaAsientoTemp();
					if (bBalance8)
					{
						strFiltro += " {TmpBalance.Tipo} > 3 ";
						strReporte = "Balance8.rpt";
					}
					else
					{
						strFiltro += " AND {Cuenta.Tipo} > 3 ";
						strReporte = "Balance.rpt";
					}
					break;
				case 5:
					if (txtDesde.Text.Trim().Length > 0)
						strFiltro += " AND {Asiento.Numero} = '" + txtDesde.Text.Trim() + "' ";
					strReporte = "TipoAsiento.rpt";
					break;
			}
			#endregion Nombre del Reporte
			#region Reporte con Procedimientos Almacenados
			if (bBalance8)
			{
				stFiltroB = stFiltroB.Replace("{", "");
				stFiltroB = stFiltroB.Replace("}", "");
				stFiltroB = stFiltroB.Replace("'", "''");
				stFiltroB = stFiltroB.Replace("*", "%");
				string stExec ="";
				if ((int) rdgTipo.EditValue == 1)
				{
					stExec ="EXEC CreaMayor '" + dtFechaIni.ToString("yyyyMMdd")
						+ "', '" + dtFechaFin.AddDays(1.0).ToString("yyyyMMdd") 
						+ "', '" + stFiltroB +"'";
				}
				else
				{
					string stTotal = "0";
					if (cmbReporte.Text.ToUpper().StartsWith("BALANCEHOR")) stTotal = "1";
					string stFiltroExt = strFiltro.Replace("{", "");
					stFiltroExt = stFiltroExt.Replace("}", "");
					stFiltroExt = stFiltroExt.Replace("'", "''");
					stFiltroExt = stFiltroExt.Replace("*", "%");

					///     juan marcet
					string stFiltroExtN= "";   
					if (JuanM==1)
						stFiltroExtN= ",1";
					////////////juan marcet

					stExec ="EXEC CreaBalance '" + dtFechaIni.ToString("yyyyMMdd")
						+ "', '" + dtFechaFin.AddDays(1.0).ToString("yyyyMMdd") 
						+ "', '" + stFiltroB +"', " + stTotal + ", '" + stFiltroExt + "'" + stFiltroExtN ;
				}
				//stExec = string.Format("Exec GeneraBalance '20140701', '20140731'");
				Funcion.EjecutaSQL(cdsAsiento, stExec, true);
			}
			#endregion Reporte con Procedimientos Almacenados
			#region Definicion de Parametros
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();

			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "FechaIni";
			discreteVal.Value = dtFechaIni.ToString("yyyy-MM-dd");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			string stTipo = GeneraTipo();
			if (stTipo.Length > 0)
			{
				ParameterField paramField1 = new ParameterField ();
				ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
				paramField1.ParameterFieldName = "pcTipo";
				discreteVal1.Value = stTipo;
				paramField1.CurrentValues.Add (discreteVal1);
				paramFields.Add (paramField1);

				ParameterField paramField2 = new ParameterField ();
				ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
				paramField2.ParameterFieldName = "pnUtilidad";
				discreteVal2.Value = dUtilidad;
				paramField2.CurrentValues.Add (discreteVal2);
				paramFields.Add (paramField2);

				ParameterField paramField3 = new ParameterField ();
				ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
				paramField3.ParameterFieldName = "pnTipo";
				discreteVal3.Value = (int) rdgTipo.EditValue; // 3 General 4 Resultados
				paramField3.CurrentValues.Add (discreteVal3);
				paramFields.Add (paramField3);
			}
			#endregion Fin de parametros

			if (!bBalance8)
				strReporte = cmbReporte.Text.ToString().Trim() + ".rpt";
			else
        strReporte = cmbReporte.Text.ToString().Trim() + "R.rpt";

			string stFiltroRep = cmbReporte.ActiveRow.Cells["Filtro"].Value.ToString();
			if (stFiltroRep != null) strFiltro += stFiltroRep;

			miReporte = new Reporte(strReporte, strFiltro);
			miReporte.MdiParent = MdiParent;
			if ((int) rdgTipo.EditValue != 0 && (int) rdgTipo.EditValue != 5)
				miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Titulo(stTitulo);
			miReporte.Show();
			Cursor = Cursors.Default;
			BorraAsientoTemp();
			TimeSpan ts = DateTime.Now.Subtract(dtIni);
			btAceptar.ToolTip = "Tiempo de último reporte : " + ts.TotalSeconds.ToString("n2");
		}

		private bool Verifica()
		{
			bool bVerifica = true;
			BorraErrores();
			#region Seguridad
			int iPos = 148 + (int) rdgTipo.EditValue;
			if (!Funcion.Permiso(iPos.ToString(), cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a imprimir Reporte", 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return false;
			}
			#endregion Seguridad
			#region Reporte Activo
			if (cmbReporte.Value == null)
			{
				errorProvider.SetError(cmbReporte, "Ingrese Reporte");
				bVerifica = false;
			}
			#endregion Reporte Activo
			#region Rango de cuentas
			if (cmbCodDesde.Enabled)
				if (cmbCodDesde.EditValue == null)
				{
					errorProvider.SetError(cmbCodDesde, "Ingrese Código de Cuenta");
					bVerifica = false;
				}
			if (chkGrupoCuentas.Enabled && chkGrupoCuentas.Checked)
			{
				if (cmbCodHasta.EditValue == null)
				{
					errorProvider.SetError(cmbCodHasta, "Ingrese Cuenta Final");
					bVerifica = false;
				}
			}
			#endregion Rango de cuentas
			#region Tipo de Comprobante
			if (!chkComprobante.Checked)
			{
				if (cmbComprobante.EditValue == null)
				{
					errorProvider.SetError(cmbComprobante, "Ingrese Comprobante");
					bVerifica = false;
				}
			}
			#endregion Tipo de Comprobante
			#region Centro de Costo y Proyecto
			if (!chkProyecto.Checked)
			{
				if (cmbProyecto.EditValue == null)
				{
					errorProvider.SetError(cmbProyecto, "Ingrese Centro de Costo");
					bVerifica = false;
				}
			}
			if (!chkSubProyecto.Checked)
			{
				if (cmbSubProyecto.EditValue == null)
				{
					errorProvider.SetError(cmbSubProyecto, "Ingrese Proyecto");
					bVerifica = false;
				}
			}
			#endregion Centro de Costo y Proyecto
			#region Centro de Costo por Texto
			if (!chkCentroCosto.Checked)
			{
				if (txtCentroCosto.Value == null)
				{
					errorProvider.SetError(txtCentroCosto, "Ingrese Centro de Costo");
					bVerifica = false;
				}
			}
			#endregion Centro de Costo por Texto
			return bVerifica;
		}

		private void BorraErrores()
		{
			errorProvider.SetError(cmbComprobante, "");
			errorProvider.SetError(cmbCodDesde, "");
			errorProvider.SetError(cmbCodHasta, "");
			errorProvider.SetError(cmbProyecto, "");
			errorProvider.SetError(txtCentroCosto, "");
			errorProvider.SetError(cmbSubProyecto, "");
		}

		private void BorraAsientoTemp()
		{
			string stSelect = "Declare @idAsiento int;Select @idAsiento = a.idAsiento From Asiento a WHERE CodAsiento = '!@#$%123';Delete AsientoDetalle WHERE idAsiento = @idAsiento;Delete Asiento WHERE idAsiento = @idAsiento";
//			string stSelect = "Asiento a WHERE CodAsiento = '!@#$%123'";
			Funcion.EjecutaSQL(cdsAsiento, stSelect);
		}

		private string GeneraTipo()
		{
			string stTitulo = "";

			switch((int) rdgTipo.EditValue)
			{
				case 0:
				case 1:
					break;
				case 2:
					stTitulo = "Balance de Comprobación\n";
					break;
				case 3:
					stTitulo = "Balance General\n";
					break;
				case 4:
					stTitulo = "Balance de Resultados\n";
					break;
			}

			return stTitulo;
		}

		private void CreaAsientoTemp()
		{
			// Utilidad NIIF es 0
      string stNiif = "Select Count(*) From SeteoG Where Codigo = 'NGAT' And Activado = 1";
			int idTipoAsi = 0;
			int iAnioNiif = (int) Funcion.dEscalarSQL(cdsCuenta, stNiif);
			if (iAnioNiif > 0)
			{
				stNiif = "Select IsNull(Valor, 0) From SeteoG Where Codigo = 'NGAT' And Activado = 1";
				iAnioNiif = (int) Funcion.dEscalarSQL(cdsCuenta, stNiif);
				stNiif = "Select at.idTipoAsiento From AsientoTipo at Where at.Codigo = 'AJN'";
				idTipoAsi = (int) Funcion.dEscalarSQL(cdsCuenta, stNiif);
				////juan Marcet
						 /// idtipoasiento =0  para que saque balance niif o sea todo
						int JuanCod=0;
						stNiif = "Select Count(*) From SeteoG Where Codigo = 'NGATN' And Activado = 1";
						int JuanC = (int) Funcion.dEscalarSQL(cdsCuenta, stNiif);
						if (JuanC ==1)
						{
							stNiif = "Select isnull(valor,0) From SeteoG Where Codigo = 'NGATN' And Activado = 1";
					 	  JuanCod = (int) Funcion.dEscalarSQL(cdsCuenta, stNiif);
						}
							if (JuanCod==1)
									idTipoAsi=0;
						
									 ////juan Marcet

								 }

			// Calculo de Utilidad
			int IdCuentaU = (int) drSeteoC["CuentaUtilidad"];
			if (IdCuentaU < 1) 
			{
				if (DialogResult.No == MessageBox.Show(
					"No existe cuenta de utilidad desea crearla ahora",
					"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
					return;
				using (SeteoC miSeteo = new SeteoC())
				{
					miSeteo.ShowDialog();
				}
				return;
			}

			DateTime dtFechaFin = (DateTime) dtFechaHasta.EditValue;
			string stSelect = "Select sum(isnull(debe,0)-isnull(haber,0)) "
				+ "FROM Asiento INNER JOIN AsientoDetalle "
				+ "ON Asiento.idAsiento=AsientoDetalle.idAsiento " 
				+ "INNER JOIN Cuenta ON AsientoDetalle.idCuenta=Cuenta.idCuenta ";

			if ((bool) drSeteoC["ProyectoLinea"]) stSelect +=
				"Left Outer Join Proyecto On AsientoDetalle.idProyecto = Proyecto.idProyecto ";
			else stSelect +=
				"Left Outer Join Proyecto On Asiento.idProyecto = Proyecto.idProyecto ";

			if ((bool) drSeteoC["SubProyLinea"]) stSelect +=
				"Left Outer Join SubProyecto On AsientoDetalle.idSubProyecto = SubProyecto.idSubProyecto ";
			else stSelect +=
				"Left Outer Join SubProyecto On Asiento.idSubProyecto = SubProyecto.idSubProyecto ";

			stSelect += "where AsientoDetalle.Borrar = 0 And Asiento.Borrar = 0 "
				+ "and Asiento.fecha < '" + dtFechaFin.AddDays(1).ToString("yyyyMMdd")
				+ "' and (Cuenta.Tipo = 4 Or Cuenta.Tipo = 5) And Asiento.idTipoAsiento <> " + idTipoAsi;

			string stAdiciona = "";
			string stAdicionaLinea = "";
			string stValorAdi = "";
			string stValorAdiLinea = "";
			if (!chkProyecto.Checked)
			{
				if ((bool) drSeteoC["ProyectoLinea"])
				{
					stSelect += " AND AsientoDetalle.idProyecto = " + cmbProyecto.EditValue;
					stAdicionaLinea += ", idProyecto";
					stValorAdiLinea += ", " + cmbProyecto.EditValue.ToString();
				}
				else
				{
					stSelect += " AND Asiento.idProyecto = " + cmbProyecto.EditValue;
					stAdiciona += ", idProyecto";
					stValorAdi += ", " + cmbProyecto.EditValue.ToString();
				}
			}
				
			if (!chkSubProyecto.Checked)
			{
				if ((bool) drSeteoC["SubProyLinea"])
				{
					stSelect += " AND AsientoDetalle.idSubProyecto = " + cmbSubProyecto.EditValue;
					stAdicionaLinea += ", idSubProyecto";
					stValorAdiLinea += ", " + cmbSubProyecto.EditValue.ToString();
				}
				else
				{
					stSelect += " AND Asiento.idSubProyecto = " + cmbSubProyecto.EditValue;
					stAdiciona += ", idSubProyecto";
					stValorAdi += ", " + cmbSubProyecto.EditValue.ToString();
				}
			}

			// Cuando los proyectos son por linea debe variar el reporte
			if (!chkCentroCosto.Checked)
			{
				string stSel = "Select idProyecto From Proyecto "
					+ "Where Proyecto.CentroCosto Like '" + 
					txtCentroCosto.Text.Trim() + "%'";
				int IdProy = Funcion.iEscalarSQL(cdsProyecto, stSel);
				if ((bool) drSeteoC["ProyectoLinea"])
				{
					stSelect += " AND Proyecto.CentroCosto Like '" + 
						txtCentroCosto.Value.ToString() + "%'";
					stAdicionaLinea += ", idProyecto";
					stValorAdiLinea += ", " + IdProy.ToString();
				}
				else
				{
					stSelect += " AND Proyecto.CentroCosto Like '" + 
						txtCentroCosto.Text.Trim() + "%'";
					stAdiciona += ", idProyecto";
					stValorAdi += ", " + IdProy.ToString();
				}
			}

			// Genera utilidad mensual
			if (cmbReporte.Text.Trim().ToUpper().EndsWith("MES"))
			{
				DateTime dtFechaInicio = (DateTime) dtFechaDesde.EditValue;
				stSelect += " And asiento.fecha >= '" + dtFechaInicio.ToString("yyyyMMdd") +"' ";
			}
			if (iAnioNiif > 2000 && iAnioNiif < 2100) 
			{
				/// juan marcet 
				stNiif = "Select Count(*) From SeteoG Where Codigo = 'NGATN' And Activado = 1";
				int Juan = (int) Funcion.dEscalarSQL(cdsCuenta, stNiif);
	      if (Juan==1)
						iAnioNiif=iAnioNiif-1;
        ////Juan Marcet
				stSelect += string.Format(" And Asiento.fecha >= '{0}0101' ", iAnioNiif);
			}

			dUtilidad = Funcion.dEscalarSQL(cdsAsiento, stSelect, false);

			// Creacion de Asiento Temporal
			stSelect = "INSERT INTO Asiento (CodAsiento, Fecha" +  
				stAdiciona + ") Values('!@#$%123', '"
				+ dtFechaFin.ToString("yyyyMMdd") + "'" + stValorAdi + ")";
			Funcion.EjecutaSQL(cdsAsiento, stSelect);

			stSelect = "SELECT idAsiento FROM Asiento WHERE CodAsiento = '!@#$%123'";
			int IdAsi = Funcion.iEscalarSQL(cdsAsiento, stSelect);
			string stLado =" Debe";
			double dUtilTemp = dUtilidad;
			if (dUtilidad < 0) 
			{
				stLado = " Haber";
				dUtilTemp = -dUtilidad;
			}
			dUtilidad = -dUtilidad;
			stLado += stAdicionaLinea;
			stSelect = "Insert into AsientoDetalle (idAsiento, idCuenta, " + stLado
				+ ") Values (" + IdAsi.ToString() + ", " 
				+ IdCuentaU.ToString() + ", " + dUtilTemp.ToString("0.00", us)
				+ stValorAdiLinea + ")";
			Funcion.EjecutaSQL(cdsAsiento, stSelect);
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}


		#endregion Botones Principales
		#region cds
		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoC.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsAsiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsAsiento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsProyecto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsMoneda_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsMoneda.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsReporte_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
//			Funcion miFuncion = new Funcion();
//			cdsReporte.Schema.Connections[0].ConnectionString = 
//				miFuncion.AccesoBase(MenuLatinium.stDirInicio + "\\Empresa.Mdb", 1);
		}


		#endregion cds

		#region Filtro por tipo de reporte y habilitaciones
		private void rdgTipo_SelectedIndexChanged(object sender, System.EventArgs e)
		{
      int IdGrupo = (int) rdgTipo.EditValue + 1;
			if (IdGrupo == 4 || IdGrupo == 5) IdGrupo = 3;
			
			string sSQL = string.Format("Select idReporte, Reporte, Filtro From Reportes Where idTipo = 5 And idGrupo = {0}", IdGrupo);
			this.cmbReporte.DataSource = Funcion.dvProcedimiento(cdsSeteoC, sSQL);

			if (IdGrupo == 1) this.cmbReporte.Value = 49;
			if (IdGrupo == 2) this.cmbReporte.Value = 50;
			if (IdGrupo == 3 || IdGrupo == 4 || IdGrupo == 5) this.cmbReporte.Value = 51;
			if (IdGrupo == 6) this.cmbReporte.Value = 52;

//			string stFiltro = "Tipo = 4 And idGrupo = " + IdGrupo.ToString();
//			cdvReporte.RowFilter = stFiltro;
//			if (IdGrupo == 4 || IdGrupo == 5)
//			{
//				stFiltro = "Tipo = 4 And idGrupo = 3";
//				cdvReporte.RowFilter = stFiltro;
//				cmbReporte.Value = 92;
//			}
//			else
//			{
//				DataRow[] drRep = cdsReporte.StorageDataSet.Tables["Reporte"].Select(stFiltro);
//				if (drRep.Length > 0)
//					cmbReporte.Value = drRep[0]["idReporte"];		
//			}

			switch (IdGrupo-1)
			{
				case 0: //Diario General
				case 5: //Resumen de asientos
					chkCuenta.Enabled = false;
					chkComprobante.Enabled = true;
					gbxBalance.Enabled = false;
					break;
				case 1: // Mayor General
					chkCuenta.Enabled = true;
					chkGrupoCuentas.Checked = false;
					chkCuenta.Checked = false;
					cmbCodDesde.Enabled = true;
					gbxBalance.Enabled = false;
					chkComprobante.Enabled = true;
					break;
				case 2: // Balance de Comprobacion
				case 3: // Balance de General
				case 4: // Balance de Resultados
					chkCuenta.Enabled = true;
					chkCuenta.Checked = true;
					gbxBalance.Enabled = true;
					chkComprobante.Enabled = false;
					break;
			}
		}

		private void chkCuenta_CheckStateChanged(object sender, System.EventArgs e)
		{
			bool bEstado = chkCuenta.Checked;
			cmbCodDesde.Enabled = !bEstado;
			chkGrupoCuentas.Enabled = !bEstado;
			if (bEstado)
			{
				chkGrupoCuentas.Checked = false;
				cmbCodHasta.Enabled = false;
			}
			else
			{
				cmbCodDesde.Select();
			}
		}

		private void chkCuenta_EnabledChanged(object sender, System.EventArgs e)
		{
			if (chkCuenta.Enabled == false)
			{
				chkGrupoCuentas.Enabled = false;
				cmbCodDesde.Enabled = false;
				cmbCodHasta.Enabled = false;
			}
		}

		private void cmbCodDesde_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (MouseButtons.Right == e.Button)
			{
				if (cmbCodDesde.Properties.DisplayMember == "CodRapido")
				{
					cmbCodDesde.Properties.Columns.Clear();
					cmbCodDesde.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																								new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																								new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 66, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
					cmbCodDesde.Properties.DataSource = cdvCuentaNombre;
					cmbCodDesde.Properties.DisplayMember = "Descripcion";
				}
				else
				{
					cmbCodDesde.Properties.Columns.Clear();
					cmbCodDesde.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																								new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 66, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																											new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
					cmbCodDesde.Properties.DataSource = cdvCuenta;
					cmbCodDesde.Properties.DisplayMember = "CodRapido";
				}
			}
		}

		private void cmbCodHasta_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (MouseButtons.Right == e.Button)
			{
				if (cmbCodHasta.Properties.DisplayMember == "CodRapido")
				{
					cmbCodHasta.Properties.Columns.Clear();
					cmbCodHasta.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																								new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																								new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 66, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
					cmbCodDesde.Properties.DataSource = cdvCuentaNombre;
					cmbCodHasta.Properties.DisplayMember = "Descripcion";
				}
				else
				{
					cmbCodHasta.Properties.Columns.Clear();
					cmbCodHasta.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																								new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 66, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																								new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
					cmbCodDesde.Properties.DataSource = cdvCuenta;
					cmbCodHasta.Properties.DisplayMember = "CodRapido";
				}
			}		
		}

		private void chkGrupoCuentas_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCodHasta.Enabled = chkGrupoCuentas.Checked;		
			if (chkGrupoCuentas.Checked) cmbCodHasta.Select();
		}

		private void chkComprobante_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbComprobante.Enabled = !chkComprobante.Checked;
			cmbComprobante.Focus();
		}

		#endregion Filtro por tipo de reporte y habilitaciones

		#region Habilitaciones Independintes
		private void chkProyecto_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!chkProyecto.Checked) chkCentroCosto.Checked = true;
			cmbProyecto.Enabled = !chkProyecto.Checked;
			cmbProyecto.Focus();
		}

		private void chkSubProyecto_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!chkSubProyecto.Checked) chkCentroCosto.Checked = true;
			cmbSubProyecto.Enabled = !chkSubProyecto.Checked;
			cmbSubProyecto.Focus();
		}

		private void chkCentroCosto_CheckedChanged(object sender, System.EventArgs e)
		{
			txtCentroCosto.Enabled = !chkCentroCosto.Checked;
			if (!chkCentroCosto.Checked)
			{
				chkProyecto.Checked = true;
				chkSubProyecto.Checked = true;
				txtCentroCosto.Focus();
			}
		}

		private void chkSubCentro_CheckedChanged(object sender, System.EventArgs e)
		{
			txtProyecto.Enabled = !chkSubCentro.Checked;
			if (!chkCentroCosto.Checked)
			{
				chkProyecto.Checked = true;
				chkSubProyecto.Checked = true;
				txtProyecto.Focus();
			}
		}

		private void chkProyecto_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				chkProyecto.CheckState = System.Windows.Forms.CheckState.Indeterminate;
		}

		private void cmbFechas_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			int iTipo = cmbFechas.SelectedIndex;
			DateTime dtHoy = DateTime.Today;
			int iDia = dtHoy.Day;
			DateTime dtMes1 = dtHoy.AddDays(-iDia+1);
			C1.Data.C1DataRow drSeteoC = cdsSeteoC.TableViews["SeteoC"].Rows[0];
			switch(iTipo)
			{
				case 0: // Hasta Hoy
					//drSeteoC["FechaRIni"] = DateTime.Today;
					drSeteoC["FechaRFin"] = DateTime.Today;
					break;
				case 1: // Solo Hoy
					drSeteoC["FechaRIni"] = DateTime.Today;
					drSeteoC["FechaRFin"] = DateTime.Today;
					break;
				case 2: // Este Mes
					drSeteoC["FechaRIni"] = dtMes1;
					dtMes1 = dtMes1.AddMonths(1);
					drSeteoC["FechaRFin"] = dtMes1.AddDays(-1); 
					break;
				case 3: // El mes pasado
					dtMes1 = dtMes1.AddMonths(-1);
					drSeteoC["FechaRIni"] = dtMes1;
					dtMes1 = dtMes1.AddMonths(1);
					drSeteoC["FechaRFin"] = dtMes1.AddDays(-1); 
					break;
				case 4: // Este Año
					drSeteoC["FechaRIni"] = DateTime.Parse("01/01/" 
						+ DateTime.Today.Year.ToString(), us);
					drSeteoC["FechaRFin"] = DateTime.Parse("12/31/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 5: // El año pasado
					drSeteoC["FechaRIni"] = DateTime.Parse("01/01/" 
						+ DateTime.Today.AddYears(-1).Year.ToString(), us);
					drSeteoC["FechaRFin"] = DateTime.Parse("12/31/" 
						+ DateTime.Today.AddYears(-1).Year.ToString(), us);
					break;
				case 6: // El Periodo Ingresado
					drSeteoC["FechaRIni"] = Contabilidad.FechaAsiento(cdsAsiento, 0);
					drSeteoC["FechaRFin"] = Contabilidad.FechaAsiento(cdsAsiento, 1);
					break;
				case 7: // El Primer Trimestre
					drSeteoC["FechaRIni"] = DateTime.Parse("01/01/" 
						+ DateTime.Today.Year.ToString(), us);
					drSeteoC["FechaRFin"] = DateTime.Parse("03/31/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 8: // El segundo Trimestre
					drSeteoC["FechaRIni"] = DateTime.Parse("04/01/" 
						+ DateTime.Today.Year.ToString(), us);
					drSeteoC["FechaRFin"] = DateTime.Parse("06/30/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 9: // El tercer Trimestre
					drSeteoC["FechaRIni"] = DateTime.Parse("07/01/" 
						+ DateTime.Today.Year.ToString(), us);
					drSeteoC["FechaRFin"] = DateTime.Parse("09/30/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 10: // El cuarto Trimestre
					drSeteoC["FechaRIni"] = DateTime.Parse("10/01/" 
						+ DateTime.Today.Year.ToString(), us);
					drSeteoC["FechaRFin"] = DateTime.Parse("12/31/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 11: // Periodo Digitado
					break;
			}
			cdsSeteoC.Update();
		}

		private void chkUsuario_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbUsuario.Enabled = !chkUsuario.Checked;
		}

		#endregion Habilitaciones Independintes

		private void btAceptar_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			int Juan=0;
			string stSelect = "Select count(*) From SeteoG Where Codigo = 'NGATN' And Activado = 1";
			Juan = (int) Funcion.dEscalarSQL(cdsCuenta, stSelect  );
			if (Juan==1) //esta activado
			{
				stSelect="update seteog set valor=1 where codigo='ngatn'";
				Funcion.EjecutaSQL(cdsAsiento,stSelect);   
				JuanM =1;
			}			
			////juan Marcet		
			btAceptar_Click(this,e);

			if (Juan==1) //esta activado
			{
				stSelect="update seteog set valor=0 where codigo='ngatn'";
				Funcion.EjecutaSQL(cdsAsiento,stSelect);  
				JuanM =0;
			}
		}

		private void cmbCodDesde_EditValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
