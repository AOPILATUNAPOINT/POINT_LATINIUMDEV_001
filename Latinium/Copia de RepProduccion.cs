using System;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RepProduccion.
	/// </summary>
	public class RepProduccion : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.RadioGroup rdgTipo;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.ComboBoxEdit cmbFechas;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private DevExpress.XtraEditors.DateEdit dtFechaHasta;
		private DevExpress.XtraEditors.DateEdit dtFechaDesde;
		private DevExpress.XtraEditors.SpinEdit spnBodega;
		private DevExpress.XtraEditors.CheckEdit chkBodega;
		private DevExpress.XtraEditors.CheckEdit chkArticulo;
		private DevExpress.XtraEditors.LookUpEdit cmbArticulo;
		private DevExpress.XtraEditors.CheckEdit chkSeccion;
		private DevExpress.XtraEditors.CheckEdit chkClase;
		private DevExpress.XtraEditors.CheckEdit chkMarca;
		private DevExpress.XtraEditors.CheckEdit chkSubGrupo;
		private DevExpress.XtraEditors.CheckEdit chkGrupo;
		private DevExpress.XtraEditors.LookUpEdit cmbSeccion;
		private DevExpress.XtraEditors.LookUpEdit cmbClase;
		private DevExpress.XtraEditors.LookUpEdit cmbMarca;
		private DevExpress.XtraEditors.LookUpEdit cmbSubGrupo;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataView cdvClase;
		private C1.Data.C1DataSet cdsClase;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataSet cdsSubGrupo;
		private C1.Data.C1DataSet cdsMarca;
		private C1.Data.C1DataView cdvMarca;
		private C1.Data.C1DataSet cdsGrupo;
		private DevExpress.XtraEditors.LookUpEdit cmbGrupo;
		private C1.Data.C1DataView cdvGrupo;
		private C1.Data.C1DataView cdvSubGrupo;
		private C1.Data.C1DataSet cdsSeccion;
		private C1.Data.C1DataView cdvSeccion;
		private DevExpress.XtraEditors.CheckEdit chkTipoGrupo;
		private DevExpress.XtraEditors.LookUpEdit cmbTipoGrupo;
		private C1.Data.C1DataSet cdsTipoGrupo;
		private C1.Data.C1DataView cdvTipoGrupo;
		private C1.Data.C1DataSet cdsTransaccion;
		private C1.Data.C1DataView cdvTransaccion;
		private DevExpress.XtraEditors.CheckEdit chkTransaccion;
		private DevExpress.XtraEditors.LookUpEdit cmbTransaccion;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RepProduccion()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.rdgTipo = new DevExpress.XtraEditors.RadioGroup();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.cmbFechas = new DevExpress.XtraEditors.ComboBoxEdit();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.dtFechaHasta = new DevExpress.XtraEditors.DateEdit();
			this.dtFechaDesde = new DevExpress.XtraEditors.DateEdit();
			this.spnBodega = new DevExpress.XtraEditors.SpinEdit();
			this.chkBodega = new DevExpress.XtraEditors.CheckEdit();
			this.chkArticulo = new DevExpress.XtraEditors.CheckEdit();
			this.cmbArticulo = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.chkSeccion = new DevExpress.XtraEditors.CheckEdit();
			this.chkClase = new DevExpress.XtraEditors.CheckEdit();
			this.chkMarca = new DevExpress.XtraEditors.CheckEdit();
			this.chkSubGrupo = new DevExpress.XtraEditors.CheckEdit();
			this.chkGrupo = new DevExpress.XtraEditors.CheckEdit();
			this.cmbSeccion = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvSeccion = new C1.Data.C1DataView();
			this.cdsSeccion = new C1.Data.C1DataSet();
			this.cmbClase = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvClase = new C1.Data.C1DataView();
			this.cdsClase = new C1.Data.C1DataSet();
			this.cmbMarca = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvMarca = new C1.Data.C1DataView();
			this.cdsMarca = new C1.Data.C1DataSet();
			this.cmbSubGrupo = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvSubGrupo = new C1.Data.C1DataView();
			this.cdsSubGrupo = new C1.Data.C1DataSet();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsGrupo = new C1.Data.C1DataSet();
			this.cmbGrupo = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvGrupo = new C1.Data.C1DataView();
			this.chkTipoGrupo = new DevExpress.XtraEditors.CheckEdit();
			this.cmbTipoGrupo = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvTipoGrupo = new C1.Data.C1DataView();
			this.cdsTipoGrupo = new C1.Data.C1DataSet();
			this.chkTransaccion = new DevExpress.XtraEditors.CheckEdit();
			this.cmbTransaccion = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvTransaccion = new C1.Data.C1DataView();
			this.cdsTransaccion = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechas.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkBodega.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkArticulo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSeccion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkClase.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkMarca.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSeccion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSeccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbClase.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvClase)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClase)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTipoGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTransaccion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransaccion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTransaccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTransaccion)).BeginInit();
			this.SuspendLayout();
			// 
			// rdgTipo
			// 
			this.rdgTipo.EditValue = 0;
			this.rdgTipo.Location = new System.Drawing.Point(24, 16);
			this.rdgTipo.Name = "rdgTipo";
			// 
			// rdgTipo.Properties
			// 
			this.rdgTipo.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Días de Inventario"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Programa de Produccion"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Producción Explosionado"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Reportes por Bodega"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(4, "Demanda Insatisfecha"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(5, "Presupuesto de Ventas"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(6, "Tiempo de Producción"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(7, "Producción Semanal")});
			this.rdgTipo.Size = new System.Drawing.Size(160, 200);
			this.rdgTipo.TabIndex = 0;
			// 
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(544, 312);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 1;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(408, 312);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 2;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// cmbFechas
			// 
			this.cmbFechas.EditValue = "Hasta Hoy";
			this.cmbFechas.Location = new System.Drawing.Point(24, 256);
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
																															"Periodo Digitado",
																															"Periodo Ingresado"});
			this.cmbFechas.Size = new System.Drawing.Size(160, 20);
			this.cmbFechas.TabIndex = 57;
			this.cmbFechas.SelectedIndexChanged += new System.EventHandler(this.cmbFechas_SelectedIndexChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(24, 312);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 56;
			this.label6.Text = "Hasta:";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(24, 288);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(40, 16);
			this.label7.TabIndex = 55;
			this.label7.Text = "Desde:";
			// 
			// dtFechaHasta
			// 
			this.dtFechaHasta.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaHasta.Location = new System.Drawing.Point(88, 312);
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
			this.dtFechaHasta.TabIndex = 54;
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaDesde.Location = new System.Drawing.Point(88, 288);
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
			this.dtFechaDesde.TabIndex = 53;
			// 
			// spnBodega
			// 
			this.spnBodega.EditValue = new System.Decimal(new int[] {
																																0,
																																0,
																																0,
																																0});
			this.spnBodega.Location = new System.Drawing.Point(392, 112);
			this.spnBodega.Name = "spnBodega";
			// 
			// spnBodega.Properties
			// 
			this.spnBodega.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnBodega.Properties.IsFloatValue = false;
			this.spnBodega.Properties.Mask.EditMask = "N00";
			this.spnBodega.Size = new System.Drawing.Size(48, 20);
			this.spnBodega.TabIndex = 76;
			// 
			// chkBodega
			// 
			this.chkBodega.EditValue = true;
			this.chkBodega.Location = new System.Drawing.Point(248, 112);
			this.chkBodega.Name = "chkBodega";
			// 
			// chkBodega.Properties
			// 
			this.chkBodega.Properties.Caption = "Todas las Bodegas";
			this.chkBodega.Size = new System.Drawing.Size(136, 19);
			this.chkBodega.TabIndex = 75;
			this.chkBodega.CheckedChanged += new System.EventHandler(this.chkBodega_CheckedChanged);
			// 
			// chkArticulo
			// 
			this.chkArticulo.EditValue = true;
			this.chkArticulo.Location = new System.Drawing.Point(248, 80);
			this.chkArticulo.Name = "chkArticulo";
			// 
			// chkArticulo.Properties
			// 
			this.chkArticulo.Properties.Caption = "Todos los Artículos";
			this.chkArticulo.Size = new System.Drawing.Size(136, 19);
			this.chkArticulo.TabIndex = 74;
			this.chkArticulo.CheckedChanged += new System.EventHandler(this.chkArticulo_CheckedChanged);
			// 
			// cmbArticulo
			// 
			this.cmbArticulo.Location = new System.Drawing.Point(392, 80);
			this.cmbArticulo.Name = "cmbArticulo";
			// 
			// cmbArticulo.Properties
			// 
			this.cmbArticulo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbArticulo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Articulo", "Articulo", 61, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None)});
			this.cmbArticulo.Properties.DataSource = this.cdvArticulo;
			this.cmbArticulo.Properties.DisplayMember = "Articulo";
			this.cmbArticulo.Properties.NullText = "";
			this.cmbArticulo.Properties.ValueMember = "idArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(192, 20);
			this.cmbArticulo.TabIndex = 73;
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulo";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// chkSeccion
			// 
			this.chkSeccion.EditValue = true;
			this.chkSeccion.Location = new System.Drawing.Point(248, 208);
			this.chkSeccion.Name = "chkSeccion";
			// 
			// chkSeccion.Properties
			// 
			this.chkSeccion.Properties.Caption = "Todas las Secciones";
			this.chkSeccion.Size = new System.Drawing.Size(136, 19);
			this.chkSeccion.TabIndex = 72;
			this.chkSeccion.CheckedChanged += new System.EventHandler(this.chkSeccion_CheckedChanged);
			// 
			// chkClase
			// 
			this.chkClase.EditValue = true;
			this.chkClase.Location = new System.Drawing.Point(248, 176);
			this.chkClase.Name = "chkClase";
			// 
			// chkClase.Properties
			// 
			this.chkClase.Properties.Caption = "Todas las Clases";
			this.chkClase.Size = new System.Drawing.Size(136, 19);
			this.chkClase.TabIndex = 71;
			this.chkClase.CheckedChanged += new System.EventHandler(this.chkClase_CheckedChanged);
			// 
			// chkMarca
			// 
			this.chkMarca.EditValue = true;
			this.chkMarca.Location = new System.Drawing.Point(248, 144);
			this.chkMarca.Name = "chkMarca";
			// 
			// chkMarca.Properties
			// 
			this.chkMarca.Properties.Caption = "Todas los Marcas";
			this.chkMarca.Size = new System.Drawing.Size(136, 19);
			this.chkMarca.TabIndex = 70;
			this.chkMarca.CheckedChanged += new System.EventHandler(this.chkMarca_CheckedChanged);
			// 
			// chkSubGrupo
			// 
			this.chkSubGrupo.EditValue = true;
			this.chkSubGrupo.Location = new System.Drawing.Point(248, 48);
			this.chkSubGrupo.Name = "chkSubGrupo";
			// 
			// chkSubGrupo.Properties
			// 
			this.chkSubGrupo.Properties.Caption = "Todos los SubGrupos";
			this.chkSubGrupo.Size = new System.Drawing.Size(136, 19);
			this.chkSubGrupo.TabIndex = 69;
			this.chkSubGrupo.CheckedChanged += new System.EventHandler(this.chkSubGrupo_CheckedChanged);
			// 
			// chkGrupo
			// 
			this.chkGrupo.EditValue = true;
			this.chkGrupo.Location = new System.Drawing.Point(248, 16);
			this.chkGrupo.Name = "chkGrupo";
			// 
			// chkGrupo.Properties
			// 
			this.chkGrupo.Properties.Caption = "Todos los Grupos";
			this.chkGrupo.Size = new System.Drawing.Size(120, 19);
			this.chkGrupo.TabIndex = 68;
			this.chkGrupo.CheckedChanged += new System.EventHandler(this.chkGrupo_CheckedChanged);
			// 
			// cmbSeccion
			// 
			this.cmbSeccion.Location = new System.Drawing.Point(392, 208);
			this.cmbSeccion.Name = "cmbSeccion";
			// 
			// cmbSeccion.Properties
			// 
			this.cmbSeccion.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbSeccion.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Seccion", "Seccion", 50, DevExpress.Utils.FormatType.Numeric, "", true, DevExpress.Utils.HorzAlignment.Far, DevExpress.Data.ColumnSortOrder.None)});
			this.cmbSeccion.Properties.DataSource = this.cdvSeccion;
			this.cmbSeccion.Properties.DisplayMember = "Seccion";
			this.cmbSeccion.Properties.NullText = "";
			this.cmbSeccion.Properties.ValueMember = "idSeccion";
			this.cmbSeccion.Size = new System.Drawing.Size(192, 20);
			this.cmbSeccion.TabIndex = 67;
			// 
			// cdvSeccion
			// 
			this.cdvSeccion.BindingContextCtrl = this;
			this.cdvSeccion.DataSet = this.cdsSeccion;
			this.cdvSeccion.TableName = "";
			this.cdvSeccion.TableViewName = "ArticuloSeccion";
			// 
			// cdsSeccion
			// 
			this.cdsSeccion.BindingContextCtrl = this;
			this.cdsSeccion.DataLibrary = "LibFacturacion";
			this.cdsSeccion.DataLibraryUrl = "";
			this.cdsSeccion.DataSetDef = "dsArticuloTabla";
			this.cdsSeccion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeccion.Name = "cdsSeccion";
			this.cdsSeccion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeccion.SchemaDef = null;
			this.cdsSeccion.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeccion_BeforeFill);
			// 
			// cmbClase
			// 
			this.cmbClase.Location = new System.Drawing.Point(392, 176);
			this.cmbClase.Name = "cmbClase";
			// 
			// cmbClase.Properties
			// 
			this.cmbClase.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbClase.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Clase", "Clase", 51, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbClase.Properties.DataSource = this.cdvClase;
			this.cmbClase.Properties.DisplayMember = "Clase";
			this.cmbClase.Properties.NullText = "";
			this.cmbClase.Properties.ValueMember = "idClase";
			this.cmbClase.Size = new System.Drawing.Size(192, 20);
			this.cmbClase.TabIndex = 66;
			// 
			// cdvClase
			// 
			this.cdvClase.BindingContextCtrl = this;
			this.cdvClase.DataSet = this.cdsClase;
			this.cdvClase.TableName = "";
			this.cdvClase.TableViewName = "ArticuloClase";
			// 
			// cdsClase
			// 
			this.cdsClase.BindingContextCtrl = this;
			this.cdsClase.DataLibrary = "LibFacturacion";
			this.cdsClase.DataLibraryUrl = "";
			this.cdsClase.DataSetDef = "dsArticuloTabla";
			this.cdsClase.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsClase.Name = "cdsClase";
			this.cdsClase.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsClase.SchemaDef = null;
			this.cdsClase.BeforeFill += new C1.Data.FillEventHandler(this.cdsClase_BeforeFill);
			// 
			// cmbMarca
			// 
			this.cmbMarca.Location = new System.Drawing.Point(392, 144);
			this.cmbMarca.Name = "cmbMarca";
			// 
			// cmbMarca.Properties
			// 
			this.cmbMarca.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbMarca.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Marca", "Marca", 41, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbMarca.Properties.DataSource = this.cdvMarca;
			this.cmbMarca.Properties.DisplayMember = "Marca";
			this.cmbMarca.Properties.NullText = "";
			this.cmbMarca.Properties.ValueMember = "idMarca";
			this.cmbMarca.Size = new System.Drawing.Size(192, 20);
			this.cmbMarca.TabIndex = 65;
			// 
			// cdvMarca
			// 
			this.cdvMarca.BindingContextCtrl = this;
			this.cdvMarca.DataSet = this.cdsMarca;
			this.cdvMarca.TableName = "";
			this.cdvMarca.TableViewName = "ArticuloMarca";
			// 
			// cdsMarca
			// 
			this.cdsMarca.BindingContextCtrl = this;
			this.cdsMarca.DataLibrary = "LibFacturacion";
			this.cdsMarca.DataLibraryUrl = "";
			this.cdsMarca.DataSetDef = "dsArticuloTabla";
			this.cdsMarca.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsMarca.Name = "cdsMarca";
			this.cdsMarca.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMarca.SchemaDef = null;
			this.cdsMarca.BeforeFill += new C1.Data.FillEventHandler(this.cdsMarca_BeforeFill);
			// 
			// cmbSubGrupo
			// 
			this.cmbSubGrupo.Location = new System.Drawing.Point(392, 48);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			// 
			// cmbSubGrupo.Properties
			// 
			this.cmbSubGrupo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbSubGrupo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("SubGrupo", "SubGrupo", 61, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbSubGrupo.Properties.DataSource = this.cdvSubGrupo;
			this.cmbSubGrupo.Properties.DisplayMember = "SubGrupo";
			this.cmbSubGrupo.Properties.NullText = "";
			this.cmbSubGrupo.Properties.ValueMember = "idSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(192, 20);
			this.cmbSubGrupo.TabIndex = 64;
			// 
			// cdvSubGrupo
			// 
			this.cdvSubGrupo.BindingContextCtrl = this;
			this.cdvSubGrupo.DataSet = this.cdsSubGrupo;
			this.cdvSubGrupo.TableName = "";
			this.cdvSubGrupo.TableViewName = "ArticuloSubGrupo";
			// 
			// cdsSubGrupo
			// 
			this.cdsSubGrupo.BindingContextCtrl = this;
			this.cdsSubGrupo.DataLibrary = "LibFacturacion";
			this.cdsSubGrupo.DataLibraryUrl = "";
			this.cdsSubGrupo.DataSetDef = "dsArticuloTabla";
			this.cdsSubGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSubGrupo.Name = "cdsSubGrupo";
			this.cdsSubGrupo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSubGrupo.SchemaDef = null;
			this.cdsSubGrupo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSubGrupo_BeforeFill);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// cdsGrupo
			// 
			this.cdsGrupo.BindingContextCtrl = this;
			this.cdsGrupo.DataLibrary = "LibFacturacion";
			this.cdsGrupo.DataLibraryUrl = "";
			this.cdsGrupo.DataSetDef = "dsArticuloTabla";
			this.cdsGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupo.Name = "cdsGrupo";
			this.cdsGrupo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupo.SchemaDef = null;
			this.cdsGrupo.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupo_BeforeFill);
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsArticulo, "Articulo.idGrupoArticulo"));
			this.cmbGrupo.Location = new System.Drawing.Point(392, 16);
			this.cmbGrupo.Name = "cmbGrupo";
			// 
			// cmbGrupo.Properties
			// 
			this.cmbGrupo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbGrupo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Grupo", "Grupo", 41, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbGrupo.Properties.DataSource = this.cdvGrupo;
			this.cmbGrupo.Properties.DisplayMember = "Grupo";
			this.cmbGrupo.Properties.NullText = "";
			this.cmbGrupo.Properties.ValueMember = "idGrupoArticulo";
			this.cmbGrupo.Size = new System.Drawing.Size(192, 20);
			this.cmbGrupo.TabIndex = 77;
			// 
			// cdvGrupo
			// 
			this.cdvGrupo.BindingContextCtrl = this;
			this.cdvGrupo.DataSet = this.cdsGrupo;
			this.cdvGrupo.TableName = "";
			this.cdvGrupo.TableViewName = "ArticuloGrupo";
			// 
			// chkTipoGrupo
			// 
			this.chkTipoGrupo.EditValue = true;
			this.chkTipoGrupo.Location = new System.Drawing.Point(248, 240);
			this.chkTipoGrupo.Name = "chkTipoGrupo";
			// 
			// chkTipoGrupo.Properties
			// 
			this.chkTipoGrupo.Properties.Caption = "Todo Tipo de Grupo";
			this.chkTipoGrupo.Size = new System.Drawing.Size(136, 19);
			this.chkTipoGrupo.TabIndex = 79;
			this.chkTipoGrupo.CheckedChanged += new System.EventHandler(this.chkTipoGrupo_CheckedChanged);
			// 
			// cmbTipoGrupo
			// 
			this.cmbTipoGrupo.Location = new System.Drawing.Point(392, 240);
			this.cmbTipoGrupo.Name = "cmbTipoGrupo";
			// 
			// cmbTipoGrupo.Properties
			// 
			this.cmbTipoGrupo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbTipoGrupo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Tipo", "Tipo", 31, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbTipoGrupo.Properties.DataSource = this.cdvTipoGrupo;
			this.cmbTipoGrupo.Properties.DisplayMember = "Tipo";
			this.cmbTipoGrupo.Properties.NullText = "";
			this.cmbTipoGrupo.Properties.ValueMember = "idTipoGrupo";
			this.cmbTipoGrupo.Size = new System.Drawing.Size(192, 20);
			this.cmbTipoGrupo.TabIndex = 78;
			// 
			// cdvTipoGrupo
			// 
			this.cdvTipoGrupo.BindingContextCtrl = this;
			this.cdvTipoGrupo.DataSet = this.cdsTipoGrupo;
			this.cdvTipoGrupo.TableName = "";
			this.cdvTipoGrupo.TableViewName = "ArticuloTipo";
			// 
			// cdsTipoGrupo
			// 
			this.cdsTipoGrupo.BindingContextCtrl = this;
			this.cdsTipoGrupo.DataLibrary = "LibFacturacion";
			this.cdsTipoGrupo.DataLibraryUrl = "";
			this.cdsTipoGrupo.DataSetDef = "dsArticuloTabla";
			this.cdsTipoGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoGrupo.Name = "cdsTipoGrupo";
			this.cdsTipoGrupo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipoGrupo.SchemaDef = null;
			this.cdsTipoGrupo.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoGrupo_BeforeFill);
			// 
			// chkTransaccion
			// 
			this.chkTransaccion.EditValue = true;
			this.chkTransaccion.Location = new System.Drawing.Point(248, 272);
			this.chkTransaccion.Name = "chkTransaccion";
			// 
			// chkTransaccion.Properties
			// 
			this.chkTransaccion.Properties.Caption = "Toda Transacción";
			this.chkTransaccion.Size = new System.Drawing.Size(136, 19);
			this.chkTransaccion.TabIndex = 81;
			this.chkTransaccion.CheckedChanged += new System.EventHandler(this.chkTransaccion_CheckedChanged);
			// 
			// cmbTransaccion
			// 
			this.cmbTransaccion.Location = new System.Drawing.Point(392, 272);
			this.cmbTransaccion.Name = "cmbTransaccion";
			// 
			// cmbTransaccion.Properties
			// 
			this.cmbTransaccion.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbTransaccion.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbTransaccion.Properties.DataSource = this.cdvTransaccion;
			this.cmbTransaccion.Properties.DisplayMember = "Nombre";
			this.cmbTransaccion.Properties.NullText = "";
			this.cmbTransaccion.Properties.ValueMember = "idTipoFactura";
			this.cmbTransaccion.Size = new System.Drawing.Size(192, 20);
			this.cmbTransaccion.TabIndex = 80;
			// 
			// cdvTransaccion
			// 
			this.cdvTransaccion.BindingContextCtrl = this;
			this.cdvTransaccion.DataSet = this.cdsTransaccion;
			this.cdvTransaccion.RowFilter = "Grupo = 1";
			this.cdvTransaccion.Sort = "Numero";
			this.cdvTransaccion.TableName = "";
			this.cdvTransaccion.TableViewName = "CompraNumero";
			// 
			// cdsTransaccion
			// 
			this.cdsTransaccion.BindingContextCtrl = this;
			this.cdsTransaccion.DataLibrary = "LibFacturacion";
			this.cdsTransaccion.DataLibraryUrl = "";
			this.cdsTransaccion.DataSetDef = "dsCompraTabla";
			this.cdsTransaccion.Locale = new System.Globalization.CultureInfo("es-ES");
			this.cdsTransaccion.Name = "cdsTransaccion";
			this.cdsTransaccion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTransaccion.SchemaDef = null;
			this.cdsTransaccion.BeforeFill += new C1.Data.FillEventHandler(this.cdsTransaccion_BeforeFill);
			// 
			// RepProduccion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(648, 349);
			this.Controls.Add(this.chkTransaccion);
			this.Controls.Add(this.cmbTransaccion);
			this.Controls.Add(this.chkTipoGrupo);
			this.Controls.Add(this.cmbTipoGrupo);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.spnBodega);
			this.Controls.Add(this.chkBodega);
			this.Controls.Add(this.chkArticulo);
			this.Controls.Add(this.cmbArticulo);
			this.Controls.Add(this.chkSeccion);
			this.Controls.Add(this.chkClase);
			this.Controls.Add(this.chkMarca);
			this.Controls.Add(this.chkSubGrupo);
			this.Controls.Add(this.chkGrupo);
			this.Controls.Add(this.cmbSeccion);
			this.Controls.Add(this.cmbClase);
			this.Controls.Add(this.cmbMarca);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.cmbFechas);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.dtFechaHasta);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.rdgTipo);
			this.MaximizeBox = false;
			this.Name = "RepProduccion";
			this.Text = "Reportes de Producción";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RepProduccion_Closing);
			this.Load += new System.EventHandler(this.RepProduccion_Load);
			((System.ComponentModel.ISupportInitialize)(this.rdgTipo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechas.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkBodega.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkArticulo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSeccion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkClase.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkMarca.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSeccion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSeccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbClase.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvClase)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClase)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTipoGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTransaccion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransaccion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTransaccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTransaccion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private string ReporteNombre()
		{
			string stReporte = "";
			switch ((int) this.rdgTipo.EditValue)
			{
				case 0:
					stReporte = "ProdDiasInventario.rpt";
					break;
				case 1:
					stReporte = "ProdEntrega.rpt";
					break;
				case 7:
					stReporte = "ProdEntregaSemana.rpt";
					break;
				case 2:
					stReporte = "ProdEntregaMPSemana.rpt";
					break;
				case 3:
					stReporte = "ProdInvBodega.rpt";
					break;
				case 4:
					stReporte = "ProdVentaBodega.rpt";
					break;
				case 5:
					stReporte = "PresupuestoVentas.rpt";
					break;
				case 6:
					stReporte = "ProdTiempo.rpt";
					break;
			}
			return stReporte;
		}

		private string ReporteFiltro()
		{
			string stFiltro = "";
			DateTime dtDesde = (DateTime) this.dtFechaDesde.EditValue;
			DateTime dtHasta = (DateTime) this.dtFechaHasta.EditValue;
			switch ((int) this.rdgTipo.EditValue)
			{
				case 0:
					stFiltro = "{Compra.fecha} < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'";
					stFiltro = "True";
					break;
				case 1:
					stFiltro = "{Compra.Fecha} = '" + dtDesde.ToString("yyyyMMdd") + 
						"' And {Compra.Fecha} = '" + dtHasta.ToString("yyyyMMdd") + "'";
					stFiltro = "True";
					break;
				case 7:
					stFiltro = "{ProdSemana.Mes} = " + dtDesde.Month.ToString() + 
						" And {ProdSemana.Anio} = " + dtDesde.Year.ToString();
					break;
				case 2:
					stFiltro = "{ProdSemana.Mes} = " + dtDesde.Month.ToString() + 
						" And {ProdSemana.Anio} = " + dtDesde.Year.ToString();
					break;
				case 3:
				case 4:
				case 6:
					stFiltro = "{Compra.fecha} >= #" + dtDesde.ToString("MM/dd/yyyy")
						+  "# and {Compra.fecha} < #" + dtHasta.AddDays(1).ToString("MM/dd/yyyy") + "#";
					break;
			}

			// Filtros por Opciones de Chequeo
			if (this.chkArticulo.Checked == false)
			{
				stFiltro += " And {Articulo.idArticulo} = " 
					+ this.cmbArticulo.EditValue.ToString();
			}
			if (this.chkBodega.Checked == false)
			{
				stFiltro += " And {DetCompra.Bodega} = " 
					+ this.spnBodega.EditValue.ToString();
			}
			if (this.chkClase.Checked == false)
			{
				stFiltro += " And {Articulo.idClase} = " 
					+ this.cmbClase.EditValue.ToString();
			}
			if (this.chkGrupo.Checked == false)
			{
				stFiltro += " And {Articulo.idGrupoArticulo} = " 
					+ this.cmbGrupo.EditValue.ToString();
			}
			if (this.chkMarca.Checked == false)
			{
				stFiltro += " And {Articulo.idMarca} = " 
					+ this.cmbMarca.EditValue.ToString();
			}
			if (this.chkSeccion.Checked == false)
			{
				stFiltro += " And {Articulo.idSeccion} = " 
					+ this.cmbSeccion.EditValue.ToString();
			}
			if (this.chkSubGrupo.Checked == false)
			{
				stFiltro += " And {Articulo.idSubGrupo} = " 
					+ this.cmbSubGrupo.EditValue.ToString();
			}
			if (this.chkTipoGrupo.Checked == false)
			{
				stFiltro += " And {ArticuloGrupo.idTipoGrupo} = " 
					+ this.cmbTipoGrupo.EditValue.ToString();
			}
			if (this.chkTransaccion.Checked == false)
			{
				stFiltro += " And {Compra.idTipoFactura} = " 
					+ this.cmbTransaccion.EditValue.ToString();
			}

			return stFiltro;
		}

		private string ReporteTitulo()
		{
			string stTitulo = "";
			DateTime dtDesde = (DateTime) this.dtFechaDesde.EditValue;
			DateTime dtHasta = (DateTime) this.dtFechaHasta.EditValue;
			switch ((int) this.rdgTipo.EditValue)
			{
				case 0:
					stTitulo = "Días de Producción  de " + dtDesde.ToString("MMM/yyyy");
					break;
				case 1:
					stTitulo = "Produccion de " + dtDesde.ToString("MMM/yyyy");
					break;
				case 7:
					stTitulo = "Producción Semanal de " + dtDesde.ToString("MMM/yyyy") ;
					break;
				case 2:
					stTitulo = "Produccion por Semanas de " + dtDesde.ToString("MMM/yyyy");
					break;
				case 3:
					if (this.chkTransaccion.Checked)
						stTitulo = "Saldos por Bodega";
					else
						stTitulo = this.cmbTransaccion.Text.Trim();
					stTitulo += " desde " +
							dtDesde.ToString("dd/MMM/yyyy") + " hasta " + dtHasta.ToString("dd/MMM/yyyy");
					break;
				case 4:
					stTitulo = "{Compra.fecha} >= '" + dtDesde.ToString("yyyyMMdd")
						+  "' and {Compra.fecha} < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
				case 6:
					stTitulo = "{Compra.fecha} >= '" + dtDesde.ToString("yyyyMMdd")
						+  "' and {Compra.fecha} < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
			}
			return stTitulo;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string strReporte = ReporteNombre();
			string strFiltro = ReporteFiltro();
			string strTitulo = ReporteTitulo();

			// Definicion de Parametros
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			if ((int) rdgTipo.EditValue != 5)
			{

				DateTime dtMes1 = (DateTime) this.dtFechaDesde.EditValue;
				ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
				paramField.ParameterFieldName = "iMes";
				discreteVal.Value = dtMes1.Month.ToString();
				paramField.CurrentValues.Add (discreteVal);
				paramFields.Add (paramField);
				// Fin de parametros
			}
			Reporte miReporte = new Reporte(strReporte, strFiltro);
			if ((int) rdgTipo.EditValue != 3)
				if ((int) rdgTipo.EditValue != 5)
					miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.MdiParent = this.MdiParent;
			if (strTitulo.Trim().Length > 0)
				try
				{
					miReporte.Titulo(strTitulo);
				}
				catch
				{
					MessageBox.Show("No existe pcTitulo");
				}
			miReporte.Show();
		}

		private void chkSeccion_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbSeccion.Enabled = !this.chkSeccion.Checked;
		}

		private void chkGrupo_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbGrupo.Enabled = !this.chkGrupo.Checked;
		}

		private void chkArticulo_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbArticulo.Enabled = !this.chkArticulo.Checked;
		}

		private void chkBodega_CheckedChanged(object sender, System.EventArgs e)
		{
			this.spnBodega.Enabled = !this.chkBodega.Checked;
		}

		private void chkMarca_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbMarca.Enabled = !this.chkMarca.Checked;
		}

		private void chkClase_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbClase.Enabled = !this.chkClase.Checked;
		}

		private void chkSubGrupo_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbSubGrupo.Enabled = !this.chkSubGrupo.Checked;
		}


		private bool Verifica()
		{
			bool bOk = true;
			this.errorProvider.SetError(this.cmbArticulo, "");
			this.errorProvider.SetError(this.cmbClase, "");
			this.errorProvider.SetError(this.cmbGrupo, "");
			this.errorProvider.SetError(this.cmbMarca, "");
			this.errorProvider.SetError(this.cmbSeccion, "");
			this.errorProvider.SetError(this.cmbSubGrupo, "");
			this.errorProvider.SetError(this.cmbTipoGrupo, "");
			this.errorProvider.SetError(this.cmbTransaccion, "");

			if (this.chkArticulo.Checked == false)
				if (this.cmbArticulo.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbArticulo, "Ingrese Artículo");
					bOk = false;
				}
					
			if (this.chkBodega.Checked == false)
				if (this.spnBodega.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.spnBodega, "Ingrese Bodega");
					bOk = false;
				}
			if (this.chkClase.Checked == false)
				if (this.cmbClase.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbClase, "Ingrese Clase del articulo");
					bOk = false;
				}
			if (this.chkGrupo.Checked == false)
				if (this.cmbGrupo.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbGrupo, "Ingrese Grupo");
					bOk = false;
				}
			if (this.chkMarca.Checked == false)
				if (this.cmbMarca.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbMarca, "Ingrese Marca");
					bOk = false;
				}
			if (this.chkSeccion.Checked == false)
				if (this.cmbSeccion.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbSeccion, "Ingrese Sección");
					bOk = false;
				}
			if (this.chkSubGrupo.Checked == false)
				if (this.cmbSubGrupo.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbSubGrupo, "Ingrese SubGrupo");
					bOk = false;
				}
			if (this.chkTipoGrupo.Checked == false)
				if (this.cmbTipoGrupo.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbTipoGrupo, "Ingrese Tipo de Inventario");
					bOk = false;
				}
			if (this.chkTransaccion.Checked == false)
				if (this.cmbTransaccion.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbTransaccion, "Ingrese Tipo de Transacción");
					bOk = false;
				}
			return bOk;
		}

		private void cmbFechas_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			int iTipo = this.cmbFechas.SelectedIndex;
			DateTime dtHoy = DateTime.Today;
			int iDia = dtHoy.Day;
			DateTime dtMes1 = dtHoy.AddDays(-iDia+1);
			switch(iTipo)
			{
				case 0: // Hasta Hoy
					this.dtFechaDesde.EditValue = DateTime.Today;
					break;
				case 1: // Solo Hoy
					this.dtFechaDesde.EditValue = DateTime.Today;
					this.dtFechaHasta.EditValue = DateTime.Today;
					break;
				case 2: // Este Mes
					this.dtFechaDesde.EditValue = dtMes1;
					dtMes1 = dtMes1.AddMonths(1);
					this.dtFechaHasta.EditValue = dtMes1.AddDays(-1); 
					break;
				case 3: // El mes pasado
					dtMes1 = dtMes1.AddMonths(-1);
					this.dtFechaDesde.EditValue = dtMes1;
					dtMes1 = dtMes1.AddMonths(1);
					this.dtFechaHasta.EditValue = dtMes1.AddDays(-1); 
					break;
				case 4: // Este Año
					this.dtFechaDesde.EditValue = DateTime.Parse("01/01/" 
						+ DateTime.Today.Year.ToString());
					this.dtFechaHasta.EditValue = DateTime.Parse("12/31/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 5: // El año pasado
					this.dtFechaDesde.EditValue = DateTime.Parse("01/01/" 
						+ DateTime.Today.AddYears(-1).Year.ToString());
					this.dtFechaHasta.EditValue = DateTime.Parse("12/31/" 
						+ DateTime.Today.AddYears(-1).Year.ToString(), us);
					break;

			}

		}
		CultureInfo us = new CultureInfo("en-US");

		private void RepProduccion_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'RepProduccion'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			Cursor = Cursors.Default;
			this.dtFechaDesde.EditValue = DateTime.Today;
			this.dtFechaHasta.EditValue = DateTime.Today;
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void chkTipoGrupo_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbTipoGrupo.Enabled = !this.chkTipoGrupo.Checked;
		}

		private void chkTransaccion_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbTransaccion.Enabled = !this.chkTransaccion.Checked;		
		}

		private void cdsGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsMarca_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMarca.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSubGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSubGrupo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsClase_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsClase.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSeccion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeccion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsTipoGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTipoGrupo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsTransaccion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTransaccion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void RepProduccion_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'RepProduccion'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}
	}
}
