using System;
using System.Globalization;
using System.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RepInventario.
	/// </summary>
	public class RepInventario : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataSet cdsSubGrupo;
		private C1.Data.C1DataSet cdsGrupo;
		private C1.Data.C1DataSet cdsMarca;
		private C1.Data.C1DataView cdvMarca;
		private C1.Data.C1DataSet cdsClase;
		private C1.Data.C1DataView cdvClase;
		private DevExpress.XtraEditors.LookUpEdit cmbMarca;
		private DevExpress.XtraEditors.LookUpEdit cmbSubGrupo;
		private DevExpress.XtraEditors.LookUpEdit cmbGrupo;
		private DevExpress.XtraEditors.LookUpEdit cmbClase;
		private DevExpress.XtraEditors.ComboBoxEdit cmbFechas;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.CheckEdit chkSubGrupo;
		private DevExpress.XtraEditors.CheckEdit chkMarca;
		private DevExpress.XtraEditors.CheckEdit chkClase;
		private DevExpress.XtraEditors.CheckEdit chkSeccion;
		private DevExpress.XtraEditors.LookUpEdit cmbSeccion;
		private DevExpress.XtraEditors.CheckEdit chkGrupo;
		private DevExpress.XtraEditors.CheckEdit chkArticulo;
		private DevExpress.XtraEditors.SpinEdit spnBodega;
		private DevExpress.XtraEditors.LookUpEdit cmbArticulo;
		private DevExpress.XtraEditors.RadioGroup rdgOpcion;
		private C1.Data.C1DataView cdvGrupo;
		private C1.Data.C1DataView cdvSubGrupo;
		private C1.Data.C1DataSet cdsSeccion;
		private C1.Data.C1DataView cdvSeccion;
		private C1.Data.C1DataSet cdsCiudad;
		private C1.Data.C1DataView cdvReporte;
		private C1.Data.C1DataSet cdsTipo;
		private C1.Data.C1DataView cdvTipo;
		private DevExpress.XtraEditors.CheckEdit chkTipo;
		private DevExpress.XtraEditors.LookUpEdit cmbTipo;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkResumen;
		private DevExpress.XtraEditors.CheckEdit chkTipoTrans;
		private DevExpress.XtraEditors.LookUpEdit cmbTipoFact;
		private C1.Data.C1DataSet cdsTipoFact;
		private C1.Data.C1DataView cdvTipoFact;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEntreFechas;
		private C1.Data.C1DataView cdvArtCodigo;
		private DevExpress.XtraEditors.LookUpEdit cmbArtCod;
		private C1.Data.C1DataSet cdsReporte;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private DevExpress.XtraEditors.CheckEdit chkColor;
		private DevExpress.XtraEditors.CheckEdit chkModelo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodArt;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsBodegas;
		private DevExpress.XtraEditors.LookUpEdit cmbModelo;
		private DevExpress.XtraEditors.LookUpEdit cmbColor;
		private DevExpress.XtraEditors.CheckEdit chkKardex;
		private DevExpress.XtraEditors.CheckEdit chkBodega;
		private DevExpress.XtraEditors.DateEdit dtFechaDesde;
		private DevExpress.XtraEditors.DateEdit dtFechaHasta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RepInventario()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.rdgOpcion = new DevExpress.XtraEditors.RadioGroup();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdsSubGrupo = new C1.Data.C1DataSet();
			this.cdsGrupo = new C1.Data.C1DataSet();
			this.cdsMarca = new C1.Data.C1DataSet();
			this.cdvMarca = new C1.Data.C1DataView();
			this.cdsClase = new C1.Data.C1DataSet();
			this.cdvClase = new C1.Data.C1DataView();
			this.cmbMarca = new DevExpress.XtraEditors.LookUpEdit();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbSubGrupo = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvSubGrupo = new C1.Data.C1DataView();
			this.cmbGrupo = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvGrupo = new C1.Data.C1DataView();
			this.cmbSeccion = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvSeccion = new C1.Data.C1DataView();
			this.cdsSeccion = new C1.Data.C1DataSet();
			this.cmbClase = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbFechas = new DevExpress.XtraEditors.ComboBoxEdit();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.chkGrupo = new DevExpress.XtraEditors.CheckEdit();
			this.chkSubGrupo = new DevExpress.XtraEditors.CheckEdit();
			this.chkMarca = new DevExpress.XtraEditors.CheckEdit();
			this.chkClase = new DevExpress.XtraEditors.CheckEdit();
			this.chkSeccion = new DevExpress.XtraEditors.CheckEdit();
			this.chkArticulo = new DevExpress.XtraEditors.CheckEdit();
			this.cmbArticulo = new DevExpress.XtraEditors.LookUpEdit();
			this.spnBodega = new DevExpress.XtraEditors.SpinEdit();
			this.cdvReporte = new C1.Data.C1DataView();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.cdsCiudad = new C1.Data.C1DataSet();
			this.cdsTipo = new C1.Data.C1DataSet();
			this.cdvTipo = new C1.Data.C1DataView();
			this.chkTipo = new DevExpress.XtraEditors.CheckEdit();
			this.cmbTipo = new DevExpress.XtraEditors.LookUpEdit();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.chkResumen = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkTipoTrans = new DevExpress.XtraEditors.CheckEdit();
			this.cmbTipoFact = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvTipoFact = new C1.Data.C1DataView();
			this.cdsTipoFact = new C1.Data.C1DataSet();
			this.chkEntreFechas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdvArtCodigo = new C1.Data.C1DataView();
			this.cmbArtCod = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkColor = new DevExpress.XtraEditors.CheckEdit();
			this.chkModelo = new DevExpress.XtraEditors.CheckEdit();
			this.txtCodArt = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsBodegas = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkBodega = new DevExpress.XtraEditors.CheckEdit();
			this.cmbModelo = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbColor = new DevExpress.XtraEditors.LookUpEdit();
			this.chkKardex = new DevExpress.XtraEditors.CheckEdit();
			this.dtFechaDesde = new DevExpress.XtraEditors.DateEdit();
			this.dtFechaHasta = new DevExpress.XtraEditors.DateEdit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClase)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvClase)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSeccion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSeccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbClase.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechas.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkMarca.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkClase.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSeccion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkArticulo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTipo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTipoTrans.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFact.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArtCod.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkColor.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkModelo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBodegas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkBodega.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModelo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkKardex.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btAceptar.Location = new System.Drawing.Point(24, 345);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 24);
			this.btAceptar.TabIndex = 0;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(120, 345);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 1;
			this.btCancelar.Text = "&Salir";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// rdgOpcion
			// 
			this.rdgOpcion.EditValue = 0;
			this.rdgOpcion.Location = new System.Drawing.Point(16, 17);
			this.rdgOpcion.Name = "rdgOpcion";
			// 
			// rdgOpcion.Properties
			// 
			this.rdgOpcion.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Inventario"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Kardex"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Utilidad, Lista de Precios")});
			this.rdgOpcion.Size = new System.Drawing.Size(176, 95);
			this.rdgOpcion.TabIndex = 2;
			this.rdgOpcion.EditValueChanged += new System.EventHandler(this.rdgOpcion_EditValueChanged);
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
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
			// cdvMarca
			// 
			this.cdvMarca.BindingContextCtrl = this;
			this.cdvMarca.DataSet = this.cdsMarca;
			this.cdvMarca.TableName = "";
			this.cdvMarca.TableViewName = "ArticuloMarca";
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
			// cdvClase
			// 
			this.cdvClase.BindingContextCtrl = this;
			this.cdvClase.DataSet = this.cdsClase;
			this.cdvClase.TableName = "";
			this.cdvClase.TableViewName = "ArticuloClase";
			// 
			// cmbMarca
			// 
			this.cmbMarca.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Marca"));
			this.cmbMarca.Location = new System.Drawing.Point(368, 112);
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
			this.cmbMarca.Properties.PopupWidth = 300;
			this.cmbMarca.Properties.ValueMember = "idMarca";
			this.cmbMarca.Size = new System.Drawing.Size(192, 20);
			this.cmbMarca.TabIndex = 28;
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("ro-RO");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// cmbSubGrupo
			// 
			this.cmbSubGrupo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.SubGrupo"));
			this.cmbSubGrupo.Location = new System.Drawing.Point(368, 40);
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
			this.cmbSubGrupo.Properties.PopupWidth = 300;
			this.cmbSubGrupo.Properties.ValueMember = "idSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(192, 20);
			this.cmbSubGrupo.TabIndex = 27;
			// 
			// cdvSubGrupo
			// 
			this.cdvSubGrupo.BindingContextCtrl = this;
			this.cdvSubGrupo.DataSet = this.cdsSubGrupo;
			this.cdvSubGrupo.TableName = "";
			this.cdvSubGrupo.TableViewName = "ArticuloSubGrupo";
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Grupo"));
			this.cmbGrupo.Location = new System.Drawing.Point(368, 16);
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
			this.cmbGrupo.Properties.PopupWidth = 300;
			this.cmbGrupo.Properties.ValueMember = "idGrupoArticulo";
			this.cmbGrupo.Size = new System.Drawing.Size(192, 20);
			this.cmbGrupo.TabIndex = 26;
			// 
			// cdvGrupo
			// 
			this.cdvGrupo.BindingContextCtrl = this;
			this.cdvGrupo.DataSet = this.cdsGrupo;
			this.cdvGrupo.TableName = "";
			this.cdvGrupo.TableViewName = "ArticuloGrupo";
			// 
			// cmbSeccion
			// 
			this.cmbSeccion.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Tipo"));
			this.cmbSeccion.Location = new System.Drawing.Point(368, 160);
			this.cmbSeccion.Name = "cmbSeccion";
			// 
			// cmbSeccion.Properties
			// 
			this.cmbSeccion.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbSeccion.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Seccion", "Seccion", 50, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbSeccion.Properties.DataSource = this.cdvSeccion;
			this.cmbSeccion.Properties.DisplayMember = "Seccion";
			this.cmbSeccion.Properties.NullText = "";
			this.cmbSeccion.Properties.PopupWidth = 300;
			this.cmbSeccion.Properties.ValueMember = "idSeccion";
			this.cmbSeccion.Size = new System.Drawing.Size(192, 20);
			this.cmbSeccion.TabIndex = 38;
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
			this.cmbClase.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Clase"));
			this.cmbClase.Location = new System.Drawing.Point(368, 136);
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
			this.cmbClase.Properties.PopupWidth = 300;
			this.cmbClase.Properties.ValueMember = "idClase";
			this.cmbClase.Size = new System.Drawing.Size(192, 20);
			this.cmbClase.TabIndex = 37;
			// 
			// cmbFechas
			// 
			this.cmbFechas.EditValue = "Hasta Hoy";
			this.cmbFechas.Location = new System.Drawing.Point(24, 241);
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
																															"El año pasado",
																															"Periodo Ingresado",
																															"Primer Trimestre",
																															"SegundoTrimestre",
																															"TercerTrimestre",
																															"CuartoTrimestre",
																															"Fechas Digitadas"});
			this.cmbFechas.Size = new System.Drawing.Size(160, 20);
			this.cmbFechas.TabIndex = 52;
			this.cmbFechas.SelectedIndexChanged += new System.EventHandler(this.cmbFechas_SelectedIndexChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(24, 302);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 51;
			this.label6.Text = "Hasta:";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(24, 276);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(39, 16);
			this.label7.TabIndex = 50;
			this.label7.Text = "Desde:";
			// 
			// chkGrupo
			// 
			this.chkGrupo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bGrupo"));
			this.chkGrupo.EditValue = true;
			this.chkGrupo.Location = new System.Drawing.Point(224, 17);
			this.chkGrupo.Name = "chkGrupo";
			// 
			// chkGrupo.Properties
			// 
			this.chkGrupo.Properties.Caption = "Todos los Grupos";
			this.chkGrupo.Size = new System.Drawing.Size(120, 18);
			this.chkGrupo.TabIndex = 53;
			this.chkGrupo.CheckedChanged += new System.EventHandler(this.chkGrupo_CheckedChanged);
			// 
			// chkSubGrupo
			// 
			this.chkSubGrupo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bSubGrupo"));
			this.chkSubGrupo.EditValue = true;
			this.chkSubGrupo.Location = new System.Drawing.Point(224, 40);
			this.chkSubGrupo.Name = "chkSubGrupo";
			// 
			// chkSubGrupo.Properties
			// 
			this.chkSubGrupo.Properties.Caption = "Todos los SubGrupos";
			this.chkSubGrupo.Size = new System.Drawing.Size(136, 18);
			this.chkSubGrupo.TabIndex = 54;
			this.chkSubGrupo.CheckedChanged += new System.EventHandler(this.chkSubGrupo_CheckedChanged);
			// 
			// chkMarca
			// 
			this.chkMarca.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bMarca"));
			this.chkMarca.EditValue = true;
			this.chkMarca.Location = new System.Drawing.Point(224, 112);
			this.chkMarca.Name = "chkMarca";
			// 
			// chkMarca.Properties
			// 
			this.chkMarca.Properties.Caption = "Todas los Marcas";
			this.chkMarca.Size = new System.Drawing.Size(136, 18);
			this.chkMarca.TabIndex = 55;
			this.chkMarca.CheckedChanged += new System.EventHandler(this.chkMarca_CheckedChanged);
			// 
			// chkClase
			// 
			this.chkClase.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bClase"));
			this.chkClase.EditValue = true;
			this.chkClase.Location = new System.Drawing.Point(224, 136);
			this.chkClase.Name = "chkClase";
			// 
			// chkClase.Properties
			// 
			this.chkClase.Properties.Caption = "Todas las Clases";
			this.chkClase.Size = new System.Drawing.Size(136, 18);
			this.chkClase.TabIndex = 56;
			this.chkClase.CheckedChanged += new System.EventHandler(this.chkClase_CheckedChanged);
			// 
			// chkSeccion
			// 
			this.chkSeccion.EditValue = true;
			this.chkSeccion.Location = new System.Drawing.Point(224, 160);
			this.chkSeccion.Name = "chkSeccion";
			// 
			// chkSeccion.Properties
			// 
			this.chkSeccion.Properties.Caption = "Todas las Secciones";
			this.chkSeccion.Size = new System.Drawing.Size(136, 18);
			this.chkSeccion.TabIndex = 57;
			this.chkSeccion.CheckedChanged += new System.EventHandler(this.chkSeccion_CheckedChanged);
			// 
			// chkArticulo
			// 
			this.chkArticulo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bArticulo"));
			this.chkArticulo.EditValue = true;
			this.chkArticulo.Location = new System.Drawing.Point(224, 64);
			this.chkArticulo.Name = "chkArticulo";
			// 
			// chkArticulo.Properties
			// 
			this.chkArticulo.Properties.Caption = "Todos los Artículos";
			this.chkArticulo.Size = new System.Drawing.Size(136, 18);
			this.chkArticulo.TabIndex = 59;
			this.chkArticulo.CheckedChanged += new System.EventHandler(this.chkArticulo_CheckedChanged);
			// 
			// cmbArticulo
			// 
			this.cmbArticulo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Articulo"));
			this.cmbArticulo.Location = new System.Drawing.Point(368, 64);
			this.cmbArticulo.Name = "cmbArticulo";
			// 
			// cmbArticulo.Properties
			// 
			this.cmbArticulo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbArticulo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Articulo", "Articulo", 180, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 80, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbArticulo.Properties.DataSource = this.cdvArticulo;
			this.cmbArticulo.Properties.DisplayMember = "Articulo";
			this.cmbArticulo.Properties.NullText = "";
			this.cmbArticulo.Properties.PopupWidth = 300;
			this.cmbArticulo.Properties.ValueMember = "idArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(192, 20);
			this.cmbArticulo.TabIndex = 58;
			this.cmbArticulo.DoubleClick += new System.EventHandler(this.cmbArticulo_DoubleClick);
			// 
			// spnBodega
			// 
			this.spnBodega.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Bodega"));
			this.spnBodega.EditValue = new System.Decimal(new int[] {
																																0,
																																0,
																																0,
																																0});
			this.spnBodega.Location = new System.Drawing.Point(368, 88);
			this.spnBodega.Name = "spnBodega";
			// 
			// spnBodega.Properties
			// 
			this.spnBodega.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnBodega.Properties.IsFloatValue = false;
			this.spnBodega.Properties.Mask.EditMask = "N00";
			this.spnBodega.Size = new System.Drawing.Size(48, 20);
			this.spnBodega.TabIndex = 62;
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
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
			// cdsCiudad
			// 
			this.cdsCiudad.BindingContextCtrl = this;
			this.cdsCiudad.DataLibrary = "LibFacturacion";
			this.cdsCiudad.DataLibraryUrl = "";
			this.cdsCiudad.DataSetDef = "dsClienteTabla";
			this.cdsCiudad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCiudad.Name = "cdsCiudad";
			this.cdsCiudad.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCiudad.SchemaDef = null;
			this.cdsCiudad.BeforeFill += new C1.Data.FillEventHandler(this.cdsCiudad_BeforeFill);
			// 
			// cdsTipo
			// 
			this.cdsTipo.BindingContextCtrl = this;
			this.cdsTipo.DataLibrary = "LibFacturacion";
			this.cdsTipo.DataLibraryUrl = "";
			this.cdsTipo.DataSetDef = "dsArticuloTabla";
			this.cdsTipo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipo.Name = "cdsTipo";
			this.cdsTipo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipo.SchemaDef = null;
			this.cdsTipo.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipo_BeforeFill);
			// 
			// cdvTipo
			// 
			this.cdvTipo.BindingContextCtrl = this;
			this.cdvTipo.DataSet = this.cdsTipo;
			this.cdvTipo.TableName = "";
			this.cdvTipo.TableViewName = "ArticuloTipo";
			// 
			// chkTipo
			// 
			this.chkTipo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bTipo"));
			this.chkTipo.EditValue = true;
			this.chkTipo.Location = new System.Drawing.Point(224, 184);
			this.chkTipo.Name = "chkTipo";
			// 
			// chkTipo.Properties
			// 
			this.chkTipo.Properties.Caption = "Todos los Tipos";
			this.chkTipo.Size = new System.Drawing.Size(136, 18);
			this.chkTipo.TabIndex = 78;
			this.chkTipo.CheckedChanged += new System.EventHandler(this.chkTipo_CheckedChanged);
			// 
			// cmbTipo
			// 
			this.cmbTipo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Tipo"));
			this.cmbTipo.Location = new System.Drawing.Point(368, 184);
			this.cmbTipo.Name = "cmbTipo";
			// 
			// cmbTipo.Properties
			// 
			this.cmbTipo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbTipo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Tipo", "Tipo", 31, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbTipo.Properties.DataSource = this.cdvTipo;
			this.cmbTipo.Properties.DisplayMember = "Tipo";
			this.cmbTipo.Properties.NullText = "";
			this.cmbTipo.Properties.PopupWidth = 300;
			this.cmbTipo.Properties.ValueMember = "idTipoGrupo";
			this.cmbTipo.Size = new System.Drawing.Size(192, 20);
			this.cmbTipo.TabIndex = 77;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// chkResumen
			// 
			this.chkResumen.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoRepFact.Resumen"));
			this.chkResumen.Location = new System.Drawing.Point(24, 164);
			this.chkResumen.Name = "chkResumen";
			this.chkResumen.Size = new System.Drawing.Size(120, 21);
			this.chkResumen.TabIndex = 79;
			this.chkResumen.Text = "Resumen";
			// 
			// chkTipoTrans
			// 
			this.chkTipoTrans.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bTransaccion"));
			this.chkTipoTrans.EditValue = true;
			this.chkTipoTrans.Location = new System.Drawing.Point(224, 208);
			this.chkTipoTrans.Name = "chkTipoTrans";
			// 
			// chkTipoTrans.Properties
			// 
			this.chkTipoTrans.Properties.Caption = "Toda Transacción";
			this.chkTipoTrans.Size = new System.Drawing.Size(128, 18);
			this.chkTipoTrans.TabIndex = 81;
			this.chkTipoTrans.CheckedChanged += new System.EventHandler(this.chkTipoTrans_CheckedChanged);
			// 
			// cmbTipoFact
			// 
			this.cmbTipoFact.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Transaccion"));
			this.cmbTipoFact.Location = new System.Drawing.Point(368, 208);
			this.cmbTipoFact.Name = "cmbTipoFact";
			// 
			// cmbTipoFact.Properties
			// 
			this.cmbTipoFact.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbTipoFact.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbTipoFact.Properties.DataSource = this.cdvTipoFact;
			this.cmbTipoFact.Properties.DisplayMember = "Nombre";
			this.cmbTipoFact.Properties.NullText = "";
			this.cmbTipoFact.Properties.PopupWidth = 300;
			this.cmbTipoFact.Properties.ShowFooter = false;
			this.cmbTipoFact.Properties.ShowHeader = false;
			this.cmbTipoFact.Properties.ValueMember = "idTipoFactura";
			this.cmbTipoFact.Size = new System.Drawing.Size(192, 20);
			this.cmbTipoFact.TabIndex = 80;
			// 
			// cdvTipoFact
			// 
			this.cdvTipoFact.BindingContextCtrl = this;
			this.cdvTipoFact.DataSet = this.cdsTipoFact;
			this.cdvTipoFact.RowFilter = "Signo = 1 Or Signo = -1";
			this.cdvTipoFact.TableName = "";
			this.cdvTipoFact.TableViewName = "CompraNumero";
			// 
			// cdsTipoFact
			// 
			this.cdsTipoFact.BindingContextCtrl = this;
			this.cdsTipoFact.DataLibrary = "LibFacturacion";
			this.cdsTipoFact.DataLibraryUrl = "";
			this.cdsTipoFact.DataSetDef = "dsCompraTabla";
			this.cdsTipoFact.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoFact.Name = "cdsTipoFact";
			this.cdsTipoFact.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipoFact.SchemaDef = null;
			this.cdsTipoFact.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoFact_BeforeFill);
			// 
			// chkEntreFechas
			// 
			this.chkEntreFechas.Location = new System.Drawing.Point(24, 190);
			this.chkEntreFechas.Name = "chkEntreFechas";
			this.chkEntreFechas.Size = new System.Drawing.Size(120, 21);
			this.chkEntreFechas.TabIndex = 82;
			this.chkEntreFechas.Text = "Entre Fechas";
			// 
			// cdvArtCodigo
			// 
			this.cdvArtCodigo.BindingContextCtrl = this;
			this.cdvArtCodigo.DataSet = this.cdsArticulo;
			this.cdvArtCodigo.Sort = "Codigo";
			this.cdvArtCodigo.TableName = "";
			this.cdvArtCodigo.TableViewName = "Articulo";
			// 
			// cmbArtCod
			// 
			this.cmbArtCod.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Articulo"));
			this.cmbArtCod.Location = new System.Drawing.Point(368, 64);
			this.cmbArtCod.Name = "cmbArtCod";
			// 
			// cmbArtCod.Properties
			// 
			this.cmbArtCod.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbArtCod.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Código", 80, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Articulo", "Artículo", 180, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbArtCod.Properties.DataSource = this.cdvArtCodigo;
			this.cmbArtCod.Properties.DisplayMember = "Codigo";
			this.cmbArtCod.Properties.NullText = "";
			this.cmbArtCod.Properties.PopupWidth = 400;
			this.cmbArtCod.Properties.ValueMember = "idArticulo";
			this.cmbArtCod.Size = new System.Drawing.Size(192, 20);
			this.cmbArtCod.TabIndex = 83;
			this.cmbArtCod.Visible = false;
			this.cmbArtCod.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbArtCod_MouseUp);
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DataSource = this.cdvReporte;
			ultraGridBand1.AddButtonCaption = "Reporte";
			this.cmbReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbReporte.DisplayMember = "Reporte";
			this.cmbReporte.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbReporte.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(24, 121);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(160, 21);
			this.cmbReporte.TabIndex = 84;
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// chkColor
			// 
			this.chkColor.EditValue = true;
			this.chkColor.Location = new System.Drawing.Point(224, 256);
			this.chkColor.Name = "chkColor";
			// 
			// chkColor.Properties
			// 
			this.chkColor.Properties.Caption = "Todo Color";
			this.chkColor.Size = new System.Drawing.Size(120, 18);
			this.chkColor.TabIndex = 101;
			this.chkColor.CheckedChanged += new System.EventHandler(this.chkColor_CheckedChanged);
			// 
			// chkModelo
			// 
			this.chkModelo.EditValue = true;
			this.chkModelo.Location = new System.Drawing.Point(224, 232);
			this.chkModelo.Name = "chkModelo";
			// 
			// chkModelo.Properties
			// 
			this.chkModelo.Properties.Caption = "Todo Modelo";
			this.chkModelo.Size = new System.Drawing.Size(120, 18);
			this.chkModelo.TabIndex = 98;
			this.chkModelo.CheckedChanged += new System.EventHandler(this.chkModelo_CheckedChanged);
			// 
			// txtCodArt
			// 
			this.txtCodArt.Location = new System.Drawing.Point(368, 64);
			this.txtCodArt.Name = "txtCodArt";
			this.txtCodArt.Size = new System.Drawing.Size(176, 22);
			this.txtCodArt.TabIndex = 105;
			this.txtCodArt.Visible = false;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoRepFact.Bodega"));
			this.cmbBodega.DataSource = this.udsBodegas;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(424, 88);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(88, 21);
			this.cmbBodega.TabIndex = 128;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// udsBodegas
			// 
			ultraDataColumn2.DataType = typeof(int);
			this.udsBodegas.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn1,
																													 ultraDataColumn2});
			// 
			// chkBodega
			// 
			this.chkBodega.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bBodega"));
			this.chkBodega.EditValue = true;
			this.chkBodega.Location = new System.Drawing.Point(224, 88);
			this.chkBodega.Name = "chkBodega";
			// 
			// chkBodega.Properties
			// 
			this.chkBodega.Properties.Caption = "Todas las Bodegas";
			this.chkBodega.Size = new System.Drawing.Size(136, 18);
			this.chkBodega.TabIndex = 129;
			this.chkBodega.CheckedChanged += new System.EventHandler(this.chkBodega_CheckedChanged);
			// 
			// cmbModelo
			// 
			this.cmbModelo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Transaccion"));
			this.cmbModelo.Location = new System.Drawing.Point(368, 232);
			this.cmbModelo.Name = "cmbModelo";
			// 
			// cmbModelo.Properties
			// 
			this.cmbModelo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbModelo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Modelo", "Modelo", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbModelo.Properties.DisplayMember = "Modelo";
			this.cmbModelo.Properties.NullText = "";
			this.cmbModelo.Properties.PopupWidth = 300;
			this.cmbModelo.Properties.ShowFooter = false;
			this.cmbModelo.Properties.ShowHeader = false;
			this.cmbModelo.Properties.ValueMember = "idModelo";
			this.cmbModelo.Size = new System.Drawing.Size(192, 20);
			this.cmbModelo.TabIndex = 130;
			// 
			// cmbColor
			// 
			this.cmbColor.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Transaccion"));
			this.cmbColor.Location = new System.Drawing.Point(368, 256);
			this.cmbColor.Name = "cmbColor";
			// 
			// cmbColor.Properties
			// 
			this.cmbColor.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbColor.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Color", "Color", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbColor.Properties.DisplayMember = "Color";
			this.cmbColor.Properties.NullText = "";
			this.cmbColor.Properties.PopupWidth = 300;
			this.cmbColor.Properties.ShowFooter = false;
			this.cmbColor.Properties.ShowHeader = false;
			this.cmbColor.Properties.ValueMember = "idColor";
			this.cmbColor.Size = new System.Drawing.Size(192, 20);
			this.cmbColor.TabIndex = 131;
			// 
			// chkKardex
			// 
			this.chkKardex.Location = new System.Drawing.Point(224, 296);
			this.chkKardex.Name = "chkKardex";
			// 
			// chkKardex.Properties
			// 
			this.chkKardex.Properties.Caption = "Kardex Alterno";
			this.chkKardex.Size = new System.Drawing.Size(120, 18);
			this.chkKardex.TabIndex = 132;
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.FechaIni"));
			this.dtFechaDesde.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaDesde.Location = new System.Drawing.Point(88, 272);
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
			this.dtFechaDesde.TabIndex = 133;
			// 
			// dtFechaHasta
			// 
			this.dtFechaHasta.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.FechaFin"));
			this.dtFechaHasta.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaHasta.Location = new System.Drawing.Point(88, 304);
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
			this.dtFechaHasta.TabIndex = 134;
			// 
			// RepInventario
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(584, 402);
			this.Controls.Add(this.dtFechaHasta);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.chkKardex);
			this.Controls.Add(this.cmbColor);
			this.Controls.Add(this.cmbModelo);
			this.Controls.Add(this.chkBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtCodArt);
			this.Controls.Add(this.chkColor);
			this.Controls.Add(this.chkModelo);
			this.Controls.Add(this.cmbReporte);
			this.Controls.Add(this.cmbArtCod);
			this.Controls.Add(this.chkEntreFechas);
			this.Controls.Add(this.chkTipoTrans);
			this.Controls.Add(this.cmbTipoFact);
			this.Controls.Add(this.chkResumen);
			this.Controls.Add(this.chkTipo);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.spnBodega);
			this.Controls.Add(this.chkArticulo);
			this.Controls.Add(this.cmbArticulo);
			this.Controls.Add(this.chkSeccion);
			this.Controls.Add(this.chkClase);
			this.Controls.Add(this.chkMarca);
			this.Controls.Add(this.chkSubGrupo);
			this.Controls.Add(this.chkGrupo);
			this.Controls.Add(this.cmbFechas);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.cmbSeccion);
			this.Controls.Add(this.cmbClase);
			this.Controls.Add(this.cmbMarca);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.MaximizeBox = false;
			this.Name = "RepInventario";
			this.Text = "Reporte de Inventario";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RepInventario_Closing);
			this.Load += new System.EventHandler(this.RepInventario_Load);
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClase)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvClase)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSeccion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSeccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbClase.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechas.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkMarca.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkClase.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSeccion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkArticulo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTipo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTipoTrans.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFact.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArtCod.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkColor.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkModelo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBodegas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkBodega.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModelo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkKardex.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private bool Verifica()
		{
			bool bOk = true;
			this.errorProvider.SetError(this.txtCodArt, "");
			this.errorProvider.SetError(this.cmbArticulo, "");
			this.errorProvider.SetError(this.cmbArtCod, "");
			this.errorProvider.SetError(this.cmbClase, "");
			this.errorProvider.SetError(this.cmbGrupo, "");
			this.errorProvider.SetError(this.cmbMarca, "");
			this.errorProvider.SetError(this.cmbReporte, "");
			this.errorProvider.SetError(this.cmbSeccion, "");
			this.errorProvider.SetError(this.cmbSubGrupo, "");
			this.errorProvider.SetError(this.cmbTipo, "");

			if (this.chkArticulo.Checked == false)
			{
				if (!bCargaArticulo)
				{
					string stBuscaArt = "Select idArticulo From Articulo Where Codigo = '"
						+ txtCodArt.Text.Trim() + "'";
					IdArtxCod = Funcion.iEscalarSQL(cdsArticulo, stBuscaArt, false);
					if (IdArtxCod < 1)
					{
						this.errorProvider.SetError(this.txtCodArt, "Ingrese Articulo");
						bOk = false;
					}
				}
				else
				{
					if (this.cmbArticulo.EditValue == System.DBNull.Value)
					{
						this.errorProvider.SetError(this.cmbArticulo, "Ingrese Articulo");
						this.errorProvider.SetError(this.cmbArtCod, "Ingrese Articulo");
						bOk = false;
					}
				}
			}
					
			if (this.chkClase.Checked == false)
				if (this.cmbClase.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbClase, "Ingrese Clase");
					bOk = false;
				}

				if (this.cmbReporte.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbReporte, "Ingrese Clase");
					bOk = false;
				}

			if (this.chkGrupo.Checked == false)
				if (this.cmbGrupo.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbGrupo, "Ingrese Grupo de Articulo");
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
					this.errorProvider.SetError(this.cmbSeccion, "Ingrese Seccion de Articulo");
					bOk = false;
				}
			if (this.chkSubGrupo.Checked == false)
				if (this.cmbSubGrupo.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbSubGrupo, "Ingrese SubGrupo de Articulo");
					bOk = false;
				}
			if (this.chkTipo.Checked == false)
				if (this.cmbTipo.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbTipo, "Ingrese Tipo de Articulo");
					bOk = false;
				}
			if (this.chkBodega.Checked == false)
				if ((int.Parse(this.spnBodega.EditValue.ToString())) < 1)
				{
					this.errorProvider.SetError(this.spnBodega, "Ingrese Bodega");
					bOk = false;
				}
			return bOk;
		}

		int IdArtxCod = 0;
		private string[] FiltroReporte()
		{
			// Filtros por Opciones de Chequeo
			string stFiltro = "";
			string stTitulo = "";
			
			DateTime dtHasta = (DateTime) this.dtFechaHasta.EditValue;
			DateTime dtDesde = (DateTime) this.dtFechaDesde.EditValue;
			if (this.chkEntreFechas.Checked)
			{
				stFiltro = "{Compra.fecha} < #" + dtHasta.AddDays(1).ToString("MM/dd/yyyy", us) 
					+ "# And {Compra.fecha} >= #" + dtDesde.ToString("MM/dd/yyyy", us) 
					+ "# And {Compra.Borrar} = false "
					+ " And ({CompraNumero.Signo} = 1 OR {CompraNumero.Signo} = -1) ";
				stTitulo = "Desde " + dtDesde.ToString("dd/MMM/yyyy") + 
					" Hasta " + dtHasta.ToString("dd/MMM/yyyy") + " ";
			}
			else
			{
				stFiltro = "{Compra.fecha} < #" + dtHasta.AddDays(1).ToString("MM/dd/yyyy", us) 
					+ "# And {Compra.Borrar} = false "
					+ " And ({CompraNumero.Signo} = 1 OR {CompraNumero.Signo} = -1) ";
				stTitulo = " Hasta " + dtHasta.ToString("dd/MMM/yyyy") + " ";
			}
			stTitulo = "Desde " + dtDesde.ToString("dd/MMM/yyyy") + 
				" Hasta " + dtHasta.ToString("dd/MMM/yyyy") + " ";

			if (this.chkTipoTrans.Checked == false)
			{
				stFiltro += " And {Compra.idTipoFactura} = " + cmbTipoFact.EditValue.ToString();
				if ((int) rdgOpcion.EditValue != 2)
					stTitulo += "Trans. " + cmbTipoFact.Text + " ";
				else
					stTitulo += "Ventas ";
			}
			if ((int) rdgOpcion.EditValue == 2)
			{
				stFiltro += " And {Compra.idTipoFactura} = 1";
			}

			if (cmbReporte.Text == "Lista Precios") stFiltro = "true ";

			if (chkTipo.Checked == false)
			{
				stFiltro += " And {Articulo.idTipoGrupo} = " + cmbTipo.EditValue.ToString();
				stTitulo += "Tipo " + cmbTipo.Text + " ";
			}

			if (chkGrupo.Checked == false)
			{
				stFiltro += " And {Articulo.idGrupoArticulo} = " + cmbGrupo.EditValue.ToString();
				stTitulo += "Grupo " + cmbGrupo.Text + " ";
			}
			if (chkSubGrupo.Checked == false)
			{
				stFiltro += " And {Articulo.idSubGrupo} = " + cmbSubGrupo.EditValue.ToString();
				stTitulo += "SubGrupo " + cmbSubGrupo.Text + " ";
			}
			if (chkArticulo.Checked == false)
			{
				if (!bCargaArticulo)
				{
					stFiltro += " And {Articulo.idArticulo} = " + IdArtxCod.ToString();
					stTitulo += "Articulo " + txtCodArt.Text + " ";
				}
				else
				{
					if (cmbArticulo.Visible)
					{
						stFiltro += " And {Articulo.idArticulo} = " + cmbArticulo.EditValue.ToString();
						stTitulo += "Articulo " + cmbArticulo.Text + " ";
					}
					else
					{
						stFiltro += " And {Articulo.idArticulo} = " + cmbArtCod.EditValue.ToString();
						stTitulo += "Articulo " + cmbArtCod.Text + " ";
					}
				}
			}
			if (chkBodega.Checked == false)
			{
				stFiltro += " And {DetCompra.Bodega} = " + spnBodega.EditValue.ToString();
				stTitulo += "Bodega " + spnBodega.Value.ToString() + " ";
			}
			if (chkBodega.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {DetCompra.Bodega} <> " + spnBodega.EditValue.ToString();
				stTitulo += "Bodega No es " + spnBodega.Value.ToString() + " ";
			}
			if (chkClase.Checked == false)
			{
				stFiltro += " And {Articulo.idClase} = " + cmbClase.EditValue.ToString();
				stTitulo += "Clase " + cmbClase.Text + " ";
			}
			if (chkMarca.Checked == false)
			{
				stFiltro += " And {Articulo.idMarca} = " + cmbMarca.EditValue.ToString();
				stTitulo += "Marca " + cmbMarca.Text + " ";
			}
			if (chkSeccion.Checked == false)
			{
				stFiltro += " And {Articulo.idSeccion} = " + cmbSeccion.EditValue.ToString();
				stTitulo += "Seccion " + cmbSeccion.Text + " ";
			}
			if (chkModelo.Checked == false)
			{
				stFiltro += " And {Articulo.idModelo} = " + cmbModelo.EditValue.ToString();
				stTitulo += "Modelo " + cmbModelo.Text + " ";
			}
			if (chkColor.Checked == false)
			{
				stFiltro += " And {Articulo.idColor} = " + cmbColor.EditValue.ToString();
				stTitulo += "Color " + cmbColor.Text + " ";
			}

			string[] stFiltro1 = new string [2];
			stFiltro1[0] = stFiltro;
			stFiltro1[1] = stTitulo;

			return stFiltro1;
		}

		CultureInfo us = new CultureInfo("en-US");
		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			DateTime dtIni = DateTime.Now;

			if (Funcion.bEjecutaSQL(cdsCiudad, "Exec SeteoGExiste 'GEUR'"))
			{
				try
				{
					this.cdsSeteoF.Update();
				}
				catch
				{
				}
			}
			dr = this.cdsSeteoF.TableViews["SeteoRepFact"].Rows[0];
			this.Cursor = Cursors.WaitCursor;
			// Filtro de Ingresos y Egresos
			DateTime dtFechaIni = (DateTime) this.dtFechaDesde.EditValue;
			DateTime dtFechaFin = (DateTime) this.dtFechaHasta.EditValue;
			string[] strFiltro1 = new string[2];
			strFiltro1 =	FiltroReporte();
			string strFiltro = strFiltro1[0];
			
			string stTitulo = "Desde " + dtFechaIni.ToString("dd/MMM/yyyy") +
				" Hasta " + dtFechaFin.ToString("dd/MMM/yyyy");
			stTitulo = strFiltro1[1];

			// Definicion de Parametros
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();

			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "FechaIni";
			discreteVal.Value = dtFechaIni.ToString("yyyy-MM-dd");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
			// Fin de parametros

			string stReporte = this.cmbReporte.Text.ToString().Trim() + ".rpt";
			string stFiltroRep = cmbReporte.ActiveRow.Cells["Filtro"].Value.ToString();
			if (stFiltroRep != null && stFiltroRep.Trim().Length > 0) 
				strFiltro += " And " + stFiltroRep;
			//	MessageBox.Show(cmbReporte.ActiveRow.Cells["Filtro"].Value.ToString());
			
			if (chkKardex.Checked && (int) rdgOpcion.EditValue == 1)
			{
				stReporte = "KardexR.rpt";
				strFiltro = ", 0";
				if (!chkArticulo.Checked) 
					if (bCargaArticulo)
						strFiltro = ", " + cmbArticulo.EditValue.ToString();
					else
						strFiltro = ", " + IdArtxCod.ToString();

				if (!chkBodega.Checked) strFiltro += ", " + spnBodega.Value.ToString();
				string stExec = "Exec CreaKardex '" + dtFechaIni.ToString("yyyyMMdd")
					+ "', '" + dtFechaFin.AddDays(1).ToString("yyyyMMdd") + "'" + strFiltro;
				Funcion.EjecutaSQL(cdsArticulo, stExec, true);
				#region Filtro de seccion
				strFiltro = "{Articulo.idArticulo} > 0 ";
				if (this.chkClase.Checked == false)
					strFiltro += " And {Articulo.idClase} = " 
						+ this.cmbClase.EditValue.ToString();
				if (this.chkMarca.Checked == false)
					strFiltro += " And {Articulo.idMarca} = " 
						+ this.cmbMarca.EditValue.ToString();
				if (this.chkSeccion.Checked == false)
					strFiltro += " And {Articulo.idSeccion} = " 
						+ this.cmbSeccion.EditValue.ToString();
				if (this.chkModelo.Checked == false)
					strFiltro += " And {Articulo.idModelo} = " 
						+ this.cmbModelo.EditValue.ToString();
				if (this.chkColor.Checked == false)
					strFiltro += " And {Articulo.idColor} = " 
						+ this.cmbColor.EditValue.ToString();
				if (this.chkTipo.Checked == false)
					strFiltro += " And {Articulo.idTipoGrupo} = " 
						+ this.cmbTipo.EditValue.ToString();
				if (this.chkGrupo.Checked == false)
					strFiltro += " And {Articulo.idGrupoArticulo} = " 
						+ this.cmbGrupo.EditValue.ToString();
				if (this.chkSubGrupo.Checked == false)
					strFiltro += " And {Articulo.idSubGrupo} = " 
						+ this.cmbSubGrupo.EditValue.ToString();
				#endregion Filtro de seccion
				#region Visibilidad de Combo
				chkArticulo_CheckedChanged(this, e);
				chkBodega_CheckedChanged(this, e);
				chkClase_CheckedChanged(this, e);
				chkColor_CheckedChanged(this, e);
				chkGrupo_CheckedChanged(this, e);
				chkMarca_CheckedChanged(this, e);
				chkModelo_CheckedChanged(this, e);
				chkSeccion_CheckedChanged(this, e);
				chkSubGrupo_CheckedChanged(this, e);
				chkTipo_CheckedChanged(this, e);
				chkTipoTrans_CheckedChanged(this, e);
				#endregion Visibilidad de Combo
			}

			Reporte miReporte = new Reporte(stReporte, strFiltro);
			//	MessageBox.Show(strFiltro);
			miReporte.crVista.ParameterFieldInfo = paramFields;
			if (this.chkResumen.Checked) miReporte.Resumen(1);
			miReporte.Titulo(stTitulo);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			this.Cursor = Cursors.Default;
			TimeSpan ts = DateTime.Now.Subtract(dtIni);
			btAceptar.ToolTip = "Tiempo de ultimo reporte : " + ts.TotalSeconds.ToString("n2");
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
			this.cmbArtCod.Enabled = !this.chkArticulo.Checked;
			txtCodArt.Enabled = !this.chkArticulo.Checked;
		}

		private void chkBodega_CheckedChanged(object sender, System.EventArgs e)
		{
			this.spnBodega.Enabled = !this.chkBodega.Checked;
			this.cmbBodega.Enabled = !this.chkBodega.Checked;
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

		private void cmbFechas_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (dr == null) return;
			dr["FiltroFecha"] = (int) this.cmbFechas.SelectedIndex;
			int iTipo = (int) this.cmbFechas.SelectedIndex;
			DateTime dtHoy = DateTime.Today;
			int iDia = dtHoy.Day;
			DateTime dtMes1 = dtHoy.AddDays(-iDia+1);
			DateTime dtDesde = (DateTime) dr["FechaIni"];
			DateTime dtHasta = (DateTime) dr["FechaFin"];
			switch(iTipo)
			{
				case 0: // Hasta Hoy
					dtHasta = DateTime.Today;
					dtDesde = (DateTime) this.dtFechaDesde.EditValue;
					break;
				case 1: // Solo Hoy
					dtDesde = DateTime.Today;
					dtHasta = DateTime.Today;
					break;
				case 2: // Este Mes
					dtDesde = dtMes1;
					dtMes1 = dtMes1.AddMonths(1);
					dtHasta = dtMes1.AddDays(-1); 
					break;
				case 3: // El mes pasado
					dtMes1 = dtMes1.AddMonths(-1);
					dtDesde = dtMes1;
					dtMes1 = dtMes1.AddMonths(1);
					dtHasta = dtMes1.AddDays(-1); 
					break;
				case 4: // Este Año
					dtDesde = DateTime.Parse("01/01/" 
						+ DateTime.Today.Year.ToString(), us);
					dtHasta = DateTime.Parse("12/31/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 5: // El Año pasado
					dtDesde = DateTime.Parse("01/01/" 
						+ DateTime.Today.AddYears(-1).Year.ToString(), us);
					dtHasta = DateTime.Parse("12/31/" 
						+ DateTime.Today.AddYears(-1).Year.ToString(), us);
					break;
				case 6: // El Periodo Ingresado
					dtDesde = Funcion.FechaCompra(this.cdsArticulo, 0);
					dtHasta = Funcion.FechaCompra(this.cdsArticulo, 1);
					break;
				case 7: // El Primer Trimestre
					dtDesde = DateTime.Parse("01/01/" 
						+ DateTime.Today.Year.ToString(), us);
					dtHasta = DateTime.Parse("03/31/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 8: // El segundo Trimestre
					dtDesde = DateTime.Parse("04/01/" 
						+ DateTime.Today.Year.ToString(), us);
					dtHasta = DateTime.Parse("06/30/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 9: // El tercer Trimestre
					dtDesde = DateTime.Parse("07/01/" 
						+ DateTime.Today.Year.ToString(), us);
					dtHasta = DateTime.Parse("09/30/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 10: // El cuarto Trimestre
					dtDesde = DateTime.Parse("10/01/" 
						+ DateTime.Today.Year.ToString(), us);
					dtHasta = DateTime.Parse("12/31/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 11: // Periodo Digitado
					break;
			}
			dr["FechaIni"] = dtDesde;
			dr["FechaFin"] = dtHasta;
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
//				DateTime dtIni = DateTime.Now;
//
//		string		stSelect = "EXEC InventarioCrea 0, '20070201', '"
//						 + "20070301', 1";
//			Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
//			TimeSpan ts = DateTime.Now.Subtract(dtIni);
//			MessageBox.Show("Procesado en " + ts.Minutes.ToString()+ ":" +ts.Seconds.ToString());
//
			this.Close();
		}

		private void chkTipo_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbTipo.Enabled = !this.chkTipo.Checked;		
		}

		C1.Data.C1DataRow dr;
		C1.Data.C1DataRow drSeteoF;
		private void RepInventario_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'RepInventario'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			Cursor = Cursors.Default;
			#region Tres Estados en Check
			if (Funcion.bEjecutaSQL(cdsCiudad, "Exec SeteoGExiste 'FTMTE'")) 				
			{
				chkGrupo.Properties.AllowGrayed = true;
				chkSubGrupo.Properties.AllowGrayed = true;
				chkArticulo.Properties.AllowGrayed = true;
				chkBodega.Properties.AllowGrayed = true;
				chkMarca.Properties.AllowGrayed = true;
				chkClase.Properties.AllowGrayed = true;
				chkSeccion.Properties.AllowGrayed = true;
				chkTipo.Properties.AllowGrayed = true;
				chkTipoTrans.Properties.AllowGrayed = true;
				chkModelo.Properties.AllowGrayed = true;
				chkColor.Properties.AllowGrayed = true;
			}
			#endregion Tres Estados en Check
			#region Seteo
			dr = this.cdsSeteoF.TableViews["SeteoRepFact"].Rows[0];
			drSeteoF = this.cdsSeteoF.TableViews["SeteoF"].Rows[0];
			if (dr["FiltroFecha"] == System.DBNull.Value)
				cmbFechas.SelectedIndex = 0;
			else
        cmbFechas.SelectedIndex = (int) dr["FiltroFecha"];
			CambioGrupo();
			if (!bCargaArticulo)
			{
				cmbArtCod.Visible = false;
				cmbArticulo.Visible = false;
				txtCodArt.Visible = true;
			}
			else
			{
				txtCodArt.Visible = false;
			}
			#endregion Seteo

			#region Bodega por Nombre
			if (drSeteoF["VerBodegaNombre"] != DBNull.Value && (bool) drSeteoF["VerBodegaNombre"] == true)
			{
				spnBodega.Width = 0;
				udsBodegas.Rows.Clear();
				string stExecBodega = "Exec BodegaUsuario 0"	;//+ IdTipoFactura.ToString();
				SqlDataReader drBod = Funcion.rEscalarSQL(cdsArticulo, stExecBodega, true);
				int iNumFilas = 0;
				while(drBod.Read())
				{
					udsBodegas.Rows.Add();
					udsBodegas.Rows[iNumFilas]["Nombre"]=drBod.GetString(0);
					udsBodegas.Rows[iNumFilas]["Bodega"]=drBod.GetInt32(1);
					iNumFilas++;
				}
			}
			else
			{
				cmbBodega.Width = 0;
			}
			#endregion Bodega por Nombre
			#region Carga Combos
			cmbModelo.Properties.DataSource = Funcion.dvProcedimiento(cdsArticulo,
				"Select idModelo, Modelo From ArticuloModelo");
			cmbColor.Properties.DataSource = Funcion.dvProcedimiento(cdsArticulo,
				"Select idColor, Color From ArticuloColor");

			#endregion Carga Combos
			#region Visibilidad de Combo
			chkArticulo_CheckedChanged(this, e);
			chkBodega_CheckedChanged(this, e);
			chkClase_CheckedChanged(this, e);
			chkColor_CheckedChanged(this, e);
			chkGrupo_CheckedChanged(this, e);
			chkMarca_CheckedChanged(this, e);
			chkModelo_CheckedChanged(this, e);
			chkSeccion_CheckedChanged(this, e);
			chkSubGrupo_CheckedChanged(this, e);
			chkTipo_CheckedChanged(this, e);
			chkTipoTrans_CheckedChanged(this, e);
			#endregion Visibilidad de Combo
		}

		private void cmbArticulo_DoubleClick(object sender, System.EventArgs e)
		{
			this.cmbArtCod.Visible = true;
			this.cmbArtCod.Enabled = true;
			this.cmbArticulo.EditValue = this.cmbArtCod.EditValue;
			this.cmbArticulo.Visible = false;
		}

		private void chkTipoTrans_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbTipoFact.Enabled = !this.chkTipoTrans.Checked;				
		}

		private void cdsGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		bool bCargaArticulo = true;
		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			string stSelect = "Select IsNull(VerDesplegable, 1) From SeteoF";
			string stFiltro = "";
			bCargaArticulo = Funcion.bEscalarSQL(cdsArticulo, stSelect, true);
			if (!bCargaArticulo)
			{
				stFiltro = "idArticulo < 2";
				C1.Data.SchemaObjects.DataSetDef dataSetDef = 
					e.DataSet.Schema.DataSetDefs["dsArticulos"];
				e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Articulo"],
					stFiltro));
			}
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

		private void cdsTipoFact_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTipoFact.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCiudad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCiudad.Schema.Connections[0].ConnectionString = 
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

		private void cdsTipo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTipo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsReporte_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
//			Funcion miFuncion = new Funcion();
//			this.cdsReporte.Schema.Connections[0].ConnectionString = 
//				miFuncion.AccesoBase(MenuLatinium.stDirInicio);
		}

		private void rdgOpcion_EditValueChanged(object sender, System.EventArgs e)
		{
			CambioGrupo();
		}

		private void CambioGrupo()
		{
			int IdGrupo = (int) this.rdgOpcion.EditValue + 1;
			if (IdGrupo == 2) chkKardex.Enabled = true;
			else chkKardex.Enabled = false;
			string stFiltro = "Tipo = 3 And idGrupo = " + IdGrupo.ToString();
			this.cdvReporte.RowFilter = stFiltro;
			DataRow[] dr = this.cdsReporte.StorageDataSet.Tables["Reporte"].Select(stFiltro);
			this.cmbReporte.Value = dr[0]["idReporte"];
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void chkModelo_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbModelo.Enabled = !chkModelo.Checked;
			if (!chkModelo.Checked) cmbModelo.Select();		
		}

		private void chkColor_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbColor.Enabled = !chkColor.Checked;
			if (!chkColor.Checked) cmbColor.Select();
		}

		private void cmbArtCod_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			this.cmbArticulo.Visible = true;
			this.cmbArticulo.Enabled = true;
			this.cmbArtCod.EditValue = this.cmbArticulo.EditValue;
			this.cmbArtCod.Visible = false;

		}

		private void RepInventario_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'RepInventario'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void chkTodaBodega_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkBodega.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			spnBodega.Enabled = true;
		}

	}
}
