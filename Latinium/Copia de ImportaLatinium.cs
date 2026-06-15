using System;
using System.IO;
using System.Globalization;
using System.Data.SqlClient;
using System.Data;
using C1.C1Excel;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ImportarExcel.
	/// </summary>
	public class ImportarLatinium : DevExpress.XtraEditors.XtraForm
	{
		private C1.C1Excel.C1XLBook c1XLBook;
		private C1.Data.C1DataSet cdsCuentas;
		private C1.Data.C1DataView cdvCuentas;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private C1.Data.C1DataSet cdsAsiento;
		private C1.Data.C1DataView cdvAsiento;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataSet cdsTipoCompra;
		private C1.Data.C1DataView cdvTipoCompra;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataView cdvCompra;
		private C1.Data.C1DataSet cdsGrupoArt;
		private C1.Data.C1DataView cdvGrupo;
		private C1.Data.C1DataSet cdsGrupoCliente;
		private C1.Data.C1DataView cdvClienteGrupo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMensaje;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkModificaExiste;
		private System.Windows.Forms.Button button1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactura;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private Infragistics.Win.Misc.UltraButton btVenta;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraButton btCompra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClave;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.Misc.UltraButton btRetencion;
		private Infragistics.Win.Misc.UltraButton btFacturaLinea;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRucRep;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEmpresa;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private C1.Data.C1DataSet cdsEmpresa;
		private C1.Data.C1DataView cdvEmpresa;
		private Infragistics.Win.Misc.UltraButton btContabilidad;
		private Infragistics.Win.Misc.UltraButton btFacturacion;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private System.Windows.Forms.ErrorProvider errorProvider1;
		private System.ComponentModel.IContainer components;

		public ImportarLatinium()
		{
			InitializeComponent();
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Empresa", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contabilidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numeracion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Roles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoDB");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reportes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Servidor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fox");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruta");
			this.c1XLBook = new C1.C1Excel.C1XLBook();
			this.cdsCuentas = new C1.Data.C1DataSet();
			this.cdvCuentas = new C1.Data.C1DataView();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.cdvAsiento = new C1.Data.C1DataView();
			this.btContabilidad = new Infragistics.Win.Misc.UltraButton();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsTipoCompra = new C1.Data.C1DataSet();
			this.cdvTipoCompra = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cdvCompra = new C1.Data.C1DataView();
			this.cdsGrupoArt = new C1.Data.C1DataSet();
			this.cdvGrupo = new C1.Data.C1DataView();
			this.cdsGrupoCliente = new C1.Data.C1DataSet();
			this.cdvClienteGrupo = new C1.Data.C1DataView();
			this.txtMensaje = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkModificaExiste = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.button1 = new System.Windows.Forms.Button();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.btVenta = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.btCompra = new Infragistics.Win.Misc.UltraButton();
			this.txtClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btRetencion = new Infragistics.Win.Misc.UltraButton();
			this.btFacturaLinea = new Infragistics.Win.Misc.UltraButton();
			this.chkRucRep = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btFacturacion = new Infragistics.Win.Misc.UltraButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbEmpresa = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvEmpresa = new C1.Data.C1DataView();
			this.cdsEmpresa = new C1.Data.C1DataSet();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.errorProvider1 = new System.Windows.Forms.ErrorProvider();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvClienteGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMensaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsCuentas
			// 
			this.cdsCuentas.BindingContextCtrl = this;
			this.cdsCuentas.DataLibrary = "LibContabilidad";
			this.cdsCuentas.DataLibraryUrl = "";
			this.cdsCuentas.DataSetDef = "dsCuenta";
			this.cdsCuentas.FillOnRequest = false;
			this.cdsCuentas.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentas.Name = "cdsCuentas";
			this.cdsCuentas.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentas.SchemaDef = null;
			this.cdsCuentas.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuentas_BeforeFill);
			// 
			// cdvCuentas
			// 
			this.cdvCuentas.BindingContextCtrl = this;
			this.cdvCuentas.DataSet = this.cdsCuentas;
			this.cdvCuentas.Sort = "Codigo";
			this.cdvCuentas.TableName = "";
			this.cdvCuentas.TableViewName = "Cuenta";
			// 
			// openFileDialog
			// 
			this.openFileDialog.DefaultExt = "xls";
			this.openFileDialog.Filter = "Comprimido (*.zip)|*.zip";
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
			// cdvAsiento
			// 
			this.cdvAsiento.BindingContextCtrl = this;
			this.cdvAsiento.DataSet = this.cdsAsiento;
			this.cdvAsiento.Sort = "CodAsiento";
			this.cdvAsiento.TableName = "";
			this.cdvAsiento.TableViewName = "Asiento";
			// 
			// btContabilidad
			// 
			this.btContabilidad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btContabilidad.Location = new System.Drawing.Point(48, 60);
			this.btContabilidad.Name = "btContabilidad";
			this.btContabilidad.Size = new System.Drawing.Size(88, 25);
			this.btContabilidad.TabIndex = 5;
			this.btContabilidad.Text = "Contabilidad";
			this.btContabilidad.Click += new System.EventHandler(this.btContabilidad_Click);
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulo";
			this.cdsArticulo.FillOnRequest = false;
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Codigo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdsTipoCompra
			// 
			this.cdsTipoCompra.BindingContextCtrl = this;
			this.cdsTipoCompra.DataLibrary = "LibFacturacion";
			this.cdsTipoCompra.DataLibraryUrl = "";
			this.cdsTipoCompra.DataSetDef = "dsCompraTabla";
			this.cdsTipoCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoCompra.Name = "cdsTipoCompra";
			this.cdsTipoCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipoCompra.SchemaDef = null;
			this.cdsTipoCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoCompra_BeforeFill);
			// 
			// cdvTipoCompra
			// 
			this.cdvTipoCompra.BindingContextCtrl = this;
			this.cdvTipoCompra.DataSet = this.cdsTipoCompra;
			this.cdvTipoCompra.TableName = "";
			this.cdvTipoCompra.TableViewName = "CompraNumero";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.FillOnRequest = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// cdvCompra
			// 
			this.cdvCompra.BindingContextCtrl = this;
			this.cdvCompra.DataSet = this.cdsCompra;
			this.cdvCompra.Sort = "Numero";
			this.cdvCompra.TableName = "";
			this.cdvCompra.TableViewName = "Compra";
			// 
			// cdsGrupoArt
			// 
			this.cdsGrupoArt.BindingContextCtrl = this;
			this.cdsGrupoArt.DataLibrary = "LibFacturacion";
			this.cdsGrupoArt.DataLibraryUrl = "";
			this.cdsGrupoArt.DataSetDef = "dsArticuloTabla";
			this.cdsGrupoArt.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupoArt.Name = "cdsGrupoArt";
			this.cdsGrupoArt.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupoArt.SchemaDef = null;
			this.cdsGrupoArt.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupoArt_BeforeFill);
			// 
			// cdvGrupo
			// 
			this.cdvGrupo.BindingContextCtrl = this;
			this.cdvGrupo.DataSet = this.cdsGrupoArt;
			this.cdvGrupo.Sort = "Grupo";
			this.cdvGrupo.TableName = "";
			this.cdvGrupo.TableViewName = "ArticuloGrupo";
			// 
			// cdsGrupoCliente
			// 
			this.cdsGrupoCliente.BindingContextCtrl = this;
			this.cdsGrupoCliente.DataLibrary = "LibFacturacion";
			this.cdsGrupoCliente.DataLibraryUrl = "";
			this.cdsGrupoCliente.DataSetDef = "dsClienteTabla";
			this.cdsGrupoCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupoCliente.Name = "cdsGrupoCliente";
			this.cdsGrupoCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupoCliente.SchemaDef = null;
			// 
			// cdvClienteGrupo
			// 
			this.cdvClienteGrupo.BindingContextCtrl = this;
			this.cdvClienteGrupo.DataSet = this.cdsGrupoCliente;
			this.cdvClienteGrupo.TableName = "";
			this.cdvClienteGrupo.TableViewName = "ClienteGrupo";
			// 
			// txtMensaje
			// 
			this.txtMensaje.AcceptsReturn = true;
			this.txtMensaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMensaje.Location = new System.Drawing.Point(16, 241);
			this.txtMensaje.Multiline = true;
			this.txtMensaje.Name = "txtMensaje";
			this.txtMensaje.Size = new System.Drawing.Size(456, 104);
			this.txtMensaje.TabIndex = 17;
			// 
			// chkModificaExiste
			// 
			this.chkModificaExiste.Location = new System.Drawing.Point(128, 164);
			this.chkModificaExiste.Name = "chkModificaExiste";
			this.chkModificaExiste.Size = new System.Drawing.Size(128, 21);
			this.chkModificaExiste.TabIndex = 18;
			this.chkModificaExiste.Text = "Modificar Existentes";
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(376, 138);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(75, 25);
			this.button1.TabIndex = 19;
			this.button1.Text = "Factura";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// cmbTipoFactura
			// 
			this.cmbTipoFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFactura.DataMember = "CompraNumero";
			this.cmbTipoFactura.DataSource = this.cdsTipoCompra;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn3.Header.VisiblePosition = 7;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
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
																										 ultraGridColumn17});
			this.cmbTipoFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoFactura.DisplayMember = "Nombre";
			this.cmbTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFactura.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoFactura.Location = new System.Drawing.Point(288, 26);
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(144, 22);
			this.cmbTipoFactura.TabIndex = 20;
			this.cmbTipoFactura.ValueMember = "idTipoFactura";
			// 
			// btFactura
			// 
			this.btFactura.Location = new System.Drawing.Point(16, 129);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(96, 25);
			this.btFactura.TabIndex = 21;
			this.btFactura.Text = "Factura";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// btVenta
			// 
			this.btVenta.Location = new System.Drawing.Point(128, 207);
			this.btVenta.Name = "btVenta";
			this.btVenta.Size = new System.Drawing.Size(75, 25);
			this.btVenta.TabIndex = 22;
			this.btVenta.Text = "Venta";
			this.btVenta.Click += new System.EventHandler(this.btVenta_Click);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(8, 207);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(107, 15);
			this.ultraLabel1.TabIndex = 23;
			this.ultraLabel1.Text = "Anexo Transaccional";
			// 
			// btCompra
			// 
			this.btCompra.Location = new System.Drawing.Point(224, 207);
			this.btCompra.Name = "btCompra";
			this.btCompra.Size = new System.Drawing.Size(75, 25);
			this.btCompra.TabIndex = 24;
			this.btCompra.Text = "Compra";
			this.btCompra.Click += new System.EventHandler(this.btCompra_Click);
			this.btCompra.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCompra_MouseDown);
			// 
			// txtClave
			// 
			this.txtClave.Location = new System.Drawing.Point(392, 207);
			this.txtClave.Name = "txtClave";
			this.txtClave.Size = new System.Drawing.Size(72, 22);
			this.txtClave.TabIndex = 25;
			this.toolTip1.SetToolTip(this.txtClave, "Ingrese Clave para Borrar Compra de este mes Con click derecho en compras");
			// 
			// btRetencion
			// 
			this.btRetencion.Location = new System.Drawing.Point(312, 207);
			this.btRetencion.Name = "btRetencion";
			this.btRetencion.Size = new System.Drawing.Size(75, 25);
			this.btRetencion.TabIndex = 26;
			this.btRetencion.Text = "Retención";
			this.btRetencion.Click += new System.EventHandler(this.btRetencion_Click);
			// 
			// btFacturaLinea
			// 
			this.btFacturaLinea.Location = new System.Drawing.Point(16, 164);
			this.btFacturaLinea.Name = "btFacturaLinea";
			this.btFacturaLinea.Size = new System.Drawing.Size(96, 24);
			this.btFacturaLinea.TabIndex = 27;
			this.btFacturaLinea.Text = "Factura Línea";
			this.btFacturaLinea.Click += new System.EventHandler(this.btFacturaLinea_Click);
			// 
			// chkRucRep
			// 
			this.chkRucRep.Location = new System.Drawing.Point(328, 181);
			this.chkRucRep.Name = "chkRucRep";
			this.chkRucRep.Size = new System.Drawing.Size(136, 21);
			this.chkRucRep.TabIndex = 28;
			this.chkRucRep.Text = "Verifica Ruc Repetidos";
			// 
			// btFacturacion
			// 
			this.btFacturacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFacturacion.Location = new System.Drawing.Point(280, 60);
			this.btFacturacion.Name = "btFacturacion";
			this.btFacturacion.Size = new System.Drawing.Size(88, 25);
			this.btFacturacion.TabIndex = 29;
			this.btFacturacion.Text = "Facturación";
			this.btFacturacion.Click += new System.EventHandler(this.btFacturación_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.ultraLabel3);
			this.groupBox1.Controls.Add(this.cmbEmpresa);
			this.groupBox1.Controls.Add(this.ultraLabel2);
			this.groupBox1.Controls.Add(this.btFacturacion);
			this.groupBox1.Controls.Add(this.btContabilidad);
			this.groupBox1.Controls.Add(this.cmbTipoFactura);
			this.groupBox1.Location = new System.Drawing.Point(16, 9);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(448, 103);
			this.groupBox1.TabIndex = 30;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Sincronizar";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(248, 26);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(26, 15);
			this.ultraLabel3.TabIndex = 33;
			this.ultraLabel3.Text = "Tipo";
			// 
			// cmbEmpresa
			// 
			this.cmbEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEmpresa.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresa.DataSource = this.cdvEmpresa;
			this.cmbEmpresa.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Width = 165;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 3;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 4;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 5;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 6;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 7;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 8;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 9;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 10;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 11;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 19;
			ultraGridColumn30.Header.VisiblePosition = 12;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 56;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
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
			this.cmbEmpresa.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEmpresa.DisplayMember = "Empresa";
			this.cmbEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEmpresa.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresa.Location = new System.Drawing.Point(48, 26);
			this.cmbEmpresa.Name = "cmbEmpresa";
			this.cmbEmpresa.Size = new System.Drawing.Size(184, 21);
			this.cmbEmpresa.TabIndex = 31;
			this.cmbEmpresa.ValueMember = "Contabilidad";
			// 
			// cdvEmpresa
			// 
			this.cdvEmpresa.BindingContextCtrl = this;
			this.cdvEmpresa.DataSet = this.cdsEmpresa;
			this.cdvEmpresa.Sort = "Empresa";
			this.cdvEmpresa.TableName = "";
			this.cdvEmpresa.TableViewName = "Empresa";
			// 
			// cdsEmpresa
			// 
			this.cdsEmpresa.BindingContextCtrl = this;
			this.cdsEmpresa.DataLibrary = "LibEmpresa";
			this.cdsEmpresa.DataLibraryUrl = "";
			this.cdsEmpresa.DataSetDef = "dsEmpresa";
			this.cdsEmpresa.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEmpresa.Name = "cdsEmpresa";
			this.cdsEmpresa.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsEmpresa.SchemaDef = null;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(8, 26);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(37, 15);
			this.ultraLabel2.TabIndex = 30;
			this.ultraLabel2.Text = "Origen";
			// 
			// errorProvider1
			// 
			this.errorProvider1.ContainerControl = this;
			// 
			// ImportarLatinium
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(488, 351);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.chkRucRep);
			this.Controls.Add(this.btFacturaLinea);
			this.Controls.Add(this.btRetencion);
			this.Controls.Add(this.txtClave);
			this.Controls.Add(this.btCompra);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.btVenta);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.chkModificaExiste);
			this.Controls.Add(this.txtMensaje);
			this.Name = "ImportarLatinium";
			this.Text = "Importar de Latinium";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ImportarExcel_Closing);
			this.Load += new System.EventHandler(this.ImportarExcel_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvClienteGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMensaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsCuentas_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuentas.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void btAsientos_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Formato Asiento: \nCódigo, Descripción, Fecha, ";
			this.txtMensaje.Text += "Codigo Contable, Debe, Haber, Centro de Costo";
			if (MessageBox.Show("¿Desea Subir Asiento de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stArchivo = "";
			this.openFileDialog.InitialDirectory = Application.ExecutablePath;
			this.openFileDialog.InitialDirectory = @"C:\Ciad2000"; // Borrar
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				stArchivo = this.openFileDialog.FileName.ToString();
			}
			else
				MessageBox.Show("Cancelado por el usuario");

			C1.C1Excel.C1XLBook miAsiento = new C1.C1Excel.C1XLBook();
			try
			{
				miAsiento.Load(@stArchivo);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message);
				return;
			}
			XLSheet sheet = miAsiento.Sheets[0];

			string stCodigo, stNombre;
			int iCuenta = 0, iTipo;

			C1.Data.C1DataRow drAsiento = this.cdsAsiento.TableViews["Asiento"].AddNew();
			C1.Data.C1DataRow drDetAsiento= this.cdsAsiento.TableViews["AsientoDetalle"].AddNew();
			string [] stTitulo = new string[100];
			// Lectura de Titulos de columnas
			for (int i = 0; i < sheet.Columns.Count; i++)
			{
				XLCell cell = sheet.GetCell(0, i);
				//				if (cell != null && cell.Value != null)
			{
				stTitulo[i] = (string) cell.Value;
				try
				{
					drAsiento[stTitulo[i]] = System.DBNull.Value;
				}
				catch
				{
					try
					{
						drDetAsiento[stTitulo[i]]=System.DBNull.Value;
					}
					catch
					{
						MessageBox.Show("No existe campo " + stTitulo[i].Trim()
							+ ".\nPrimera fila debe tener los campos a importar.",
							"Error");
						return;
					} // Fin de segundo try
				} // Fin de Primer try
			} // Fin de If
			} // Fin de For de primera fila
			this.cdsAsiento.RejectChanges();

			for (int f = 1; f < sheet.Rows.Count; f++)
			{
				XLCell cell = sheet.GetCell(f, 0);
				if (cell != null && cell.Value != null)
				{
					stCodigo = (string) cell.Value;
					int iBusca = this.cdvAsiento.Find(stCodigo);
					if (iBusca < 0)
					{
						iCuenta ++;
						drAsiento = this.cdsAsiento.TableViews["Asiento"].AddNew();
						drAsiento["CodAsiento"] = stCodigo;
						cell = sheet.GetCell(f, 1);
						if (cell != null && cell.Value != null)
						{
							stNombre = (string) cell.Value;
							drAsiento[stTitulo[1]] = stNombre; // Descripcion
						}
						cell = sheet.GetCell(f, 2);
						if (cell != null && cell.Value != null)
						{
							DateTime dtFecha = (DateTime) sheet.GetCell(f, 2).Value;
							drAsiento[stTitulo[2]] = dtFecha;
						}
						cell = sheet.GetCell(f, 3);
						if (cell != null && cell.Value != null)
						{
							iTipo = int.Parse(sheet.GetCell(f, 3).Value.ToString());
							drAsiento["idTipoAsiento"] = iTipo;
						}
						cell = sheet.GetCell(f, 4);
						if (cell != null && cell.Value != null)
						{
							stNombre = (string) sheet.GetCell(f, 4).Value;
							drAsiento["Cheque"] = stNombre;
						}
						cell = sheet.GetCell(f, 5);
						if (cell != null && cell.Value != null)
						{
							stNombre = (string) sheet.GetCell(f, 5).Value;
							drAsiento["Notas"] = stNombre;
						}
						this.cdsAsiento.Update();

						// Grabacion de detalle de asientos
						while(f < sheet.Rows.Count && (string) sheet.GetCell(f, 0).Value == stCodigo)
						{
							drDetAsiento = this.cdsAsiento.TableViews["AsientoDetalle"].AddNew();
							cell = sheet.GetCell(f, 6);
							if (cell != null && cell.Value != null)
							{
								string stCodCuenta = (string) sheet.GetCell(f, 6).Value;
								int iBuscaCuenta = this.cdvCuentas.Find(stCodCuenta);
								if (iBuscaCuenta > 0)
								{
									C1.Data.C1DataRow drCuenta = 
										this.cdsCuentas.TableViews["Cuenta"].Rows[iBuscaCuenta-1];
									drDetAsiento["idCuenta"] = drCuenta["idCuenta"];
								}
							}
							cell = sheet.GetCell(f, 7);
							if (cell != null && cell.Value != null)
							{
								drDetAsiento["Debe"] = (double) sheet.GetCell(f, 7).Value;
							}
							cell = sheet.GetCell(f, 8);
							if (cell != null && cell.Value != null)
							{
								drDetAsiento["Haber"] = (double) sheet.GetCell(f, 8).Value;
							}
							drDetAsiento["idAsiento"] = drAsiento["idAsiento"];
							f++;
						} // fin de detAsiento
						f --;
					} // Fin de Asiento Encontrado
				} // Fin de Linea
			} // Fin de For

			if (iCuenta == 0)
			{
				MessageBox.Show("No hay Asientos para actualizar", 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.cdsAsiento.Update();
			MessageBox.Show(iCuenta.ToString().Trim() + " Asientos Ingresados desde Excel", 
				"Información", MessageBoxButtons.OK,
				MessageBoxIcon.Information);
		}

		private void cdsAsiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAsiento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsTipoCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTipoCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private string stArchivo = "";
		CultureInfo us = new CultureInfo("en-US");
		private bool LeeExcel(C1.Data.C1DataSet ds, C1.Data.C1DataTable dtArchivo, 
			C1.Data.C1DataView dv, int iTipo)
		{
			string stTabla = dtArchivo.Name.ToString();
			this.Cursor = Cursors.WaitCursor;
			if (stArchivo.Trim().Length == 0)
			{
				this.openFileDialog.Filter = "Excel (*.Xls)|*.xls";
				this.openFileDialog.InitialDirectory = Application.ExecutablePath;
				if (DialogResult.OK == this.openFileDialog.ShowDialog())
				{
					stArchivo = this.openFileDialog.FileName.ToString();
				}
				else
				{
					MessageBox.Show("Cancelado por el usuario");
					this.Cursor = Cursors.Default;
					return false;
				}
			}

			C1.C1Excel.C1XLBook miArticulo = new C1.C1Excel.C1XLBook();
			try
			{
				miArticulo.Load(@stArchivo);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error Abriendo el Archivo");
				this.Cursor = Cursors.Default;
				return false;
			}
			XLSheet sheet = miArticulo.Sheets[0];
			string stCodigo;
			int iCuenta = 0;

			// Verificacion que primera fila tiene los nombres de los campos a subirse
			C1.Data.C1DataRow drArticulo = dtArchivo.AddNew();
			string [] stTitulo = new string[100];
			// Lectura de Titulos de columnas
			for (int i = 0; i < sheet.Columns.Count; i++)
			{
				XLCell cell = sheet.GetCell(0, i);
				//				if (cell != null && cell.Value != null)
			{
				try
				{
					stTitulo[i] = "";
					stTitulo[i] = (string) cell.Value;
					drArticulo[stTitulo[i]].GetType();
				}
				catch
				{
					if (i == 0)
					{
						int ColError = i + 1;
						MessageBox.Show("No existe campo " + stTitulo[i]
							+ " en columna " + ColError.ToString().Trim()
							+ ".\nPrimera fila debe tener los campos a importar.",
							"Error en nombres de Campos");
						this.Cursor = Cursors.Default;
						return false;
					}
					i = sheet.Columns.Count;
				} // Fin de Primer try
			} // Fin de If
			} // Fin de For de primera fila
			ds.RejectChanges();
			// FIN de Verificacion que primera fila tiene los nombres de los campos a subirse

			// Lectura de datos desde segunda fila
			this.Cursor = Cursors.WaitCursor;
			//			if (File.Exists("SalidaError.txt")) File.Delete("SalidaError.txt");
			StreamWriter sr= File.CreateText("SalidaError.XLS");

			dv.Sort = stTitulo[0];
			for (int f = 1; f < sheet.Rows.Count; f++)
			{
				// Primera columna tiene el codigo que no se debe repetir
				int iBusca = -1;
				stCodigo = "";
				XLCell cell = sheet.GetCell(f, 0);
				if (cell != null && cell.Value != null)
				{
					if (cell.Value.GetType().Name == "String")
						stCodigo = (string) cell.Value;
					if (cell.Value.GetType().Name == "Double")
						stCodigo = (string) cell.Value.ToString();

					stCodigo = stCodigo.Trim();
					string stSelect = "SELECT COUNT(*) FROM "
						+ stTabla + " WHERE " + stTitulo[0] + " = '" 
						+ stCodigo + "'";
					iBusca = Funcion.iEscalarSQL(cdsArticulo, stSelect);

					if (iBusca <= 0 || chkModificaExiste.Checked) // Articulo no encontrado o modifica existente
					{
						iCuenta ++;

						//						drArticulo[stTitulo[0]] = stCodigo;
						//TipoDefinido(drArticulo, iTipo);

						#region Lectura de columnas
						string stInsert = "";
						bool bModifica = false;
						if (chkModificaExiste.Checked && iBusca > 0) 
							bModifica = true;

						if (bModifica)
							stInsert = "UPDATE " + stTabla + " SET ";
						else
							stInsert = "INSERT INTO " + stTabla + " (";

						string stValor = " VALUES (";
						DateTime dtFechaLeida;
						for (int j = 0; j < sheet.Columns.Count; j++)
						{
							cell = sheet.GetCell(f, j);
							if (cell != null && cell.Value != null)
							{
								try
								{
									stInsert += stTitulo[j];
									if (cell.Value.GetType().Name == "String")
									{
										string stCelda = (string) cell.Value;
										if (stCelda.ToUpper() == "TRUE" || stCelda.ToUpper() == "FALSE")
										{
											string stBool = "0";
											if (stCelda.ToUpper() == "TRUE") stBool = "1";
											stValor += stBool;
											if (bModifica) stInsert += " = " + stBool;
										}
										else 
										{
											stValor += "'" + stCelda +"'";
											if (bModifica) stInsert += " = " + "'" + stCelda +"'";
										}
									}
									if (cell.Value.GetType().Name == "DateTime")
									{
										dtFechaLeida = (DateTime) cell.Value;
										stValor += "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
										if (bModifica) stInsert += " = " + "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
									}
									if (cell.Value.GetType().Name == "Double")
									{
										double dValor = (double) cell.Value;
										stValor += dValor.ToString("0.00", us);
										if (bModifica) stInsert += " = " + dValor.ToString("0.00", us);
									}
									if (j!=sheet.Columns.Count -1) 
									{
										stInsert+= ", ";
										stValor += ", ";
									}
								}
								catch(System.Exception ex)
								{
									int ColError = j + 1;
									int FilaError = f + 1;
									DialogResult drOpcion;
									drOpcion = MessageBox.Show("Error en Fila: " + FilaError.ToString() 
										+ " Columna: " + ColError.ToString() + " " 
										+ ex.Message + "\n¿Desea continuar con la subida de datos?", 
										"Error en Celda", MessageBoxButtons.YesNo,
										MessageBoxIcon.Question);
									if (drOpcion == DialogResult.No)
									{
										this.Cursor = Cursors.Default;
										return false;
									}
								} // Fin de Catch
							} // Fin de verificacion de valor nulo
						} // Fin de lectura de columnas for
						if (stInsert.Trim().EndsWith(","))
						{
							stInsert = stInsert.Substring(0,stInsert.Length - 2);
							stValor = stValor.Substring(0,stValor.Length - 2);
						}
						if (!bModifica) stInsert+= ") ";
						stValor += ") ";
						#endregion
					
						try
						{
							if (bModifica) 
							{
								stValor = "";
								stInsert += " WHERE " + stTitulo[0] + " = '" + stCodigo +"'";
							}
							Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT " + stTabla + " on", false);
							Funcion.EjecutaSQL(cdsArticulo, stInsert + stValor, false);
						}
						catch (System.Exception ex)
						{
							int f1 = f+1;
							sr.WriteLine(ex.Message + " Codigo " + stCodigo 
								+ " Linea " + f1.ToString());
						}
					} // Fin de Articulo no Encontrado
					else
					{
						int f2 = f+1;
						sr.WriteLine("Repetido: " + stCodigo 
							+ " Linea " + f2.ToString());
					}
				} // Fin de Linea
			} // Fin de For

			sr.Close();
			this.Cursor = Cursors.Default;
			if (iCuenta == 0)
			{
				MessageBox.Show("No hay Registros para Actualizar en " + stTabla, 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return false;
			}
			MessageBox.Show(iCuenta.ToString().Trim() + " Registros Migrados desde Excel", 
				"Información", MessageBoxButtons.OK,
				MessageBoxIcon.Information);
			return true;
		}


		private void AsignaCampo(C1.Data.C1DataRow dr, C1.C1Excel.XLCell cell, 
			string stCampo, int i, int j)
		{
			if (cell != null && cell.Value != null)
			{
				try
				{
					if (cell.Value.GetType().Name == "String")
						dr[stCampo] = (string) cell.Value;
					if (cell.Value.GetType().Name == "DateTime")
						dr[stCampo] = (DateTime) cell.Value;
					if (cell.Value.GetType().Name == "Double")
						dr[stCampo] = (Double) cell.Value;
				}
				catch(System.Exception ex)
				{
					int ColError = j + 1;
					int FilaError = i + 1;
					DialogResult drOpcion;
					drOpcion = MessageBox.Show("Error en Fila: " + FilaError.ToString() 
						+ " Columna: " + ColError.ToString() + " " 
						+ ex.Message + "\n¿Desea continuar con la subida de datos?", 
						"Error en Celda", MessageBoxButtons.YesNo,
						MessageBoxIcon.Question);
					if (drOpcion == DialogResult.No)
					{
						this.Cursor = Cursors.Default;
						return;
					}
				}
			}
		}
		
		private void BuscaCliente(C1.Data.C1DataRow dr, C1.C1Excel.XLCell cell)
		{
			if (cell == null) return;
			string stCliente = (string) cell.Value;

			string stFiltro = "Nombre = '" + stCliente + "'";
			DataRow [] drCliente = 
				cdsCliente.StorageDataSet.Tables["Cliente"].Select(stFiltro);
			if (drCliente.Length == 0) return;
			dr["idCliente"] = drCliente[0]["idCliente"];
		}
		private void BuscaArticulo(C1.Data.C1DataRow dr, C1.C1Excel.XLCell cell)
		{
			if (cell == null) return;
			string stArticulo = (string) cell.Value;
			string stFiltro = "Codigo = '" + stArticulo + "'";
			DataRow [] drArticulo = 
				cdsArticulo.StorageDataSet.Tables["Articulo"].Select(stFiltro);
			if (drArticulo.Length == 0) return;
			dr["idArticulo"] = drArticulo[0]["idArticulo"];
		}

		private void btClientes_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Ruc, Nombre, Direccion, ";
			this.txtMensaje.Text += "Telefono, EMail, idGrupoCliente, Fax, Representante, etc";
			if (MessageBox.Show("¿Desea Subir Clientes de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			this.LeeExcel(this.cdsCliente, this.cdsCliente.TableViews["Cliente"], 
				this.cdvCliente, 1);
		}

		private void btProveedor_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Ruc, Nombre, Direccion, ";
			this.txtMensaje.Text += "Telefono, EMail, idGrupoCliente, Fax, Representante, etc";
			if (MessageBox.Show("¿Desea Subir Proveedores de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			this.LeeExcel(this.cdsCliente, this.cdsCliente.TableViews["Cliente"], 
				this.cdvCliente, 2);
		}

		private void btCuentas_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Codigo, Descripcion";
			if (MessageBox.Show("¿Desea Subir Plan de Cuentas de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			bool bExito = this.LeeExcel(this.cdsCuentas, this.cdsCuentas.TableViews["Cuenta"], this.cdvCuentas, 0);
			if (bExito)
			{
				this.Cursor = Cursors.WaitCursor;
				Contabilidad.GeneraArbol(this.cdsCuentas);
				this.Cursor = Cursors.Default;
			}
		}


		private void cdsGrupoArt_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupoArt.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}


		private void ImportarExcel_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Importa Latinium'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			Cursor = Cursors.Default;
		}


		private void ImportarExcel_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Importa Latinium'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			SqlCommand cmdActualiza;
			this.cdsCompra.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = 
				this.cdsCompra.Schema.Connections[0];

			// Borra blancos de Asiento
			miConeccion = this.cdsCompra.Schema.Connections[0];
			string stSelect = "ALTER TABLE DetCompra ENABLE TRIGGER SaldoCosto";
			//			Notas = RTrim(Notas),
			cmdActualiza = new SqlCommand (stSelect
				, (SqlConnection) miConeccion.DbConnection);
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Subir Facturas de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			//			string stArchivo = "";
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = Application.ExecutablePath;
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				stArchivo = this.openFileDialog.FileName.ToString();
			}
			else
			{
				MessageBox.Show("Cancelado por el usuario");
				this.Cursor = Cursors.Default;
				return;
			}

			stArchivo = stArchivo.ToUpper();

			C1.C1Excel.C1XLBook miCompra = new C1.C1Excel.C1XLBook();
			try
			{
				miCompra.Load(@stArchivo);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message);
				return;
			}
			XLSheet sheet = miCompra.Sheets[0];

			string stCodigo;
			int iCuenta = 0;

			C1.Data.C1DataRow drCompra = this.cdsCompra.TableViews["Compra"].AddNew();
			C1.Data.C1DataRow drDetCompra = this.cdsCompra.TableViews["DetCompra"].AddNew();
			string [] stTitulo = new string[200];
			// Lectura de Titulos de columnas de Compras
			XLCell cell = sheet.GetCell(0, 0);

			int iCamposCompra = 0;
			for (int i = 0; i < sheet.Columns.Count; i++)
			{
				cell = sheet.GetCell(0, i);
				try
				{
					stTitulo[i] = (string) cell.Value;
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error al leer titulo");
					return;
				}
				if (stTitulo[i].ToUpper().Trim() == "CLIENTE") continue;

				try
				{
					drCompra[stTitulo[i]].GetType();
					iCamposCompra ++;
				}
				catch
				{
					MessageBox.Show("No existe campo " + stTitulo[i].Trim()
						+ ".\nPrimera fila debe tener los campos de Compras a importar.",
						"Error");
					return;
				} // Fin de try
			} // Fin de For de primera fila
			//			this.cdsCompra.RejectChanges();

			// Lectura de Titulos de columnas de DetCompras
			string [] stTitulo1 = new string[200];
			int iCamposDet = 0;
			cell = sheet.GetCell(3, 0);
			if ((cell != null && cell.Value != null))
			{
				MessageBox.Show("Columna A4 debe estar vacia");
				return;
			}
			for (int i = 1; i < sheet.Columns.Count; i++)
			{
				cell = sheet.GetCell(1, i);
				try
				{
					stTitulo1[i] = (string) cell.Value;
					if (stTitulo1[i].ToUpper().Trim() == "CODIGO")
					{
						iCamposDet ++;
						continue;
					}
				}
				catch(System.Exception ex)
				{
					Console.WriteLine(ex.Message + "Segunda Fila Titulo");
					break;
				}
				try
				{
					drDetCompra[stTitulo1[i]].GetType();
					iCamposDet ++;
				}
				catch
				{
					MessageBox.Show("No existe campo " + stTitulo1[i].Trim()
						+ ".\nSegunda fila debe tener los campos de Detalle de Compras a importar.",
						"Error");
					return;
				} // Fin de try
			} // Fin de For de primera fila
			this.cdsCompra.RejectChanges();

			// Ingreso de Compras
			StreamWriter sr= File.CreateText("LatiniumError.xls");
			string stClave = "";
			for (int f = 1; f < sheet.Rows.Count; f++)
			{
				cell = sheet.GetCell(f, 0);
				if (cell != null && cell.Value != null) // Linea de Compra
				{
					double dIdCompra = 0;
					int iBusca = 0;
					if (stTitulo[0].Trim().ToUpper() == "IDCOMRPA")
					{
						dIdCompra = (double) cell.Value;
						stCodigo = "SELECT COUNT(*) FROM Compra WHERE idCompra = "
							+ dIdCompra.ToString();				
						stClave = dIdCompra.ToString();
					}
					else
					{
						string stNumero = cell.Value.ToString().Trim();
						stCodigo = "SELECT COUNT(*) FROM Compra WHERE Numero = '"
							+ stNumero + "'";
						stClave = stNumero;
					}

					iBusca = Funcion.iEscalarSQL(cdsArticulo, stCodigo);

					string stInsert = "";
					bool bModifica = false;
					if (chkModificaExiste.Checked && iBusca > 0) 
						bModifica = true;

					if (bModifica)
						stInsert = "UPDATE Compra SET ";
					else
						stInsert = "INSERT INTO Compra (";

					string stValor = " VALUES (";
					if (iBusca <= 0 || bModifica) // Compra no encontrada o Modificacion
					{
						iCuenta ++;

						DateTime dtFechaLeida;
						#region Lectura de Columnas de Compra
						for (int j = 0; j < iCamposCompra; j++)
						{
							cell = sheet.GetCell(f, j);
							if (cell != null && cell.Value != null)
							{
								try
								{
									stInsert += stTitulo[j];
									if (cell.Value.GetType().Name == "String")
									{
										string stCelda = (string) cell.Value;
										
										if (stCelda.ToUpper() == "TRUE" || stCelda.ToUpper() == "FALSE")
										{
											string stBool = "0";
											if (stCelda.ToUpper() == "TRUE") stBool = "1";
											stValor += stBool;
											if (bModifica) stInsert += " = " + stBool;
										}
										else 
										{
											stValor += "'" + stCelda +"'";
											if (bModifica) stInsert += " = " + "'" + stCelda +"'";
										}
									}
									if (cell.Value.GetType().Name == "DateTime")
									{
										dtFechaLeida = (DateTime) cell.Value;
										stValor += "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
										if (bModifica) stInsert += " = " + "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
									}
									if (cell.Value.GetType().Name == "Double")
									{
										double dValor = (double) cell.Value;
										stValor += dValor.ToString("0.00", us);
										if (bModifica) stInsert += " = " + dValor.ToString("0.00", us);
									}
									if (j!=sheet.Columns.Count -1) 
									{
										stInsert+= ", ";
										stValor += ", ";
									}
								}
								catch(System.Exception ex)
								{
									int ColError = j + 1;
									int FilaError = f + 1;
									DialogResult drOpcion;
									drOpcion = MessageBox.Show("Error en Fila: " + FilaError.ToString() 
										+ " Columna: " + ColError.ToString() + " " 
										+ ex.Message + "\n¿Desea continuar con la subida de datos?", 
										"Error en Celda", MessageBoxButtons.YesNo,
										MessageBoxIcon.Question);
									if (drOpcion == DialogResult.No)
									{
										this.Cursor = Cursors.Default;
										return;
									}
								} // Fin de Catch
							} // Fin de verificacion de valor nulo
						} // Fin de lectura de columnas for
						#endregion
						if (stInsert.Trim().EndsWith(","))
						{
							stInsert = stInsert.Substring(0,stInsert.Length - 2);
							stValor = stValor.Substring(0,stValor.Length - 2);
						}
						if (!bModifica) stInsert+= ") ";
						stValor += ") ";
						f++;

						// Grabacion de Compra
						#region Compra Grabacion

						try
						{
							if (bModifica) 
							{
								stValor = "";
								if (stTitulo[0].ToUpper().Trim() == "NUMERO")
									stInsert += " WHERE Numero = '" + stClave +"'";
								else
									stInsert += " WHERE idCompra = " + stClave;
							}
							Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT Compra on", false);
							Funcion.EjecutaSQL(cdsArticulo, stInsert + stValor, false);
						}
						catch (System.Exception ex)
						{
							int f1 = f+1;
							sr.WriteLine(ex.Message + " Linea " + f1.ToString());
						}
						#endregion
					} // Fin de compra no encontrada
					else continue;

					if (bModifica) 
						Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT Compra off", false);

					#region Grabacion de detalle de Compras
					if (bModifica)
					{
						stInsert = "DELETE DetCompra WHERE idCompra = " + stClave;
						Funcion.EjecutaSQL(cdsArticulo, stInsert);
					}

					f ++; // Incrementa linea que deja vacia el programa
					while(f < sheet.Rows.Count && sheet.GetCell(f, 1) != null)
					{
						stInsert = "INSERT INTO DetCompra (idCompra, ";
						stValor = " VALUES (" + stClave + ", ";
						for (int j=1; j<= iCamposDet; j++)
						{

							cell = sheet.GetCell(f, j);
							if (cell != null && cell.Value != null)
							{
								try
								{
									if (stTitulo1[j].ToUpper().Trim() == "CODIGO")
									{
										string stSelect = "Select idArticulo From Articulo Where Codigo = '"
											+ cell.Value.ToString() + "'";
										int iValor = Funcion.iEscalarSQL(cdsArticulo, stSelect, false);
										stInsert += "idArticulo";
										stValor += iValor.ToString("0.00", us);
									}
									else 
									{
										stInsert += stTitulo1[j];
										if (cell.Value.GetType().Name == "String")
										{
											string stCelda = (string) cell.Value;
										
											if (stCelda.ToUpper() == "TRUE" || stCelda.ToUpper() == "FALSE")
											{
												if (stCelda.ToUpper() == "TRUE") stValor += "1";
												else stValor += "0";
											}
											else 
												stValor += "'" + stCelda +"'";
										}
										if (cell.Value.GetType().Name == "DateTime")
										{
											DateTime dtFechaLeida = (DateTime) cell.Value;
											stValor += "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
										}
										if (cell.Value.GetType().Name == "Double")
										{
											double dValor = (double) cell.Value;
											stValor += dValor.ToString("0.00", us);
										}
									} // Fin de no codigo
									if (j!=sheet.Columns.Count -1) 
									{
										stInsert+= ", ";
										stValor += ", ";
									}	
								}
								catch(System.Exception ex)
								{
									int ColError = j + 1;
									int FilaError = f + 1;
									DialogResult drOpcion;
									drOpcion = MessageBox.Show("Error en Fila: " + FilaError.ToString() 
										+ " Columna: " + ColError.ToString() + " " 
										+ ex.Message + "\n¿Desea continuar con la subida de datos?", 
										"Error en Celda", MessageBoxButtons.YesNo,
										MessageBoxIcon.Question);
									if (drOpcion == DialogResult.No)
									{
										this.Cursor = Cursors.Default;
										return;
									}
								} // Fin de Catch
							} // Fin de verificacion de valor nulo
						} // Fin de While
						f++;
						#region Grabacion de DetCompra
						if (stInsert.Trim().EndsWith(","))
						{
							stInsert = stInsert.Substring(0,stInsert.Length - 2);
							stValor = stValor.Substring(0,stValor.Length - 2);
						}
						stInsert+= ") ";
						stValor += ") ";

						try
						{
							//							Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT DetCompra on", false);
							Funcion.EjecutaSQL(cdsArticulo, stInsert + stValor, true);
						}
						catch (System.Exception ex)
						{
							int f1 = f+1;
							sr.WriteLine(ex.Message + " Linea " + f1.ToString());
						}
						#endregion
					} // Fin de DetCompra
					//					f ++; // Ignora linea de titulos
					//					Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT DetCompra off", false);
					#endregion
				} // Fin de Compra Encontrada
			} // Fin de For
			sr.Close();
			MessageBox.Show("Factura Importada", "Información");
			return;
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			#region Verificacion
			if (this.cmbTipoFactura.Value == null)
			{
				MessageBox.Show("Ingrese Tipo de Factura", "Información");
				return;
			}
			string stTipoFactura = this.cmbTipoFactura.Value.ToString();
			if (MessageBox.Show("¿Desea Subir Facturas de Excel?", 
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No)
			{
				MessageBox.Show("Cancelado por el usuario");
				return;
			}
			#endregion Fin de Verificacion

			#region Apertura de Archivo de Excel
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = Application.ExecutablePath;
			if (DialogResult.OK != this.openFileDialog.ShowDialog())
				return;

			this.stArchivo = this.openFileDialog.FileName.ToString();
			this.stArchivo = this.stArchivo.ToUpper();
			C1XLBook book1 = new C1XLBook();
			try
			{
				book1.Load(this.stArchivo);
			}
			catch (Exception exception1)
			{
				MessageBox.Show(exception1.Message);
				return;
			}
			#endregion

			#region Verificacion de Titulos
			XLSheet sheet1 = book1.Sheets[0];
			int num1 = 0;
			C1.Data.C1DataRow drCompra = this.cdsCompra.TableViews["Compra"].AddNew();
			C1.Data.C1DataRow drDetCompra = this.cdsCompra.TableViews["DetCompra"].AddNew();
			string[] stCampo = new string[200];
			XLCell cell1 = sheet1.GetCell(0, 0);
			int iNumColumnas = 0;
			for (int i = 0; i < sheet1.Columns.Count; i++)
			{
				cell1 = sheet1.GetCell(0, i);
				if (cell1 != null)
				{
					try
					{
						stCampo[i] = (string) cell1.Value;
					}
					catch (Exception exception2)
					{
						MessageBox.Show(exception2.Message, "Error al leer titulo");
						return;
					}
					if (stCampo[i].ToUpper().Trim() == "CLIENTE")
					{
						stCampo[i] = "idCliente";
					}
					try
					{
						drCompra[stCampo[i]].GetType();
						iNumColumnas++;
					}
					catch
					{
						MessageBox.Show("No existe campo " + stCampo[i].Trim() 
							+ ".\nPrimera fila debe tener los campos de Compras a importar.", "Error");
						return;
					}
				}
			}
			#endregion Verificacion de Titulos

			#region Verificacion de Fila de Detalle de Compra
			string[] textArray2 = new string[200];
			int iNumColDet = 0;
			cell1 = sheet1.GetCell(3, 0);
			if ((cell1 != null) && (cell1.Value != null))
			{
				MessageBox.Show("Columna A4 debe estar vacia");
			}
			else
			{
				for (int i = 1; i < sheet1.Columns.Count; i++)
				{
					cell1 = sheet1.GetCell(1, i);
					try
					{
						textArray2[i] = (string) cell1.Value;
						if (textArray2[i].ToUpper().Trim() == "CODIGO")
						{
							iNumColDet++;
							continue;
						}
					}
					catch (Exception exception3)
					{
						Console.WriteLine(exception3.Message + "Segunda Fila Titulo");
						break;
					}
					try
					{
						drDetCompra[textArray2[i]].GetType();
						iNumColDet++;
					}
					catch
					{
						MessageBox.Show("No existe campo " + textArray2[i].Trim() 
							+ ".\nSegunda fila debe tener los campos de Detalle de Compras a importar.", 
							"Error");
						return;
					}
				}
				this.cdsCompra.RejectChanges();
				#endregion Verificacion de Fila de Detalle de Compra

				#region Variables de Inicio
				StreamWriter stArchError = File.CreateText("LatiniumError.xls");
				string stClavePrinc = "";
				int IdCompra = 0;
				#endregion Variables de Inicio
				#region Lectura de Filas
				for (int iFila = 1; iFila < sheet1.Rows.Count; iFila++)
				{
					cell1 = sheet1.GetCell(iFila, 0);
					if ((cell1 == null) || (cell1.Value == null)) continue;

					#region Primera columna Clave Principal Verifacion de Existentes
					string stSelect;
					double num8 = 0;
					int iCuentaExistentes = 0;
					if (stCampo[0].Trim().ToUpper() == "IDCOMPRA")
					{
						num8 = (double) cell1.Value;
						stSelect = "SELECT COUNT(*) FROM Compra WHERE idCompra = " 
							+ num8.ToString();
						stClavePrinc = num8.ToString();
					}
					else
					{
						string text4 = cell1.Value.ToString().Trim();
						stSelect = "SELECT COUNT(*) FROM Compra WHERE Numero = '"
							+ text4 + "' And idTipoFactura = " + stTipoFactura;
						stClavePrinc = text4;
					}
					iCuentaExistentes = Funcion.iEscalarSQL(cdsArticulo, stSelect);
					string stInsert = "";
					bool bExiste = false;
					if (iCuentaExistentes > 0)
					{
						if (!this.chkModificaExiste.Checked) continue;
						bExiste = true;
						if (stCampo[0].Trim().ToUpper() != "IDCOMPRA")
						{
							string text6 = cell1.Value.ToString().Trim();
							stSelect = "SELECT idCompra FROM Compra WHERE Numero = '" 
								+ text6 + "' And idTipoFactura = " + stTipoFactura;
							IdCompra = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
						}
					}
					if (bExiste)
					{
						stInsert = "UPDATE Compra SET ";
					}
					else
					{
						stInsert = "INSERT INTO Compra (";
					}
					#endregion Primera columna Clave Principal Verifacion de Existentes

					string stValores = " VALUES (";
					if ((iCuentaExistentes <= 0) || bExiste)
					{
						num1++;
						#region Lectura de Columnas
						for (int iCol = 0; iCol < iNumColumnas; iCol++)
						{
							cell1 = sheet1.GetCell(iFila, iCol);
							if ((cell1 == null) || (cell1.Value == null)) continue;
							try
							{
								stInsert = stInsert + stCampo[iCol];
								#region Caso de String
								if (((cell1.Value.GetType().Name == "String") 
									|| (stCampo[iCol].ToUpper() == "NUMERO")) 
									|| (stCampo[iCol].ToUpper() == "IDCLIENTE"))
								{
									string stCelda = cell1.Value.ToString();
									if (stCampo[iCol].ToUpper() == "IDCLIENTE")
									{
										string text9 = "Select idCliente From Cliente Where Nombre like '" + stCelda.Trim() + "%'";
										int IdCliente = Funcion.iEscalarSQL(this.cdsArticulo, text9, false);
										if (IdCliente == 0)
										{
											int ITipoF = (int) cmbTipoFactura.Value;
											string stProv = "0";
											if (ITipoF == 4 || ITipoF == 2 || ITipoF == 6 || ITipoF == 9
												|| ITipoF == 26 || ITipoF == 25) stProv = "1";
											string stCreaNombre = "Insert Into Cliente (Nombre, Proveedor) Values ('"
												+ stCelda.Trim() + "', " + stProv + ")";
											Funcion.EjecutaSQL(cdsCliente, stCreaNombre, true);
											text9 = "Select idCliente From Cliente Where Nombre like '" + stCelda.Trim() + "%'";
											IdCliente = Funcion.iEscalarSQL(this.cdsArticulo, text9, false);
										}
										stValores = stValores + IdCliente.ToString();
										if (bExiste)
										{
											stInsert = stInsert + " = " + IdCliente.ToString();
										}
									}
									else if ((stCelda.ToUpper() == "TRUE") || (stCelda.ToUpper() == "FALSE"))
									{
										string text10 = "0";
										if (stCelda.ToUpper() == "TRUE")
										{
											text10 = "1";
										}
										stValores = stValores + text10;
										if (bExiste)
										{
											stInsert = stInsert + " = " + text10;
										}
									}
									else
									{
										stValores = stValores + "'" + stCelda + "'";
										if (bExiste)
										{
											stInsert = stInsert + " = '" + stCelda + "'";
										}
									}
								}
									#endregion Caso de String								
									#region Caso de Fecha
								else if (cell1.Value.GetType().Name == "DateTime")
								{
									DateTime time1 = (DateTime) cell1.Value;
									stValores = stValores + "'" + time1.ToString("yyyMMdd") + "'";
									if (bExiste)
									{
										stInsert = stInsert + " = '" + time1.ToString("yyyMMdd") + "'";
									}
								}
									#endregion Caso de Fecha							
									#region Caso de Doble
								else if (cell1.Value.GetType().Name == "Double")
								{
									double num12 = (double) cell1.Value;
									stValores = stValores + num12.ToString("0.0000", this.us);
									if (bExiste)
									{
										stInsert = stInsert + " = " + num12.ToString("0.0000", this.us);
									}
								}
								#endregion Caso de Doble
								if (iCol != (sheet1.Columns.Count - 1))
								{
									stInsert = stInsert + ", ";
									stValores = stValores + ", ";
								}
							}
								#region Manejo de Error
							catch (Exception ex)
							{
								int num13 = iCol + 1;
								int num14 = iFila + 1;
								DialogResult result1 = MessageBox.Show("Error en Fila: " + num14.ToString() + " Columna: " + num13.ToString() + " " + ex.Message 
									+ "\n¿Desea continuar con la subida de datos?", "Error en Celda", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
								if (result1 == DialogResult.No)
								{
									this.Cursor = Cursors.Default;
									return;
								}
							}
							#endregion Manejo de Error
						}
						#region Manejo de Texto
						if (stInsert.Trim().EndsWith(","))
						{
							stInsert = stInsert.Substring(0, stInsert.Length - 2);
							stValores = stValores.Substring(0, stValores.Length - 2);
						}
						if (!bExiste)
						{
							stInsert = stInsert + ", idTipoFactura) ";
						}
						stValores = stValores + ", " + stTipoFactura + ") ";
						#endregion Manejo de Texto
						#endregion Lectura de Columnas
						iFila++;
						#region Grabacion de Compra
						try
						{
							#region Crea filtro de linea
							if (bExiste)
							{
								stValores = "";
								if (stCampo[0].ToUpper().Trim() == "NUMERO")
								{
									stInsert = stInsert + " WHERE Numero = '" + stClavePrinc + "'";
								}
								else
								{
									stInsert = stInsert + " WHERE idCompra = " + stClavePrinc;
									IdCompra = int.Parse(stClavePrinc);
								}
							}
							#endregion Crea filtro de linea
							#region crea fila y devuelve idcompra
							string stCreaFila = stInsert + stValores;
							Funcion.EjecutaSQL(cdsArticulo, stCreaFila, true);
							if (!bExiste)
							{
								stCreaFila = "Select Top 1 idCompra From Compra Order By idCompra DESC";
								IdCompra = Funcion.iEscalarSQL(cdsArticulo, stCreaFila, true);
							}
							#endregion crea fila y devuelve idcompra
						}
						catch (Exception exception5)
						{
							stArchError.WriteLine(exception5.Message + " Linea " 
								+ ((iFila + 1)).ToString());
						}
						if (bExiste)
						{
							Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT Compra off", false);
							stInsert = "DELETE DetCompra WHERE idCompra = " + IdCompra;
							Funcion.EjecutaSQL(cdsArticulo, stInsert);
						}
						#endregion Grabacion de Compra

						#region Crea detalle de compra
						while (((iFila < sheet1.Rows.Count) && ((sheet1.GetCell(iFila, 0) == null
							|| sheet1.GetCell(iFila, 0).Value == null) && (sheet1.GetCell(iFila, 1) != null) 
							&& (sheet1.GetCell(iFila, 1).Value != null))))
						{
							stInsert = "INSERT INTO DetCompra (idCompra, ";
							stValores = " VALUES (" + IdCompra.ToString() + ", ";
							#region Lectura de Columnas de DetCompra
							for (int iColDet = 1; iColDet <= iNumColDet; iColDet++)
							{
								cell1 = sheet1.GetCell(iFila, iColDet);
								if (cell1 == null || cell1.Value == null)
									continue;
								#region Manejo de Celda
								try
								{
									if (textArray2[iColDet].ToUpper().Trim() == "CODIGO")
									{
										string text12 = "Select idArticulo From Articulo Where Codigo = '" + cell1.Value.ToString() + "'";
										int num17 = Funcion.iEscalarSQL(this.cdsArticulo, text12, false);
										stInsert = stInsert + "idArticulo";
										stValores = stValores + num17.ToString("0.0000", this.us);
									}
									else
									{
										stInsert = stInsert + textArray2[iColDet];
										#region Manejo de Texto
										if (cell1.Value.GetType().Name == "String")
										{
											string text13 = (string) cell1.Value;
											if ((text13.ToUpper() == "TRUE") || (text13.ToUpper() == "FALSE"))
											{
												if (text13.ToUpper() == "TRUE")
												{
													stValores = stValores + "1";
												}
												else
												{
													stValores = stValores + "0";
												}
											}
											else
											{
												stValores = stValores + "'" + text13 + "'";
											}
										}
										#endregion Manejo de Texto
										#region Manejo de Fecha
										if (cell1.Value.GetType().Name == "DateTime")
										{
											stValores = stValores + "'" + ((DateTime) cell1.Value).ToString("yyyMMdd") + "'";
										}
										#endregion Manejo de Fecha
										#region Manejo de Numero
										if (cell1.Value.GetType().Name == "Double")
										{
											stValores = stValores + ((double) cell1.Value).ToString("0.00", this.us);
										}
										#endregion Manejo de Numero
									}
									#region Manejo de texto
									if (iColDet != (sheet1.Columns.Count - 1))
									{
										stInsert = stInsert + ", ";
										stValores = stValores + ", ";
									}
									#endregion Manejo de texto
								}
									#region manejo de error
								catch (Exception ex1)
								{
									int num19 = iColDet + 1;
									int num20 = iFila + 1;
									DialogResult result2 = MessageBox.Show("Error en Fila: " 
										+ num20.ToString() + " Columna: " + num19.ToString() + " " 
										+ ex1.Message + "\n¿Desea continuar con la subida de datos?", 
										"Error en Celda", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
									if (result2 == DialogResult.No)
									{
										this.Cursor = Cursors.Default;
										return;
									}
								}
								#endregion manejo de error
								#endregion Manejo de Celda
							}
							iFila++;
							#endregion Lectura de Columnas de DetCompra
							#region Manejo de Texto
							if (stInsert.Trim().EndsWith(","))
							{
								stInsert = stInsert.Substring(0, stInsert.Length - 2);
								stValores = stValores.Substring(0, stValores.Length - 2);
							}
							stInsert = stInsert + ") ";
							stValores = stValores + ") ";
							#endregion Manejo de Texto
							#region Grabacion de DetCompra
							try
							{
								Funcion.EjecutaSQL(cdsArticulo, stInsert + stValores, true);
								continue;
							}
							catch (Exception exception7)
							{
								stArchError.WriteLine(exception7.Message + " Linea " + ((iFila + 1)).ToString());
								continue;
							}
							#endregion Grabacion de DetCompra
						}
						#endregion Crea detalle de compra
						#region Grabacion de Totales
						string stCalculoTotal = "EXEC FacturaTotal " + IdCompra.ToString();
						Funcion.EjecutaSQL(this.cdsArticulo, stCalculoTotal, true);
						#endregion Grabacion de Totales
						iFila--;
					}	
				}
				#endregion Lectura de Filas
				stArchError.Close();
				MessageBox.Show("Factura Importada", "Información");
			}
		}

		private void btVenta_Click(object sender, System.EventArgs e)
		{
			ImportaAnexoT(1);
		}

		private void ImportaAnexoT(int IdTipo)
		{
			#region Creacion de mensaje
			if (IdTipo == 4)
			{
				txtMensaje.Text = "Primera fila de Excel: RUC idCliente idComprobante NumComprobantes "
					+ "SubTotalExcento SubTotalIva Iva Numero Fecha FechaCaducidad AutFactura "
					+ "idCredTributario \n"
					+	"idComprobante: 15 = Documentos Autorizados, 4 = Nota de Credito, "
					+ "5 = nota de Debito, 18 = Comprobante de Venta por reembolso\n"
					+ "Numero : 0010010000006\n"
					+ "idCredTributario: 1 Bienes y Servicios (NO inventario) CON crédito tributario, 2 ... \n";
			}
			else
			{
				txtMensaje.Text = "Primera fila de Excel: RUC idComprobante NumComprobantes SubTotalExcento SubTotalIva Iva\n"
					+	"idComprobante: 15 = Documentos Autorizados, 4 = Nota de Credito, "
					+ "5 = nota de Debito, 18 = Comprobante de Venta por reembolso";
			}
			#endregion Creacion de mensaje
			#region Verificacion
			//			if (this.cmbTipoFactura.Value == null)
			//			{
			//				MessageBox.Show("Ingrese Tipo de Factura", "Información");
			//				return;
			//			}
			//			string stTipoFactura = this.cmbTipoFactura.Value.ToString();
			if (MessageBox.Show("¿Desea Subir Facturas de Excel?", 
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No)
				return;
			#endregion Fin de Verificacion
			#region Apertura de Archivo de Excel
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
			if (DialogResult.OK != this.openFileDialog.ShowDialog())
				return;

			stArchivo = openFileDialog.FileName.ToString();
			stArchivo = stArchivo.ToUpper();
			C1XLBook book1 = new C1XLBook();
			try
			{
				book1.Load(stArchivo);
			}
			catch (Exception exception1)
			{
				MessageBox.Show(exception1.Message);
				return;
			}
			#endregion		
			#region Verificacion de Titulos
			XLSheet sheet1 = book1.Sheets[0];
			int num1 = 0;
			C1.Data.C1DataRow drCompra = cdsCompra.TableViews["Compra"].AddNew();
			C1.Data.C1DataRow drDetCompra = cdsCompra.TableViews["DetCompra"].AddNew();
			string[] stCampo = new string[200];
			XLCell cell1 = sheet1.GetCell(0, 0);
			int iNumColumnas = 0;
			#region Lectura de Titulos
			for (int i = 0; i < sheet1.Columns.Count; i++)
			{
				cell1 = sheet1.GetCell(0, i);
				if (cell1 == null) continue;

				try
				{
					stCampo[i] = cell1.Value.ToString().Trim();
				}
				catch (Exception exception2)
				{
					MessageBox.Show(exception2.Message, "Error al leer titulo");
					return;
				}
				if (stCampo[i].ToUpper().Trim() == "RUC")
				{
					continue;
				}
				if (stCampo[i].ToUpper().Trim() == "IVABIENES")
				{
					continue;
				}
				if (stCampo[i].ToUpper().Trim() == "IVASERVICIOS")
				{
					continue;
				}
				if (stCampo[i].ToUpper().Trim() == "NUMCOMPROBANTES")
				{
					stCampo[i] = "Bodega";
				}
				try
				{
					drCompra[stCampo[i]].GetType();
					iNumColumnas++;
				}
				catch
				{
					MessageBox.Show("No existe campo " + stCampo[i].Trim() 
						+ ".\nPrimera fila debe tener los campos de Compras a importar.", "Error");
					return;
				}
			}
			#endregion Lectura de Titulos
			
			#endregion Verificacion de Titulos
			#region Variables de Inicio
			StreamWriter stArchError = File.CreateText("LatiniumError.xls");
			string stClavePrinc = "";
			string stRuc = "";
			int IdCompra = 0;
			double dIvaBienes = 0;
			double dIvaServicios = 0;
			#endregion Variables de Inicio
			#region Lectura de Filas
			for (int iFila = 1; iFila < sheet1.Rows.Count; iFila++)
			{
				cell1 = sheet1.GetCell(iFila, 0);
				if ((cell1 == null) || (cell1.Value == null)) continue;

				#region Primera columna Clave Principal Verificacion de Existentes
				string stSelect;
				int IdCliente = 0;
				int iCuentaExistentes = 0;
				//				if (stCampo[0].Trim().ToUpper() == "IDCLIENTE")
				//				{
				//					IdCliente = (double) cell1.Value;
				//					stSelect = "SELECT COUNT(*) FROM Compra WHERE idCompra = " 
				//						+ IdCliente.ToString();
				//					stClavePrinc = num8.ToString();
				//				}
				//				iCuentaExistentes = Funcion.iEscalarSQL(cdsArticulo, stSelect);
				string stInsert = "";
				bool bExiste = false;
				//				if (iCuentaExistentes > 0)
				//				{
				//					if (!this.chkModificaExiste.Checked) continue;
				//					bExiste = true;
				//					if (stCampo[0].Trim().ToUpper() != "IDCOMPRA")
				//					{
				//						string text6 = cell1.Value.ToString().Trim();
				//						stSelect = "SELECT idCompra FROM Compra WHERE Numero = '" 
				//							+ text6 + "' And idTipoFactura = " + stTipoFactura;
				//						IdCompra = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
				//					}
				//				}
				//				if (bExiste)
				//				{
				//					stInsert = "UPDATE Compra SET ";
				//				}
				//				else
			{
				stInsert = "INSERT INTO Compra (";
			}
				#endregion Primera columna Clave Principal Verifacion de Existentes

				string stValores = " VALUES (";
				if ((iCuentaExistentes <= 0) || bExiste)
				{
					num1++;
					#region Lectura de Columnas
					double dExento = 0, dSubT = 0;
					int iNumComprobantes = 1;
					for (int iCol = 0; iCol < iNumColumnas+3; iCol++)
					{
						cell1 = sheet1.GetCell(iFila, iCol);
						if ((cell1 == null) || (cell1.Value == null)) continue;
						try
						{
							if(!(stCampo[iCol].ToUpper() == "RUC" || stCampo[iCol].ToUpper() == "IVABIENES"
								|| stCampo[iCol].ToUpper() == "IVASERVICIOS"))
								stInsert = stInsert + stCampo[iCol];
							#region Caso de String
							if (cell1.Value.GetType().Name == "String"
								|| stCampo[iCol].ToUpper() == "BODEGA" 
								|| stCampo[iCol].ToUpper() == "RUC"
								|| stCampo[iCol].ToUpper() == "IDCLIENTE")
							{
								string stCelda = cell1.Value.ToString();
								#region Caso de IdCliente
								if (stCampo[iCol].ToUpper() == "RUC")
								{
									stRuc = stCelda.Trim();
									continue;
								}
								if (stCampo[iCol].ToUpper() == "IDCLIENTE")
								{
//									stRuc = stCelda.Trim();
									if (stRuc.Length < 10)
									{
										MessageBox.Show("Error en RUC O Cedula debe tener 10 o mas digitos", "Información");
										continue;
									}
									string stSelCliente = "Select Count(*) From Cliente Where RUC like '" 
										+ stRuc + "%'";
									int iNumClientes = Funcion.iEscalarSQL(cdsArticulo, stSelCliente, false);
									if (chkRucRep.Checked)
									{
										if (iNumClientes > 1)
										{
											MessageBox.Show("Existe mas de un cliente con ese RUC: " + stRuc, "Información");
											continue;
										}
									}
									if (iNumClientes == 0)
									{
										int iTipoRuc = 6; //Venta con pasapporte
										if (stRuc.Length == 10) iTipoRuc = 5; // Cedula
										if (stRuc.Length == 13) iTipoRuc = 4; // Ruc
										int IProv = 0;
										if (IdTipo == 4) 
										{
											iTipoRuc -= 3;
											IProv = 1;
										}
										stSelCliente = "Insert Into Cliente (RUC, Nombre, idTipoRuc, Proveedor) Values ('"
											+ stRuc + "', '" + stCelda.Trim() + "', " 
											+ iTipoRuc.ToString()+ ", " + IProv.ToString() + ")";
										Funcion.EjecutaSQL(cdsCompra, stSelCliente, true);
									}

									stSelCliente = "Select idCliente From Cliente Where RUC like '" 
										+ stRuc + "%'";
									IdCliente = Funcion.iEscalarSQL(cdsArticulo, stSelCliente, false);
									stValores = stValores + IdCliente.ToString();
									if (bExiste)
									{
										stInsert = stInsert + " = " + IdCliente.ToString();
									}
								}
									#endregion Caso de IdCliente
									#region Caso de Bodega
								else if (stCampo[iCol].ToUpper() == "BODEGA")
								{
									stValores = stValores + int.Parse(stCelda);
									if (bExiste)
									{
										stInsert = stInsert + " = " + int.Parse(stCelda) + "";
									}
								}
									#endregion Caso de Bodega
									#region Otro caso
								else
								{
									stValores = stValores + "'" + stCelda + "'";
									if (bExiste)
									{
										stInsert = stInsert + " = '" + stCelda + "'";
									}
								}
								#endregion Otro caso
							}
								#endregion Caso de String								
								#region Caso de Fecha
							else if (cell1.Value.GetType().Name == "DateTime")
							{
								DateTime time1 = (DateTime) cell1.Value;
								stValores = stValores + "'" + time1.ToString("yyyMMdd") + "'";
								if (bExiste)
								{
									stInsert = stInsert + " = '" + time1.ToString("yyyMMdd") + "'";
								}
							}
								#endregion Caso de Fecha							
								#region Caso de Doble
							else if (cell1.Value.GetType().Name == "Double")
							{
								double dValor = (double) cell1.Value;
								if (stCampo[iCol].ToUpper() == "NUMCOMPROBANTES")
									iNumComprobantes = int.Parse(dValor.ToString());
								if (stCampo[iCol].ToUpper() == "SUBTOTALEXCENTO")
									dExento = dValor;
								if (stCampo[iCol].ToUpper() == "SUBTOTALIVA")
									dSubT = dValor;
								if (stCampo[iCol].ToUpper() == "IVABIENES")
								{
									dIvaBienes = dValor;
									continue;
								}
								if (stCampo[iCol].ToUpper() == "IVASERVICIOS")
								{
									dIvaServicios = dValor;
									continue;
								}

								stValores = stValores + dValor.ToString("0.00", this.us);
								if (bExiste)
								{
									stInsert = stInsert + " = " + dValor.ToString("0.00", this.us);
								}
							}
							#endregion Caso de Doble
							if (iCol != (sheet1.Columns.Count - 1))
							{
								stInsert = stInsert + ", ";
								stValores = stValores + ", ";
							}
						}
							#region Manejo de Error
						catch (Exception exception4)
						{
							int num13 = iCol + 1;
							int num14 = iFila + 1;
							DialogResult result1 = MessageBox.Show("Error en Fila: " + num14.ToString() + " Columna: " + num13.ToString() + " " 
								+ exception4.Message + "\n¿Desea continuar con la subida de datos?", "Error en Celda", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
							if (result1 == DialogResult.No)
							{
								this.Cursor = Cursors.Default;
								return;
							}
						}
						#endregion Manejo de Error
					}
					#region Manejo de Texto
					if (stInsert.Trim().EndsWith(","))
					{
						stInsert = stInsert.Substring(0, stInsert.Length - 2);
						stValores = stValores.Substring(0, stValores.Length - 2);
					}
					if (!bExiste)
					{
						stInsert = stInsert + ", idTipoFactura) ";
					}
					stValores = stValores + ", " + IdTipo.ToString() + ") ";
					#endregion Manejo de Texto
					#endregion Lectura de Columnas
					iFila++;
					#region Grabacion de Compra
					try
					{
						#region Crea filtro de linea
						if (bExiste)
						{
							stValores = "";
							if (stCampo[0].ToUpper().Trim() == "NUMERO")
							{
								stInsert = stInsert + " WHERE Numero = '" + stClavePrinc + "'";
							}
							else
							{
								stInsert = stInsert + " WHERE idCompra = " + stClavePrinc;
								IdCompra = int.Parse(stClavePrinc);
							}
						}
						#endregion Crea filtro de linea
						#region crea fila y devuelve idcompra
						string stCreaFila = stInsert + stValores;
						for (int iNumFact = 0; iNumFact < iNumComprobantes; iNumFact ++)
							Funcion.EjecutaSQL(cdsArticulo, stCreaFila, true);
						if (!bExiste)
						{
							stCreaFila = "Select Top 1 idCompra From Compra Order By idCompra DESC";
							IdCompra = Funcion.iEscalarSQL(cdsArticulo, stCreaFila, true);
						}
						#endregion crea fila y devuelve idcompra
					}
					catch (Exception exception5)
					{
						stArchError.WriteLine(exception5.Message + " Linea " 
							+ ((iFila + 1)).ToString());
					}
					if (bExiste)
					{
						Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT Compra off", false);
						stInsert = "DELETE DetCompra WHERE idCompra = " + IdCompra;
						Funcion.EjecutaSQL(cdsArticulo, stInsert);
					}
					#region Grabacion de Detalle de Compra
					if (dSubT >= .01)
					{
						double dIvaBienes1 = dIvaBienes;
						double dIvaServicios1 = dIvaServicios;
						dIvaBienes = Math.Round(dIvaBienes/.12, 2);
						dIvaServicios = Math.Round(dIvaServicios/.12, 2);
						if (Math.Round(dSubT, 2) != Math.Round(dIvaBienes + dIvaServicios, 2))
						{
							if (dIvaServicios == 0) dIvaBienes = dSubT;
							if (dIvaBienes == 0) dIvaServicios = dSubT;
							double dDiferencia = Math.Round(dSubT, 2) - Math.Round(dIvaBienes + dIvaServicios, 2);
							if (dDiferencia >= .01)
							{
								MessageBox.Show("Diferencia: " + dDiferencia.ToString("0.00") 
									+ " en " + stClavePrinc);
							}
						}

						if (dIvaBienes >= 0.01)
						{
							stSelect = "Insert Into DetCompra (IdCompra, IdArticulo, Precio, Cantidad, "
								+ "Impuesto) Values("	+ IdCompra.ToString() + ", 1, " 
								+ dIvaBienes.ToString("0.00", us) + ", 1, 12)";
							Funcion.EjecutaSQL(cdsCompra, stSelect, true);
						}
						if (dIvaServicios >= 0.01)
						{
							stSelect = "Insert Into DetCompra (IdCompra, IdArticulo, Precio, Cantidad, "
								+ "Impuesto) Values("	+ IdCompra.ToString() + ", 2, " 
								+ dIvaServicios.ToString("0.00", us) + ", 1, 12)";
							Funcion.EjecutaSQL(cdsCompra, stSelect, true);
						}
					}
					if (dExento >= .01)
					{
						stSelect = "Insert Into DetCompra (IdCompra, IdArticulo, Precio, Cantidad, "
							+ "Impuesto) Values("	+ IdCompra.ToString() + ", 2, " 
							+ dExento.ToString("0.00", us) + ", 1, 0)";
						Funcion.EjecutaSQL(cdsCompra, stSelect, true);
					}
					#endregion Grabacion de Detalle de Compra

					#endregion Grabacion de Compra
					#region Grabacion de Totales
					string stCalculoTotal = "EXEC FacturaTotal " + IdCompra.ToString();
					Funcion.EjecutaSQL(this.cdsArticulo, stCalculoTotal, true);
					#endregion Grabacion de Totales
					iFila--;
				}	
			}
			#endregion Lectura de Filas
			#region Fecha de Caducidad
			string stFechaCad = "update compra set fechacaducidad = "
				+ "dateadd(day, -1, dateadd(month, 1, fechacaducidad)) From Compra "
				+ "where datepart(day, fechacaducidad) = 1 ";
			Funcion.EjecutaSQL(cdsCompra, stFechaCad, true);
			#endregion Fecha de Caducidad
			stArchError.Close();
			MessageBox.Show("Archivo de Anexo Transaccional Subido", "Información");		
		}

		private void btCompra_Click(object sender, System.EventArgs e)
		{
			ImportaAnexoT(4);
		}

		private void btCompra_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (txtClave.Text.Trim() != "2243")
			{
				MessageBox.Show("Ingrese Clave", "Información");
				return;
			}
			DateTime dtFechaIni = DateTime.Today;
			dtFechaIni = dtFechaIni.AddDays(-dtFechaIni.Day + 1);
			DateTime dtFechaFin = dtFechaIni.AddMonths(1);
			string stSelect = "Select Count(*) From Compra Where Fecha >= '" 
				+ dtFechaIni.ToString("yyyyMMdd") + "' And Fecha < '" 
				+ dtFechaFin.ToString("yyyyMMdd") + "'";
			int idCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
			if (DialogResult.No == MessageBox.Show("Desea Borrar " + idCuenta 
				+ " Registros?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			stSelect = "Delete Compra Where Fecha >= '" 
				+ dtFechaIni.ToString("yyyyMMdd") + "' And Fecha < '" 
				+ dtFechaFin.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);
		}

		private void btRetencion_Click(object sender, System.EventArgs e)
		{
			#region Creacion de mensaje
				txtMensaje.Text = "Primera fila de Excel: Numero, RUC, Fecha, IvaTotal, "
					+ "IdRetIva, IvaRetenido, IvaTotal1, IdRetIva1, IvaRetenido1 "
					+ ", IngresoEgreso, AutorizaRet, SerieRet, Codigo, Retenido, CodRetRenta, Base, FechaContable";
			#endregion Creacion de mensaje
			#region Verificacion
			if (MessageBox.Show("Desea Subir Retenciones de Excel?", 
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No)
				return;
			#endregion Fin de Verificacion
			#region Apertura de Archivo de Excel
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
			if (DialogResult.OK != this.openFileDialog.ShowDialog())
				return;

			stArchivo = openFileDialog.FileName.ToString();
			stArchivo = stArchivo.ToUpper();
			C1XLBook book1 = new C1XLBook();
			try
			{
				book1.Load(stArchivo);
			}
			catch (Exception exception1)
			{
				MessageBox.Show(exception1.Message);
				return;
			}
			#endregion		
			#region Verificacion de Titulos
			XLSheet sheet1 = book1.Sheets[0];
			int num1 = 0;
			string[] stCampo = new string[200];
			XLCell cell1 = sheet1.GetCell(0, 0);
			#region Lectura de Titulos
			if (sheet1.GetCell(0, 0).Value.ToString().Trim().ToUpper() != "NUMERO")
			{
				MessageBox.Show("Primera Columna debe ser Numero", "Información");
				return;
			}
			if (sheet1.GetCell(0, 1).Value.ToString().Trim().ToUpper() != "RUC")
			{
				MessageBox.Show("Segunda Columna debe ser RUC", "Información");
				return;
			}
			if (sheet1.GetCell(0, 2).Value.ToString().Trim().ToUpper() != "FECHA")
			{
				MessageBox.Show("Tercera Columna debe ser Fecha", "Información");
				return;
			}
			if (sheet1.GetCell(0, 3).Value.ToString().Trim().ToUpper() != "IVATOTAL")
			{
				MessageBox.Show("Cuarta Columna debe ser IvaTotal", "Información");
				return;
			}
			if (sheet1.GetCell(0, 4).Value.ToString().Trim().ToUpper() != "IDRETIVA")
			{
				MessageBox.Show("Quinta Columna debe ser idRetIva", "Información");
				return;
			}
			if (sheet1.GetCell(0, 5).Value.ToString().Trim().ToUpper() != "IVARETENIDO")
			{
				MessageBox.Show("Sexta Columna debe ser IvaRetenido", "Información");
				return;
			}
			if (sheet1.GetCell(0, 6).Value.ToString().Trim().ToUpper() != "IVATOTAL1")
			{
				MessageBox.Show("Septima Columna debe ser IvaTotal1", "Información");
				return;
			}
			if (sheet1.GetCell(0, 7).Value.ToString().Trim().ToUpper() != "IDRETIVA1")
			{
				MessageBox.Show("Octava Columna debe ser IdRetIva1", "Información");
				return;
			}
			if (sheet1.GetCell(0, 8).Value.ToString().Trim().ToUpper() != "IVARETENIDO1")
			{
				MessageBox.Show("Novena Columna debe ser IvaRetenido1", "Información");
				return;
			}
			if (sheet1.GetCell(0, 9).Value.ToString().Trim().ToUpper() != "INGRESOEGRESO")
			{
				MessageBox.Show("Columna 10 debe ser IngresoEgreso", "Información");
				return;
			}
			if (sheet1.GetCell(0, 10).Value.ToString().Trim().ToUpper() != "AUTORIZARET")
			{
				MessageBox.Show("Columna 11 debe ser AutorizaRet", "Información");
				return;
			}
			if (sheet1.GetCell(0, 11).Value.ToString().Trim().ToUpper() != "SERIERET")
			{
				MessageBox.Show("Columna 12 debe ser SerieRet", "Información");
				return;
			}
			if (sheet1.GetCell(0, 12).Value.ToString().Trim().ToUpper() != "CODIGO")
			{
				MessageBox.Show("Columna 13 debe ser Codigo", "Información");
				return;
			}
			if (sheet1.GetCell(0, 13).Value.ToString().Trim().ToUpper() != "RETENIDO")
			{
				MessageBox.Show("Columna 14 debe ser Retenido", "Información");
				return;
			}
			if (sheet1.GetCell(0, 14).Value.ToString().Trim().ToUpper() != "IDRETRENTA")
			{
				MessageBox.Show("Columna 15 debe ser IdRetRenta", "Información");
				return;
			}
			if (sheet1.GetCell(0, 15).Value.ToString().Trim().ToUpper() != "BASE")
			{
				MessageBox.Show("Columna 16 debe ser Base", "Información");
				return;
			}
			#endregion Lectura de Titulos
			
			#endregion Verificacion de Titulos
			#region Variables de Inicio
			StreamWriter stArchError = File.CreateText("LatiniumError.xls");
			int IdCompra = 0;
			#endregion Variables de Inicio
			#region Lectura de Filas
			for (int iFila = 1; iFila < sheet1.Rows.Count; iFila++)
			{
				#region Verificacion de existencia de Celda
				cell1 = sheet1.GetCell(iFila, 0);
				if ((cell1 == null) || (cell1.Value == null)) continue;
				string stNumero = cell1.Value.ToString();
				cell1 = sheet1.GetCell(iFila, 1);
				string stRuc = cell1.Value.ToString();
				#endregion Verificacion de existencia de Celda
				#region Primera columna Clave Principal Verificacion de Existentes
				string stSelect = "Select IdCompra From Compra Inner Join Cliente On " +
					"Compra.idCliente = Cliente.idCliente " +
					"Where Numero = '" + stNumero.Trim() + "' And Ruc = '" + stRuc.Trim() + "'";
				IdCompra = Funcion.iEscalarSQL(cdsArticulo, stSelect, false);
				if (IdCompra == 0) continue;
				// Si ya existe retencion no sube
				stSelect = "Select Count(*) From Retencion " +
					"Where IdCompra = " + IdCompra.ToString();
				int iCuentaRet = Funcion.iEscalarSQL(cdsArticulo, stSelect, false);
				if (iCuentaRet > 0) continue;

				stSelect = "Select Nombre From Cliente Where RUC = '" +
					stRuc.Trim() + "'";
				string stClienteNombre = Funcion.sEscalarSQL(cdsArticulo, stSelect, true);

				stSelect = "Select IdCliente From Cliente Where RUC = '" +
					stRuc.Trim() + "'";
				int IdCliente = Funcion.iEscalarSQL(cdsArticulo, stSelect, false);

				string stInsert = "";
				stInsert = "INSERT INTO Retencion (idCompra, idCliente, Fecha, IvaTotal, "
					+ "idRetIva, IvaRetenido, IvaTotal1, idRetIva1, IvaRetenido1, "
					+ "IngresoEgreso, AutorizaRet, SerieRet, Codigo, RentaRetenida, "
					+ "FechaContable, NumFactura, Beneficiario) ";

				string stValores = " VALUES (" + IdCompra.ToString() + ", "
					+ IdCliente.ToString() + ", ";
				#endregion Primera columna Clave Principal Verificacion de Existentes
				num1++;
				#region Lectura de Columnas
				for (int iCol = 2; iCol < 15; iCol++)
				{
					if (iCol==14) 
						iCol=16;
					cell1 = sheet1.GetCell(iFila, iCol);
					if ((cell1 == null || cell1.Value == null) && iCol < 10) continue;
					try
					{
						#region Caso de String
						if ((cell1.Value.GetType().Name == "String" && iCol != 17 && iCol != 2)
							|| iCol == 10 || iCol == 11 || iCol == 12)
						{
							string stCelda = "";
							if (cell1 != null && cell1.Value != null) 
								stCelda = cell1.Value.ToString();
							stValores = stValores + "'" + stCelda + "'";
						}
							#endregion Caso de String								
						#region Caso de Fecha
						else if (cell1.Value.GetType().Name == "DateTime" || iCol == 17 || iCol == 2)
						{
							DateTime time1 = DateTime.Parse(cell1.Value.ToString());
							//							DateTime time1 = (DateTime) cell1.Value;
							stValores = stValores + "'" + time1.ToString("yyyMMdd") + "'";
						}
							#endregion Caso de Fecha							
						#region Caso de Doble
						else if (cell1.Value.GetType().Name == "Double")
						{
							double dValor = (double) cell1.Value;
//							if (iCol == 4 || iCol == 7)
//							{
//								dValor += 1;	
//							}
							stValores = stValores + dValor.ToString("0.00", this.us);
						}
						#endregion Caso de Doble
						if (iCol != (sheet1.Columns.Count - 1))
						{
							stValores = stValores + ", ";
						}
					}
					#region Manejo de Error
					catch (Exception exception4)
					{
						int num13 = iCol + 1;
						int num14 = iFila + 1;
						DialogResult result1 = MessageBox.Show("Error en Fila: " + num14.ToString() + " Columna: " + num13.ToString() + " " + exception4.Message 
							+ "\n¿Desea continuar con la subida de datos?", "Error en Celda", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
						if (result1 == DialogResult.No)
						{
							this.Cursor = Cursors.Default;
							return;
						}
					}
					#endregion Manejo de Error
				}
				#endregion Lectura de Columnas
				if (stValores.Trim().EndsWith(","))
					stValores = stValores.Substring(0, stValores.Length - 2);
				#region Grabacion de Retencion
				stSelect = stInsert + stValores + ", '" + stNumero + "', '" 
					+ stClienteNombre + "')";
				Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
				stSelect = "Select Top 1 IdRetencion From Retencion Order By IdRetencion DESC";
				double dRetenido = (double) sheet1.GetCell(iFila, 13).Value;
				try
				{
					if (dRetenido > 0 && sheet1.GetCell(iFila, 14) != null)
					{
						int IdRetencion = Funcion.iEscalarSQL(cdsArticulo, stSelect, false);
						string stRetRenta = sheet1.GetCell(iFila, 14).Value.ToString();
						stSelect = "Select idArticuloSRI From ArticuloSRI Where Codigo = '"
							+ stRetRenta + "'";
						int IdRetRenta = Funcion.iEscalarSQL(cdsArticulo, stSelect, false);
						double dBase = (double) sheet1.GetCell(iFila, 15).Value;
						stSelect = "Insert Into RetDetalle(idRetencion, idRetRenta, Base, Retenido) "
							+ " Values (" + IdRetencion.ToString() + ", " + IdRetRenta.ToString() + ", "
							+ dBase.ToString("0.00", us) + ", " + dRetenido.ToString("0.00", us) + ")";
						Funcion.EjecutaSQL(cdsArticulo, stSelect, false);

						#region 2da retencion si numero y ruc coinciden
						dRetenido = (double) sheet1.GetCell(iFila + 1, 13).Value;
						if (stNumero.Trim() == sheet1.GetCell(iFila + 1, 0).Value.ToString().Trim()
							&& sheet1.GetCell(iFila + 1, 1).Value.ToString() == stRuc)
						{
							iFila ++;
							dRetenido = (double) sheet1.GetCell(iFila, 13).Value;
							stRetRenta = sheet1.GetCell(iFila, 14).Value.ToString();
							stSelect = "Select idArticuloSRI From ArticuloSRI Where Codigo = '"
								+ stRetRenta + "'";
							IdRetRenta = Funcion.iEscalarSQL(cdsArticulo, stSelect, false);
							dBase = (double) sheet1.GetCell(iFila, 15).Value;
							stSelect = "Insert Into RetDetalle(idRetencion, idRetRenta, Base, Retenido) "
								+ " Values (" + IdRetencion.ToString() + ", " + IdRetRenta.ToString() + ", "
								+ dBase.ToString("0.00", us) + ", " + dRetenido.ToString("0.00", us) + ")";
							Funcion.EjecutaSQL(cdsArticulo, stSelect, false);
						}
						#endregion 2da retencion si numero y ruc coinciden
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error en RetencionDetalle");
				}
				#endregion Grabacion de Retencion
			}
			#endregion Lectura de Filas
			#region Actualizaciones Globales
			string stUpdate = "Update Retencion Set PorcIva = 12 Where IsNull(PorcIva, 0) = 0";
			Funcion.EjecutaSQL(cdsArticulo, stUpdate, true);
			stUpdate = "Update Retencion Set FacturaIva12 = IvaTotal / .12 Where IsNull(FacturaIva12, 0) = 0";
			Funcion.EjecutaSQL(cdsArticulo, stUpdate, true);
			stUpdate = "Update Retencion Set FacturaIva121 = IvaTotal1 / .12 Where IsNull(FacturaIva121, 0) = 0";
			Funcion.EjecutaSQL(cdsArticulo, stUpdate, true);
			stUpdate = "Update Retencion Set FacturaIva121 = IvaTotal1 / .12 Where IsNull(FacturaIva121, 0) = 0";
			Funcion.EjecutaSQL(cdsArticulo, stUpdate, true);
			stUpdate = "Update Retencion Set RentaRetenida = "
				+ "(Select Sum(rd.Retenido) from RetDetalle rd Where rd.idRetencion = Retencion.idRetencion)";
			Funcion.EjecutaSQL(cdsArticulo, stUpdate, true);

			Funcion.EjecutaSQL(cdsArticulo, stUpdate, true);

			#endregion Actualizaciones Globales
			stArchError.Close();
			MessageBox.Show("Archivo de Retenciones de Anexo Transaccional Subido", "Información");		
		}

		private void btFacturaLinea_Click(object sender, System.EventArgs e)
		{
			#region Verificacion
			if (this.cmbTipoFactura.Value == null)
			{
				MessageBox.Show("Ingrese Tipo de Factura", "Información");
				return;
			}
			string stTipoFactura = this.cmbTipoFactura.Value.ToString();
			if (MessageBox.Show("¿Desea Subir Facturas de Excel?", 
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No)
			{
				MessageBox.Show("Cancelado por el usuario");
				return;
			}
			#endregion Fin de Verificacion
			#region Creacion de mensaje
			txtMensaje.Text = "Primera fila de Excel: NUMERO, FECHA, CODARTICULO, CODCLIENTE, VALOR, IVA, COMPROBANTE, ContadoCredito, DiasCredito"
				+ "\nColumnas CodArticulo y CodCliente. Deben ser Texto. Comprobante el campo Otro. Contado = 1 Credito = 2";
			if (MessageBox.Show("Desea Subir Facturas por Linea de Excel de Archivo? \nEste Archivo debe Estar Cerrado", 
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No)
				return;
			#endregion Creacion de mensaje

			#region Apertura de Archivo de Excel
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
			if (DialogResult.OK != this.openFileDialog.ShowDialog())
				return;

			string stArchivo = openFileDialog.FileName.ToString();
			stArchivo = stArchivo.ToUpper();
			C1XLBook book1 = new C1XLBook();
			try
			{
				book1.Load(stArchivo);
			}
			catch (Exception exception1)
			{
				MessageBox.Show(exception1.Message);
				return;
			}
			#endregion		
			#region Lectura de Titulos
			XLSheet sheet1 = book1.Sheets[0];
			if (sheet1.GetCell(0, 0).Value.ToString().ToUpper().Trim() != "NUMERO")
			{
				MessageBox.Show("Primera Columna debe ser Numero", "Información");
				return;
			}
			if (sheet1.GetCell(0, 1).Value.ToString().ToUpper().Trim() != "FECHA")
			{
				MessageBox.Show("Segunda Columna debe ser Fecha", "Información");
				return;
			}
			if (sheet1.GetCell(0, 2).Value.ToString().ToUpper().Trim() != "CODARTICULO")
			{
				MessageBox.Show("Tercera Columna debe ser CodArticulo", "Información");
				return;
			}
			if (sheet1.GetCell(0, 3).Value.ToString().ToUpper().Trim() != "CODCLIENTE")
			{
				MessageBox.Show("Cuarta Columna debe ser CodCliente", "Información");
				return;
			}
			if (sheet1.GetCell(0, 4).Value.ToString().ToUpper().Trim() != "VALOR")
			{
				MessageBox.Show("Quinta Columna debe ser Valor", "Información");
				return;
			}
			if (sheet1.GetCell(0, 5).Value.ToString().ToUpper().Trim() != "IVA")
			{
				MessageBox.Show("Sexta Columna debe ser Iva", "Información");
				return;
			}
			if (sheet1.GetCell(0, 6).Value.ToString().ToUpper().Trim() != "COMPROBANTE")
			{
				MessageBox.Show("Septima Columna debe ser Comprobante", "Información");
				return;
			}
			if (sheet1.GetCell(0, 7).Value.ToString().ToUpper().Trim() != "CONTADOCREDITO")
			{
				MessageBox.Show("Octava Columna debe ser ContadoCredito", "Información");
				return;
			}
			if (sheet1.GetCell(0, 8).Value.ToString().ToUpper().Trim() != "DIASCREDITO")
			{
				MessageBox.Show("Novena Columna debe ser DiasCredito", "Información");
				return;
			}
	
			#endregion Lectura de Titulos
			#region Lectura de Filas
			for (int iFila = 1; iFila < sheet1.Rows.Count; iFila++)
			{
				try
				{
					#region Verificacion de existencia de Celda
					XLCell cell1 = sheet1.GetCell(iFila, 0);
					if ((cell1 == null) || (cell1.Value == null)) continue;
					string stNumero = cell1.Value.ToString();
					cell1 = sheet1.GetCell(iFila, 3);
					string stRuc = cell1.Value.ToString();
					#endregion Verificacion de existencia de Celda
					#region Primera columna Clave Principal Verificacion de Existentes
					string stSelect = "Select IdCompra From Compra Inner Join Cliente On " +
						"Compra.idCliente = Cliente.idCliente " +
						"Where Compra.Numero = '" + stNumero.Trim() + "' And Cliente.Ruc = '" + stRuc.Trim() + "'";
					int IdCompra = Funcion.iEscalarSQL(cdsArticulo, stSelect, false);
					if (IdCompra > 0) continue; // Factura ya Existe
					stSelect = "Select IdCliente From Cliente Where Codigo = '" +
						stRuc.Trim() + "'";
					int IdCliente = Funcion.iEscalarSQL(cdsArticulo, stSelect, false);
					cell1 = sheet1.GetCell(iFila, 1);
					DateTime dtFecha = (DateTime) cell1.Value;
					cell1 = sheet1.GetCell(iFila, 2);
					string stCodArticulo = cell1.Value.ToString();
					cell1 = sheet1.GetCell(iFila, 4);
					double dValor = (double) cell1.Value;
					cell1 = sheet1.GetCell(iFila, 5);
					double dIva = (double) cell1.Value;
					cell1 = sheet1.GetCell(iFila, 6);
					string stComprobante = cell1.Value.ToString().Trim();
					cell1 = sheet1.GetCell(iFila, 7);
					int iContadoCredito = int.Parse(cell1.Value.ToString());
					cell1 = sheet1.GetCell(iFila, 8);
					int iDiasCredito = int.Parse(cell1.Value.ToString());

					int iCentroCosto = 0;
					int iSubProyecto = 0;
					string stFinInsertC = ")";
					string stFinInsertV = ")";
				
					string stInsert = "Insert Into Compra (Numero, Fecha, idTipoFactura, idCliente, "
						+ "Otro, ContadoCredito, DiasCredito" + stFinInsertC
						+ "	Values ('" + stNumero + "', '" + dtFecha.ToString("yyyyMMdd") + "', " 
						+ stTipoFactura + ", " + IdCliente + ", '" + stComprobante + "', " 
						+ iContadoCredito.ToString() + ", " + iDiasCredito.ToString() + stFinInsertV;
					Funcion.EjecutaSQL(cdsArticulo, stInsert, true);

					stInsert = "Select top 1 IdCompra From Compra Order By IdCompra Desc ";
					IdCompra = Funcion.iEscalarSQL(cdsArticulo, stInsert, true);

					stInsert = "Select idArticulo From Articulo Where Codigo = '" + stCodArticulo + "'";
					int IdArticulo = Funcion.iEscalarSQL(cdsArticulo, stInsert, true);

					stInsert = "Insert Into DetCompra (IdCompra, Cantidad, Impuesto, Precio, IdArticulo) "
						+ "	Values (" + IdCompra.ToString() + ", 1, " + dIva.ToString("0.0000", us) + ", " 
						+ dValor.ToString("0.0000", us) + ", "  + IdArticulo.ToString() + ")";
					Funcion.EjecutaSQL(cdsArticulo, stInsert, true);

					stInsert = "Exec FacturaTotal " + IdCompra.ToString();
					Funcion.EjecutaSQL(cdsArticulo, stInsert, true);
					#endregion Primera columna Clave Principal Verificacion de Existentes
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error Linea " + iFila.ToString());
				}
			}
//			stArchError.Close();
			#endregion Lectura de Filas
			
			MessageBox.Show("Facturas por Linea Subidas");
			Cursor = Cursors.Default;
		}

		private void RestauraProcedimiento(string Procedimiento)
		{
			string stEmpresa = cmbEmpresa.Value.ToString().Trim();
			string stExec = string.Format("EXEC sp_helptext '{0}'", Procedimiento);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsCompra, stExec, true);

			string stArchivo = "";
			while (dr.Read())
			{
				string stLinea = dr.GetString(0);
				stLinea = stLinea.ToUpper().Replace("CREATE PROCEDURE", "ALTER PROCEDURE");
				stLinea = stLinea.Replace(stEmpresa, "SUPERTODOEJEMPLO");
				stArchivo += stLinea;
			}
			dr.Close();
			Funcion.EjecutaSQL(cdsCompra, stArchivo);
		}

		private void ModificaProcedimiento(string Procedimiento)
		{
			string stEmpresa = cmbEmpresa.Value.ToString().Trim();
			string stExec = string.Format("EXEC sp_helptext '{0}'", Procedimiento);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsCompra, stExec, true);

			string stArchivo = "";
			while (dr.Read())
			{
				string stLinea = dr.GetString(0);
				stLinea = stLinea.ToUpper().Replace("CREATE PROCEDURE", "ALTER PROCEDURE");
				stLinea = stLinea.Replace("SUPERTODOEJEMPLO", stEmpresa);
				stArchivo += stLinea;
			}
			dr.Close();
			Funcion.EjecutaSQL(cdsCompra, stArchivo);
		}

		private void btFacturacion_Click(object sender, System.EventArgs e)
		{
			if (!Verifica(1)) return;

			ModificaProcedimiento("FacturaSincroniza");
			int iCuenta = Funcion.iEscalarSQL(cdsCompra, "Exec FacturaSincroniza " + cmbTipoFactura.Value);
			RestauraProcedimiento("FacturaSincroniza");
			MessageBox.Show(string.Format("{0} Datos Sincronizados", iCuenta), "Información");
		}

		private bool Verifica(int iTipo)
		{
			errorProvider1.SetError(cmbEmpresa, "");
			errorProvider1.SetError(cmbTipoFactura, "");
			bool bOk = true;
			if (cmbEmpresa.Value == null) 
			{
				errorProvider1.SetError(cmbEmpresa, "Ingrese Empresa");
				bOk = false;
			}
			if (iTipo == 1 && cmbTipoFactura.Value == null) 
			{
				errorProvider1.SetError(cmbTipoFactura, "Ingrese Tipo de Factura");
				bOk = false;
			}
			return bOk;
		}

		private void btContabilidad_Click(object sender, System.EventArgs e)
		{
			if (!Verifica(0)) return;
			ModificaProcedimiento("SincronizaAsiento");
			int iCuenta = 0;//Funcion.iEscalarSQL(cdsCompra, "Exec SincronizaAsiento");
			RestauraProcedimiento("SincronizaAsiento");
			MessageBox.Show(string.Format("{0} Datos Sincronizados", iCuenta), "Información");
		}

		private void btFacturación_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}