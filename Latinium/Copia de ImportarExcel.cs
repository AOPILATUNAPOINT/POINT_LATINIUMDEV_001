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
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ImportarExcel.
	/// </summary>
	public class ImportarExcel : DevExpress.XtraEditors.XtraForm
	{
		private C1.C1Excel.C1XLBook c1XLBook;
		private C1.Data.C1DataSet cdsCuentas;
		private C1.Data.C1DataView cdvCuentas;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private Infragistics.Win.Misc.UltraButton btAsientos;
		private C1.Data.C1DataSet cdsAsiento;
		private C1.Data.C1DataView cdvAsiento;
		private Infragistics.Win.Misc.UltraButton btArticulo;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataSet cdsTipoCompra;
		private C1.Data.C1DataView cdvTipoCompra;
		private Infragistics.Win.Misc.UltraButton btClientes;
		private Infragistics.Win.Misc.UltraButton btProveedor;
		private Infragistics.Win.Misc.UltraButton btInventario;
		private Infragistics.Win.Misc.UltraButton btActivoFijo;
		private Infragistics.Win.Misc.UltraButton btCuentasxC;
		private Infragistics.Win.Misc.UltraButton btCuentasxP;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactura;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private Infragistics.Win.Misc.UltraButton btCuentas;
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataView cdvCompra;
		private C1.Data.C1DataSet cdsGrupoArt;
		private Infragistics.Win.Misc.UltraButton btGrupoArt;
		private C1.Data.C1DataView cdvGrupo;
		private Infragistics.Win.Misc.UltraButton btGrupoCliente;
		private C1.Data.C1DataSet cdsGrupoCliente;
		private C1.Data.C1DataView cdvClienteGrupo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMensaje;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupoArt;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupoCliente;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkModificaExiste;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkClavePrimaria;
		private Infragistics.Win.Misc.UltraButton btCargo;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataSet cdsCargo;
		private C1.Data.C1DataView cdvPersonal;
		private C1.Data.C1DataView cdvCargo;
		private Infragistics.Win.Misc.UltraButton btDepartamento;
		private C1.Data.C1DataView cdvDepartamento;
		private Infragistics.Win.Misc.UltraButton btPersonal;
		private Infragistics.Win.Misc.UltraButton btCarro;
		private C1.Data.C1DataView cdvCarros;
		private C1.Data.C1DataSet cdsCarros;
		private Infragistics.Win.Misc.UltraButton btOrdenMov;
		private C1.Data.C1DataView cdvOrdenMov;
		private Infragistics.Win.Misc.UltraButton btSucursal;
		private C1.Data.C1DataView cdvSucursal;
		private Infragistics.Win.Misc.UltraButton btMarca;
		private Infragistics.Win.Misc.UltraButton btArancel;
		private C1.Data.C1DataView cdvMarca;
		private C1.Data.C1DataView cdvArancel;
		private Infragistics.Win.Misc.UltraButton btUnidad;
		private C1.Data.C1DataView cdvUnidad;
		private Infragistics.Win.Misc.UltraButton btInvSaldo;
		private C1.Data.C1DataSet cdsPagos;
		private Infragistics.Win.Misc.UltraButton btRolConcepto;
		private C1.Data.C1DataSet cdsRolConcepto;
		private C1.Data.C1DataView cdvRolConcepto;
		private Infragistics.Win.Misc.UltraButton btTomaFisica;
		private C1.Data.C1DataSet cdsToma;
		private C1.Data.C1DataView cdvTomaFisica;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private C1.Data.C1DataView cdvRolIngreso;
		private C1.Data.C1DataSet cdsRolIngreso;
		private Infragistics.Win.Misc.UltraButton btPagos;
		private Infragistics.Win.Misc.UltraButton btCobros;
		private Infragistics.Win.Misc.UltraButton btAvanzado;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ImportarExcel()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas8");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			this.c1XLBook = new C1.C1Excel.C1XLBook();
			this.btCuentas = new Infragistics.Win.Misc.UltraButton();
			this.cdsCuentas = new C1.Data.C1DataSet();
			this.cdvCuentas = new C1.Data.C1DataView();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.btAsientos = new Infragistics.Win.Misc.UltraButton();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.cdvAsiento = new C1.Data.C1DataView();
			this.btArticulo = new Infragistics.Win.Misc.UltraButton();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.btClientes = new Infragistics.Win.Misc.UltraButton();
			this.btProveedor = new Infragistics.Win.Misc.UltraButton();
			this.btInventario = new Infragistics.Win.Misc.UltraButton();
			this.btActivoFijo = new Infragistics.Win.Misc.UltraButton();
			this.btCuentasxC = new Infragistics.Win.Misc.UltraButton();
			this.btCuentasxP = new Infragistics.Win.Misc.UltraButton();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsTipoCompra = new C1.Data.C1DataSet();
			this.cdvTipoCompra = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cdvCompra = new C1.Data.C1DataView();
			this.cdsGrupoArt = new C1.Data.C1DataSet();
			this.btGrupoArt = new Infragistics.Win.Misc.UltraButton();
			this.cdvGrupo = new C1.Data.C1DataView();
			this.cmbGrupoArt = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbGrupoCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsGrupoCliente = new C1.Data.C1DataSet();
			this.btGrupoCliente = new Infragistics.Win.Misc.UltraButton();
			this.cdvClienteGrupo = new C1.Data.C1DataView();
			this.txtMensaje = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkModificaExiste = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkClavePrimaria = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btCargo = new Infragistics.Win.Misc.UltraButton();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdsCargo = new C1.Data.C1DataSet();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdvCargo = new C1.Data.C1DataView();
			this.btDepartamento = new Infragistics.Win.Misc.UltraButton();
			this.cdvDepartamento = new C1.Data.C1DataView();
			this.btPersonal = new Infragistics.Win.Misc.UltraButton();
			this.btCarro = new Infragistics.Win.Misc.UltraButton();
			this.cdvCarros = new C1.Data.C1DataView();
			this.cdsCarros = new C1.Data.C1DataSet();
			this.btOrdenMov = new Infragistics.Win.Misc.UltraButton();
			this.cdvOrdenMov = new C1.Data.C1DataView();
			this.btSucursal = new Infragistics.Win.Misc.UltraButton();
			this.cdvSucursal = new C1.Data.C1DataView();
			this.btMarca = new Infragistics.Win.Misc.UltraButton();
			this.btArancel = new Infragistics.Win.Misc.UltraButton();
			this.cdvMarca = new C1.Data.C1DataView();
			this.cdvArancel = new C1.Data.C1DataView();
			this.btUnidad = new Infragistics.Win.Misc.UltraButton();
			this.cdvUnidad = new C1.Data.C1DataView();
			this.btInvSaldo = new Infragistics.Win.Misc.UltraButton();
			this.cdsPagos = new C1.Data.C1DataSet();
			this.btRolConcepto = new Infragistics.Win.Misc.UltraButton();
			this.cdsRolConcepto = new C1.Data.C1DataSet();
			this.cdvRolConcepto = new C1.Data.C1DataView();
			this.btTomaFisica = new Infragistics.Win.Misc.UltraButton();
			this.cdsToma = new C1.Data.C1DataSet();
			this.cdvTomaFisica = new C1.Data.C1DataView();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.cdvRolIngreso = new C1.Data.C1DataView();
			this.cdsRolIngreso = new C1.Data.C1DataSet();
			this.btPagos = new Infragistics.Win.Misc.UltraButton();
			this.btCobros = new Infragistics.Win.Misc.UltraButton();
			this.btAvanzado = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvClienteGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMensaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCarros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCarros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvOrdenMov)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArancel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvUnidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsToma)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTomaFisica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolIngreso)).BeginInit();
			this.SuspendLayout();
			// 
			// btCuentas
			// 
			this.btCuentas.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCuentas.Location = new System.Drawing.Point(16, 17);
			this.btCuentas.Name = "btCuentas";
			this.btCuentas.Size = new System.Drawing.Size(96, 25);
			this.btCuentas.TabIndex = 0;
			this.btCuentas.Text = "Plan de Cuentas";
			this.btCuentas.Click += new System.EventHandler(this.btCuentas_Click);
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
			this.cdvCuentas.Sort = "idCuenta";
			this.cdvCuentas.TableName = "";
			this.cdvCuentas.TableViewName = "Cuenta";
			// 
			// openFileDialog
			// 
			this.openFileDialog.DefaultExt = "xls";
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			// 
			// btAsientos
			// 
			this.btAsientos.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAsientos.Location = new System.Drawing.Point(136, 17);
			this.btAsientos.Name = "btAsientos";
			this.btAsientos.Size = new System.Drawing.Size(96, 25);
			this.btAsientos.TabIndex = 2;
			this.btAsientos.Text = "Asientos";
			this.btAsientos.Click += new System.EventHandler(this.btAsientos_Click);
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
			// btArticulo
			// 
			this.btArticulo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btArticulo.Location = new System.Drawing.Point(16, 60);
			this.btArticulo.Name = "btArticulo";
			this.btArticulo.Size = new System.Drawing.Size(96, 25);
			this.btArticulo.TabIndex = 5;
			this.btArticulo.Text = "Artículos";
			this.btArticulo.Click += new System.EventHandler(this.btArticulo_Click);
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
			// btClientes
			// 
			this.btClientes.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btClientes.Location = new System.Drawing.Point(256, 60);
			this.btClientes.Name = "btClientes";
			this.btClientes.Size = new System.Drawing.Size(96, 25);
			this.btClientes.TabIndex = 6;
			this.btClientes.Text = "Clientes";
			this.btClientes.Click += new System.EventHandler(this.btClientes_Click);
			// 
			// btProveedor
			// 
			this.btProveedor.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btProveedor.Location = new System.Drawing.Point(376, 60);
			this.btProveedor.Name = "btProveedor";
			this.btProveedor.Size = new System.Drawing.Size(96, 25);
			this.btProveedor.TabIndex = 7;
			this.btProveedor.Text = "Proveedores";
			this.btProveedor.Click += new System.EventHandler(this.btProveedor_Click);
			// 
			// btInventario
			// 
			this.btInventario.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btInventario.Location = new System.Drawing.Point(256, 138);
			this.btInventario.Name = "btInventario";
			this.btInventario.Size = new System.Drawing.Size(96, 26);
			this.btInventario.TabIndex = 8;
			this.btInventario.Text = "Facturas";
			this.btInventario.Click += new System.EventHandler(this.btInventario_Click);
			// 
			// btActivoFijo
			// 
			this.btActivoFijo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btActivoFijo.Location = new System.Drawing.Point(136, 60);
			this.btActivoFijo.Name = "btActivoFijo";
			this.btActivoFijo.Size = new System.Drawing.Size(96, 26);
			this.btActivoFijo.TabIndex = 9;
			this.btActivoFijo.Text = "Activo Fijo";
			this.btActivoFijo.Click += new System.EventHandler(this.btActivoFijo_Click);
			// 
			// btCuentasxC
			// 
			this.btCuentasxC.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCuentasxC.Location = new System.Drawing.Point(256, 304);
			this.btCuentasxC.Name = "btCuentasxC";
			this.btCuentasxC.Size = new System.Drawing.Size(96, 26);
			this.btCuentasxC.TabIndex = 10;
			this.btCuentasxC.Text = "Cuentas x P";
			this.btCuentasxC.Visible = false;
			this.btCuentasxC.Click += new System.EventHandler(this.btCuentaxC_Click);
			// 
			// btCuentasxP
			// 
			this.btCuentasxP.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCuentasxP.Location = new System.Drawing.Point(136, 138);
			this.btCuentasxP.Name = "btCuentasxP";
			this.btCuentasxP.Size = new System.Drawing.Size(96, 26);
			this.btCuentasxP.TabIndex = 11;
			this.btCuentasxP.Text = "Cuentas x C";
			this.btCuentasxP.Click += new System.EventHandler(this.btCuentasxP_Click);
			// 
			// cmbTipoFactura
			// 
			this.cmbTipoFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipoFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoFactura.DisplayMember = "Nombre";
			this.cmbTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFactura.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoFactura.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbTipoFactura.Location = new System.Drawing.Point(376, 138);
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(96, 22);
			this.cmbTipoFactura.TabIndex = 12;
			this.cmbTipoFactura.ValueMember = "idTipoFactura";
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
			// btGrupoArt
			// 
			this.btGrupoArt.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGrupoArt.Location = new System.Drawing.Point(16, 181);
			this.btGrupoArt.Name = "btGrupoArt";
			this.btGrupoArt.Size = new System.Drawing.Size(96, 26);
			this.btGrupoArt.TabIndex = 13;
			this.btGrupoArt.Text = "Grupo Artículos";
			this.btGrupoArt.Click += new System.EventHandler(this.btGrupoArt_Click);
			// 
			// cdvGrupo
			// 
			this.cdvGrupo.BindingContextCtrl = this;
			this.cdvGrupo.DataSet = this.cdsGrupoArt;
			this.cdvGrupo.Sort = "Grupo";
			this.cdvGrupo.TableName = "";
			this.cdvGrupo.TableViewName = "ArticuloGrupo";
			// 
			// cmbGrupoArt
			// 
			this.cmbGrupoArt.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupoArt.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupoArt.DataMember = "ArticuloGrupo";
			this.cmbGrupoArt.DataSource = this.cdsGrupoArt;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn4.Header.VisiblePosition = 7;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 9;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 12;
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
																										 ultraGridColumn15});
			this.cmbGrupoArt.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbGrupoArt.DisplayMember = "Grupo";
			this.cmbGrupoArt.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupoArt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupoArt.Location = new System.Drawing.Point(136, 181);
			this.cmbGrupoArt.Name = "cmbGrupoArt";
			this.cmbGrupoArt.Size = new System.Drawing.Size(96, 21);
			this.cmbGrupoArt.TabIndex = 14;
			this.cmbGrupoArt.ValueMember = "";
			// 
			// cmbGrupoCliente
			// 
			this.cmbGrupoCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupoCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupoCliente.DataMember = "ClienteGrupo";
			this.cmbGrupoCliente.DataSource = this.cdsGrupoCliente;
			ultraGridBand3.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 112;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Width = 260;
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.cmbGrupoCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbGrupoCliente.DisplayMember = "Grupo";
			this.cmbGrupoCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupoCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupoCliente.Location = new System.Drawing.Point(376, 181);
			this.cmbGrupoCliente.Name = "cmbGrupoCliente";
			this.cmbGrupoCliente.Size = new System.Drawing.Size(96, 21);
			this.cmbGrupoCliente.TabIndex = 16;
			this.cmbGrupoCliente.ValueMember = "";
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
			this.cdsGrupoCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupoCliente_BeforeFill);
			// 
			// btGrupoCliente
			// 
			this.btGrupoCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGrupoCliente.Location = new System.Drawing.Point(256, 181);
			this.btGrupoCliente.Name = "btGrupoCliente";
			this.btGrupoCliente.Size = new System.Drawing.Size(96, 26);
			this.btGrupoCliente.TabIndex = 15;
			this.btGrupoCliente.Text = "Grupo Clientes";
			this.btGrupoCliente.Click += new System.EventHandler(this.btGrupoCliente_Click);
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
			this.txtMensaje.Location = new System.Drawing.Point(16, 336);
			this.txtMensaje.Multiline = true;
			this.txtMensaje.Name = "txtMensaje";
			this.txtMensaje.Size = new System.Drawing.Size(456, 78);
			this.txtMensaje.TabIndex = 17;
			// 
			// chkModificaExiste
			// 
			this.chkModificaExiste.Location = new System.Drawing.Point(256, 17);
			this.chkModificaExiste.Name = "chkModificaExiste";
			this.chkModificaExiste.Size = new System.Drawing.Size(88, 35);
			this.chkModificaExiste.TabIndex = 18;
			this.chkModificaExiste.Text = "Modificar Existentes";
			// 
			// chkClavePrimaria
			// 
			this.chkClavePrimaria.Location = new System.Drawing.Point(360, 17);
			this.chkClavePrimaria.Name = "chkClavePrimaria";
			this.chkClavePrimaria.Size = new System.Drawing.Size(112, 35);
			this.chkClavePrimaria.TabIndex = 19;
			this.chkClavePrimaria.Text = "Ingreso de Clave Principal";
			// 
			// btCargo
			// 
			this.btCargo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCargo.Location = new System.Drawing.Point(16, 224);
			this.btCargo.Name = "btCargo";
			this.btCargo.Size = new System.Drawing.Size(96, 26);
			this.btCargo.TabIndex = 20;
			this.btCargo.Text = "Cargos";
			this.btCargo.Click += new System.EventHandler(this.btCargo_Click);
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// cdsCargo
			// 
			this.cdsCargo.BindingContextCtrl = this;
			this.cdsCargo.DataLibrary = "LibPersonal";
			this.cdsCargo.DataLibraryUrl = "";
			this.cdsCargo.DataSetDef = "dsRolTablas";
			this.cdsCargo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCargo.Name = "cdsCargo";
			this.cdsCargo.SchemaClassName = "LibPersonal.DataClass";
			this.cdsCargo.SchemaDef = null;
			this.cdsCargo.BeforeFill += new C1.Data.FillEventHandler(this.cdsCargo_BeforeFill);
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.Sort = "idPersonal";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
			// 
			// cdvCargo
			// 
			this.cdvCargo.BindingContextCtrl = this;
			this.cdvCargo.DataSet = this.cdsCargo;
			this.cdvCargo.Sort = "idCargo";
			this.cdvCargo.TableName = "";
			this.cdvCargo.TableViewName = "RolCargo";
			// 
			// btDepartamento
			// 
			this.btDepartamento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btDepartamento.Location = new System.Drawing.Point(136, 224);
			this.btDepartamento.Name = "btDepartamento";
			this.btDepartamento.Size = new System.Drawing.Size(96, 26);
			this.btDepartamento.TabIndex = 21;
			this.btDepartamento.Text = "Departamento";
			this.btDepartamento.Click += new System.EventHandler(this.btDepartamento_Click);
			// 
			// cdvDepartamento
			// 
			this.cdvDepartamento.BindingContextCtrl = this;
			this.cdvDepartamento.DataSet = this.cdsCargo;
			this.cdvDepartamento.Sort = "idDepartamento";
			this.cdvDepartamento.TableName = "";
			this.cdvDepartamento.TableViewName = "RolDepartamento";
			// 
			// btPersonal
			// 
			this.btPersonal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPersonal.Location = new System.Drawing.Point(256, 224);
			this.btPersonal.Name = "btPersonal";
			this.btPersonal.Size = new System.Drawing.Size(96, 26);
			this.btPersonal.TabIndex = 22;
			this.btPersonal.Text = "Personal";
			this.btPersonal.Click += new System.EventHandler(this.btPersonal_Click);
			// 
			// btCarro
			// 
			this.btCarro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCarro.Location = new System.Drawing.Point(16, 267);
			this.btCarro.Name = "btCarro";
			this.btCarro.Size = new System.Drawing.Size(96, 26);
			this.btCarro.TabIndex = 23;
			this.btCarro.Text = "Carros";
			this.btCarro.Click += new System.EventHandler(this.btCarro_Click);
			// 
			// cdvCarros
			// 
			this.cdvCarros.BindingContextCtrl = this;
			this.cdvCarros.DataSet = this.cdsCarros;
			this.cdvCarros.Sort = "idCarro";
			this.cdvCarros.TableName = "";
			this.cdvCarros.TableViewName = "Carros";
			// 
			// cdsCarros
			// 
			this.cdsCarros.BindingContextCtrl = this;
			this.cdsCarros.DataLibrary = "LibFacturacion";
			this.cdsCarros.DataLibraryUrl = "";
			this.cdsCarros.DataSetDef = "dsOrdenMov";
			this.cdsCarros.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCarros.Name = "cdsCarros";
			this.cdsCarros.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCarros.SchemaDef = null;
			this.cdsCarros.BeforeFill += new C1.Data.FillEventHandler(this.cdsCarros_BeforeFill);
			// 
			// btOrdenMov
			// 
			this.btOrdenMov.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btOrdenMov.Location = new System.Drawing.Point(136, 267);
			this.btOrdenMov.Name = "btOrdenMov";
			this.btOrdenMov.Size = new System.Drawing.Size(96, 26);
			this.btOrdenMov.TabIndex = 24;
			this.btOrdenMov.Text = "Orden Mov.";
			this.btOrdenMov.Click += new System.EventHandler(this.btOrdenMov_Click);
			// 
			// cdvOrdenMov
			// 
			this.cdvOrdenMov.BindingContextCtrl = this;
			this.cdvOrdenMov.DataSet = this.cdsCarros;
			this.cdvOrdenMov.Sort = "CodArt";
			this.cdvOrdenMov.TableName = "";
			this.cdvOrdenMov.TableViewName = "OrdenMovilizacion";
			// 
			// btSucursal
			// 
			this.btSucursal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSucursal.Location = new System.Drawing.Point(376, 224);
			this.btSucursal.Name = "btSucursal";
			this.btSucursal.Size = new System.Drawing.Size(96, 26);
			this.btSucursal.TabIndex = 25;
			this.btSucursal.Text = "Sucursal";
			this.btSucursal.Click += new System.EventHandler(this.btSucursal_Click);
			// 
			// cdvSucursal
			// 
			this.cdvSucursal.BindingContextCtrl = this;
			this.cdvSucursal.DataSet = this.cdsCargo;
			this.cdvSucursal.Sort = "idRolSucursal";
			this.cdvSucursal.TableName = "";
			this.cdvSucursal.TableViewName = "RolSucursal";
			// 
			// btMarca
			// 
			this.btMarca.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btMarca.Location = new System.Drawing.Point(16, 95);
			this.btMarca.Name = "btMarca";
			this.btMarca.Size = new System.Drawing.Size(96, 26);
			this.btMarca.TabIndex = 26;
			this.btMarca.Text = "Marca";
			this.btMarca.Click += new System.EventHandler(this.btMarca_Click);
			// 
			// btArancel
			// 
			this.btArancel.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btArancel.Location = new System.Drawing.Point(136, 95);
			this.btArancel.Name = "btArancel";
			this.btArancel.Size = new System.Drawing.Size(96, 26);
			this.btArancel.TabIndex = 27;
			this.btArancel.Text = "Arancel";
			this.btArancel.Click += new System.EventHandler(this.btArancel_Click);
			// 
			// cdvMarca
			// 
			this.cdvMarca.BindingContextCtrl = this;
			this.cdvMarca.DataSet = this.cdsGrupoArt;
			this.cdvMarca.Sort = "idMarca";
			this.cdvMarca.TableName = "";
			this.cdvMarca.TableViewName = "ArticuloMarca";
			// 
			// cdvArancel
			// 
			this.cdvArancel.BindingContextCtrl = this;
			this.cdvArancel.DataSet = this.cdsGrupoArt;
			this.cdvArancel.Sort = "idArancel";
			this.cdvArancel.TableName = "";
			this.cdvArancel.TableViewName = "ArticuloArancel";
			// 
			// btUnidad
			// 
			this.btUnidad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btUnidad.Location = new System.Drawing.Point(256, 95);
			this.btUnidad.Name = "btUnidad";
			this.btUnidad.Size = new System.Drawing.Size(96, 26);
			this.btUnidad.TabIndex = 28;
			this.btUnidad.Text = "Unidad";
			this.btUnidad.Click += new System.EventHandler(this.btUnidad_Click);
			// 
			// cdvUnidad
			// 
			this.cdvUnidad.BindingContextCtrl = this;
			this.cdvUnidad.DataSet = this.cdsGrupoArt;
			this.cdvUnidad.Sort = "idUnidad";
			this.cdvUnidad.TableName = "";
			this.cdvUnidad.TableViewName = "ArticuloUnidad";
			// 
			// btInvSaldo
			// 
			this.btInvSaldo.Location = new System.Drawing.Point(376, 96);
			this.btInvSaldo.Name = "btInvSaldo";
			this.btInvSaldo.Size = new System.Drawing.Size(96, 25);
			this.btInvSaldo.TabIndex = 29;
			this.btInvSaldo.Text = "Inventario";
			this.btInvSaldo.Click += new System.EventHandler(this.btInvSaldo_Click);
			// 
			// cdsPagos
			// 
			this.cdsPagos.BindingContextCtrl = this;
			this.cdsPagos.DataLibrary = "LibFacturacion";
			this.cdsPagos.DataLibraryUrl = "";
			this.cdsPagos.DataSetDef = "dsPago";
			this.cdsPagos.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPagos.Name = "cdsPagos";
			this.cdsPagos.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPagos.SchemaDef = null;
			// 
			// btRolConcepto
			// 
			this.btRolConcepto.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btRolConcepto.Location = new System.Drawing.Point(256, 267);
			this.btRolConcepto.Name = "btRolConcepto";
			this.btRolConcepto.Size = new System.Drawing.Size(96, 26);
			this.btRolConcepto.TabIndex = 30;
			this.btRolConcepto.Text = "Rol Concepto";
			this.btRolConcepto.Click += new System.EventHandler(this.btRolConcepto_Click);
			// 
			// cdsRolConcepto
			// 
			this.cdsRolConcepto.BindingContextCtrl = this;
			this.cdsRolConcepto.DataLibrary = "LibPersonal";
			this.cdsRolConcepto.DataLibraryUrl = "";
			this.cdsRolConcepto.DataSetDef = "dsRolGrupo";
			this.cdsRolConcepto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolConcepto.Name = "cdsRolConcepto";
			this.cdsRolConcepto.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolConcepto.SchemaDef = null;
			// 
			// cdvRolConcepto
			// 
			this.cdvRolConcepto.BindingContextCtrl = this;
			this.cdvRolConcepto.DataSet = this.cdsRolConcepto;
			this.cdvRolConcepto.TableName = "";
			this.cdvRolConcepto.TableViewName = "RolConcepto";
			// 
			// btTomaFisica
			// 
			this.btTomaFisica.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btTomaFisica.Location = new System.Drawing.Point(376, 267);
			this.btTomaFisica.Name = "btTomaFisica";
			this.btTomaFisica.Size = new System.Drawing.Size(96, 26);
			this.btTomaFisica.TabIndex = 31;
			this.btTomaFisica.Text = "Toma Física";
			this.btTomaFisica.Click += new System.EventHandler(this.btTomaFisica_Click);
			// 
			// cdsToma
			// 
			this.cdsToma.BindingContextCtrl = this;
			this.cdsToma.DataLibrary = "LibFacturacion";
			this.cdsToma.DataLibraryUrl = "";
			this.cdsToma.DataSetDef = "dsTomaFisica";
			this.cdsToma.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsToma.Name = "cdsToma";
			this.cdsToma.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsToma.SchemaDef = null;
			// 
			// cdvTomaFisica
			// 
			this.cdvTomaFisica.BindingContextCtrl = this;
			this.cdvTomaFisica.DataSet = this.cdsToma;
			this.cdvTomaFisica.Sort = "idTmpTomaDetalle";
			this.cdvTomaFisica.TableName = "";
			this.cdvTomaFisica.TableViewName = "TmpTomaDetalle";
			// 
			// ultraButton1
			// 
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton1.Location = new System.Drawing.Point(16, 302);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(96, 25);
			this.ultraButton1.TabIndex = 32;
			this.ultraButton1.Text = "Rol Ingreso";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// cdvRolIngreso
			// 
			this.cdvRolIngreso.BindingContextCtrl = this;
			this.cdvRolIngreso.DataSet = this.cdsRolIngreso;
			this.cdvRolIngreso.TableName = "";
			this.cdvRolIngreso.TableViewName = "RolPago";
			// 
			// cdsRolIngreso
			// 
			this.cdsRolIngreso.BindingContextCtrl = this;
			this.cdsRolIngreso.DataLibrary = "LibPersonal";
			this.cdsRolIngreso.DataLibraryUrl = "";
			this.cdsRolIngreso.DataSetDef = "dsRolGrupo";
			this.cdsRolIngreso.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolIngreso.Name = "cdsRolIngreso";
			this.cdsRolIngreso.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolIngreso.SchemaDef = null;
			// 
			// btPagos
			// 
			this.btPagos.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPagos.Location = new System.Drawing.Point(16, 136);
			this.btPagos.Name = "btPagos";
			this.btPagos.Size = new System.Drawing.Size(96, 25);
			this.btPagos.TabIndex = 33;
			this.btPagos.Text = "Pagos";
			this.btPagos.Click += new System.EventHandler(this.btPagos_Click);
			// 
			// btCobros
			// 
			this.btCobros.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCobros.Location = new System.Drawing.Point(136, 302);
			this.btCobros.Name = "btCobros";
			this.btCobros.Size = new System.Drawing.Size(96, 25);
			this.btCobros.TabIndex = 34;
			this.btCobros.Text = "Cobros";
			this.btCobros.Click += new System.EventHandler(this.btCobros_Click);
			// 
			// btAvanzado
			// 
			this.btAvanzado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAvanzado.Location = new System.Drawing.Point(376, 304);
			this.btAvanzado.Name = "btAvanzado";
			this.btAvanzado.Size = new System.Drawing.Size(96, 26);
			this.btAvanzado.TabIndex = 35;
			this.btAvanzado.Text = "Avanzado ...";
			this.btAvanzado.Click += new System.EventHandler(this.btAvanzado_Click);
			// 
			// ImportarExcel
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(488, 428);
			this.Controls.Add(this.btAvanzado);
			this.Controls.Add(this.btCobros);
			this.Controls.Add(this.btPagos);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.btTomaFisica);
			this.Controls.Add(this.btRolConcepto);
			this.Controls.Add(this.btInvSaldo);
			this.Controls.Add(this.btUnidad);
			this.Controls.Add(this.btArancel);
			this.Controls.Add(this.btMarca);
			this.Controls.Add(this.btSucursal);
			this.Controls.Add(this.btOrdenMov);
			this.Controls.Add(this.btCarro);
			this.Controls.Add(this.btPersonal);
			this.Controls.Add(this.btDepartamento);
			this.Controls.Add(this.btCargo);
			this.Controls.Add(this.chkClavePrimaria);
			this.Controls.Add(this.chkModificaExiste);
			this.Controls.Add(this.txtMensaje);
			this.Controls.Add(this.cmbGrupoCliente);
			this.Controls.Add(this.btGrupoCliente);
			this.Controls.Add(this.cmbGrupoArt);
			this.Controls.Add(this.btGrupoArt);
			this.Controls.Add(this.cmbTipoFactura);
			this.Controls.Add(this.btCuentasxP);
			this.Controls.Add(this.btCuentasxC);
			this.Controls.Add(this.btActivoFijo);
			this.Controls.Add(this.btInventario);
			this.Controls.Add(this.btProveedor);
			this.Controls.Add(this.btClientes);
			this.Controls.Add(this.btArticulo);
			this.Controls.Add(this.btAsientos);
			this.Controls.Add(this.btCuentas);
			this.Name = "ImportarExcel";
			this.Text = "Importar de Excel";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ImportarExcel_Closing);
			this.Load += new System.EventHandler(this.ImportarExcel_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvClienteGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMensaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCarros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCarros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvOrdenMov)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArancel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvUnidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsToma)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTomaFisica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolIngreso)).EndInit();
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
			#region Formato y Lectura de Archivo
			this.txtMensaje.Text = "Formato Asiento: \n\rCodAsiento, Descripcion, Fecha, ";
			this.txtMensaje.Text += "\n\r--- CodCuenta, Debe, Haber, Centro de Costo";
			this.txtMensaje.Text += "\n\rExisten 2 columnas con los nombres de las columnas deseadas. La celda A2 esta vacia.";
			if (MessageBox.Show("¿Desea Subir Asiento de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stArchivo = "";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				stArchivo = this.openFileDialog.FileName.ToString();
			}
			else
				MessageBox.Show("Cancelado por el usuario");

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Asientos'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

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
			#endregion Formato y Lectura de Archivo
			#region Lectura de Columnas de Asiento
			string stCodigo;
			int iCuenta = 0;

			C1.Data.C1DataRow drAsiento = this.cdsAsiento.TableViews["Asiento"].AddNew();
			string [] stTitulo = new string[100];
			XLCell cell = sheet.GetCell(0, 0);
			int iCuentaColsAsi = 0;

			// Lectura de Titulos de columnas
			for (int i = 0; i < sheet.Columns.Count; i++)
			{
				cell = sheet.GetCell(0, i);
				if (cell != null && cell.Value != null)
				{
					stTitulo[i] = cell.Value.ToString().Trim();
					try
					{
						drAsiento[stTitulo[i]] = System.DBNull.Value;
						iCuentaColsAsi ++;
					}
					catch
					{
						MessageBox.Show("No existe campo " + stTitulo[i].Trim()
							+ ".\nPrimera fila debe tener los campos a importar.",
							"Error");
						return;
					} // Fin de Primer try
				} // Fin de If
			} // Fin de For de primera fila
			this.cdsAsiento.RejectChanges();
			#endregion Lectura de Columnas de Asiento
			#region Lectura de Columnas de Detalle
			C1.Data.C1DataRow drDetAsiento= this.cdsAsiento.TableViews["AsientoDetalle"].AddNew();
			string [] stTitulo1 = new string[100];
			int iCuentaColsDetAsi = 0;
			cell = sheet.GetCell(1, 0);
			if ((cell != null && cell.Value != null))
			{
				MessageBox.Show("Columna A2 debe estar vacia");
				return;
			}
			for (int i = 1; i < sheet.Columns.Count; i++)
			{
				cell = sheet.GetCell(1, i);
				if (cell != null && cell.Value != null)
				{
					try
					{
						stTitulo1[i] = cell.Value.ToString().Trim();
					}
					catch(System.Exception ex)
					{
						Console.WriteLine(ex.Message+ "Segunda Fila Titulo");
						break;
					}
					if (stTitulo1[i].ToUpper().Trim() == "CODCUENTA") 
					{
						iCuentaColsDetAsi ++;
						continue;
					}
					if (stTitulo1[i].ToUpper().Trim() == "CODSUBPROYECTO") 
					{
						iCuentaColsDetAsi ++;
						continue;
					}
					try
					{
						drDetAsiento[stTitulo1[i]].GetType();
						iCuentaColsDetAsi ++;
					}
					catch
					{
						MessageBox.Show("No existe campo " + stTitulo1[i].Trim()
							+ ".\nSegunda fila debe tener los campos de Detalle de Compras a importar.",
							"Error");
						return;
					} // Fin de If
				} // Fin de try
			} // Fin de For de primera fila
			this.cdsAsiento.RejectChanges();
			if (stTitulo[0].ToUpper() != "CODASIENTO")
			{
				MessageBox.Show("Primera columna debe ser CodAsiento", "Información");
				return;
			}
			#endregion Lectura de Columnas de Detalle
			#region Lectura de Filas
			for (int f = 2; f < sheet.Rows.Count; f++)
			{
				cell = sheet.GetCell(f, 0);
				if (cell == null || cell.Value == null) continue;
				#region Busqueda de asiento existente
				stCodigo = cell.Value.ToString();
				int iBusca = this.cdvAsiento.Find(stCodigo);
				if (iBusca >= 0) continue;
				#endregion Busqueda de asiento existente
				#region Numeracion de Asiento
				iCuenta ++;
				drAsiento = this.cdsAsiento.TableViews["Asiento"].AddNew();
				if (stCodigo.ToUpper() == "TMPASI") // Crea numeracion de asiento
				{
					stCodigo = Contabilidad.NumeroAsiento(cdsAsiento, 1);
				}
				if (stCodigo.ToUpper() == "TMPING") // Crea numeracion de Ingreso
				{
					stCodigo = Contabilidad.NumeroAsiento(cdsAsiento, 2);
				}
				if (stCodigo.ToUpper() == "TMPEGR") // Crea numeracion de Egreso
				{
					stCodigo = Contabilidad.NumeroAsiento(cdsAsiento, 3);
				}
				drAsiento["CodAsiento"] = stCodigo;
				#endregion Numeracion de Asiento
				#region Lectura de columnas y grabacion de asiento
				for(int j=1; j< iCuentaColsAsi; j++)
				{
					cell = sheet.GetCell(f, j);
					if (cell == null || cell.Value == null) continue;

					try
					{
						if (cell.Value.GetType().Name == "String")
							drAsiento[stTitulo[j]] = (string) cell.Value; // Texto
						if (cell.Value.GetType().Name == "Double")
							drAsiento[stTitulo[j]] = (double) cell.Value; // Valor
						if (cell.Value.GetType().Name == "DateTime")
							drAsiento[stTitulo[j]] = (DateTime) cell.Value; // Fecha	
					}
					catch(System.Exception ex)
					{
						int f1 = f + 1;
						int j1 = j+1;
						MessageBox.Show(ex.Message, "Error en Linea " + f1.ToString() + " columna " + j1.ToString());
					}
				}
				cdsAsiento.Update();
				f++;
				#endregion Lectura de columnas y grabacion de asiento			
				#region Grabacion de detalle de asientos
				while(f < sheet.Rows.Count && 
					(sheet.GetCell(f, 0) == null || sheet.GetCell(f, 0).Value == null || sheet.GetCell(f, 0).Value.ToString().Trim() == ""))
				{
					try
					{
						string stCelda = sheet.GetCell(f, 1).Value.ToString();
						#region Crea Nueva Linea e idAsiento
						drDetAsiento = this.cdsAsiento.TableViews["AsientoDetalle"].AddNew();
						drDetAsiento["idAsiento"] = drAsiento["idAsiento"];
						#endregion Crea Nueva Linea e idAsiento
						#region Lectura de Columnas
						for(int j=1; j<= iCuentaColsDetAsi; j++)
						{
							cell = sheet.GetCell(f, j);
							if (cell == null || cell.Value == null) continue;

							#region Lectura de Columnas Especiales
							if (stTitulo1[j].ToUpper() == "CODCUENTA")
							{
								string stBuscaCuenta = "SELECT idCuenta FROM Cuenta "
									+ " Where Codigo = '" + cell.Value.ToString().Trim() + "'";
								int IdCuenta = Funcion.iEscalarSQL(cdsAsiento, stBuscaCuenta, false);
								drDetAsiento["idCuenta"] = IdCuenta;
								continue;
							}
							if (stTitulo1[j].ToUpper() == "CODSUBPROYECTO")
							{
								string stBuscaCuenta = "SELECT idSubProyecto FROM SubProyecto "
									+ " Where CodSubproyecto = '" + cell.Value.ToString().Trim() + "'";
								int IdCuenta = Funcion.iEscalarSQL(cdsAsiento, stBuscaCuenta, false);
								drDetAsiento["idSubProyecto"] = IdCuenta;
								continue;
							}
							#endregion Lectura de Columnas Especiales
							#region Grabacion de Valores
							if (cell.Value.GetType().Name == "String")
								drDetAsiento[stTitulo1[j]] = (string) cell.Value; // Texto
							if (cell.Value.GetType().Name == "Double")
								drDetAsiento[stTitulo1[j]] = (double) cell.Value; // Valor
							#endregion Grabacion de Valores
						}
						f++;
						#endregion Lectura de Columnas
					}
					catch(System.Exception ex)
					{
						f++;
						MessageBox.Show(ex.Message, "Error en fila " + f.ToString());
					}
					this.cdsAsiento.Update();
				} // fin de detAsiento
				f --;
				//						this.cdsAsiento.Update();
				#endregion Grabacion de detalle de asientos
			} // Fin de For
			#endregion Lectura de Filas
			#region Mensajes
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
			#endregion Mensajes
		}

		private void btArticulo_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Codigo, Articulo, CodAlterno, ";
			this.txtMensaje.Text += "Caracteristica, Precio1 al 6, Iva, Ice, Ubicacion, etc";
			if (MessageBox.Show("¿Desea Subir Artículos de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;
			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Articulo'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			this.LeeExcel(this.cdsArticulo, 
				this.cdsArticulo.TableViews["Articulo"], this.cdvArticulo, 0);
			string stExec = "EXEC ArticuloActualizaTipo";
			Funcion.EjecutaSQL(cdsArticulo, stExec, true);

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

		private void TipoDefinido(C1.Data.C1DataRow dr, int iTipo)
		{
			switch(iTipo)
			{
				case 1:
					dr["Proveedor"] = false;
					break;
				case 2:
					dr["Proveedor"] = true;
					break;
			}
		}
		private string stArchivo = "";
		CultureInfo us = new CultureInfo("en-US");
		private bool LeeExcel(C1.Data.C1DataSet ds, C1.Data.C1DataTable dtArchivo, 
			C1.Data.C1DataView dv, int iTipo)
		{
			string stTabla = dtArchivo.Name.ToString();
			this.Cursor = Cursors.WaitCursor;
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
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
			string [] stTitulo = new string[500];
			string [] stTipoDato = new string[500];
			// Lectura de Titulos de columnas
			for (int i = 0; i < sheet.Columns.Count; i++)
			{
				XLCell cell = sheet.GetCell(0, i);
				//				if (cell != null && cell.Value != null)
			{
				try
				{
					stTitulo[i] = "";
					stTitulo[i] = cell.Value.ToString().Trim();

					drArticulo[stTitulo[i]].GetType();
					try
					{
						drArticulo[stTitulo[i]] = "a";
						stTipoDato[i] = "String";
					}
					catch
					{
						try
						{
							drArticulo[stTitulo[i]] = 1;
							stTipoDato[i] = "Double";
						}
						catch
						{
							stTipoDato[i] = "Fecha";
						}
					}
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
					else
					{
						int ColError = i + 1;
						MessageBox.Show("No existe campo " + stTitulo[i]
							+ " en columna " + ColError.ToString().Trim(),
							"Error en nombres de Campos");
						this.Cursor = Cursors.Default;
						return false;
					}
					//					i = sheet.Columns.Count;
				} // Fin de Primer try
			} // Fin de If
			} // Fin de For de primera fila
			ds.RejectChanges();
			// FIN de Verificacion que primera fila tiene los nombres de los campos a subirse

			// Lectura de datos desde segunda fila
			this.Cursor = Cursors.WaitCursor;
			//			if (File.Exists("SalidaError.txt")) File.Delete("SalidaError.txt");
			StreamWriter sr;
			try
			{
				sr= File.CreateText("SalidaError.XLS");
			}
			catch
			{
				MessageBox.Show("Esta abierto el archivo SalidaError.xls");
				Cursor = Cursors.Default;
				return false;
			}

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
					stCodigo = stCodigo.Replace("'", "''");
					string stSelect = "SELECT COUNT(*) FROM "
						+ stTabla + " WHERE " + stTitulo[0] + " = '" 
						+ stCodigo + "'";
					iBusca = Funcion.iEscalarSQL(cdsArticulo, stSelect);

					if (iBusca <= 0 || chkModificaExiste.Checked) // Articulo no encontrado o modifica existente
					{
						iCuenta ++;

						//						drArticulo[stTitulo[0]] = stCodigo;

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
									// Numero gabado en campo de texto
									if (stTipoDato[j] == "String" && cell.Value.GetType().Name == "Double")
									{
										string stTexto = cell.Value.ToString();
										stValor += "'" + stTexto + "'";
										if (bModifica) stInsert += " = " + "'" + stTexto + "'";
									}
									else if (cell.Value.GetType().Name == "String")
									{
										string stTexto = (string) cell.Value;
										stTexto = stTexto.Replace("'", "''");
										stValor += "'" + stTexto + "'";
										if (bModifica) stInsert += " = " + "'" + stTexto + "'";
									}
									else if (cell.Value.GetType().Name == "DateTime")
									{
										dtFechaLeida = (DateTime) cell.Value;
										stValor += "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
										if (bModifica) stInsert += " = " + "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
									}
									else if (cell.Value.GetType().Name == "Double")
									{
										double dValor = (double) cell.Value;
										stValor += dValor.ToString("0.000000", us);
										if (bModifica) stInsert += " = " + dValor.ToString("0.000000", us);
									}
									else if (cell.Value.GetType().Name == "Boolean")
									{
										string stTexto = "0";
										if ((bool) cell.Value)
											stTexto = "1";
										stValor += stTexto;
										if (bModifica) stInsert += " = " + stTexto;
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
						if (iTipo > 0) // Cliente o Proveedor
						{
							int iProv = iTipo - 1;
							if (!bModifica) stInsert += ", Proveedor";
							stValor += ", " + iProv.ToString();
						}
						if (!bModifica) stInsert+= ") ";
						stValor += ") ";
						#endregion

						//						if (iCuenta == 6)
						//						{
						//							MessageBox.Show(stInsert);
						//						}

						try
						{
							if (bModifica) 
							{
								stValor = "";
								stInsert += " WHERE " + stTitulo[0] + " = '" + stCodigo +"'";
							}
							if (chkClavePrimaria.Checked)
							{
								Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT " + stTabla + " on", false);
							}

							stInsert += stValor;
							Funcion.EjecutaSQL(cdsArticulo, stInsert, true);
							if (chkClavePrimaria.Checked)
							{
								Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT " + stTabla + " off", false);
							}
							//cdsArticulo.Update();
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
				MessageBox.Show("No hay Registros para Actualizar", 
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

		private void btInventario_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Formato: \nNumero, Fecha, Cliente/CodCliente, ContadoCredito"
				+ "\n ---- Cantidad, Codigo, Precio";
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
			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Facturas'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			#region Apertura de Archivo de Excel
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
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

			bool bCodigo = false;
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
				if (cell1 != null && cell1.Value != null)
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
					if (stCampo[i].ToUpper().Trim() == "CODCLIENTE")
					{
						stCampo[i] = "idCliente";
						bCodigo = true;
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
										string text9 = "";
										if (bCodigo)
											text9 = "Select idCliente From Cliente Where Codigo = '" + stCelda.Trim() + "'";
										else 
											text9 = "Select idCliente From Cliente Where Nombre = '" + stCelda.Trim() + "'";
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
									stValores = stValores + num12.ToString("0.00000", this.us);
									if (bExiste)
									{
										stInsert = stInsert + " = " + num12.ToString("0.00000", this.us);
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
									stInsert = stInsert + " WHERE Numero = '" + stClavePrinc 
										+ "' And idTipoFactura = " + stTipoFactura;
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
							int idCompra1 = Funcion.iEscalarSQL(cdsArticulo, 
								stCreaFila + ";Select cast(IsNull(Scope_Identity(), 0)as int)",true);
							if (!bExiste)
							{
								IdCompra = idCompra1;
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
						while (((iFila < sheet1.Rows.Count) && (sheet1.GetCell(iFila, 0) == null ||
							sheet1.GetCell(iFila, 0).Value == null)) 
							&& ((sheet1.GetCell(iFila, 1) != null) && (sheet1.GetCell(iFila, 1).Value != null)))
						{
							stInsert = "INSERT INTO DetCompra (idCompra, ";
							stValores = " VALUES (" + IdCompra.ToString() + ", ";
							#region Lectura de Columnas de DetCompra
							for (int iColDet = 1; iColDet <= iNumColDet; iColDet++)
							{
								cell1 = sheet1.GetCell(iFila, iColDet);
								if (cell1 == null || cell1.Value == null) return;
								#region Manejo de Celda
								try
								{
									if (textArray2[iColDet].ToUpper().Trim() == "CODIGO")
									{
										string text12 = "Select idArticulo From Articulo Where Codigo = '" + cell1.Value.ToString() + "'";
										int num17 = Funcion.iEscalarSQL(this.cdsArticulo, text12, false);
										stInsert = stInsert + "idArticulo";
										stValores = stValores + num17.ToString("0.00", this.us);
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
											stValores = stValores + ((double) cell1.Value).ToString("0.000000", this.us);
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
								catch (Exception exception6)
								{
									int num19 = iColDet + 1;
									int num20 = iFila + 1;
									DialogResult result2 = MessageBox.Show("Error en Fila: " 
										+ num20.ToString() + " Columna: " + num19.ToString() + " " 
										+ exception6.Message + "\n¿Desea continuar con la subida de datos?", 
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

		private void btClientes_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Ruc, Nombre, Direccion, ";
			this.txtMensaje.Text += "Telefono, EMail, idGrupoCliente, Fax, Representante, etc";
			if (MessageBox.Show("¿Desea Subir Clientes de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Clientes'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			LeeExcel(cdsCliente, cdsCliente.TableViews["Cliente"], cdvCliente, 1);
		}

		private void btProveedor_Click(object sender, System.EventArgs e)
		{
			txtMensaje.Text = "Campos: Ruc, Nombre, Direccion, ";
			txtMensaje.Text += "Telefono, EMail, idGrupoCliente, Fax, Representante, etc";
			if (MessageBox.Show("¿Desea Subir Proveedores de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Proveedor'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			LeeExcel(cdsCliente, cdsCliente.TableViews["Cliente"], cdvCliente, 2);
		}

		private void btCuentas_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Codigo, Descripcion";
			if (MessageBox.Show("¿Desea Subir Plan de Cuentas de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:PlanCuentas'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			bool bExito = this.LeeExcel(this.cdsCuentas, this.cdsCuentas.TableViews["Cuenta"], this.cdvCuentas, 0);
			if (bExito)
			{
				this.Cursor = Cursors.WaitCursor;
				Contabilidad.GeneraArbol(this.cdsCuentas);
				this.Cursor = Cursors.Default;
			}
		}

		private void btActivoFijo_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Codigo, Articulo, FechaUltimo, ";
			this.txtMensaje.Text += "CostoUltimo, Saldo, IdGrupoArticulo, Depreciacion, etc";
			if (MessageBox.Show("¿Desea Subir Artículos de Activos Fijos de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:ActivoF'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			stArchivo = "";
			bool bExito = this.LeeExcel(this.cdsArticulo, 
				this.cdsArticulo.TableViews["Articulo"], this.cdvArticulo, 0);
			if (stArchivo.Length == 0) return;

			if (MessageBox.Show("¿Desea Subir Activos Fijos de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;
			// Cod_Art, Desc_Art, Saldo, Fecha_u1, Costo_u1, depreciacion, IdArticuloGrupo
			// Lectura de datos desde segunda fila
			C1.C1Excel.C1XLBook miArticulo = new C1.C1Excel.C1XLBook();
			try
			{
				miArticulo.Load(@stArchivo);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error Abriendo el Archivo");
				this.Cursor = Cursors.Default;
			}
			XLSheet sheet = miArticulo.Sheets[0];
			int iCuenta = 0;

			for (int f = 1; f < sheet.Rows.Count; f++)
			{
				// Primera columna tiene el codigo que debe Existir

				XLCell cell = sheet.GetCell(f, 0);
				if (cell != null && cell.Value != null)
				{
					string stCodigo = (string) cell.Value;
					string stBusca = "Select idArticulo From Articulo "
						+ "Where Codigo = '" + stCodigo + "'";
					int iBusca = Funcion.iEscalarSQL(cdsArticulo, stBusca, false);
					if (iBusca > 0) // Articulo encontrado
					{
						#region Articulo Existe
						iCuenta ++;
						try
						{
							C1.Data.C1DataRow drCompra = this.cdsCompra.TableViews["Compra"].AddNew();
							drCompra["idTipoFactura"] = 4;
							DateTime dtFecha = DateTime.Today;
							if (sheet.GetCell(f, 2).Value != null)
								dtFecha= (DateTime) sheet.GetCell(f, 2).Value;
							drCompra["Fecha"] = dtFecha;
							drCompra["Numero"] = "A.F.";
							this.cdsCompra.Update();
							int lnIdCompra = (int) drCompra["idCompra"];

							// Creacion de Detalle de Compra
							bool bAumenta = false;
							while ( f <= sheet.Rows.Count && sheet.GetCell(f, 2) != null 
								&& (DateTime) sheet.GetCell(f, 2).Value == dtFecha
								&& sheet.GetCell(f, 0).Value.ToString().Trim().Length > 0)
							{
								stCodigo = sheet.GetCell(f, 0).Value.ToString().Trim();
								stBusca = "Select idArticulo From Articulo "
									+ "Where Codigo = '" + stCodigo + "'";
								iBusca = Funcion.iEscalarSQL(cdsArticulo, stBusca, false);

								C1.Data.C1DataRow drDetCompra = this.cdsCompra.TableViews["DetCompra"].AddNew();
								drDetCompra["idCompra"] = lnIdCompra;
								drDetCompra["idArticulo"] = iBusca;
								drDetCompra["Cantidad"] = sheet.GetCell(f, 4).Value;
								drDetCompra["Precio"] = sheet.GetCell(f, 3).Value;
								bAumenta = true;
								f ++;
							} //Fin de detalle de compra
							if  (bAumenta) f--;
						}
						catch (System.Exception ex)
						{
							int FilaError = f + 1;
							MessageBox.Show("Error en Fila " + FilaError.ToString()
								+ " " + ex.Message, "Error Creando Compra de Activo Fijo");
							return;
						}
						//f --;
						#endregion  Articulo Existente
					} // Fin de Creacion de Compra
					else
					{
						MessageBox.Show("Artículo de Codigo: " + stCodigo + " no existe");
					}
				} // Fin de lectura de Fila
			} // Fin de Lectura de Archivo
			this.cdsCompra.Update();

			string stExec = "EXEC ArticuloActualizaTipo";
			Funcion.EjecutaSQL(cdsArticulo, stExec, true);
			MessageBox.Show("Activos Fijos Creados",
				"Confirmación", MessageBoxButtons.OK, 
				MessageBoxIcon.Information);
		}



		private void btGrupoArt_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Grupo, idTipoGrupo";
			if (MessageBox.Show("¿Desea Subir Grupos de Artículos de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:GrupoArt'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			this.LeeExcel(this.cdsGrupoArt, this.cdsGrupoArt.TableViews["ArticuloGrupo"], 
				this.cdvGrupo, 0);		
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

		private void btCuentasxC_Click(object sender, System.EventArgs e)
		{
			#region Apertura de Archivo
			this.txtMensaje.Text = "Formato: Numero, CodigoCliente, Fecha, FechaVencimiento, Tipo, Total";
			if (MessageBox.Show("¿Desea Subir CxC de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stArchivo = "";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
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
			#endregion Apertura Archivo

			#region Lectura de Titulos
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
			string [] stTitulo = new string[100];
			string [] stTipoDato = new string[100];
			// Lectura de Titulos de columnas de Compras
			XLCell cell = sheet.GetCell(0, 0);

			int iCamposCompra = 0;
			for (int i = 0; i < sheet.Columns.Count; i++)
			{
				cell = sheet.GetCell(0, i);
				if (cell == null) continue;
				try
				{
					stTitulo[i] = (string) cell.Value;
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error al leer titulo");
					return;
				}
				iCamposCompra++;
				if (stTitulo[i].ToUpper().Trim() == "CODIGOCLIENTE") 
				{
					stTipoDato[i] = "String";
					continue;
				}
				if (stTitulo[i].ToUpper().Trim() == "TIPO") 
				{
					stTipoDato[i] = "String";
					continue;
				}

				//completo
				try
				{
					stTitulo[i] = "";
					stTitulo[i] = (string) cell.Value;

					drCompra[stTitulo[i]].GetType();
					try
					{
						drCompra[stTitulo[i]] = "a";
						stTipoDato[i] = "String";
					}
					catch
					{
						try
						{
							drCompra[stTitulo[i]] = 1;
							stTipoDato[i] = "Double";
						}
						catch
						{
							stTipoDato[i] = "DateTime";
						}
					}
				}
				catch
				{
					MessageBox.Show("No existe campo " + stTitulo[i].Trim()
						+ ".\nPrimera fila debe tener los campos de Compras a importar.",
						"Error");
					return;
				} // Fin de try
			} // Fin de For de primera fila
			this.cdsCompra.RejectChanges();
			if (stTitulo[0].ToString().Trim().ToUpper() != "NUMERO")
			{
				MessageBox.Show("Primera columna debe ser Numero", "Información");
				return;
			}

			#endregion Lectura de Titulos

			// Actualizacion de Articulos
			string stSelect = "Select Count(*) From Articulo "
				+ "Where Articulo like 'Saldo%'";
			int IdArt = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
			Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			if (IdArt <= 0)
			{
				MessageBox.Show("No existe articulo Saldo inicial");
				return;
			}
			stSelect = "Select idArticulo From Articulo "
				+ "Where Articulo like 'Saldo%'";
			IdArt = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);

			#region Ingreso de Compras a Credito
			StreamWriter sr= File.CreateText("SalidaError.xls");
			string stClave = "";
			string stCodCliente = "";
			for (int f = 1; f < sheet.Rows.Count; f++)
			{
				cell = sheet.GetCell(f, 0);
				#region Existe Valor de Celda
				if (cell != null && cell.Value != null) // Linea de Compra
				{
					#region Verifica Nuevo o Modificacion
					string stNumero = "";
					if (cell.Value.GetType().Name.ToString().ToUpper() == "DOUBLE")
						stNumero = cell.Value.ToString();
					else stNumero =	(string) cell.Value;
				
					stClave = stNumero;
					if (bCxP)
					{
						stCodCliente = sheet.GetCell(f, 1).Value.ToString();
						stCodigo = "SELECT COUNT(*) FROM Compra WHERE Numero = '" 
							+ stNumero + "' And Otro = '" + stCodCliente + "'";
					}
					else
						stCodigo = "SELECT COUNT(*) FROM Compra WHERE Numero = '" + stNumero + "'";
				
					int iBusca = 0;
					iBusca = Funcion.iEscalarSQL(cdsArticulo, stCodigo);
					string stInsert = "";
					bool bModifica = false;
					if (chkModificaExiste.Checked && iBusca == 0) 
						bModifica = true;

					if (bModifica)
						stInsert = "UPDATE Compra SET ";
					else
						stInsert = "INSERT INTO Compra (ContadoCredito,";

					string stValor = " VALUES (2,";
					#endregion Verifica Nuevo o Modificacion

					if (iBusca <= 0 || bModifica) // Inserta o Modifica Compra
					{
						#region Lectura de Columnas de Compra
						iCuenta ++;
						DateTime dtFechaLeida;
						double dTotal = 0;
						for (int j = 0; j < iCamposCompra; j++)
						{
							cell = sheet.GetCell(f, j);
							if (cell != null && cell.Value != null)
							{
								#region try 
								try
								{
									if (stTitulo[j].Trim().ToUpper() == "CODIGOCLIENTE")
										stInsert += "Otro";
									else if (stTitulo[j].Trim().ToUpper() == "TIPO")
									{
										stInsert += "idTipoFactura, ";
										string stCelda = cell.Value.ToString();
										string stCodCV = "1";
										string stCodDCDV = "5";
										if (bCxP) 
										{
											stCodCV = "4";
											stCodDCDV = "6";
										}
										if (stCelda == "FC") stValor += "'" + stCodCV +"', ";
										else stValor += "'" + stCodDCDV +"', ";
										if (bModifica) stInsert += " = " + "'" + stCelda +"'";
										continue;
									}
									else
									{
										if (stTitulo[j].Trim().ToUpper() == "TOTAL") 
											dTotal = (double) cell.Value;
										stInsert += stTitulo[j];
									}
									if (stTipoDato[j] == "String")
									{
										string stCelda = "";
										if (cell.Value.GetType().Name == "Double")
											stCelda = cell.Value.ToString();
										else
											stCelda = (string) cell.Value;
										
										stValor += "'" + stCelda +"'";
										if (bModifica) stInsert += " = " + "'" + stCelda +"'";
									}
									if (stTipoDato[j] == "DateTime")
									{
										dtFechaLeida = (DateTime) cell.Value;
										stValor += "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
										if (bModifica) stInsert += " = " + "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
									}
									if (stTipoDato[j] == "Double")
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
								} 
								#endregion  Fin de Try Catch
							} // Fin de verificacion de valor nulo
						} // Fin de lectura de columnas for

						if (stInsert.Trim().EndsWith(","))
						{
							stInsert = stInsert.Substring(0,stInsert.Length - 2);
							stValor = stValor.Substring(0,stValor.Length - 2);
						}
						if (!bModifica) stInsert+= ") ";
						stValor += ") ";
						f++;
						#endregion Lectura de Columnas de compra

						#region Compra Grabacion

						try
						{
							if (bModifica) 
							{
								stValor = "";
								stInsert += " WHERE Numero = '" + stClave +"'";
								if (bCxP) 
									stInsert += " And Otro = '" + stCodCliente +"'";
							}
							//							Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT Compra on", false);
							string stTodo = stInsert + stValor;
							Funcion.EjecutaSQL(cdsArticulo, stTodo, true);
							int iCuenta1 = Funcion.iEscalarSQL(cdsArticulo, "Select Count(*) From Compra", true);
							//							if (iCuenta1 == 1832)
							//							{
							//								MessageBox.Show("1832");
							//							}
						}
						catch (System.Exception ex)
						{
							int f1 = f+1;
							sr.WriteLine(ex.Message + " Linea " + f1.ToString());
						}
						if (chkClavePrimaria.Checked)
						{
							Funcion.EjecutaSQL(cdsArticulo, "Set IDENTITY_INSERT Compra off", false);
						}
						#endregion Compra Grabacion

						#region Grabacion de detalle de Compras
						if (bModifica)
						{
							stInsert = "DELETE DetCompra WHERE Numero = '" + stClave +"'";
						}

						stSelect = "Select idCompra from Compra Where Numero = '" + stClave + "'";
						int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect);
						stInsert = "INSERT INTO DetCompra (Cantidad, Impuesto, Precio, idCompra, idArticulo) ";
						stValor = " VALUES (1, 0, " + dTotal.ToString("0.00", us) + ", " 
							+ IdCompra.ToString() + ", " + IdArt.ToString() + ")";
						if (stInsert.Trim().EndsWith(",")) 
						{
							stInsert = stInsert.Substring(0,stInsert.Length - 2);
							stValor = stValor.Substring(0,stValor.Length - 2);
						}
						//						stInsert+= ") ";
						//						stValor += ") ";

						try
						{
							// Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT DetCompra on", false);
							Funcion.EjecutaSQL(cdsArticulo, stInsert + stValor, true);
						}
						catch (System.Exception ex)
						{
							int f1 = f+1;
							sr.WriteLine(ex.Message + " Linea " + f1.ToString());
						}
						#endregion Grabacion de Detalle de compras
					} // Inserta o Modifica Compra
					else 
					{
						MessageBox.Show("Codigo " + stClave + " ya existe");
						f++;
					}
					f --;
					Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT DetCompra off", false);
				} // Fin de Fila de Compra Existente
				#endregion Existe Valor de Celda
			} // Fin de for de lectura de filas
			#endregion Ingreso de Compras a Credito

			// Actualizacion de clientes
			string stProveedor = "0";
			if (bCxP) stProveedor = "1";
			stSelect = "update compra set Compra.idCliente = "
				+ "Cliente.idCliente From Compra Inner Join Cliente "
				+ "On Compra.Otro = Cliente.Codigo "
				+ "Where Proveedor = " + stProveedor;
			Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			
			// Actualiza notas de Debito
			if (!bCxP)
			{
				stSelect = "Update compra set compra.idTipoFactura = 27 " +
					"Where compra.idTipoFactura = 5 And compra.Numero like 'ND%'";
				Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
				stSelect = "Update compra set compra.idTipoFactura = 1 " +
					"Where compra.idTipoFactura = 0";
				Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			}
			else
			{
				stSelect = "Update compra set compra.idTipoFactura = 28 " +
					"WHere compra.idTipoFactura = 6 And compra.Numero like 'ND%'";
				Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
				stSelect = "Update compra set compra.idTipoFactura = 4 " +
					"WHere compra.idTipoFactura = 0";
				Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			}

			MessageBox.Show("Cuentas por Cobrar Terminado", "Información");
		}

		private bool bCxP = false;
		private void btCuentasxP_Click(object sender, System.EventArgs e)
		{

			#region Apertura de Archivo
			//			bCxP = true;
			//			btCuentasxC_Click(this, e);
			//			bCxP = false;
			this.txtMensaje.Text = "Formato: NumeroIngreso, NumFactura, idCobrador, Fecha, Concepto, idFormaPago, IdCuentaCorriente, Cheque, Pago, Interes";
			if (MessageBox.Show("¿Desea Subir Cobros de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stArchivo = "";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
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
			#endregion Apertura Archivo

			#region Lectura de Titulos
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

			C1.Data.C1DataRow drPago = cdsPagos.TableViews["Pago"].AddNew();
			string [] stTitulo = new string[100];
			string [] stTipoDato = new string[100];
			// Lectura de Titulos de columnas de Pagos
			XLCell cell = sheet.GetCell(0, 0);

			int iCamposCompra = 0;
			for (int i = 0; i < sheet.Columns.Count; i++)
			{
				cell = sheet.GetCell(0, i);
				if (cell == null) continue;
				try
				{
					stTitulo[i] = (string) cell.Value;
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error al leer titulo");
					return;
				}
				iCamposCompra++;
				if (stTitulo[i].ToUpper().Trim() == "NUMFACTURA") 
				{
					stTipoDato[i] = "String";
					continue;
				}
				//completo
				try
				{
					stTitulo[i] = "";
					stTitulo[i] = (string) cell.Value;

					drPago[stTitulo[i]].GetType();
					try
					{
						drPago[stTitulo[i]] = "a";
						stTipoDato[i] = "String";
					}
					catch
					{
						try
						{
							drPago[stTitulo[i]] = 1;
							stTipoDato[i] = "Double";
						}
						catch
						{
							stTipoDato[i] = "DateTime";
						}
					}
				}
				catch
				{
					MessageBox.Show("No existe campo " + stTitulo[i].Trim()
						+ ".\nPrimera fila debe tener los campos de Compras a importar.",
						"Error");
					return;
				} // Fin de try
			} // Fin de For de primera fila
			cdsPagos.RejectChanges();
			if (stTitulo[0].ToString().Trim().ToUpper() != "NUMEROINGRESO")
			{
				MessageBox.Show("Primera columna debe ser Numero de Ingreso", "Información");
				return;
			}

			#endregion Lectura de Titulos

			#region Actualizacion de Articulos
			string stSelect = "Select Count(*) From Articulo "
				+ "Where Articulo like 'Saldo%'";
			int IdArt = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
			Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			if (IdArt <= 0)
			{
				MessageBox.Show("No existe articulo Saldo inicial");
				return;
			}
			stSelect = "Select idArticulo From Articulo "
				+ "Where Articulo like 'Saldo%'";
			IdArt = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
			#endregion Actualizacion de Articulos

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Cobros'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			#region Ingreso de Cobros a Credito
			StreamWriter sr= File.CreateText("SalidaError.xls");
			string stClave = "";
			string stCodCliente = "";
			//Lectura de filas
			for (int f = 1; f < sheet.Rows.Count; f++)
			{
				cell = sheet.GetCell(f, 0);
				#region Existe Valor de Celda
				if (cell == null || cell.Value == null) continue;// Linea de Compra
				#region Verifica Nuevo o Modificacion
				string stNumero = "";
				if (cell.Value.GetType().Name.ToString().ToUpper() == "DOUBLE")
					stNumero = cell.Value.ToString();
				else stNumero =	(string) cell.Value;
				
				stClave = stNumero.Trim();
				stCodigo = "SELECT COUNT(*) FROM Pago WHERE NumeroIngreso = '" + stNumero + "'";
				
				int iBusca = 0;
				iBusca = Funcion.iEscalarSQL(cdsArticulo, stCodigo);
				string stInsert = "";
				bool bModifica = false;
				if (chkModificaExiste.Checked && iBusca == 0) 
					bModifica = true;

				if (bModifica)
					stInsert = "UPDATE Pago SET ";
				else
					stInsert = "INSERT INTO Pago (";

				string stValor = " VALUES (";
				#endregion Verifica Nuevo o Modificacion

				if (iBusca <= 0 || bModifica) // Inserta o Modifica Pago
				{
					#region Lectura de Columnas de Pago
					iCuenta ++;
					DateTime dtFechaLeida;
					for (int j = 0; j < iCamposCompra; j++)
					{
						cell = sheet.GetCell(f, j);
						if (cell != null && cell.Value != null)
						{
							#region try
							try
							{
								if (stTitulo[j].Trim().ToUpper() == "NUMFACTURA")
								{
									stInsert += "idCompra,";
									string stCelda = cell.Value.ToString().Trim();
									string stBusca = "Select idCompra From Compra Where Numero = '" + stCelda + "' And idTipoFactura = 1";
									int iNumFact = Funcion.iEscalarSQL(cdsCompra, stBusca, false);
									if (iNumFact == 0)
									{
										sr.WriteLine("No existe factura de Cobro " + stCelda + " en Linea " + f.ToString());
										continue;
									}
									stValor += iNumFact.ToString() + ", ";

									if (bModifica) stInsert += " = " + "'" + stCelda +"'";
									continue;
								}
								stInsert += stTitulo[j].Trim();

								if (stTipoDato[j] == "String")
								{
									string stCelda = "";
									if (cell.Value.GetType().Name == "Double")
										stCelda = cell.Value.ToString();
									else if (cell.Value.GetType().Name == "DateTime")
										stCelda = cell.Value.ToString().Substring(0,10);
									else
										stCelda = (string) cell.Value;
										
									stValor += "'" + stCelda +"'";
									if (bModifica) stInsert += " = " + "'" + stCelda +"'";
								}
								if (stTipoDato[j] == "DateTime")
								{
									dtFechaLeida = (DateTime) cell.Value;
									stValor += "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
									if (bModifica) stInsert += " = " + "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
								}
								if (stTipoDato[j] == "Double")
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
							} 
							#endregion  Fin de Try Catch
						} // Fin de verificacion de valor nulo
					} // Fin de lectura de columnas for

					if (stInsert.Trim().EndsWith(","))
					{
						stInsert = stInsert.Substring(0,stInsert.Length - 2);
						stValor = stValor.Substring(0,stValor.Length - 2);
					}
					if (!bModifica) stInsert+= ") ";
					stValor += ") ";
					//f++;
					#endregion Lectura de Columnas de compra

					#region Compra Grabacion
					try
					{
						if (bModifica) 
						{
							stValor = "";
							stInsert += " WHERE Numero = '" + stClave +"'";
							if (bCxP) 
								stInsert += " And Otro = '" + stCodCliente +"'";
						}
						string stTodo = stInsert + stValor;
						Funcion.EjecutaSQL(cdsArticulo, stTodo, true);
					}
					catch (System.Exception ex)
					{
						int f1 = f+1;
						sr.WriteLine(ex.Message + " Linea " + f1.ToString());
					}
					#endregion Compra Grabacion
				} // Fin de for de lectura de filas
				#endregion Ingreso de Compras a Credito

			}
			MessageBox.Show("Cuentas por Cobrar Terminado", "Información");
			#endregion Ingreso de Cobros a credito
		}

		private void DeshabilitaTrigger()
		{
			SqlCommand cmdActualiza;
			this.cdsCompra.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = 
				this.cdsCompra.Schema.Connections[0];

			// Borra blancos de Asiento
			miConeccion = this.cdsCompra.Schema.Connections[0];
			string stSelect = "ALTER TABLE DetCompra DISABLE TRIGGER SaldoCosto";
			//			Notas = RTrim(Notas),
			cmdActualiza = new SqlCommand (stSelect
				, (SqlConnection) miConeccion.DbConnection);
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();		
		}

		private void ImportarExcel_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Importa Excel'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			Cursor = Cursors.Default;
			cmbTipoFactura.DataSource = Funcion.dvProcedimiento(cdsArticulo,
				"Select idTipoFactura, Nombre From CompraNumero Where Grupo != 3 Order by Nombre");
			cmbTipoFactura.Value = 4;
		}

		private void btGrupoCliente_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Grupo";
			if (MessageBox.Show("¿Desea Subir Grupos de Clientes de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:GrupoCli'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			this.LeeExcel(this.cdsGrupoCliente, this.cdsGrupoCliente.TableViews["ClienteGrupo"], 
				this.cdvClienteGrupo, 0);
		}

		private void cdsGrupoCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupoCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void ImportarExcel_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Importa Excel'";
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

		private void btCargo_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Descripcion, Codigo";
			if (MessageBox.Show("¿Desea Subir Grupos de Cargos de Roles de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Cargo'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			this.LeeExcel(this.cdsCargo, this.cdsCargo.TableViews["RolCargo"], 
				this.cdvCargo, 0);
		}

		private void btDepartamento_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Descripcion";
			if (MessageBox.Show("¿Desea Subir Departamentos de Roles de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Depart'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			
			this.LeeExcel(this.cdsCargo, this.cdsCargo.TableViews["RolDepartamento"], 
				this.cdvDepartamento, 0);
		}

		private void btPersonal_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: cedula,	codigo,	nombre, direccion, idRolSucursal, fechanacimiento, idCargo, fechaingreso, cargasfamiliares";
			if (MessageBox.Show("¿Desea Subir Personal de Roles desde Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Personal'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			
			this.LeeExcel(this.cdsPersonal, this.cdsPersonal.TableViews["Personal"], 
				this.cdvPersonal, 0);		
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void cdsCargo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCargo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void btCarro_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Placa, Color, Anio, Marca, Moneda, "
				+ "Tamanio, Estado, Licencia, Millas";
			if (MessageBox.Show("¿Desea Subir Carros de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Carros'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			LeeExcel(cdsCarros, cdsCarros.TableViews["Carros"], cdvCarros, 0);
		}

		private void cdsCarros_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCarros.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btOrdenMov_Click(object sender, System.EventArgs e)
		{
			txtMensaje.Text = "Campos: idArticulo, idCarro, Fecha, Cabezal, Termoking, "
				+ "Galones, ValorGalon, Galon1, ValorGalon1";
			if (MessageBox.Show("¿Desea Subir Artículos de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:OrdenMov'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			LeeExcel(cdsCarros, cdsCarros.TableViews["OrdenMovilizacion"], cdvOrdenMov, 0);		
			string stSelect = "Update OrdenMovilizacion SET idArticulo = "
				+ "Articulo.idArticulo FROM OrdenMovilizacion, Articulo "
				+ "Where OrdenMovilizacion.CodArt = Articulo.Codigo";
			Funcion.EjecutaSQL(cdsArticulo, stSelect);
		}

		private void btSucursal_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: idRolSucursal, Sucursal";
			if (MessageBox.Show("¿Desea Subir Sucursales de Roles de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Sucursal'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			this.LeeExcel(this.cdsCargo, this.cdsCargo.TableViews["RolSucursal"], 
				this.cdvSucursal, 0);		
		}

		private void btMarca_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: idMarca, Marca";
			if (MessageBox.Show("¿Desea Subir Marcas de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Marca'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			this.LeeExcel(cdsGrupoArt, cdsGrupoArt.TableViews["articuloMarca"], 
				this.cdvMarca, 0);		
		}

		private void btArancel_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: idArancel, Codigo, Porcentaje, arancel";
			if (MessageBox.Show("¿Desea Subir Marcas de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Arancel'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			this.LeeExcel(cdsGrupoArt, cdsGrupoArt.TableViews["articuloArancel"], 
				this.cdvArancel, 0);				
		}

		private void btUnidad_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: idUnidad, Unidad";
			if (MessageBox.Show("¿Desea Subir Marcas de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Unidad'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			this.LeeExcel(cdsGrupoArt, cdsGrupoArt.TableViews["articuloUnidad"], 
				this.cdvUnidad, 0);		
		}

		private void btInvSaldo_Click(object sender, System.EventArgs e)
		{
			txtMensaje.Text = "Antes de usar esta opción debe crear una factura de compra ";
			if (MessageBox.Show("¿Desea Generar Inventario con Saldo y costoPromedio de Artículos?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;
			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Inventario'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			int IdCompra = Funcion.iEscalarSQL(cdsArticulo, 
				"Select Top 1 IdCompra From Compra Where idTipoFactura = 4 Order by IdCompra desc");
			string stSelect = "Insert Into DetCompra (Cantidad, Precio, idArticulo, idCompra) "
				+ " Select Articulo.Saldo, Articulo.CostoPromedio, Articulo.idArticulo, " 
				+ IdCompra.ToString() + " From Articulo";
			Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			MessageBox.Show("Inventario con Saldo y costoPromedio de Artículos Subido",
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Question);
		}

		private void btCuentaxC_Click(object sender, System.EventArgs e)
		{
			//factura, fecha, cliente, saldo, diascredito, iva
		}

		private void btRolConcepto_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Descripcion, Codigo";
			if (MessageBox.Show("¿Desea Subir Grupos de Conceptos de Roles de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:RolConc'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			this.LeeExcel(this.cdsCargo, this.cdsRolConcepto.TableViews["RolConcepto"], 
				this.cdvRolConcepto, 0);
		}

		private void btTomaFisica_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Codigo, idTmpTomaFisica, Cantidad";
			if (MessageBox.Show("¿Desea Subir Toma Física?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:TomaFis'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			
			LeeExcel(cdsToma, cdsToma.TableViews["TmpTomaDetalle"], 
				cdvTomaFisica, 0);
			string stSelect = "Update TmpTomaDetalle Set idArticulo = a.idArticulo "
				+ "From Articulo a Inner Join TmpTomaDetalle t On t.Codigo = a.Codigo ";
			Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			this.txtMensaje.Text = "Campos: Cedula, Mes, Anio, Nombre, Ingreso1,Ingreso2..., Egreso1,Egreso2..., Valore1,Valor2..., etc";
			if (MessageBox.Show("¿Desea Subir Valores para Roles de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;
			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:RolIngr'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			
			this.LeeExcelRol(this.cdsRolIngreso, 
				this.cdsRolIngreso.TableViews["RolPago"], this.cdvRolIngreso,0);
			string stExec = "Exec RolActualizaRuc ";
			Funcion.EjecutaSQL(cdsArticulo, stExec, true);
			MessageBox.Show("Valores de Rubros subido", "Información");
			//			this.LeeExcelRol(this.cdsRolIngreso, 
			//				this.cdsRolIngreso.TableViews["RolPago"], this.cdvRolIngreso,0);
		}

		private bool LeeExcelRol(C1.Data.C1DataSet ds, C1.Data.C1DataTable dtArchivo, 
			C1.Data.C1DataView dv, int iTipo)
		{
			string stTabla = dtArchivo.Name.ToString();
			this.Cursor = Cursors.WaitCursor;
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
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
			string [] stTitulo = new string[500];
			string [] stTipoDato = new string[500];
			// Lectura de Titulos de columnas
			for (int i = 1; i < sheet.Columns.Count; i++)
			{
				XLCell cell = sheet.GetCell(0, i);
				//				if (cell != null && cell.Value != null)
			{
				try
				{
					stTitulo[i] = "";
					stTitulo[i] = (string) cell.Value;

					drArticulo[stTitulo[i]].GetType();
					try
					{
						drArticulo[stTitulo[i]] = "a";
						stTipoDato[i] = "String";
					}
					catch
					{
						try
						{
							drArticulo[stTitulo[i]] = 1;
							stTipoDato[i] = "Double";
						}
						catch
						{
							stTipoDato[i] = "Fecha";
						}
					}
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
					else
					{
						int ColError = i + 1;
						MessageBox.Show("No existe campo " + stTitulo[i]
							+ " en columna " + ColError.ToString().Trim(),
							"Error en nombres de Campos");
						this.Cursor = Cursors.Default;
						return false;
					}
					//					i = sheet.Columns.Count;
				} // Fin de Primer try
			} // Fin de If
			} // Fin de For de primera fila
			ds.RejectChanges();
			// FIN de Verificacion que primera fila tiene los nombres de los campos a subirse

			// Lectura de datos desde segunda fila
			this.Cursor = Cursors.WaitCursor;
			//			if (File.Exists("SalidaError.txt")) File.Delete("SalidaError.txt");
			StreamWriter sr;
			try
			{
				sr= File.CreateText("SalidaError.XLS");
			}
			catch
			{
				MessageBox.Show("Esta abierto el archivo SalidaError.xls");
				Cursor = Cursors.Default;
				return false;
			}

			//	dv.Sort = stTitulo[0];
			for (int f = 1; f < sheet.Rows.Count; f++)
			{
				// Primera columna tiene el codigo que no se debe repetir
				int iBusca = -1;
				stCodigo = "";
				XLCell cell = sheet.GetCell(f, 0);
				XLCell cellMes = sheet.GetCell(f, 1);
				XLCell cellAnio = sheet.GetCell(f, 2);
				if (cell != null && cell.Value != null)
				{
					if (cell.Value.GetType().Name == "String")
						stCodigo = (string) cell.Value;
					if (cell.Value.GetType().Name == "Double")
						stCodigo = (string) cell.Value.ToString();

					string stMes = "";
					string stAnio = "";
					if (cellMes.Value.GetType().Name == "String" || cellMes.Value.GetType().Name == "Double")
						stMes = cellMes.Value.ToString();
					if (cellAnio.Value.GetType().Name == "String"  || cellAnio.Value.GetType().Name == "Double")
						stAnio = cellAnio.Value.ToString();

					stCodigo = stCodigo.Trim();
					stCodigo = stCodigo.Replace("'", "''");
					string stIdPersonal = "Select idPersonal From Personal Where Cedula = '"
						+ stCodigo + "'";
					int IdPersonal = Funcion.iEscalarSQL(cdsArticulo, stIdPersonal, true);
					
					string stSelect = "SELECT COUNT(*) FROM "
						+ stTabla + " WHERE idPersonal = " 
						+ IdPersonal.ToString() + " And Mes = " + stMes + " And Anio = " + stAnio;
					iBusca = Funcion.iEscalarSQL(cdsArticulo, stSelect);

					if (iBusca <= 0 || chkModificaExiste.Checked) // Articulo no encontrado o modifica existente
					{
						iCuenta ++;

						//						drArticulo[stTitulo[0]] = stCodigo;

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
									if (j==0) 
									{
										stInsert += "idPersonal";
									}
									stInsert += stTitulo[j];
									// Numero grabado en campo de texto
									if ((stTipoDato[j] == "String" && cell.Value.GetType().Name == "Double") || j == 0)
									{
										if (j!=0)
										{
											string stTexto = cell.Value.ToString();
											stValor += "'" + stTexto + "'";
											if (bModifica) stInsert += " = " + "'" + stTexto + "'";
										}
										else
										{
											string stCedula = cell.Value.ToString();
											//											stIdPersonal = "Select idPersonal From Personal where Cedula = '"
											//												+ stCedula + "'";
											//											int IdPers = Funcion.iEscalarSQL(cdsArticulo, stIdPersonal, true);
											stValor += "'" + IdPersonal + "'";
											if (bModifica) stInsert += " = " + "'" + IdPersonal + "'";
										}
									}
									else if (cell.Value.GetType().Name == "String")
									{
										string stTexto = (string) cell.Value;
										stTexto = stTexto.Replace("'", "''");
										stValor += "'" + stTexto + "'";
										if (bModifica) stInsert += " = " + "'" + stTexto + "'";
									}
									else if (cell.Value.GetType().Name == "DateTime")
									{
										dtFechaLeida = (DateTime) cell.Value;
										stValor += "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
										if (bModifica) stInsert += " = " + "'" + dtFechaLeida.ToString("yyyMMdd") + "'";
									}
									else if (cell.Value.GetType().Name == "Double")
									{
										double dValor = (double) cell.Value;
										stValor += dValor.ToString("0.00", us);
										if (bModifica) stInsert += " = " + dValor.ToString("0.00", us);
									}
									else if (cell.Value.GetType().Name == "Boolean")
									{
										string stTexto = "0";
										if ((bool) cell.Value)
											stTexto = "1";
										stValor += stTexto;
										if (bModifica) stInsert += " = " + stTexto;
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
						if (iTipo > 0) // Cliente o Proveedor
						{
							int iProv = iTipo - 1;
							if (!bModifica) stInsert += ", Proveedor";
							stValor += ", " + iProv.ToString();
						}
						if (!bModifica) stInsert+= ") ";
						stValor += ") ";
						#endregion

						//						if (iCuenta == 6)
						//						{
						//							MessageBox.Show(stInsert);
						//						}

						try
						{
							if (bModifica) 
							{
								stValor = "";
								stInsert += " WHERE idPersonal = '" + IdPersonal.ToString() + "' and Mes = '" + stMes + "' and Anio = '" + stAnio +"'";
							}
							if (chkClavePrimaria.Checked)
							{
								Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT " + stTabla + " on", false);
							}

							stInsert += stValor;
							Funcion.EjecutaSQL(cdsArticulo, stInsert, true);
							if (chkClavePrimaria.Checked)
							{
								Funcion.EjecutaSQL(cdsArticulo, "set IDENTITY_INSERT " + stTabla + " off", false);
							}
							//cdsArticulo.Update();
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
				MessageBox.Show("No hay Registros para Actualizar", 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return false;
			}
			MessageBox.Show(iCuenta.ToString().Trim() + " Registros Migrados desde Excel", 
				"Información", MessageBoxButtons.OK,
				MessageBoxIcon.Information);
			return true;
		}

		private void Cobros(int idTipoFactura)
		{
			this.txtMensaje.Text = "Campos: Factura, Codigo, Pago, Fecha, Cheque, CuentaCorriente, Asiento, FormaPago, Cobrador, Concepto, Referencia";
			if (MessageBox.Show("¿Desea Subir Cobros de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			stArchivo = "";
			#region Apertura de Archivo de Excel
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
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

			int iCuenta = 0;
			XLSheet sheet = book1.Sheets[0];

			for (int f = 1; f < sheet.Rows.Count; f++)
			{
				// Primera columna tiene el numero de factura que debe Existir
				XLCell cell = sheet.GetCell(f, 0);
				if (cell != null && cell.Value != null)
				{
					string stFactura = cell.Value.ToString();
					string stBusca = "Select idCompra From Compra "
						+ "Where Numero = '" + stFactura + "' And idTipoFactura = " + idTipoFactura;
					int iBusca = Funcion.iEscalarSQL(cdsArticulo, stBusca, false);
					if (iBusca > 0) // Compra encontrada
					{
						#region Compra Existe
						iCuenta ++;
						string stMensajeCol = "Col 2: Codigo";
						try
						{
							string stCodigo = sheet.GetCell(f, 1) == null ? "" : sheet.GetCell(f, 1).Value.ToString();
							stMensajeCol = "Col 3: Pago";
							double dPago = sheet.GetCell(f, 2) == null ? 0 : double.Parse(sheet.GetCell(f, 2).Value.ToString());
							stMensajeCol = "Col 4: Fecha";
							DateTime dtFecha = sheet.GetCell(f, 3) == null ? DateTime.Today : DateTime.Parse(sheet.GetCell(f, 3).Value.ToString());
							stMensajeCol = "Col 5: Cheque";
							string stCheque = sheet.GetCell(f, 4) == null ? "" : sheet.GetCell(f, 4).Value.ToString();
							stMensajeCol = "Col 6: Cuenta Corriente";
							string stCuentaCorr = sheet.GetCell(f, 5) == null ? "" : sheet.GetCell(f, 5).Value.ToString();
							stMensajeCol = "Col 7: Asiento";
							string stAsiento = sheet.GetCell(f, 6) == null ? "" : sheet.GetCell(f, 6).Value.ToString();
							stMensajeCol = "Col 8: FormaPago";
							string stFormaPago = sheet.GetCell(f, 7) == null ? "" : sheet.GetCell(f, 7).Value.ToString();
							stMensajeCol = "Col 9: Cobrador";
							string stCobrador = sheet.GetCell(f, 8) == null ? "" : sheet.GetCell(f, 8).Value.ToString();
							stMensajeCol = "Col 10: Concepto";
							string stConcepto = sheet.GetCell(f, 9) == null ? "" : sheet.GetCell(f, 9).Value.ToString();
							stMensajeCol = "Col 11: Referencia";
							string stReferencia = sheet.GetCell(f, 10) == null ? "" : sheet.GetCell(f, 10).Value.ToString();
							stMensajeCol = "Col 12: Banco";
							string stBanco = sheet.GetCell(f, 11) == null ? "" : sheet.GetCell(f, 11).Value.ToString();

							stMensajeCol = "Generacion de Pago";
							//	ALTER PROCEDURE [dbo].[CreaPago] @FACTURA NVARCHAR(20), @IDTIPO INT, @NOMBRECLIENTE NVARCHAR(60),
							//	@PAGO FLOAT, @CHEQUE NVARCHAR(20) = '', @FECHA DATETIME = '19000101', 
							//	@CREACLIENTE INT = 0, @CUENTACORR NVARCHAR(20) = '', @REFERENCIA NVARCHAR(20) = '', 
							//	@Numero nVarChar(20) = '', @Asiento nVarChar(20) = '', @Banco nVarChar(20) = '',
							//	@FormaPago nVarChar(20) = '', @Cobrador nVarChar(20) = '', @Concepto nVarChar(20) = ''
							string stExec = string.Format("Exec PagoCrear '{0}', {1}, '{2}', {3}, '{4}', '{5}', {6}, '{7}', "
								+ " '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}' ", 
								stFactura, idTipoFactura, "", dPago.ToString("0.00000", us), stCheque, dtFecha.ToString("yyyyMMdd"), -2, stCuentaCorr, stReferencia, 
								stCodigo, stAsiento, stBanco, stFormaPago, stCobrador, stConcepto);
							Funcion.EjecutaSQL(cdsArticulo, stExec);
						}
						catch (System.Exception ex)
						{
							int FilaError = f + 1;
							MessageBox.Show("Error en Fila " + FilaError.ToString()
								+  " " + stMensajeCol + " " + ex.Message, "Error Subiendo Pagos");
							return;
						}
						#endregion Compra Existente
					} // Fin de Creacion de Compra
					else
					{
						MessageBox.Show("Factura de Codigo: " + stFactura + " no existe");
					}
				} // Fin de lectura de Fila
			} // Fin de Lectura de Archivo

			//			string stExec = "EXEC ArticuloActualizaTipo";
			//			Funcion.EjecutaSQL(cdsArticulo, stExec, true);
			MessageBox.Show("Pagos Subidos",
				"Confirmación", MessageBoxButtons.OK, 
				MessageBoxIcon.Information);
		}

		private void btPagos_Click(object sender, System.EventArgs e)
		{
			Cobros(4);
			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Pagos'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void btCobros_Click(object sender, System.EventArgs e)
		{
			Cobros(1);		
			string stAudita = "Exec AuditaCrear 53, 6, 'Excel:Cobros'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void btCompraRet_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btAvanzado_Click(object sender, System.EventArgs e)
		{
			ImportaBlasar miAvanzado = new ImportaBlasar();
			miAvanzado.MdiParent = this.MdiParent;
			miAvanzado.Show();
		}
	}

}
