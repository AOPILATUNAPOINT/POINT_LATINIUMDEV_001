using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.OleDb;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de SaldosIniciales.
	/// </summary>
	public class CopiarArchivos : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPlanCuentas;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSeguridad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCuentaBanco;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkBancos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPropiedades;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAsignaCuenta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPersonal;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkReceta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGrupoArt;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkClientes;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGrupoCliente;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProyecto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSubProyecto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkArticulos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodo;
		private C1.Data.C1DataSet cdsContabilidad;
		private C1.Data.C1DataSet cdsFacturacion;
		private C1.Data.C1DataSet cdsRoles;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private C1.Data.C1DataSet cdsEmpresa;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataView cdvEmpresa;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEmpresa;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAsiento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAgregar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSustento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRolRubro;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFactura;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRetencion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPagos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSobrescribir;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDistribucion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTablaRet;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkImportacion;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFecha;
		private Infragistics.Win.Misc.UltraButton btBorraBase;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClave;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRolPago;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CopiarArchivos()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Empresa", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contabilidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numeracion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Roles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoDB");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reportes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Servidor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fox");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruta");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			this.chkPlanCuentas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSeguridad = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCuentaBanco = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkBancos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPropiedades = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAsignaCuenta = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPersonal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkReceta = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkGrupoArt = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkClientes = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkGrupoCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkProyecto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSubProyecto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkArticulos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.chkTodo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsContabilidad = new C1.Data.C1DataSet();
			this.cdsFacturacion = new C1.Data.C1DataSet();
			this.cdsRoles = new C1.Data.C1DataSet();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsEmpresa = new C1.Data.C1DataSet();
			this.cmbEmpresa = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvEmpresa = new C1.Data.C1DataView();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.chkAsiento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAgregar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSustento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkRolRubro = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.chkFactura = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkRetencion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPagos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSobrescribir = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkFormaPago = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDistribucion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkTablaRet = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkImportacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btBorraBase = new Infragistics.Win.Misc.UltraButton();
			this.cmbFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.txtClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkRolPago = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsContabilidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFacturacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRoles)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).BeginInit();
			this.SuspendLayout();
			// 
			// chkPlanCuentas
			// 
			this.chkPlanCuentas.Location = new System.Drawing.Point(40, 121);
			this.chkPlanCuentas.Name = "chkPlanCuentas";
			this.chkPlanCuentas.Size = new System.Drawing.Size(120, 21);
			this.chkPlanCuentas.TabIndex = 0;
			this.chkPlanCuentas.Text = "Plan de Cuentas";
			// 
			// chkSeguridad
			// 
			this.chkSeguridad.Location = new System.Drawing.Point(40, 250);
			this.chkSeguridad.Name = "chkSeguridad";
			this.chkSeguridad.Size = new System.Drawing.Size(120, 21);
			this.chkSeguridad.TabIndex = 1;
			this.chkSeguridad.Text = "Seguridades";
			this.chkSeguridad.CheckedChanged += new System.EventHandler(this.chkSeguridad_CheckedChanged);
			// 
			// chkCuentaBanco
			// 
			this.chkCuentaBanco.Location = new System.Drawing.Point(40, 224);
			this.chkCuentaBanco.Name = "chkCuentaBanco";
			this.chkCuentaBanco.Size = new System.Drawing.Size(120, 22);
			this.chkCuentaBanco.TabIndex = 2;
			this.chkCuentaBanco.Text = "Cuentas Bancarias";
			// 
			// chkBancos
			// 
			this.chkBancos.Location = new System.Drawing.Point(40, 198);
			this.chkBancos.Name = "chkBancos";
			this.chkBancos.Size = new System.Drawing.Size(120, 22);
			this.chkBancos.TabIndex = 3;
			this.chkBancos.Text = "Bancos";
			// 
			// chkPropiedades
			// 
			this.chkPropiedades.Location = new System.Drawing.Point(40, 172);
			this.chkPropiedades.Name = "chkPropiedades";
			this.chkPropiedades.Size = new System.Drawing.Size(120, 22);
			this.chkPropiedades.TabIndex = 4;
			this.chkPropiedades.Text = "Propiedades";
			// 
			// chkAsignaCuenta
			// 
			this.chkAsignaCuenta.Location = new System.Drawing.Point(40, 146);
			this.chkAsignaCuenta.Name = "chkAsignaCuenta";
			this.chkAsignaCuenta.Size = new System.Drawing.Size(208, 22);
			this.chkAsignaCuenta.TabIndex = 5;
			this.chkAsignaCuenta.Text = "Asignación de Cuentas, Distribución";
			// 
			// chkPersonal
			// 
			this.chkPersonal.Location = new System.Drawing.Point(40, 276);
			this.chkPersonal.Name = "chkPersonal";
			this.chkPersonal.Size = new System.Drawing.Size(192, 21);
			this.chkPersonal.TabIndex = 6;
			this.chkPersonal.Text = "Personal, Cargo, Departamento";
			// 
			// chkReceta
			// 
			this.chkReceta.Location = new System.Drawing.Point(272, 276);
			this.chkReceta.Name = "chkReceta";
			this.chkReceta.Size = new System.Drawing.Size(160, 21);
			this.chkReceta.TabIndex = 13;
			this.chkReceta.Text = "Recetas y Productos";
			// 
			// chkGrupoArt
			// 
			this.chkGrupoArt.Location = new System.Drawing.Point(272, 146);
			this.chkGrupoArt.Name = "chkGrupoArt";
			this.chkGrupoArt.Size = new System.Drawing.Size(144, 22);
			this.chkGrupoArt.TabIndex = 12;
			this.chkGrupoArt.Text = "Grupo de Artículos";
			// 
			// chkClientes
			// 
			this.chkClientes.Location = new System.Drawing.Point(272, 172);
			this.chkClientes.Name = "chkClientes";
			this.chkClientes.Size = new System.Drawing.Size(144, 22);
			this.chkClientes.TabIndex = 11;
			this.chkClientes.Text = "Distribuidores y Clientes";
			// 
			// chkGrupoCliente
			// 
			this.chkGrupoCliente.Location = new System.Drawing.Point(272, 198);
			this.chkGrupoCliente.Name = "chkGrupoCliente";
			this.chkGrupoCliente.Size = new System.Drawing.Size(120, 22);
			this.chkGrupoCliente.TabIndex = 10;
			this.chkGrupoCliente.Text = "Grupo de Clientes";
			// 
			// chkProyecto
			// 
			this.chkProyecto.Location = new System.Drawing.Point(272, 224);
			this.chkProyecto.Name = "chkProyecto";
			this.chkProyecto.Size = new System.Drawing.Size(120, 22);
			this.chkProyecto.TabIndex = 9;
			this.chkProyecto.Text = "Centros de Costo";
			// 
			// chkSubProyecto
			// 
			this.chkSubProyecto.Location = new System.Drawing.Point(272, 250);
			this.chkSubProyecto.Name = "chkSubProyecto";
			this.chkSubProyecto.Size = new System.Drawing.Size(120, 21);
			this.chkSubProyecto.TabIndex = 8;
			this.chkSubProyecto.Text = "Proyecto";
			// 
			// chkArticulos
			// 
			this.chkArticulos.Location = new System.Drawing.Point(272, 121);
			this.chkArticulos.Name = "chkArticulos";
			this.chkArticulos.Size = new System.Drawing.Size(136, 21);
			this.chkArticulos.TabIndex = 7;
			this.chkArticulos.Text = "Artículos";
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(120, 388);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 24);
			this.btAceptar.TabIndex = 14;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(424, 388);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 15;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// chkTodo
			// 
			appearance1.FontData.BoldAsString = "True";
			this.chkTodo.Appearance = appearance1;
			this.chkTodo.Location = new System.Drawing.Point(64, 95);
			this.chkTodo.Name = "chkTodo";
			this.chkTodo.Size = new System.Drawing.Size(120, 21);
			this.chkTodo.TabIndex = 16;
			this.chkTodo.Text = "Seleccionar Todo";
			this.chkTodo.CheckedChanged += new System.EventHandler(this.chkTodo_CheckedChanged);
			// 
			// cdsContabilidad
			// 
			this.cdsContabilidad.BindingContextCtrl = this;
			this.cdsContabilidad.DataLibrary = "LibContabilidad";
			this.cdsContabilidad.DataLibraryUrl = "";
			this.cdsContabilidad.DataSetDef = "dsSeteoC";
			this.cdsContabilidad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsContabilidad.Name = "cdsContabilidad";
			this.cdsContabilidad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsContabilidad.SchemaDef = null;
			this.cdsContabilidad.BeforeFill += new C1.Data.FillEventHandler(this.cdsContabilidad_BeforeFill);
			// 
			// cdsFacturacion
			// 
			this.cdsFacturacion.BindingContextCtrl = this;
			this.cdsFacturacion.DataLibrary = "LibFacturacion";
			this.cdsFacturacion.DataLibraryUrl = "";
			this.cdsFacturacion.DataSetDef = "dsSeteoF";
			this.cdsFacturacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsFacturacion.Name = "cdsFacturacion";
			this.cdsFacturacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFacturacion.SchemaDef = null;
			this.cdsFacturacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsFacturacion_BeforeFill);
			// 
			// cdsRoles
			// 
			this.cdsRoles.BindingContextCtrl = this;
			this.cdsRoles.DataLibrary = "LibPersonal";
			this.cdsRoles.DataLibraryUrl = "";
			this.cdsRoles.DataSetDef = "dsPersonal";
			this.cdsRoles.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRoles.Name = "cdsRoles";
			this.cdsRoles.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRoles.SchemaDef = null;
			this.cdsRoles.BeforeFill += new C1.Data.FillEventHandler(this.cdsRoles_BeforeFill);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(40, 43);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(101, 15);
			this.ultraLabel1.TabIndex = 18;
			this.ultraLabel1.Text = "Copiar Archivos de:";
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
			// cmbEmpresa
			// 
			this.cmbEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEmpresa.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresa.DataSource = this.cdvEmpresa;
			this.cmbEmpresa.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 165;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 19;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 56;
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
																										 ultraGridColumn13});
			this.cmbEmpresa.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEmpresa.DisplayMember = "Empresa";
			this.cmbEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEmpresa.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresa.Location = new System.Drawing.Point(144, 43);
			this.cmbEmpresa.Name = "cmbEmpresa";
			this.cmbEmpresa.Size = new System.Drawing.Size(184, 21);
			this.cmbEmpresa.TabIndex = 19;
			this.cmbEmpresa.ValueMember = "";
			this.cmbEmpresa.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEmpresa_InitializeLayout);
			// 
			// cdvEmpresa
			// 
			this.cdvEmpresa.BindingContextCtrl = this;
			this.cdvEmpresa.DataSet = this.cdsEmpresa;
			this.cdvEmpresa.Sort = "Empresa";
			this.cdvEmpresa.TableName = "";
			this.cdvEmpresa.TableViewName = "Empresa";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// chkAsiento
			// 
			this.chkAsiento.Location = new System.Drawing.Point(40, 327);
			this.chkAsiento.Name = "chkAsiento";
			this.chkAsiento.Size = new System.Drawing.Size(80, 22);
			this.chkAsiento.TabIndex = 20;
			this.chkAsiento.Text = "Asientos";
			// 
			// chkAgregar
			// 
			appearance2.FontData.BoldAsString = "True";
			this.chkAgregar.Appearance = appearance2;
			this.chkAgregar.Location = new System.Drawing.Point(448, 345);
			this.chkAgregar.Name = "chkAgregar";
			this.chkAgregar.Size = new System.Drawing.Size(120, 21);
			this.chkAgregar.TabIndex = 21;
			this.chkAgregar.Text = "Agregar Registros";
			this.chkAgregar.CheckedChanged += new System.EventHandler(this.chkAgregar_CheckedChanged);
			// 
			// chkSustento
			// 
			this.chkSustento.Location = new System.Drawing.Point(448, 198);
			this.chkSustento.Name = "chkSustento";
			this.chkSustento.Size = new System.Drawing.Size(144, 22);
			this.chkSustento.TabIndex = 22;
			this.chkSustento.Text = "Sustento Comprobante";
			// 
			// chkRolRubro
			// 
			this.chkRolRubro.Location = new System.Drawing.Point(448, 224);
			this.chkRolRubro.Name = "chkRolRubro";
			this.chkRolRubro.Size = new System.Drawing.Size(120, 22);
			this.chkRolRubro.TabIndex = 23;
			this.chkRolRubro.Text = "Rol Rubros";
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(24, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(584, 25);
			this.label1.TabIndex = 24;
			this.label1.Text = "Para usar esta opción es necesario obtener respaldos de la información de Origen " +
				"y Destino";
			// 
			// chkFactura
			// 
			this.chkFactura.Location = new System.Drawing.Point(272, 327);
			this.chkFactura.Name = "chkFactura";
			this.chkFactura.Size = new System.Drawing.Size(80, 22);
			this.chkFactura.TabIndex = 25;
			this.chkFactura.Text = "Facturas";
			// 
			// chkRetencion
			// 
			this.chkRetencion.Location = new System.Drawing.Point(272, 353);
			this.chkRetencion.Name = "chkRetencion";
			this.chkRetencion.Size = new System.Drawing.Size(104, 22);
			this.chkRetencion.TabIndex = 27;
			this.chkRetencion.Text = "Retenciones";
			// 
			// chkPagos
			// 
			this.chkPagos.Location = new System.Drawing.Point(40, 353);
			this.chkPagos.Name = "chkPagos";
			this.chkPagos.Size = new System.Drawing.Size(80, 22);
			this.chkPagos.TabIndex = 26;
			this.chkPagos.Text = "Pagos";
			// 
			// chkSobrescribir
			// 
			appearance3.FontData.BoldAsString = "True";
			this.chkSobrescribir.Appearance = appearance3;
			this.chkSobrescribir.Location = new System.Drawing.Point(424, 95);
			this.chkSobrescribir.Name = "chkSobrescribir";
			this.chkSobrescribir.Size = new System.Drawing.Size(104, 21);
			this.chkSobrescribir.TabIndex = 28;
			this.chkSobrescribir.Text = "Sobreescribir";
			// 
			// chkFormaPago
			// 
			this.chkFormaPago.Location = new System.Drawing.Point(448, 172);
			this.chkFormaPago.Name = "chkFormaPago";
			this.chkFormaPago.Size = new System.Drawing.Size(120, 22);
			this.chkFormaPago.TabIndex = 29;
			this.chkFormaPago.Text = "Forma de Pago";
			// 
			// chkDistribucion
			// 
			this.chkDistribucion.Location = new System.Drawing.Point(448, 146);
			this.chkDistribucion.Name = "chkDistribucion";
			this.chkDistribucion.Size = new System.Drawing.Size(152, 22);
			this.chkDistribucion.TabIndex = 30;
			this.chkDistribucion.Text = "Distribución  Porcentual";
			// 
			// chkTablaRet
			// 
			this.chkTablaRet.Location = new System.Drawing.Point(448, 121);
			this.chkTablaRet.Name = "chkTablaRet";
			this.chkTablaRet.Size = new System.Drawing.Size(144, 21);
			this.chkTablaRet.TabIndex = 31;
			this.chkTablaRet.Text = "Retención  Iva Renta";
			// 
			// chkImportacion
			// 
			this.chkImportacion.Location = new System.Drawing.Point(448, 250);
			this.chkImportacion.Name = "chkImportacion";
			this.chkImportacion.Size = new System.Drawing.Size(120, 21);
			this.chkImportacion.TabIndex = 32;
			this.chkImportacion.Text = "Importaciones";
			// 
			// btBorraBase
			// 
			this.btBorraBase.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBorraBase.Location = new System.Drawing.Point(368, 43);
			this.btBorraBase.Name = "btBorraBase";
			this.btBorraBase.Size = new System.Drawing.Size(112, 25);
			this.btBorraBase.TabIndex = 33;
			this.btBorraBase.Text = "Borrar Antes de ...";
			this.btBorraBase.Click += new System.EventHandler(this.ultraButton1_Click);
			this.btBorraBase.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btBorraBase_MouseUp);
			// 
			// cmbFecha
			// 
			this.cmbFecha.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFecha.FormatString = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(496, 43);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.Size = new System.Drawing.Size(96, 22);
			this.cmbFecha.TabIndex = 34;
			// 
			// txtClave
			// 
			this.txtClave.Location = new System.Drawing.Point(496, 69);
			this.txtClave.Name = "txtClave";
			this.txtClave.PasswordChar = '*';
			this.txtClave.Size = new System.Drawing.Size(96, 22);
			this.txtClave.TabIndex = 35;
			this.txtClave.Visible = false;
			// 
			// chkRolPago
			// 
			this.chkRolPago.Location = new System.Drawing.Point(448, 276);
			this.chkRolPago.Name = "chkRolPago";
			this.chkRolPago.Size = new System.Drawing.Size(120, 21);
			this.chkRolPago.TabIndex = 36;
			this.chkRolPago.Text = "Rol Pago";
			// 
			// CopiarArchivos
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(632, 436);
			this.Controls.Add(this.chkRolPago);
			this.Controls.Add(this.txtClave);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btBorraBase);
			this.Controls.Add(this.chkImportacion);
			this.Controls.Add(this.chkTablaRet);
			this.Controls.Add(this.chkDistribucion);
			this.Controls.Add(this.chkFormaPago);
			this.Controls.Add(this.chkSobrescribir);
			this.Controls.Add(this.chkRetencion);
			this.Controls.Add(this.chkPagos);
			this.Controls.Add(this.chkFactura);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.chkRolRubro);
			this.Controls.Add(this.chkSustento);
			this.Controls.Add(this.chkAgregar);
			this.Controls.Add(this.chkAsiento);
			this.Controls.Add(this.cmbEmpresa);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.chkTodo);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.chkReceta);
			this.Controls.Add(this.chkGrupoArt);
			this.Controls.Add(this.chkClientes);
			this.Controls.Add(this.chkGrupoCliente);
			this.Controls.Add(this.chkProyecto);
			this.Controls.Add(this.chkSubProyecto);
			this.Controls.Add(this.chkArticulos);
			this.Controls.Add(this.chkPersonal);
			this.Controls.Add(this.chkAsignaCuenta);
			this.Controls.Add(this.chkPropiedades);
			this.Controls.Add(this.chkBancos);
			this.Controls.Add(this.chkCuentaBanco);
			this.Controls.Add(this.chkSeguridad);
			this.Controls.Add(this.chkPlanCuentas);
			this.MaximizeBox = false;
			this.Name = "CopiarArchivos";
			this.Text = "Copiar Archivos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.CopiarArchivos_Closing);
			this.Load += new System.EventHandler(this.SaldosIniciales_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsContabilidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFacturacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRoles)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void chkTodo_CheckedChanged(object sender, System.EventArgs e)
		{
			bool bHabilita = chkTodo.Checked;
			chkPlanCuentas.Checked = bHabilita;
			chkTablaRet.Checked = bHabilita;
			chkDistribucion.Checked = bHabilita;
			chkImportacion.Checked = bHabilita;
			chkFormaPago.Checked = bHabilita;
			chkArticulos.Checked = bHabilita;
			chkAsignaCuenta.Checked = bHabilita;
			chkBancos.Checked = bHabilita;
			chkClientes.Checked = bHabilita;
			chkCuentaBanco.Checked = bHabilita;
			chkGrupoArt.Checked = bHabilita;
			chkGrupoCliente.Checked = bHabilita;
			chkPersonal.Checked = bHabilita;
			chkPropiedades.Checked = bHabilita;
			chkProyecto.Checked = bHabilita;
			chkReceta.Checked = bHabilita;
			chkSeguridad.Checked = bHabilita;
			chkSubProyecto.Checked = bHabilita;
			chkSustento.Checked = bHabilita;
			chkRolRubro.Checked = bHabilita;
			chkRolPago.Checked = bHabilita;
		}

		private void SaldosIniciales_Load(object sender, System.EventArgs e)
		{
			cmbEmpresa.Text = "";
			string stAudita = "Exec AuditaCrear 53, 6, 'Copiar Archivos'";
			Funcion.EjecutaSQL(cdsContabilidad, stAudita, true);
			DateTime dtFecha = new DateTime(DateTime.Today.Year, 1, 1);
			cmbFecha.Value = dtFecha;

			if (System.IO.File.Exists(@"C:\Latinium\LibAnexos1.dll"))
			{
				llenaEmpresaDB();
			}
		}

		private void llenaEmpresaDB()
		{
			
			try
			{
				string Servidor = MenuLatinium.stDirServidor;
				string stConeccion = string.Format(
					"Data Source={0};Initial Catalog=EmpresasBase;Persist Security Info=True;User ID=infoelect;Password=Bl45o6$9;",Servidor);

			
				using (SqlConnection conn = new SqlConnection(stConeccion))
				{
					conn.Open();
					string stSelect = "select *  from empresa order by empresa";
										

					DataSet dsArticulo = new DataSet();
					SqlDataAdapter myData = new SqlDataAdapter(stSelect,conn.ConnectionString);
					myData.SelectCommand.CommandTimeout = 0;
					myData.TableMappings.Add("Tabla", "ExcelSube");
					dsArticulo.Reset();
					dsArticulo.Clear();
					try
					{
						myData.Fill(dsArticulo);
					}
					catch (Exception ex)
					{
						MessageBox.Show(ex.Message, "Error al ejecutar procedimiento: ");
					}					
					cmbEmpresa.DataSource =  (DataView)dsArticulo.Tables[0].DefaultView;
					
				}
			}
			catch (SqlException ex)
			{
				MessageBox.Show("Error:"+ ex.Message);
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}
		private void GeneraTabla(C1.Data.C1DataSet ds, string stTabla)
		{
			GeneraTabla(ds, stTabla, 0);
		}

		private void GeneraTabla(C1.Data.C1DataSet ds, string stTabla, int iMinimo)
		{
			Cursor = Cursors.WaitCursor;
			int iCuenta = Funcion.iEscalarSQL(ds, "Select Count(*) From " + stTabla);
			//MessageBox.Show(stTabla + " " + iCuenta.ToString());
			if (iCuenta > iMinimo && !chkSobrescribir.Checked)
			{
				if (DialogResult.Yes != MessageBox.Show("Existen " + iCuenta.ToString() + 
					" Registros en " + stTabla + ". ¿Desea Sobreescribirlos?", "Confirmación",
					MessageBoxButtons.YesNo, 
					MessageBoxIcon.Question,
					MessageBoxDefaultButton.Button2) )
					return;
			}
			string txtEmpresa = cmbEmpresa.ActiveRow.Cells["Facturacion"].Value.ToString();
			// Borrar y Copia datos de Tabla Anterior
		//	string stSelect = "SET IDENTITY_INSERT dbo." + stTabla + " ON;";
			string stSelect = "Declare @i int; Select @i = IsNull(IDENT_incr('" + stTabla + "'), -1); Select IsNull(@i, -1);";

			int iExistePrimaria = Funcion.iEscalarSQL(ds, stSelect, true);
			string stAgrega = ", 1";
			if (chkAgregar.Checked) stAgrega = ", 0";
			stAgrega = "";
			if (iExistePrimaria>=0)
			{
				stSelect = "SET IDENTITY_INSERT dbo." + stTabla + " ON;";
				stSelect += "EXEC CopiaTabla " + stTabla + ", " 
					+ txtEmpresa + stAgrega;
				stSelect += "; SET IDENTITY_INSERT dbo." + stTabla + " OFF;";
			}
			else
				stSelect = "EXEC CopiaTabla " + stTabla + ", " 
					+ txtEmpresa + stAgrega;

			try
			{
				Funcion.EjecutaSQL(ds, stSelect, true);
			}
			catch(System.Exception ex1)
			{
				if (iExistePrimaria>=0)
				{
					stSelect = "SET IDENTITY_INSERT dbo." + stTabla + " OFF;";
					Funcion.EjecutaSQL(cdsContabilidad, stSelect, true);
				}
				MessageBox.Show(ex1.Message + " " + stSelect, "Copiar Archivos");
			}
			Cursor = Cursors.Default;
		}

		private void BorrarErrores()
		{
			errorProvider.SetError(cmbEmpresa, "");
			errorProvider.SetError(chkTodo, "");
		}

		private int Verifica()
		{
			int iNumTablas = 0;
			BorrarErrores();
			if (chkImportacion.Checked) iNumTablas ++;
			if (chkDistribucion.Checked) iNumTablas ++;
			if (chkFormaPago.Checked) iNumTablas ++;
			if (chkTablaRet.Checked) iNumTablas ++;
			if (chkAsiento.Checked) iNumTablas ++;
			if (chkPagos.Checked) iNumTablas ++;
			if (chkRetencion.Checked) iNumTablas ++;
			if (chkFactura.Checked) iNumTablas ++;
			if (chkArticulos.Checked) iNumTablas ++;
			if (chkAsignaCuenta.Checked) iNumTablas ++;
			if (chkBancos.Checked) iNumTablas ++;
			if (chkClientes.Checked) iNumTablas ++;
			if (chkCuentaBanco.Checked) iNumTablas ++;
			if (chkGrupoArt.Checked) iNumTablas ++;
			if (chkGrupoCliente.Checked) iNumTablas ++;
			if (chkPersonal.Checked) iNumTablas ++;
			if (chkRolRubro.Checked) iNumTablas ++;
			if (chkPlanCuentas.Checked) iNumTablas ++;
			if (chkPropiedades.Checked) iNumTablas ++;
			if (chkProyecto.Checked) iNumTablas ++;
			if (chkReceta.Checked) iNumTablas ++;
			if (chkSeguridad.Checked) iNumTablas ++;
			if (chkSubProyecto.Checked) iNumTablas ++;
			if (chkSustento.Checked) iNumTablas ++;
			if (chkRolPago.Checked) iNumTablas ++;

			if (iNumTablas == 0)
			{
				iNumTablas = 0;
				errorProvider.SetError(chkTodo, "Ingrese Tablas a Copiar");
			}
			if (cmbEmpresa.Text.Trim().Length == 0)
			{
				iNumTablas = 0;
				errorProvider.SetError(cmbEmpresa, "Ingrese Empresa");
			}
			if (cmbEmpresa.Text.Trim().ToUpper() == MenuLatinium.stDirConta.Trim().ToUpper())
			{
				iNumTablas = 0;
				errorProvider.SetError(cmbEmpresa, "Empresa de Origen y Destino no pueden ser las mismas");
			}
			return iNumTablas;
		}


		private void CambiaProcedimiento ()
		{
			string stEmpresa = cmbEmpresa.Text.Trim().ToUpper();
			string stEjemplo = "ECLARE CASIENTO CURSOR LOCAL READ_ONLY FOR SELECT IDASIENTO FROM ";
			stEjemplo = stEjemplo.ToUpper();

			string stExec = "EXEC sp_helptext 'AsientoSincroniza'";
			SqlDataReader dr = Funcion.rEscalarSQL(cdsContabilidad, stExec, true);

			string stArchivo = "";
			while (dr.Read())
			{
				string stLinea = dr.GetString(0);
				int iPos = stLinea.ToUpper().IndexOf(stEjemplo);
				if (iPos > 0)
					stLinea = "DECLARE CASIENTO CURSOR LOCAL READ_ONLY FOR SELECT IDASIENTO FROM "
						+ stEmpresa + ".DBO.ASIENTO\n";
				stLinea = stLinea.ToUpper().Replace("CREATE PROCEDURE", "ALTER PROCEDURE");
				stArchivo += stLinea;
			}
			dr.Close();
			Funcion.EjecutaSQL(cdsContabilidad, stArchivo);
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			int iNumTablas = Verifica();
			if (iNumTablas == 0) return;
			if (DialogResult.No == MessageBox.Show("Desea copiar " 
				+ iNumTablas + " tablas de " 
				+ cmbEmpresa.Text + " a "+ MenuLatinium.stDirConta,
				"Confirmación", MessageBoxButtons.YesNo,
				MessageBoxIcon.Question)) return;

			if (chkArticulos.Checked)
			{
				GeneraTabla(cdsContabilidad, "Articulo");
				GeneraTabla(cdsContabilidad, "ArticuloArancel");
				GeneraTabla(cdsContabilidad, "ArticuloArancelDetalle");
				GeneraTabla(cdsContabilidad, "ArticuloClase");
				GeneraTabla(cdsContabilidad, "ArticuloColor");
				GeneraTabla(cdsContabilidad, "ArticuloGrupo");
				GeneraTabla(cdsContabilidad, "ArticuloMarca");
				GeneraTabla(cdsContabilidad, "ArticuloModelo");
				GeneraTabla(cdsContabilidad, "ArticuloPrecios");
				GeneraTabla(cdsContabilidad, "ArticuloSeccion");
				GeneraTabla(cdsContabilidad, "ArticuloSRI");
				GeneraTabla(cdsContabilidad, "ArticuloSubGrupo");
//				GeneraTabla(cdsContabilidad, "ArticuloTipo");
				GeneraTabla(cdsContabilidad, "ArticuloUnidad");
			}

			if (chkReceta.Checked)
			{
				GeneraTabla(cdsFacturacion, "Producto");
				string stBase = cmbEmpresa.ActiveRow.Cells["Facturacion"].Value.ToString().Trim() + ".dbo.";
				string stBorrar = "Delete " + stBase + "ProductoDetalle From " + stBase + "ProductoDetalle "
					+ "p Left Join " + stBase + "Articulo a On a.idArticulo = p.idArticulo "
					+ "Where IsNull(a.idArticulo, -1) = -1";
				Funcion.EjecutaSQL(cdsFacturacion, stBorrar, true);
				GeneraTabla(cdsFacturacion, "ProductoDetalle");
			}
			if (chkSustento.Checked) 
			{
				GeneraTabla(cdsContabilidad, "CompraCredTribut");
				GeneraTabla(cdsContabilidad, "CompraComprobante");
				GeneraTabla(cdsContabilidad, "PorcIce");
			}

			if (chkFormaPago.Checked) 
			{
				GeneraTabla(cdsContabilidad, "CompraFormaPago");
				GeneraTabla(cdsContabilidad, "ReservaC");
			}
			if (chkImportacion.Checked) 
			{
				GeneraTabla(cdsContabilidad, "CompraImportacion");
				GeneraTabla(cdsContabilidad, "CompraDetImp");
			}
			if (chkAsignaCuenta.Checked)
			{
				GeneraTabla(cdsContabilidad, "AsignaCuenta", 50);
				GeneraTabla(cdsContabilidad, "AsientoTipo");
				GeneraTabla(cdsContabilidad, "IndiceFinanciero");
			}
			if (chkDistribucion.Checked)
			{
				GeneraTabla(cdsContabilidad, "AsientoDistGrupo");
				GeneraTabla(cdsContabilidad, "AsientoDist");
				GeneraTabla(cdsContabilidad, "FacturaNumero");
			}
			if (chkTablaRet.Checked)
			{
				GeneraTabla(cdsContabilidad, "RetencionIva");
				GeneraTabla(cdsContabilidad, "RetencionRenta");
			}
			if (chkBancos.Checked) GeneraTabla(cdsContabilidad, "Banco", 30);
			if (chkClientes.Checked)
			{
				GeneraTabla(cdsFacturacion, "Cliente");
				string stSelect = "ALTER TABLE ClienteCiudad DISABLE TRIGGER BorraCiudad";
				Funcion.EjecutaSQL(cdsContabilidad, stSelect, true);
				GeneraTabla(cdsFacturacion, "ClienteCiudad");
				stSelect = "ALTER TABLE ClienteCiudad ENABLE TRIGGER BorraCiudad";
				Funcion.EjecutaSQL(cdsContabilidad, stSelect, true);
				GeneraTabla(cdsFacturacion, "ClienteListaPrecio");
				GeneraTabla(cdsFacturacion, "ClientePais");
				GeneraTabla(cdsFacturacion, "ClienteSector");
				GeneraTabla(cdsFacturacion, "ClienteTipoRuc");
				GeneraTabla(cdsFacturacion, "ClientePublicidad");
				GeneraTabla(cdsFacturacion, "ClientePrecio");
			}
			if (chkCuentaBanco.Checked) GeneraTabla(cdsContabilidad, "CuentaCorriente");
			if (chkGrupoArt.Checked) GeneraTabla(cdsFacturacion, "ArticuloGrupo");
			if (chkGrupoCliente.Checked) GeneraTabla(cdsFacturacion, "ClienteGrupo");
			if (chkPersonal.Checked) 
			{
				// Borra Personal
				string stBorraPers = "Delete Personal;Delete RolConcepto;Delete PersEstudios;"
					+ "Delete PersFamiliar;Delete PersNivel;Delete PersTitulo";

				Funcion.EjecutaSQL(cdsContabilidad, stBorraPers, true);
				GeneraTabla(cdsRoles, "RolGrupo");
				GeneraTabla(cdsRoles, "PersParentesco");
				GeneraTabla(cdsRoles, "Personal");
				GeneraTabla(cdsRoles, "PersEstudios");
				GeneraTabla(cdsRoles, "PersNivel");
				GeneraTabla(cdsRoles, "PersTitulo");
				GeneraTabla(cdsRoles, "PersFamiliar");

				GeneraTabla(cdsRoles, "RolSucursal");
				GeneraTabla(cdsRoles, "RolCargo");
				GeneraTabla(cdsRoles, "RolCuentas");
				GeneraTabla(cdsRoles, "RolConcepto");
//				GeneraTabla(cdsRoles, "RolGrupo");
				GeneraTabla(cdsRoles, "RolRelacionTrabajo");
				GeneraTabla(cdsRoles, "RolRenta");
				GeneraTabla(cdsRoles, "RolDepartamento");
			}
			if (chkRolRubro.Checked) 
			{
//				GeneraTabla(cdsRoles, "RolGrupo");
				GeneraTabla(cdsRoles, "RolConcepto");
			}

			if (chkPlanCuentas.Checked) GeneraTabla(cdsContabilidad, "Cuenta", 6);
			if (chkAsiento.Checked) 
			{
				string stEmpresa = cmbEmpresa.ActiveRow.Cells["Contabilidad"].Value.ToString().Trim() + ".dbo.";
				string stDelete = "delete " + stEmpresa + "asientodetalle "
					+ "from " + stEmpresa + "asientodetalle ad left join " + stEmpresa + "asiento a "
					+ "on a.idasiento = ad.idasiento where isnull(a.idasiento, 0) = 0";
				Funcion.EjecutaSQL(cdsContabilidad, stDelete, true);

				if (chkAgregar.Checked)
				{
					CambiaProcedimiento();

					string stExecAsiento = "Exec AsientoSincroniza " + cmbEmpresa.Text.Trim();
					Funcion.EjecutaSQL(cdsContabilidad, stExecAsiento, true);
					MessageBox.Show("Asientos Agregados", "Información");
					return;
				}
				else
				{
					GeneraTabla(cdsContabilidad, "Asiento");
					GeneraTabla(cdsContabilidad, "AsientoDetalle");
				}
			}

			if (chkFactura.Checked) 
			{
				string stEmpresa = cmbEmpresa.ActiveRow.Cells["Contabilidad"].Value.ToString().Trim() + ".dbo.";
				string stDelete = "delete " + stEmpresa + "detCompra "
					+ "from " + stEmpresa + "DetCompra ad left join " + stEmpresa + "Compra a "
					+ "on a.idCompra = ad.idCompra where isnull(a.idCompra, 0) = 0";
				Funcion.EjecutaSQL(cdsContabilidad, stDelete, true);

				if (chkAgregar.Checked)
				{
					CambiaProcedimiento();

					string stExecAsiento = "Exec FacturaSincroniza " + cmbEmpresa.Text.Trim();
					Funcion.EjecutaSQL(cdsContabilidad, stExecAsiento, true);
					MessageBox.Show("Faturas Agregadas", "Información");
					return;
				}
				else
				{
					GeneraTabla(cdsContabilidad, "Compra");
					GeneraTabla(cdsContabilidad, "DetCompra");
				}
			}

			if (chkPagos.Checked) 
			{
				if (chkAgregar.Checked)
				{
					CambiaProcedimiento();

					string stExecAsiento = "Exec PagoSincroniza " + cmbEmpresa.Text.Trim();
					Funcion.EjecutaSQL(cdsContabilidad, stExecAsiento, true);
					MessageBox.Show("Pagos Agregados", "Información");
					return;
				}
				else
				{
					GeneraTabla(cdsContabilidad, "PagoLote");
					GeneraTabla(cdsContabilidad, "Pago");
				}
			}

			if (chkRetencion.Checked) 
			{
				if (chkAgregar.Checked)
				{
					CambiaProcedimiento();

					string stExecAsiento = "Exec RetencionSincroniza " + cmbEmpresa.Text.Trim();
					Funcion.EjecutaSQL(cdsContabilidad, stExecAsiento, true);
					MessageBox.Show("Retenciones Agregadas", "Información");
					return;
				}
				else
				{
					GeneraTabla(cdsContabilidad, "Retencion");
				}
			}

			if (chkPropiedades.Checked) 
			{
				GeneraTabla(cdsContabilidad, "Seteo", 1);
				GeneraTabla(cdsContabilidad, "SeteoC", 1);
				GeneraTabla(cdsContabilidad, "SeteoF", 1);
				GeneraTabla(cdsContabilidad, "SeteoG");
			}
			if (chkProyecto.Checked) GeneraTabla(cdsContabilidad, "Proyecto");
			if (chkSeguridad.Checked) GeneraTabla(cdsContabilidad, "Seguridad", 600);
			if (chkSubProyecto.Checked) GeneraTabla(cdsFacturacion, "SubProyecto");
			if (chkRolPago.Checked) 
				GeneraTabla(cdsRoles, "RolPago");

			MessageBox.Show("Copia de Archivos Terminada", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void cdsContabilidad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsContabilidad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsFacturacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsFacturacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsRoles_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRoles.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void chkAgregar_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void CopiarArchivos_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Copiar Archivos'";
			Funcion.EjecutaSQL(cdsContabilidad, stAudita, true);
		}

		private void chkSeguridad_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			if (txtClave.Text.Trim() != "2243") 
			{
				MessageBox.Show("No tiene acceso a esta opción", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			if (DialogResult.No == MessageBox.Show("¿Desea elimininar registros anteriores al "
				+ dtFecha.ToString("dd/MMM/yyyy") + "?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stAudita = "Exec AuditaCrear 53, 7, 'Borrar Archivos'";
			Funcion.EjecutaSQL(cdsContabilidad, stAudita, true);

			string stExec = "Exec BaseDatosBorraAntes '" + dtFecha.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsContabilidad, stExec, true);
			MessageBox.Show("Registros Borrados, necesita inventario y Balance", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btBorraBase_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			txtClave.Visible = true;
		}

		private void cmbEmpresa_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

//		private void cdsEmpresaDB_BeforeFill(object sender, C1.Data.FillEventArgs e)
//		{
//			Funcion miFuncion = new Funcion();
//			cdsEmpresa.Schema.Connections[0].ConnectionString = miFuncion.AccesoEmpresaBase();
//		}
	}
}
