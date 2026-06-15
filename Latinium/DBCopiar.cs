using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for DBCopiar.
	/// </summary>
	public class DBCopiar : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsEmpresa;
		private DevExpress.XtraGrid.GridControl gridControl1;
		private C1.Data.C1DataSet cdsEmpresaAccess;
		private DevExpress.XtraGrid.Columns.GridColumn colEmpresa;
		private DevExpress.XtraGrid.Columns.GridColumn colidEmpresa;
		private DevExpress.XtraGrid.Columns.GridColumn colPrincipal;
		private DevExpress.XtraGrid.Columns.GridColumn colReportes;
		private DevExpress.XtraGrid.Columns.GridColumn colServidor;
		private DevExpress.XtraGrid.Columns.GridColumn colFacturacion;
		private DevExpress.XtraTab.XtraTabControl xtraTabControl1;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage1;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage2;
		private DevExpress.XtraEditors.SimpleButton btQuitarTodo;
		private DevExpress.XtraEditors.SimpleButton btQuitarUno;
		private DevExpress.XtraEditors.SimpleButton btAgregaUno;
		private DevExpress.XtraEditors.SimpleButton btAgregaTodo;
		private DevExpress.XtraGrid.GridControl gridControl2;
		private DevExpress.XtraGrid.Columns.GridColumn colidEmpresa1;
		private DevExpress.XtraGrid.Columns.GridColumn colEmpresa1;
		private DevExpress.XtraGrid.Columns.GridColumn colServidor1;
		private DevExpress.XtraGrid.Columns.GridColumn colFacturacion1;
		private DevExpress.XtraGrid.Columns.GridColumn colReportes1;
		private DevExpress.XtraGrid.Views.Grid.GridView grvEmpresaAccess;
		private DevExpress.XtraGrid.Views.Grid.GridView grvEmpresa;
		private DevExpress.XtraGrid.GridControl gridControl3;
		private DevExpress.XtraGrid.Views.Grid.GridView grvUsuario;
		private DevExpress.XtraEditors.SimpleButton btQuitarUsuarioT;
		private DevExpress.XtraEditors.SimpleButton btQuitarUsuarioUno;
		private DevExpress.XtraEditors.SimpleButton btAgregarUsuarioUno;
		private DevExpress.XtraEditors.SimpleButton btAgregarUsuarioT;
		private DevExpress.XtraGrid.GridControl gridControl4;
		private DevExpress.XtraGrid.Views.Grid.GridView grvUsuarioAccess;
		private C1.Data.C1DataSet cdsUsuarioAccess;
		private DevExpress.XtraGrid.Columns.GridColumn colClave;
		private DevExpress.XtraGrid.Columns.GridColumn colidGrupo;
		private DevExpress.XtraGrid.Columns.GridColumn colidUsuario;
		private DevExpress.XtraGrid.Columns.GridColumn colNombre;
		private DevExpress.XtraGrid.Columns.GridColumn colidUsuario1;
		private DevExpress.XtraGrid.Columns.GridColumn colNombre1;
		private DevExpress.XtraGrid.Columns.GridColumn colClave1;
		private DevExpress.XtraGrid.Columns.GridColumn colidGrupo1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DBCopiar()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.cdsEmpresa = new C1.Data.C1DataSet();
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.cdsEmpresaAccess = new C1.Data.C1DataSet();
			this.grvEmpresaAccess = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colidEmpresa = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colEmpresa = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colPrincipal = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colReportes = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colServidor = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colFacturacion = new DevExpress.XtraGrid.Columns.GridColumn();
			this.xtraTabControl1 = new DevExpress.XtraTab.XtraTabControl();
			this.xtraTabPage1 = new DevExpress.XtraTab.XtraTabPage();
			this.gridControl2 = new DevExpress.XtraGrid.GridControl();
			this.grvEmpresa = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colidEmpresa1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colEmpresa1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colServidor1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colFacturacion1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colReportes1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.btQuitarTodo = new DevExpress.XtraEditors.SimpleButton();
			this.btQuitarUno = new DevExpress.XtraEditors.SimpleButton();
			this.btAgregaUno = new DevExpress.XtraEditors.SimpleButton();
			this.btAgregaTodo = new DevExpress.XtraEditors.SimpleButton();
			this.xtraTabPage2 = new DevExpress.XtraTab.XtraTabPage();
			this.gridControl3 = new DevExpress.XtraGrid.GridControl();
			this.grvUsuario = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colidUsuario1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colNombre1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colClave1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colidGrupo1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.btQuitarUsuarioT = new DevExpress.XtraEditors.SimpleButton();
			this.btQuitarUsuarioUno = new DevExpress.XtraEditors.SimpleButton();
			this.btAgregarUsuarioUno = new DevExpress.XtraEditors.SimpleButton();
			this.btAgregarUsuarioT = new DevExpress.XtraEditors.SimpleButton();
			this.gridControl4 = new DevExpress.XtraGrid.GridControl();
			this.cdsUsuarioAccess = new C1.Data.C1DataSet();
			this.grvUsuarioAccess = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colClave = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colidGrupo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colidUsuario = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colNombre = new DevExpress.XtraGrid.Columns.GridColumn();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresaAccess)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvEmpresaAccess)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).BeginInit();
			this.xtraTabControl1.SuspendLayout();
			this.xtraTabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridControl2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvEmpresa)).BeginInit();
			this.xtraTabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridControl3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUsuarioAccess)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvUsuarioAccess)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsEmpresa
			// 
			this.cdsEmpresa.BindingContextCtrl = this;
			this.cdsEmpresa.DataLibrary = "LibDBEmpresa";
			this.cdsEmpresa.DataLibraryUrl = "";
			this.cdsEmpresa.DataSetDef = "dsEmpresa";
			this.cdsEmpresa.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEmpresa.Name = "cdsEmpresa";
			this.cdsEmpresa.SchemaClassName = "LibDBEmpresa.DataClass";
			this.cdsEmpresa.SchemaDef = null;
			this.cdsEmpresa.BeforeFill += new C1.Data.FillEventHandler(this.cdsEmpresa_BeforeFill);
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl1.DataMember = "Empresa";
			this.gridControl1.DataSource = this.cdsEmpresaAccess;
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(8, 8);
			this.gridControl1.MainView = this.grvEmpresaAccess;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.Size = new System.Drawing.Size(488, 136);
			this.gridControl1.TabIndex = 0;
			this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.grvEmpresaAccess});
			// 
			// cdsEmpresaAccess
			// 
			this.cdsEmpresaAccess.BindingContextCtrl = this;
			this.cdsEmpresaAccess.DataLibrary = "LibEmpresa";
			this.cdsEmpresaAccess.DataLibraryUrl = "";
			this.cdsEmpresaAccess.DataSetDef = "dsEmpresa";
			this.cdsEmpresaAccess.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEmpresaAccess.Name = "cdsEmpresaAccess";
			this.cdsEmpresaAccess.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsEmpresaAccess.SchemaDef = null;
			// 
			// grvEmpresaAccess
			// 
			this.grvEmpresaAccess.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																														this.colidEmpresa,
																																														this.colEmpresa,
																																														this.colPrincipal,
																																														this.colReportes,
																																														this.colServidor,
																																														this.colFacturacion});
			this.grvEmpresaAccess.GridControl = this.gridControl1;
			this.grvEmpresaAccess.Name = "grvEmpresaAccess";
			this.grvEmpresaAccess.OptionsView.ShowGroupPanel = false;
			// 
			// colidEmpresa
			// 
			this.colidEmpresa.Caption = "idEmpresa";
			this.colidEmpresa.FieldName = "idEmpresa";
			this.colidEmpresa.Name = "colidEmpresa";
			this.colidEmpresa.OptionsColumn.AllowEdit = false;
			// 
			// colEmpresa
			// 
			this.colEmpresa.Caption = "Empresa Access";
			this.colEmpresa.FieldName = "Empresa";
			this.colEmpresa.Name = "colEmpresa";
			this.colEmpresa.OptionsColumn.AllowEdit = false;
			this.colEmpresa.Visible = true;
			this.colEmpresa.VisibleIndex = 0;
			// 
			// colPrincipal
			// 
			this.colPrincipal.Caption = "Principal";
			this.colPrincipal.FieldName = "Principal";
			this.colPrincipal.Name = "colPrincipal";
			// 
			// colReportes
			// 
			this.colReportes.Caption = "Reportes";
			this.colReportes.FieldName = "Reportes";
			this.colReportes.Name = "colReportes";
			this.colReportes.OptionsColumn.AllowEdit = false;
			this.colReportes.Visible = true;
			this.colReportes.VisibleIndex = 3;
			// 
			// colServidor
			// 
			this.colServidor.Caption = "Servidor";
			this.colServidor.FieldName = "Servidor";
			this.colServidor.Name = "colServidor";
			this.colServidor.OptionsColumn.AllowEdit = false;
			this.colServidor.Visible = true;
			this.colServidor.VisibleIndex = 1;
			// 
			// colFacturacion
			// 
			this.colFacturacion.Caption = "Facturacion";
			this.colFacturacion.FieldName = "Facturacion";
			this.colFacturacion.Name = "colFacturacion";
			this.colFacturacion.OptionsColumn.AllowEdit = false;
			this.colFacturacion.Visible = true;
			this.colFacturacion.VisibleIndex = 2;
			// 
			// xtraTabControl1
			// 
			this.xtraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.xtraTabControl1.Location = new System.Drawing.Point(16, 16);
			this.xtraTabControl1.Name = "xtraTabControl1";
			this.xtraTabControl1.SelectedTabPage = this.xtraTabPage1;
			this.xtraTabControl1.Size = new System.Drawing.Size(520, 360);
			this.xtraTabControl1.TabIndex = 1;
			this.xtraTabControl1.TabPages.AddRange(new DevExpress.XtraTab.XtraTabPage[] {
																																										this.xtraTabPage1,
																																										this.xtraTabPage2});
			this.xtraTabControl1.Text = "xtraTabControl1";
			// 
			// xtraTabPage1
			// 
			this.xtraTabPage1.Controls.Add(this.gridControl2);
			this.xtraTabPage1.Controls.Add(this.btQuitarTodo);
			this.xtraTabPage1.Controls.Add(this.btQuitarUno);
			this.xtraTabPage1.Controls.Add(this.btAgregaUno);
			this.xtraTabPage1.Controls.Add(this.btAgregaTodo);
			this.xtraTabPage1.Controls.Add(this.gridControl1);
			this.xtraTabPage1.Name = "xtraTabPage1";
			this.xtraTabPage1.Size = new System.Drawing.Size(511, 329);
			this.xtraTabPage1.Text = "Empresa";
			// 
			// gridControl2
			// 
			this.gridControl2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl2.DataMember = "Empresa";
			this.gridControl2.DataSource = this.cdsEmpresa;
			// 
			// gridControl2.EmbeddedNavigator
			// 
			this.gridControl2.EmbeddedNavigator.Name = "";
			this.gridControl2.Location = new System.Drawing.Point(8, 184);
			this.gridControl2.MainView = this.grvEmpresa;
			this.gridControl2.Name = "gridControl2";
			this.gridControl2.Size = new System.Drawing.Size(488, 136);
			this.gridControl2.TabIndex = 13;
			this.gridControl2.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.grvEmpresa});
			// 
			// grvEmpresa
			// 
			this.grvEmpresa.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																											this.colidEmpresa1,
																																											this.colEmpresa1,
																																											this.colServidor1,
																																											this.colFacturacion1,
																																											this.colReportes1});
			this.grvEmpresa.GridControl = this.gridControl2;
			this.grvEmpresa.Name = "grvEmpresa";
			this.grvEmpresa.OptionsView.ShowGroupPanel = false;
			// 
			// colidEmpresa1
			// 
			this.colidEmpresa1.Caption = "idEmpresa";
			this.colidEmpresa1.FieldName = "idEmpresa";
			this.colidEmpresa1.Name = "colidEmpresa1";
			// 
			// colEmpresa1
			// 
			this.colEmpresa1.Caption = "Empresa SQL";
			this.colEmpresa1.FieldName = "Empresa";
			this.colEmpresa1.Name = "colEmpresa1";
			this.colEmpresa1.OptionsColumn.AllowEdit = false;
			this.colEmpresa1.Visible = true;
			this.colEmpresa1.VisibleIndex = 0;
			// 
			// colServidor1
			// 
			this.colServidor1.Caption = "Servidor";
			this.colServidor1.FieldName = "Servidor";
			this.colServidor1.Name = "colServidor1";
			this.colServidor1.OptionsColumn.AllowEdit = false;
			this.colServidor1.Visible = true;
			this.colServidor1.VisibleIndex = 1;
			// 
			// colFacturacion1
			// 
			this.colFacturacion1.Caption = "Facturacion";
			this.colFacturacion1.FieldName = "Facturacion";
			this.colFacturacion1.Name = "colFacturacion1";
			this.colFacturacion1.OptionsColumn.AllowEdit = false;
			this.colFacturacion1.Visible = true;
			this.colFacturacion1.VisibleIndex = 2;
			// 
			// colReportes1
			// 
			this.colReportes1.Caption = "Reportes";
			this.colReportes1.FieldName = "Reportes";
			this.colReportes1.Name = "colReportes1";
			this.colReportes1.OptionsColumn.AllowEdit = false;
			this.colReportes1.Visible = true;
			this.colReportes1.VisibleIndex = 3;
			// 
			// btQuitarTodo
			// 
			this.btQuitarTodo.Location = new System.Drawing.Point(8, 152);
			this.btQuitarTodo.Name = "btQuitarTodo";
			this.btQuitarTodo.Size = new System.Drawing.Size(32, 23);
			this.btQuitarTodo.TabIndex = 12;
			this.btQuitarTodo.Text = "<<";
			this.btQuitarTodo.ToolTip = "Quitar todos las empresas de SQL de Abajo";
			this.btQuitarTodo.Click += new System.EventHandler(this.btQuitarTodo_Click);
			// 
			// btQuitarUno
			// 
			this.btQuitarUno.Location = new System.Drawing.Point(56, 152);
			this.btQuitarUno.Name = "btQuitarUno";
			this.btQuitarUno.Size = new System.Drawing.Size(32, 23);
			this.btQuitarUno.TabIndex = 11;
			this.btQuitarUno.Text = "<=";
			this.btQuitarUno.ToolTip = "Quitar una empresa de SQL de Abajo";
			this.btQuitarUno.Click += new System.EventHandler(this.btQuitarUno_Click);
			// 
			// btAgregaUno
			// 
			this.btAgregaUno.Location = new System.Drawing.Point(112, 152);
			this.btAgregaUno.Name = "btAgregaUno";
			this.btAgregaUno.Size = new System.Drawing.Size(32, 23);
			this.btAgregaUno.TabIndex = 10;
			this.btAgregaUno.Text = "=>";
			this.btAgregaUno.ToolTip = "Agrega una empresa de SQL de Abajo";
			this.btAgregaUno.Click += new System.EventHandler(this.btAgregaUno_Click);
			// 
			// btAgregaTodo
			// 
			this.btAgregaTodo.Location = new System.Drawing.Point(160, 152);
			this.btAgregaTodo.Name = "btAgregaTodo";
			this.btAgregaTodo.Size = new System.Drawing.Size(32, 23);
			this.btAgregaTodo.TabIndex = 9;
			this.btAgregaTodo.Text = ">>";
			this.btAgregaTodo.ToolTip = "Agrega todas las empresas de SQL de Abajo";
			this.btAgregaTodo.Click += new System.EventHandler(this.btAgregaTodo_Click);
			// 
			// xtraTabPage2
			// 
			this.xtraTabPage2.Controls.Add(this.gridControl3);
			this.xtraTabPage2.Controls.Add(this.btQuitarUsuarioT);
			this.xtraTabPage2.Controls.Add(this.btQuitarUsuarioUno);
			this.xtraTabPage2.Controls.Add(this.btAgregarUsuarioUno);
			this.xtraTabPage2.Controls.Add(this.btAgregarUsuarioT);
			this.xtraTabPage2.Controls.Add(this.gridControl4);
			this.xtraTabPage2.Name = "xtraTabPage2";
			this.xtraTabPage2.Size = new System.Drawing.Size(511, 329);
			this.xtraTabPage2.Text = "Usuarios";
			// 
			// gridControl3
			// 
			this.gridControl3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl3.DataMember = "Usuario";
			this.gridControl3.DataSource = this.cdsEmpresa;
			// 
			// gridControl3.EmbeddedNavigator
			// 
			this.gridControl3.EmbeddedNavigator.Name = "";
			this.gridControl3.Location = new System.Drawing.Point(11, 184);
			this.gridControl3.MainView = this.grvUsuario;
			this.gridControl3.Name = "gridControl3";
			this.gridControl3.Size = new System.Drawing.Size(488, 136);
			this.gridControl3.TabIndex = 19;
			this.gridControl3.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.grvUsuario});
			// 
			// grvUsuario
			// 
			this.grvUsuario.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																											this.colidUsuario1,
																																											this.colNombre1,
																																											this.colClave1,
																																											this.colidGrupo1});
			this.grvUsuario.GridControl = this.gridControl3;
			this.grvUsuario.Name = "grvUsuario";
			this.grvUsuario.OptionsView.ShowGroupPanel = false;
			// 
			// colidUsuario1
			// 
			this.colidUsuario1.Caption = "idUsuario";
			this.colidUsuario1.FieldName = "idUsuario";
			this.colidUsuario1.Name = "colidUsuario1";
			// 
			// colNombre1
			// 
			this.colNombre1.Caption = "Nombre";
			this.colNombre1.FieldName = "Nombre";
			this.colNombre1.Name = "colNombre1";
			this.colNombre1.Visible = true;
			this.colNombre1.VisibleIndex = 0;
			// 
			// colClave1
			// 
			this.colClave1.Caption = "Clave";
			this.colClave1.FieldName = "Clave";
			this.colClave1.Name = "colClave1";
			// 
			// colidGrupo1
			// 
			this.colidGrupo1.Caption = "idGrupo";
			this.colidGrupo1.FieldName = "idGrupo";
			this.colidGrupo1.Name = "colidGrupo1";
			this.colidGrupo1.Visible = true;
			this.colidGrupo1.VisibleIndex = 1;
			// 
			// btQuitarUsuarioT
			// 
			this.btQuitarUsuarioT.Location = new System.Drawing.Point(11, 152);
			this.btQuitarUsuarioT.Name = "btQuitarUsuarioT";
			this.btQuitarUsuarioT.Size = new System.Drawing.Size(32, 23);
			this.btQuitarUsuarioT.TabIndex = 18;
			this.btQuitarUsuarioT.Text = "<<";
			this.btQuitarUsuarioT.ToolTip = "Quitar todos los usuario de SQL de Abajo";
			this.btQuitarUsuarioT.Click += new System.EventHandler(this.btQuitarUsuarioT_Click);
			// 
			// btQuitarUsuarioUno
			// 
			this.btQuitarUsuarioUno.Location = new System.Drawing.Point(56, 152);
			this.btQuitarUsuarioUno.Name = "btQuitarUsuarioUno";
			this.btQuitarUsuarioUno.Size = new System.Drawing.Size(32, 23);
			this.btQuitarUsuarioUno.TabIndex = 17;
			this.btQuitarUsuarioUno.Text = "<=";
			this.btQuitarUsuarioUno.ToolTip = "Quitar un usuario de SQL de Abajo";
			this.btQuitarUsuarioUno.Click += new System.EventHandler(this.btQuitarUsuarioUno_Click);
			// 
			// btAgregarUsuarioUno
			// 
			this.btAgregarUsuarioUno.Location = new System.Drawing.Point(112, 152);
			this.btAgregarUsuarioUno.Name = "btAgregarUsuarioUno";
			this.btAgregarUsuarioUno.Size = new System.Drawing.Size(32, 23);
			this.btAgregarUsuarioUno.TabIndex = 16;
			this.btAgregarUsuarioUno.Text = "=>";
			this.btAgregarUsuarioUno.ToolTip = "Agrega un usuario de SQL de Abajo";
			this.btAgregarUsuarioUno.Click += new System.EventHandler(this.btAgregarUsuarioUno_Click);
			// 
			// btAgregarUsuarioT
			// 
			this.btAgregarUsuarioT.Location = new System.Drawing.Point(160, 152);
			this.btAgregarUsuarioT.Name = "btAgregarUsuarioT";
			this.btAgregarUsuarioT.Size = new System.Drawing.Size(32, 23);
			this.btAgregarUsuarioT.TabIndex = 15;
			this.btAgregarUsuarioT.Text = ">>";
			this.btAgregarUsuarioT.ToolTip = "Agrega todos los usuarios de SQL de Abajo";
			this.btAgregarUsuarioT.Click += new System.EventHandler(this.btAgregarUsuarioT_Click);
			// 
			// gridControl4
			// 
			this.gridControl4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl4.DataMember = "Usuario";
			this.gridControl4.DataSource = this.cdsUsuarioAccess;
			// 
			// gridControl4.EmbeddedNavigator
			// 
			this.gridControl4.EmbeddedNavigator.Name = "";
			this.gridControl4.Location = new System.Drawing.Point(11, 8);
			this.gridControl4.MainView = this.grvUsuarioAccess;
			this.gridControl4.Name = "gridControl4";
			this.gridControl4.Size = new System.Drawing.Size(488, 136);
			this.gridControl4.TabIndex = 14;
			this.gridControl4.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.grvUsuarioAccess});
			// 
			// cdsUsuarioAccess
			// 
			this.cdsUsuarioAccess.BindingContextCtrl = this;
			this.cdsUsuarioAccess.DataLibrary = "LibEmpresa";
			this.cdsUsuarioAccess.DataLibraryUrl = "";
			this.cdsUsuarioAccess.DataSetDef = "dsUsuario";
			this.cdsUsuarioAccess.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsUsuarioAccess.Name = "cdsUsuarioAccess";
			this.cdsUsuarioAccess.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsUsuarioAccess.SchemaDef = null;
			// 
			// grvUsuarioAccess
			// 
			this.grvUsuarioAccess.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																														this.colClave,
																																														this.colidGrupo,
																																														this.colidUsuario,
																																														this.colNombre});
			this.grvUsuarioAccess.GridControl = this.gridControl4;
			this.grvUsuarioAccess.Name = "grvUsuarioAccess";
			this.grvUsuarioAccess.OptionsView.ShowGroupPanel = false;
			// 
			// colClave
			// 
			this.colClave.Caption = "Clave";
			this.colClave.FieldName = "Clave";
			this.colClave.Name = "colClave";
			this.colClave.Visible = true;
			this.colClave.VisibleIndex = 2;
			// 
			// colidGrupo
			// 
			this.colidGrupo.Caption = "idGrupo";
			this.colidGrupo.FieldName = "idGrupo";
			this.colidGrupo.Name = "colidGrupo";
			this.colidGrupo.Visible = true;
			this.colidGrupo.VisibleIndex = 1;
			// 
			// colidUsuario
			// 
			this.colidUsuario.Caption = "idUsuario";
			this.colidUsuario.FieldName = "idUsuario";
			this.colidUsuario.Name = "colidUsuario";
			// 
			// colNombre
			// 
			this.colNombre.Caption = "Nombre";
			this.colNombre.FieldName = "Nombre";
			this.colNombre.Name = "colNombre";
			this.colNombre.Visible = true;
			this.colNombre.VisibleIndex = 0;
			// 
			// DBCopiar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(552, 390);
			this.Controls.Add(this.xtraTabControl1);
			this.Name = "DBCopiar";
			this.Text = "Copiar de Access";
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresaAccess)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvEmpresaAccess)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).EndInit();
			this.xtraTabControl1.ResumeLayout(false);
			this.xtraTabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridControl2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvEmpresa)).EndInit();
			this.xtraTabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridControl3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUsuarioAccess)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvUsuarioAccess)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAgregaUno_Click(object sender, System.EventArgs e)
		{
			int iFila = grvEmpresaAccess.FocusedRowHandle;
			if (iFila < 0) 
			{
				MessageBox.Show("Seleccione Fila de la grilla de arriba con la empresa a copiar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			EmpresaCrear(iFila);
			cdsEmpresa.Update();
			cdsEmpresa.Clear();
			cdsEmpresa.Fill();
		}

		private void EmpresaCrear(int iFila)
		{			
			string stEmpresa = (string) grvEmpresaAccess.GetRowCellValue(iFila, colEmpresa);
			string stServidor = (string) grvEmpresaAccess.GetRowCellValue(iFila, colServidor);
			string stFacturacion = (string) grvEmpresaAccess.GetRowCellValue(iFila, colFacturacion);
			string stReportes = (string) grvEmpresaAccess.GetRowCellValue(iFila, colReportes);
			string stInserta = string.Format("Exec EmpresaCrear '{0}', '{1}', '{2}', '{3}'",
				stEmpresa, stServidor, stFacturacion, stReportes);
			Funcion.EjecutaSQL(cdsEmpresa, stInserta);
		}

		private void cdsEmpresa_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsEmpresa.Schema.Connections[0].ConnectionString = miFuncion.AccesoEmpresaBase();
		}

		private void btAgregaTodo_Click(object sender, System.EventArgs e)
		{
			if (grvEmpresaAccess.RowCount == 0)
			{
				MessageBox.Show("No existen empresas para pasar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			for (int i= 0; i < grvEmpresaAccess.RowCount; i++)
			{
				EmpresaCrear(i);
			}		
			cdsEmpresa.Update();
			cdsEmpresa.Clear();
			cdsEmpresa.Fill();
		}

		private void btQuitarUno_Click(object sender, System.EventArgs e)
		{
			int iFila = grvEmpresa.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione Fila de la grilla de abajo con la empresa a quitar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			EmpresaBorrar(iFila);
			cdsEmpresa.Update();
			cdsEmpresa.Clear();
			cdsEmpresa.Fill();
		}

		private void EmpresaBorrar(int iFila)
		{
			int idEmpresa = (int) grvEmpresa.GetRowCellValue(iFila, colidEmpresa1);
			string stInserta = string.Format("Delete Empresa Where idEmpresa = {0}", idEmpresa);
			Funcion.EjecutaSQL(cdsEmpresa, stInserta);
		}

		private void btQuitarTodo_Click(object sender, System.EventArgs e)
		{
			if (grvEmpresa.RowCount == 0)
			{
				MessageBox.Show("No existen empresas abajo para borrar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			for (int i= 0; i < grvEmpresa.RowCount; i++)
			{
				EmpresaBorrar(i);
			}
			cdsEmpresa.Update();
			cdsEmpresa.Clear();
			cdsEmpresa.Fill();
		}

		private void btQuitarUsuarioUno_Click(object sender, System.EventArgs e)
		{
			int iFila = grvEmpresa.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione Fila de la grilla de abajo con el a quitar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			UsuarioBorrar(iFila);
			cdsEmpresa.Update();
			cdsEmpresa.Clear();
			cdsEmpresa.Fill();
		}

		private void btAgregarUsuarioUno_Click(object sender, System.EventArgs e)
		{
			int iFila = grvUsuarioAccess.FocusedRowHandle;
			if (iFila < 0) 
			{
				MessageBox.Show("Seleccione la Fila de la grilla de arriba con el usuario a copiar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			UsuarioCrear(iFila);
			cdsEmpresa.Update();
			cdsEmpresa.Clear();
			cdsEmpresa.Fill();
		}

		private void UsuarioCrear(int iFila)
		{			
			try
			{
				string stNombre = grvUsuarioAccess.GetRowCellValue(iFila, colNombre).ToString().Trim();
				int idGrupo = (int) grvUsuarioAccess.GetRowCellValue(iFila, colidGrupo);
				string stClaveBasica = grvUsuarioAccess.GetRowCellValue(iFila, colClave).ToString();
				string stClave = Funcion.CreaHash(stClaveBasica+stNombre, idGrupo);
				string stInserta = string.Format("Exec UsuarioCrear '{0}', '{1}', {2}, '{3}'",
					stNombre, stClave, idGrupo, stClaveBasica);
				Funcion.EjecutaSQL(cdsEmpresa, stInserta);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error");
			}
		}

		private void btAgregarUsuarioT_Click(object sender, System.EventArgs e)
		{
			if (grvUsuarioAccess.RowCount == 0)
			{
				MessageBox.Show("No existen usuarios para pasar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			for (int i= 0; i < grvUsuarioAccess.RowCount; i++)
			{
				UsuarioCrear(i);
			}		
			cdsEmpresa.Update();
			cdsEmpresa.Clear();
			cdsEmpresa.Fill();
		}

		private void UsuarioBorrar(int iFila)
		{
			int idUsuario = (int) grvUsuario.GetRowCellValue(iFila, colidUsuario1);
			string stInserta = string.Format("Delete Usuario Where idUsuario = {0}", idUsuario);
			Funcion.EjecutaSQL(cdsEmpresa, stInserta);
		}

		private void btQuitarUsuarioT_Click(object sender, System.EventArgs e)
		{
			if (grvUsuario.RowCount == 0)
			{
				MessageBox.Show("No existen empresas abajo para borrar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			for (int i= 0; i < grvUsuario.RowCount; i++)
			{
				UsuarioBorrar(i);
			}
			cdsEmpresa.Update();
			cdsEmpresa.Clear();
			cdsEmpresa.Fill();		}
	}
}
