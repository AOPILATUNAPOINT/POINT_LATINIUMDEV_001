using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Diagnostics;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSeguimientoMantenimientoLocal.
	/// </summary>
	public class frmSeguimientoMantenimientoLocal : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private C1.Data.C1DataSet cdsSeguridad;
		private System.Windows.Forms.Button btnSalir;
		private System.ComponentModel.IContainer components;
		int idResponsable = 0;
		int idAsignado = 0;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTitulo;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaTentativa;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaAsignacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsMenu;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.TabPage tabPage3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdDocumento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsDocumento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsResponsable;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsAsignado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEmpresa;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoTipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		int idEmpresa = 0;
		bool bVer = false;
		bool bNuevo = false;
		int iNuevo = 1;
		int iEditar = 2;
		int iEliminar = 3;
		int iVer = 4;
		int idRequerimiento = 0;
		private System.Windows.Forms.GroupBox groupBox1;
		bool bRequerimientosDocumento = false;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.GroupBox groupBox2;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRequerimiento;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEmpresas;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAsignado;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		public System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotaIterna;
		private System.Windows.Forms.Label lblNotaInterna;
		private System.Windows.Forms.TabControl tabRequerimiento;
		string Title = "";
		public static string sconexionEmpresaSeguimiento = @"Password=Bl45o6$9;Persist Security Info=True;User ID= {0};Initial Catalog={1};Data Source={2}";
		string getsconexionEmpresaSeguimiento = "";
		int icontadoringreso = 0;
		private System.Windows.Forms.ToolTip toolTip1;
		int iTipo = 0;
		
		public frmSeguimientoMantenimientoLocal(int IdResponsable,int IdAsignado, int IdEmpresa, int ITipo)
		{
			idResponsable = IdResponsable;
			idAsignado = IdAsignado;
			idEmpresa = IdEmpresa;
			iTipo = ITipo;
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEmpresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSeguimientoMantenimientoLocal));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("id_seg");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel_ReqPantalla");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombrePadre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreHijo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreSubHijo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDocumento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeguimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idResponsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsignado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sCarpeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sUrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDocumento", 0);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEliminar", 1);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 2);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDocumento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSeguimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idResponsable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sCarpeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sUrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.lblEstado = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTitulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsAsignado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsEmpresa = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsResponsable = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.tabRequerimiento = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.txtNotaIterna = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblNotaInterna = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.dtFechaTentativa = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.dtFechaAsignacion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.udsMenu = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.ugdDocumento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsDocumento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.btnSalir = new System.Windows.Forms.Button();
			this.udsRequerimientoTipo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label12 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbTipoRequerimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbEmpresas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbAsignado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTitulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAsignado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).BeginInit();
			this.tabRequerimiento.SuspendLayout();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNotaIterna)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaTentativa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaAsignacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsMenu)).BeginInit();
			this.tabPage3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRequerimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAsignado)).BeginInit();
			this.SuspendLayout();
			// 
			// lblEstado
			// 
			this.lblEstado.Location = new System.Drawing.Point(8, 288);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.TabIndex = 15;
			this.lblEstado.Text = "* Estado:";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(8, 16);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(48, 23);
			this.label5.TabIndex = 16;
			this.label5.Text = "* Título:";
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(8, 48);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 23);
			this.label6.TabIndex = 18;
			this.label6.Text = "* Descripción:";
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.AcceptsReturn = true;
			this.txtDescripcion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance1;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(88, 48);
			this.txtDescripcion.MaxLength = 2000;
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(736, 104);
			this.txtDescripcion.TabIndex = 799;
			this.toolTip1.SetToolTip(this.txtDescripcion, "Descripción de lo que necesita que se realice para el requerimiento");
			// 
			// txtTitulo
			// 
			this.txtTitulo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTitulo.Appearance = appearance2;
			this.txtTitulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTitulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTitulo.Location = new System.Drawing.Point(88, 16);
			this.txtTitulo.MaxLength = 50;
			this.txtTitulo.Name = "txtTitulo";
			this.txtTitulo.Size = new System.Drawing.Size(736, 21);
			this.txtTitulo.TabIndex = 800;
			this.toolTip1.SetToolTip(this.txtTitulo, "Se decribe brevemente un título de su requerimiento.");
			// 
			// cmbEstado
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance3;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.udsEstado;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEstado.DisplayMember = "Descripcion";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(88, 288);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(200, 21);
			this.cmbEstado.TabIndex = 801;
			this.cmbEstado.ValueMember = "idEstadoRequerimiento";
			// 
			// udsEstado
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsEstado.Band.Columns.AddRange(new object[] {
																													ultraDataColumn1,
																													ultraDataColumn2});
			// 
			// udsAsignado
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.udsAsignado.Band.Columns.AddRange(new object[] {
																														ultraDataColumn3,
																														ultraDataColumn4});
			// 
			// udsEmpresa
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.udsEmpresa.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn5,
																													 ultraDataColumn6});
			// 
			// udsResponsable
			// 
			ultraDataColumn7.DataType = typeof(System.UInt32);
			this.udsResponsable.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn7,
																															 ultraDataColumn8});
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(792, 16);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(82, 23);
			this.btnGuardar.TabIndex = 805;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// tabRequerimiento
			// 
			this.tabRequerimiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabRequerimiento.Controls.Add(this.tabPage1);
			this.tabRequerimiento.Controls.Add(this.tabPage3);
			this.tabRequerimiento.Enabled = false;
			this.tabRequerimiento.Location = new System.Drawing.Point(16, 128);
			this.tabRequerimiento.Name = "tabRequerimiento";
			this.tabRequerimiento.SelectedIndex = 0;
			this.tabRequerimiento.Size = new System.Drawing.Size(856, 392);
			this.tabRequerimiento.TabIndex = 806;
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.txtNotaIterna);
			this.tabPage1.Controls.Add(this.lblNotaInterna);
			this.tabPage1.Controls.Add(this.label8);
			this.tabPage1.Controls.Add(this.dtFechaTentativa);
			this.tabPage1.Controls.Add(this.label7);
			this.tabPage1.Controls.Add(this.label5);
			this.tabPage1.Controls.Add(this.txtTitulo);
			this.tabPage1.Controls.Add(this.txtDescripcion);
			this.tabPage1.Controls.Add(this.label6);
			this.tabPage1.Controls.Add(this.dtFechaAsignacion);
			this.tabPage1.Controls.Add(this.cmbEstado);
			this.tabPage1.Controls.Add(this.lblEstado);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(848, 366);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "* Descripción";
			// 
			// txtNotaIterna
			// 
			this.txtNotaIterna.AcceptsReturn = true;
			this.txtNotaIterna.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotaIterna.Appearance = appearance4;
			this.txtNotaIterna.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotaIterna.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotaIterna.Location = new System.Drawing.Point(88, 200);
			this.txtNotaIterna.MaxLength = 2000;
			this.txtNotaIterna.Multiline = true;
			this.txtNotaIterna.Name = "txtNotaIterna";
			this.txtNotaIterna.Size = new System.Drawing.Size(736, 72);
			this.txtNotaIterna.TabIndex = 1004;
			// 
			// lblNotaInterna
			// 
			this.lblNotaInterna.Location = new System.Drawing.Point(8, 208);
			this.lblNotaInterna.Name = "lblNotaInterna";
			this.lblNotaInterna.Size = new System.Drawing.Size(112, 23);
			this.lblNotaInterna.TabIndex = 1002;
			this.lblNotaInterna.Text = "Nota interna:";
			this.lblNotaInterna.Click += new System.EventHandler(this.label9_Click);
			// 
			// label8
			// 
			this.label8.Location = new System.Drawing.Point(296, 168);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(120, 24);
			this.label8.TabIndex = 1001;
			this.label8.Text = "* Fecha tentativa:";
			// 
			// dtFechaTentativa
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaTentativa.Appearance = appearance5;
			this.dtFechaTentativa.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaTentativa.DateButtons.Add(dateButton1);
			this.dtFechaTentativa.Format = "dd/MM/yyyy hh:mm:ss";
			this.dtFechaTentativa.Location = new System.Drawing.Point(416, 168);
			this.dtFechaTentativa.Name = "dtFechaTentativa";
			this.dtFechaTentativa.NonAutoSizeHeight = 23;
			this.dtFechaTentativa.Size = new System.Drawing.Size(208, 21);
			this.dtFechaTentativa.SpinButtonsVisible = true;
			this.dtFechaTentativa.TabIndex = 1000;
			this.dtFechaTentativa.Value = ((object)(resources.GetObject("dtFechaTentativa.Value")));
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(8, 168);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(80, 32);
			this.label7.TabIndex = 999;
			this.label7.Text = "* Fecha asignación:";
			// 
			// dtFechaAsignacion
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaAsignacion.Appearance = appearance6;
			this.dtFechaAsignacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaAsignacion.DateButtons.Add(dateButton2);
			this.dtFechaAsignacion.Format = "dd/MM/yyyy hh:mm:ss";
			this.dtFechaAsignacion.Location = new System.Drawing.Point(88, 168);
			this.dtFechaAsignacion.Name = "dtFechaAsignacion";
			this.dtFechaAsignacion.NonAutoSizeHeight = 23;
			this.dtFechaAsignacion.ReadOnly = true;
			this.dtFechaAsignacion.Size = new System.Drawing.Size(200, 21);
			this.dtFechaAsignacion.SpinButtonsVisible = true;
			this.dtFechaAsignacion.TabIndex = 998;
			this.dtFechaAsignacion.Value = ((object)(resources.GetObject("dtFechaAsignacion.Value")));
			// 
			// udsMenu
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn9.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn10.DefaultValue = false;
			this.udsMenu.Band.Columns.AddRange(new object[] {
																												ultraDataColumn9,
																												ultraDataColumn10,
																												ultraDataColumn11,
																												ultraDataColumn12,
																												ultraDataColumn13});
			// 
			// tabPage3
			// 
			this.tabPage3.Controls.Add(this.ugdDocumento);
			this.tabPage3.Location = new System.Drawing.Point(4, 22);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Size = new System.Drawing.Size(848, 366);
			this.tabPage3.TabIndex = 2;
			this.tabPage3.Text = "Documento *";
			// 
			// ugdDocumento
			// 
			this.ugdDocumento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdDocumento.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdDocumento.DataSource = this.udsDocumento;
			appearance7.BackColor = System.Drawing.Color.White;
			this.ugdDocumento.DisplayLayout.Appearance = appearance7;
			this.ugdDocumento.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.Format = "dd/MM/yyyy hh:mm:ss";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 150;
			ultraGridColumn4.AutoEdit = false;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Format = "dd/MM/yyy hh:mm:ss";
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.Width = 352;
			ultraGridColumn5.Header.Caption = "Seguimiento";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn5.Width = 79;
			ultraGridColumn6.AutoEdit = false;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.Caption = "Dirección archivo";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 366;
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 135;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 135;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 137;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 154;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 88;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			ultraGridColumn12.CellAppearance = appearance8;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			ultraGridColumn12.CellButtonAppearance = appearance9;
			ultraGridColumn12.Header.Caption = "...";
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn12.Width = 28;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			ultraGridColumn13.CellAppearance = appearance10;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			ultraGridColumn13.CellButtonAppearance = appearance11;
			ultraGridColumn13.Header.Caption = "...";
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn13.Width = 25;
			appearance12.Image = ((object)(resources.GetObject("appearance12.Image")));
			ultraGridColumn14.CellAppearance = appearance12;
			appearance13.Image = ((object)(resources.GetObject("appearance13.Image")));
			ultraGridColumn14.CellButtonAppearance = appearance13;
			ultraGridColumn14.Header.Caption = "...";
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn14.Width = 24;
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
																										 ultraGridColumn14});
			this.ugdDocumento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ugdDocumento.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdDocumento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdDocumento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdDocumento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.ugdDocumento.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 10F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdDocumento.DisplayLayout.Override.HeaderAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.ugdDocumento.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdDocumento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdDocumento.Location = new System.Drawing.Point(16, 16);
			this.ugdDocumento.Name = "ugdDocumento";
			this.ugdDocumento.Size = new System.Drawing.Size(816, 328);
			this.ugdDocumento.TabIndex = 810;
			this.ugdDocumento.AfterRowUpdate += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ugdDocumento_AfterRowUpdate);
			this.ugdDocumento.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ugdDocumento_AfterRowInsert);
			this.ugdDocumento.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdDocumento_ClickCellButton);
			this.ugdDocumento.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdDocumento_AfterCellUpdate);
			this.ugdDocumento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdDocumento_InitializeLayout);
			// 
			// udsDocumento
			// 
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn15.DataType = typeof(System.DateTime);
			ultraDataColumn18.DataType = typeof(int);
			this.udsDocumento.Band.Columns.AddRange(new object[] {
																														 ultraDataColumn14,
																														 ultraDataColumn15,
																														 ultraDataColumn16,
																														 ultraDataColumn17,
																														 ultraDataColumn18,
																														 ultraDataColumn19,
																														 ultraDataColumn20,
																														 ultraDataColumn21,
																														 ultraDataColumn22});
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(792, 56);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 807;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnSalir, "Salir");
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// udsRequerimientoTipo
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.udsRequerimientoTipo.Band.Columns.AddRange(new object[] {
																																		 ultraDataColumn23,
																																		 ultraDataColumn24});
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-688, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(2280, 8);
			this.groupBox1.TabIndex = 1032;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.label12);
			this.groupBox2.Controls.Add(this.txtNumero);
			this.groupBox2.Controls.Add(this.cmbTipoRequerimiento);
			this.groupBox2.Controls.Add(this.label11);
			this.groupBox2.Controls.Add(this.cmbResponsable);
			this.groupBox2.Controls.Add(this.cmbEmpresas);
			this.groupBox2.Controls.Add(this.cmbAsignado);
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Location = new System.Drawing.Point(16, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(768, 100);
			this.groupBox2.TabIndex = 1033;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(544, 64);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(44, 16);
			this.label12.TabIndex = 821;
			this.label12.Text = "Numero";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance18;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(600, 64);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.ReadOnly = true;
			this.txtNumero.Size = new System.Drawing.Size(152, 21);
			this.txtNumero.TabIndex = 820;
			// 
			// cmbTipoRequerimiento
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRequerimiento.Appearance = appearance19;
			this.cmbTipoRequerimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipoRequerimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRequerimiento.DataSource = this.udsRequerimientoTipo;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.Caption = "Descripción";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 157;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbTipoRequerimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipoRequerimiento.DisplayMember = "Descripcion";
			this.cmbTipoRequerimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRequerimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRequerimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRequerimiento.Location = new System.Drawing.Point(376, 64);
			this.cmbTipoRequerimiento.Name = "cmbTipoRequerimiento";
			this.cmbTipoRequerimiento.Size = new System.Drawing.Size(152, 21);
			this.cmbTipoRequerimiento.TabIndex = 819;
			this.cmbTipoRequerimiento.ValueMember = "idRequerimientoTipo";
			this.cmbTipoRequerimiento.ValueChanged += new System.EventHandler(this.cmbTipoRequerimiento_ValueChanged);
			this.cmbTipoRequerimiento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoRequerimiento_InitializeLayout);
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(272, 64);
			this.label11.Name = "label11";
			this.label11.TabIndex = 818;
			this.label11.Text = "* Requerimiento:";
			// 
			// cmbResponsable
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance20;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResponsable.DataSource = this.udsResponsable;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 150;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResponsable.Enabled = false;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(104, 24);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.Size = new System.Drawing.Size(248, 21);
			this.cmbResponsable.TabIndex = 817;
			this.cmbResponsable.ValueMember = "idPersonal";
			// 
			// cmbEmpresas
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEmpresas.Appearance = appearance21;
			this.cmbEmpresas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEmpresas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresas.DataSource = this.udsEmpresa;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 152;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbEmpresas.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbEmpresas.DisplayMember = "Empresa";
			this.cmbEmpresas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEmpresas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEmpresas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresas.Location = new System.Drawing.Point(104, 64);
			this.cmbEmpresas.Name = "cmbEmpresas";
			this.cmbEmpresas.Size = new System.Drawing.Size(152, 21);
			this.cmbEmpresas.TabIndex = 816;
			this.cmbEmpresas.ValueMember = "idempresa";
			this.cmbEmpresas.ValueChanged += new System.EventHandler(this.cmbEmpresas_ValueChanged);
			this.cmbEmpresas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEmpresas_InitializeLayout);
			// 
			// cmbAsignado
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAsignado.Appearance = appearance22;
			this.cmbAsignado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbAsignado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAsignado.DataSource = this.udsAsignado;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 251;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbAsignado.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbAsignado.DisplayMember = "Nombre";
			this.cmbAsignado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAsignado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbAsignado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAsignado.Location = new System.Drawing.Point(512, 24);
			this.cmbAsignado.Name = "cmbAsignado";
			this.cmbAsignado.Size = new System.Drawing.Size(240, 21);
			this.cmbAsignado.TabIndex = 815;
			this.cmbAsignado.ValueMember = "idPersonal";
			this.cmbAsignado.ValueChanged += new System.EventHandler(this.cmbAsignado_ValueChanged);
			this.cmbAsignado.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbAsignado_InitializeLayout);
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(8, 64);
			this.label4.Name = "label4";
			this.label4.TabIndex = 814;
			this.label4.Text = "* Empresa:";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(400, 24);
			this.label3.Name = "label3";
			this.label3.TabIndex = 813;
			this.label3.Text = "* Asignado:";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 24);
			this.label1.Name = "label1";
			this.label1.TabIndex = 812;
			this.label1.Text = "* Responsable:";
			// 
			// frmSeguimientoMantenimientoLocal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(888, 534);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.tabRequerimiento);
			this.Controls.Add(this.btnGuardar);
			this.Name = "frmSeguimientoMantenimientoLocal";
			this.toolTip1.SetToolTip(this, "Guardar requerimiento");
			this.Load += new System.EventHandler(this.frmSeguimientoMantenimientoLocal_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTitulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAsignado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).EndInit();
			this.tabRequerimiento.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNotaIterna)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaTentativa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaAsignacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsMenu)).EndInit();
			this.tabPage3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRequerimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAsignado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		bool bEdicion = false;
		bool bImpreso = false;
		private Acceso miAcceso;
		decimal dValor = 0.00m;
		decimal dSaldo = 0.00m;		
		#endregion Variables
		string txtDirectorioarchivo = "";

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void frmSeguimientoMantenimientoLocal_Load(object sender, System.EventArgs e)
		{
				if(iTipo == iNuevo)
				{
					bNuevo = true;
					Title = "Nuevo requerimiento";
				}
				if(iTipo == iVer)
				{
					bVer = true;
					Title = "Ver requerimiento";
				}
				this.Text = Title;
				//Carga id usuario tabla Inicio
				int idUsuarioInicio = Funcion.iEjecucion(frmRequerimiento.sconexionInicio, string.Format("select top 1 idUsuario from Usuarios where NOMBRE = '{0}' Order By idUsuario", MenuLatinium.stUsuario),false);
				//Consultar el rol del usuario de ingreso
				int idrolUsuarioInicio = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, string.Format("EXEC RequerimientoRolUsuario {0}", idUsuarioInicio), false);
				//Cargar lista empresas vinculadas al usuario de ingreso
				this.cmbEmpresas.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionInicio, string.Format("EXEC RequerimientoRolEmpresa {0},{1}", idUsuarioInicio, idrolUsuarioInicio));
				//Cargar id de empresa de ingreso
				this.cmbEmpresas.Value = Funcion.iEjecucion(frmRequerimiento.sconexionInicio, string.Format("Select Top 1 e.idEmpresa From Empresas e Where e.BaseDeDatos = '{0}'", MenuLatinium.stDirFacturacion.Trim()),false);
				//Cargar el nombre de la empresa que se logeo
				string sempresa = Funcion.sEjecucion(frmRequerimiento.sconexionInicio, string.Format("Select Top 1 e.Empresa From Empresas e Where e.BaseDeDatos = '{0}'", MenuLatinium.stDirFacturacion.Trim()));
				//CargarResponsable
				this.cmbResponsable.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionInicio, "Exec RequerimientoResponsable");
				this.cmbResponsable.Value = idResponsable;
				//CargarAsignado
				this.cmbAsignado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoAsignacionCon {0} , {1}, {2}, {3}", idResponsable, 0, 0, frmRequerimiento.iRequerimientoControlMantenimientoLocales));
				//this.cmbAsignado.Value = idAsignado;
				//Cargar EstadoRequerimiento
				this.cmbEstado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoEstado {0}", frmRequerimiento.iRequerimientoControlMantenimientoLocales));
				//Cargar tipo de requermiento
				this.cmbTipoRequerimiento.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoTipoConsultar {0}", frmRequerimiento.iRequerimientoControlMantenimientoLocales));
				//Cargar la fecha de ingreso
				this.dtFechaAsignacion.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
				//this.dtFechaTentativa.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
				DateTime dtFecha = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);
				this.dtFechaAsignacion.CalendarInfo.MinDate = DateTime.Parse("01/01/2014");
				//this.dtFechaTentativa.CalendarInfo.MinDate = DateTime.Parse("01/01/2014");
				this.dtFechaAsignacion.CalendarInfo.MaxDate = dtFecha;
				//this.dtFechaTentativa.CalendarInfo.MaxDate = dtFecha;
				this.dtFechaAsignacion.Value = "01/" + dtFecha.Month.ToString() + "/" + dtFecha.Year.ToString();
				this.dtFechaAsignacion.Value = dtFecha;
				//Cargar documentos si existió algún problema con el guardado
				RequerimientoDocumentoCon(1);
				ControlNuevo();
				RequerimientoRolPermisoCon();
				//Configuración FechaTentativa
				ConfigFechaTentativa();

		}

		private void ConfigFechaTentativa()
		{
			//LA FECHA TENTATIVA DE ATENCION DEBE SER MINIMO 3 DIAS Control de ingreso tentativo a 3 días, 
			//se debería ubicar directo los 3 días al crear
			DateTime dtFechaAsignacionMin = Convert.ToDateTime(this.dtFechaAsignacion.Value);
			DateTime dtFechaAsignacionMax = Convert.ToDateTime(this.dtFechaAsignacion.Value);
			//Aumentar tres días esto se consulta en la tabla configuraciones
			//Intervalo de refresh al timer parametrizado
			int TamanoMinimo = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "TamanoMinimo", frmRequerimiento.iConfiguracionDiasMinimoFechaCreacionCOntrolMantenimiento),false);
			int TamanoMaximo = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "TamanoMaximo", frmRequerimiento.iConfiguracionDiasMinimoFechaCreacionCOntrolMantenimiento),false);

			dtFechaAsignacionMin = dtFechaAsignacionMin.AddDays(TamanoMinimo);
			
			//Mostrar hasta la fecha de tres días
			this.dtFechaTentativa.Value = dtFechaAsignacionMin;
			this.dtFechaTentativa.CalendarInfo.MinDate = dtFechaAsignacionMin;

			//Mostrar hasta fecha más de 999 días
			dtFechaAsignacionMax = dtFechaAsignacionMax.AddDays(TamanoMaximo);
			this.dtFechaTentativa.CalendarInfo.MaxDate = dtFechaAsignacionMax;

		}

		private void RequerimientoRolPermisoCon()
		{
			try
			{
				//Permiso Valores técnicos
				bool bvalorestecnicos =
					Funcion.bEjecucion(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoRolPermisoCon {0}, '{1}', {2}", 
					(int)this.cmbResponsable.Value,
					frmRequerimiento.iValoresTecnicos,
					frmRequerimiento.iRequerimientoControlMantenimientoLocales
					), false); 

				if(bvalorestecnicos == false)
				{
					this.lblNotaInterna.Visible = false;
					this.txtNotaIterna.Visible = false;
				}


			}
			catch(Exception ex)
			{
				string sms = ex.Message;
				return;
			}
			
		}


		private void ControlNuevo()
		{
			if(bNuevo)
			{
				this.cmbEstado.Value = 1; //Estado representa nuevo, y debe ser ocultado.
				this.cmbEstado.Visible = false;
				this.lblEstado.Visible = false;
			}
		}
		//
		private void ActivarCreacionRequerimiento()
		{
			bool bActivar = false;

			try
			{
				if (this.cmbAsignado.ActiveRow != null && this.cmbTipoRequerimiento.ActiveRow != null)
				{
					bActivar = true;
				}
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			if(bActivar)
			{
				this.tabRequerimiento.Enabled = true;
			}
			else
			{
				this.tabRequerimiento.Enabled = false;
			}

		}



		private void RequerimientoDocumentoCon(int Nuevo)
		{
			int setidAsignado = 0;
			try {setidAsignado = (int)this.cmbAsignado.Value;} 
			catch(Exception ex){string sms = ex.Message;}
				
			this.ugdDocumento.DataSource = 
				FuncionesProcedimientos.dtEjecucionGeneral(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoDocumentoCon {0}, {1}, {2}, '{3}'", 
				Nuevo, (int)this.cmbResponsable.Value, setidAsignado, Convert.ToDateTime(this.dtFechaAsignacion.Value).ToString("yyyyMMdd")
				)); 
		}


		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Valida campos  
			if (!Validacion.vbComboVacio(this.cmbResponsable, "Seleccione el responsable ")) return;
			if (!Validacion.vbComboVacio(this.cmbEmpresas, "Seleccione la empresa ")) return;
			if (!Validacion.vbComboVacio(this.cmbAsignado, "Seleccione el asignado ")) return;
			if (!Validacion.vbComboVacio(this.cmbEstado, "Seleccione el estado ")) return;
			if (!Validacion.vbComboVacio(this.cmbTipoRequerimiento, "Seleccione el tipo de requerimiento ")) return;
			if (!Validacion.vbTexto(this.txtTitulo, 10, 50, "Ingrese un título")) return;
			if (!Validacion.vbTexto(this.txtDescripcion, 20, 2000, "Ingrese una descripción")) return;
			if (!Validacion.vbValidaFechasDesdeHasta(this.dtFechaAsignacion, this.dtFechaTentativa)) return;
			#endregion Valida campos  

			#region Valida documentos
			if(bRequerimientosDocumento)
			{
				// Validación si se ha ingresado al menos un documentos eso solo sería de hacer un select en la tabla documentosrequerimiento con el estado 1
				int ugdDocumentotamano = this.ugdDocumento.Rows.Count;
				if (ugdDocumentotamano == 0)
				{
					MessageBox.Show(string.Format("Proceso Finalizado, se debe cargar mínimo un documento"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (ugdDocumentotamano > 3)
				{
					MessageBox.Show(string.Format("Proceso Finalizado, se debe cargar máximo 3 documentos"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			#endregion Valida documentos 
			
			#region	Generar rutas del documento
				
			string sCarpeta = "";
			string sUrlNube = "";
			string sRuta = "";
			string carpeta = "";
			string carpeta_3 = "";
			string dtFecha = "";
			string carpeta_1 = "";
			string carpeta_2 = "";
			string txtURL = "";
			string sextencion = "";
			string squery = "";
			string squery2 = "";
			string squery3 = "";

			carpeta_1 = @"\\192.168.1.10\Proyectos\RequerimientoDocumento";
			carpeta_2 = @"\\192.168.1.10\Proyectos\RequerimientoDocumento\"+MenuLatinium.stDirFacturacion;
				
			using (SqlConnection oConexion = new SqlConnection(frmRequerimiento.sconexionPoint)) //("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))					
			{
				// Mostrar el modal de "procesando"
				MenuLatinium.MostrarModalProcesando();

				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;
						
				try
				{							
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Numeracion
					if (bNuevo)
					{
						try
						{
							//Fijarse bien en el id de la numeración en producción porque puede ser otra.
							squery = string.Format("Exec NumeracionLocales 88, 0, 0");
							oCmdActualiza.CommandText = squery;
							this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
						}
						catch(Exception ex)
						{
							oTransaction.Rollback();
							Funcion.LogSistema(frmRequerimiento.sconexionInicio, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
							MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return;
						}
					}
					#endregion Numeracion

					#region Garbar requerimiento
					int idSeguimiento = 0;
					try
					{
						//Guardar requerimiento para luego darle seguimiento
						squery2 = string.Format("Exec RequerimientoCrear {0}, '{1}', {2}, '{3}', '{4}', {5}, '{6}', {7}, {8}, '{9}', '{10}', '{11}', '{12}', '{13}', {14}", 
							(int)this.cmbEmpresas.Value, 
							(string)this.txtDescripcion.Value, 
							(int)this.cmbEstado.Value,
							Convert.ToDateTime(this.dtFechaTentativa.Value).ToString("yyyyMMdd HH:mm:ss"),
							(string)this.txtNotaIterna.Value,
							(int)this.cmbResponsable.Value, 
							(string)this.txtTitulo.Value,
							(int)this.cmbTipoRequerimiento.Value,
							(int)this.cmbAsignado.Value,
							(string)this.txtNumero.Text,
							Convert.ToDateTime(this.dtFechaAsignacion.Value).ToString("yyyyMMdd HH:mm:ss"),
							(string)this.cmbResponsable.Text,
							(string)this.cmbAsignado.Text,
							(string)this.cmbEmpresas.Text,
							frmRequerimiento.iRequerimientoControlMantenimientoLocales
							);
						oCmdActualiza.CommandText = squery2;
						idSeguimiento = (int)oCmdActualiza.ExecuteScalar();
					}
					catch(Exception ex)
					{
						oTransaction.Rollback();
						Funcion.LogSistema(frmRequerimiento.sconexionInicio, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery2, Funcion.slogtipoerror);
						MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}
					#endregion Grabar requerimiento

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugdDocumento.Rows.All)	
						if ((int)dr.Cells["idRequerimientoDocumento"].Value != 0)
						{
							#region Variables
							sCarpeta = "";
							sUrlNube = "";
							sRuta = "";
							carpeta = "";
							carpeta_3 = "";
							dtFecha = "";
							txtURL = "";
							sRuta = dr.Cells["txtDirectorioarchivo"].Value.ToString();		
							dtFecha = Convert.ToDateTime(this.dtFechaAsignacion.Value).ToString("yyyyMMdd");
							#endregion Variables
								
							#region carpetas
							carpeta = @"\\192.168.1.10\Proyectos\RequerimientoDocumento\"+MenuLatinium.stDirFacturacion+@"\" + dtFecha + "_" + Convert.ToString (this.cmbResponsable.Text) + "_" + this.txtNumero.Text + "_" + idSeguimiento ;
							carpeta_3 = @""+MenuLatinium.stDirFacturacion+@"\" + dtFecha + "_" + Convert.ToString (this.cmbResponsable.Value) + "_" +  this.txtNumero.Text + "_" + idSeguimiento ;
							#endregion carpetas

							#region directorio

							Process.Start("cmd.exe", "/C" + "net use \\192.168.1.10  /user:latinium 10dgpv15.p");
							Directory.GetDirectories(@"\\192.168.1.10\Proyectos");

							if (!Directory.Exists(carpeta_1))
							{
								Directory.CreateDirectory(carpeta);
							}
							if (!Directory.Exists(carpeta_2))
							{
								Directory.CreateDirectory(carpeta);
							}
							if (!Directory.Exists(carpeta))
							{
								Directory.CreateDirectory(carpeta);
							}
							#endregion directorio
										
							#region Generar ruta documento
							txtDirectorioarchivo = sRuta;
							txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",txtDirectorioarchivo,@"\"));
							//Coger extensión del nombre del archivo para guardarlo con un nombre generado
							sextencion = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}'))) AS Archivo",txtDirectorioarchivo,"."));
							//@ARCHIVO_EXTENSION = RIGHT(@RUTA_ARCHIVO, CHARINDEX('.', REVERSE(@RUTA_ARCHIVO)))
							//Generar nombre documento
							string sSQLNumero = string.Format("Exec NumeracionLocales 89, 0, 0");
							oCmdActualiza.CommandText = sSQLNumero;
							string sNombreDocuento = oCmdActualiza.ExecuteScalar().ToString();
								
							carpeta  = (carpeta + @"\" + sNombreDocuento.Replace(" ", "_") + sextencion); // txtURL.Replace(" ", "_"));
							#region Existe Documento
							if (System.IO.File.Exists(carpeta))
							{
								MessageBox.Show(string.Format("El archivo '{0}' ya existe. ", txtURL.Replace(" ", "_")), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
								return;
							}
							else
							{
								System.IO.File.Copy (txtDirectorioarchivo ,carpeta, false);
								sCarpeta = carpeta;
								sUrlNube = carpeta_3.Replace(" ", "_")+ @"\" + sNombreDocuento.Replace(" ", "_") + sextencion;//+ txtURL.Replace(" ", "_");
								//MessageBox.Show(string.Format(" Archivo '{0}' subido . ", txtURL.Replace(" ", "_")), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
								//Si se guardo crear un registro por el nuevo guardado
								try
								{
									squery3 = string.Format("Exec RequerimientoDocumentoNuevo '{0}', {1}, {2}, {3}, '{4}', '{5}', '{6}',{7},{8}, '{9}'", 
										MenuLatinium.stUsuario, 
										0, //Significa nuevo 
										(int)this.cmbResponsable.Value,
										(int)this.cmbAsignado.Value, 
										txtDirectorioarchivo,
										sCarpeta,
										sUrlNube,
										idSeguimiento,
										dr.Cells["idRequerimientoDocumento"].Value.ToString(),
										Convert.ToDateTime(this.dtFechaAsignacion.Value).ToString("yyyyMMdd")
										);
									oCmdActualiza.CommandText = squery3;
									oCmdActualiza.ExecuteNonQuery();
								}
								catch(Exception ex)
								{
									oTransaction.Rollback();
									//Funcion.WriteLog(ex.Message);
									Funcion.LogSistema(frmRequerimiento.sconexionInicio, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery3, Funcion.slogtipoerror);
									MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
									return;
								}
							}
							#endregion Existe Documento

							#endregion Generar ruta documento
						}
					oTransaction.Commit();
					oConexion.Close();

					#region Enviar notificacion
					frmRequerimientoMantenimientoLocal.RequerimientoNotificacionCrear(
						(int)this.cmbResponsable.Value, (int)this.cmbAsignado.Value, frmRequerimiento.iRequerimientoMensajeNuevo, idSeguimiento, 1
						);
					#endregion Enviar notificacion

					MessageBox.Show(string.Format("Proceso Finalizado Correctamente, se generó el siguiente número de requerimiento: {0}", this.txtNumero.Text), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.DialogResult = DialogResult.OK;


				}
				catch (Exception ex)
				{
					this.txtNumero.Text = "";
					oTransaction.Rollback();
					MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					//Funcion.WriteLog(ex.Message);
					Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery + " || "+ squery2 + " || " + squery3, Funcion.slogtipoerror);
				}
				finally
				{
					oConexion.Close();
					// Ocultar el modal de "procesando" una vez que se haya completado la consulta
					MenuLatinium.OcultarModalProcesando();
				}
			}
			#endregion Generar rutas del documento
		}

		private void ugdMenu_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Validacion
			if (!Validacion.vbComboVacio(this.cmbResponsable, "Seleccione un responsable")) return;
			if (!Validacion.vbComboVacio(this.cmbAsignado, "Seleccione un asignado"))return;
			#endregion Validacion

			//Propiedad que selecciona la pantalla del requerimiento
			string sColumna = e.Cell.Column.Header.Caption;

			if (e.Cell.Column.Header.Caption.ToString() == sColumna)
			{
					ActualizaValores((int)e.Cell.Row.Cells["id_seg"].Value,
					(bool)e.Cell.Row.Cells["Sel_ReqPantalla"].Value,
					(string)e.Cell.Row.Cells["NombrePadre"].Value, 
					(string)e.Cell.Row.Cells["NombreHijo"].Value,
					(string)e.Cell.Row.Cells["NombreSubHijo"].Value);
			}
		}

		private void ActualizaValores(int id_seg, bool Sel_ReqPantalla, string Modulo, string Pantalla, string SubPantalla)
		{

			string sSQL = string.Format("Exec RequerimientoMenuNuevo '{0}', {1}, {2}, {3}, {4}, {5}, '{6}', '{7}', '{8}'", 
				Convert.ToDateTime(this.dtFechaAsignacion.Value).ToString("yyyyMMdd HH:mm:ss"),
				(int)this.cmbResponsable.Value,
				(int)this.cmbAsignado.Value,
				1,//Nuevo
				id_seg , 
				Sel_ReqPantalla,
				Modulo,
				Pantalla,
				SubPantalla);
			Funcion.EjecusionEjecutaSQL(frmRequerimiento.sconexionPoint, sSQL, false);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ugdDocumento_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			//e.Row.Cells["idRequerimientoDocumento"].Value = 0;
			//e.Row.Cells["UrlCompartido"].Value = "";
			//e.Row.Cells["idSeguimiento"].Value = 0;
			//e.Row.Cells["FechaDeposito"].Value = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF); //DateTime.Today;
		}

		private void ugdDocumento_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{ 
			string txtURL ="";
			long itamano = 0;
			bool bResultado = false;
			string sMensaje = "";
			DataSet dsConfiguracion = new DataSet();

			if (e.Cell.Column.ToString() == "btnDocumento")	
			{

			#region Validacion
			//if (!Validacion.vbTexto((string)this.dtFechaAsignacion.Value, "50", "50", "Fecha asignación")) return;
			if (!Validacion.vbComboVacio(this.cmbResponsable, "Seleccione un responsable")) return;
			if (!Validacion.vbComboVacio(this.cmbAsignado, "Seleccione un asignado")) return;
			#endregion Validacion

			if (bNuevo)
			{
				#region Editar google 
				try
				{
					OpenFileDialog buscar = new OpenFileDialog ();
					if (buscar.ShowDialog () == DialogResult.OK)
					{
						  //Sacar el tamaño del archivo subido
							txtDirectorioarchivo = buscar.FileName;
							System.IO.FileInfo info = new System.IO.FileInfo(txtDirectorioarchivo);
							//El lenght está en tamaños bytes
							itamano = info.Length;
						  
						  dsConfiguracion = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfgVal {0}, {1}, {2}", frmRequerimiento.iConfiguracionDocumento, itamano, frmRequerimiento.iConfiguracionTamano));
						  bResultado = bool.Parse(dsConfiguracion.Tables[0].Rows[0]["bResultado"].ToString());
							sMensaje = dsConfiguracion.Tables[0].Rows[0]["sMensaje"].ToString();
						if(bResultado)
						{
							//Consultar la parametría
							txtURL = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",txtDirectorioarchivo,@"\"));
							//Si se guardo crear un registro por el nuevo guardado
							bool bCont=Funcion.bEjecucion(frmRequerimiento.sconexionPoint, string.Format("Exec RequerimientoDocumentoNuevo '{0}', {1}, {2}, {3}, '{4}', '{5}', '{6}',{7},{8}, '{9}'", 
								MenuLatinium.stUsuario, 
								1, //Significa nuevo 
								(int)this.cmbResponsable.Value,
								(int)this.cmbAsignado.Value, 
								txtDirectorioarchivo,
								"",
								"",
								0,
								0,
								Convert.ToDateTime(this.dtFechaAsignacion.Value).ToString("yyyyMMdd")
								), false);
							if (bCont == false)
							{
								MessageBox.Show(string.Format("El archivo '{0}' ya existe. ", txtURL.Replace(" ", "_")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
						}
						else
						{
								MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
						}

						RequerimientoDocumentoCon(1);
					}	
				}
				catch (Exception ex)
				{
					MessageBox.Show ("Error: "+ ex.Message);
				}
				#endregion Editar google 
			}
			else
			{
				#region error
				try
				{
					OpenFileDialog buscar = new OpenFileDialog ();
					if (buscar.ShowDialog () == DialogResult.OK)
					{
						txtDirectorioarchivo = buscar.FileName;
					}	
				}
				catch (Exception ex)
				{
					MessageBox.Show ("Error: "+ ex.Message);

				}
				#endregion error
			}
			}

			#region Ver
			if (e.Cell.Column.ToString() == "btnVer")	
			{
				try
				{
					if(bNuevo)
					{
						txtDirectorioarchivo = (string)e.Cell.Row.Cells["txtDirectorioarchivo"].Value;
						if (txtDirectorioarchivo.Length > 10 )
						{
							Process.Start(txtDirectorioarchivo);
						}
					}
				}
				catch (Exception ex)
				{
					string sms = ex.Message;
					return;						
				}
			}
			#endregion Ver

			#region Eliminar documento
			if (e.Cell.Column.ToString() == "btnEliminar")	
			{
				try
				{
					int itamanodocumentos = this.ugdDocumento.Rows.Count;
					if(itamanodocumentos == 0)
					{
						return;
					}
					//Si se guardo crear un registro por el nuevo guardado
					bool bCont =Funcion.bEjecucion(frmRequerimiento.sconexionPoint, string.Format("Exec RequerimientoDocumentoEliminar '{0}', {1}", 
						MenuLatinium.stUsuario, 
						(int)e.Cell.Row.Cells["idRequerimientoDocumento"].Value //id del documento
						), false);
					if (bCont == true)
					{
						// Eliminar el path
						txtDirectorioarchivo = ((string)e.Cell.Row.Cells["txtDirectorioarchivo"].Value);
						txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",txtDirectorioarchivo,@"\"));
//						System.IO.File.Delete(txtDirectorioarchivo) ;	
						MessageBox.Show(string.Format("El archivo '{0}' se eliminó correctamente. ", txtURL.Replace(" ", "_")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						//return;
					}
					RequerimientoDocumentoCon(1);
				}
				catch (Exception ex)
				{
					string sms = ex.Message;
					return;						
				}
			}
			#endregion Eliminar documento
		}

		private void ugdDocumento_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ugdDocumento_AfterRowUpdate(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
		
		}

		private void ugdDocumento_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		
		}

		private void ugdMenu_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		
		}

		private void ultraTextEditor1_ValueChanged(object sender, System.EventArgs e)
		{
			//MenuSeguridad();
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void label9_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ugdMenu_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbAsignado_ValueChanged(object sender, System.EventArgs e)
		{
			RequerimientoDocumentoCon(iNuevo);
			ActivarCreacionRequerimiento();
		}

		private void cmbAsignado_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			ActivarCreacionRequerimiento();
		}

		private void cmbTipoRequerimiento_ValueChanged(object sender, System.EventArgs e)
		{
			ActivarCreacionRequerimiento();
		}

		private void cmbTipoRequerimiento_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			ActivarCreacionRequerimiento();
		}

		private void cmbEmpresas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbEmpresas_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbEmpresas.ActiveRow != null)
			{
				//Variable de ingreso ya no coge la conexión inicial, sino la cargada en este combo
				icontadoringreso = 1;
				SqlConnection Conexion = new SqlConnection(frmRequerimiento.sconexionInicio);
				//Se debe hacer la conexión a la tabla de inicio y de la empresa traer los datos del servidor y base de datos
				string sSQL = string.Format("Select Empresa, BaseDeDatos, Servidor From Empresas Where idEmpresa = {0}", (int)this.cmbEmpresas.Value);
				string Empresa = "";
				string BaseDeDatos = "";
				string Servidor = "";
				string strHostName = string.Empty;
				strHostName = System.Environment.MachineName;

				SqlCommand command = new SqlCommand(sSQL, Conexion);
				Conexion.Open();
				SqlDataReader reader = command.ExecuteReader();
				reader.Read();
				if (reader.HasRows)
				{
					Empresa = (string)reader[0];
					BaseDeDatos = (string)reader[1];
								
					if (strHostName == "IEXP" && BaseDeDatos == "POINT"  || BaseDeDatos == "CREDISOLUTIONS")
					{
						Servidor = @"192.168.1.165,56399"; 
					}
					else 
					{
						Servidor = (string)reader[2];
					}
				}
				reader.Close();
				Conexion.Close();	

				string strUsuarioSeguimiento = MenuLatinium.stUsuario;
				getsconexionEmpresaSeguimiento = String.Format(sconexionEmpresaSeguimiento, strUsuarioSeguimiento,BaseDeDatos, Servidor);
			}
		}

		private void txtBusquedaMenu_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.ValidaLetraDigito(sender,e, ' ');
		}

	}
}
