using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Drawing.Printing;
using System.IO.IsolatedStorage;
using System.Net.Sockets;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmLineaCredito.
	/// </summary>
	public class frmLineaCredito : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEmpresa;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsResponsable;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsAsignado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoTipo;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Variable del diseñador requerida.
		bool bVer = false;
		bool bNuevo = false;
		//ids de permisos
		public static int iAgregar = 1;
		public static int iAsignar = 2;
		public static int iSeguimiento = 3;
		public static int iPermisos = 4;
		public static int iConfiguracion = 5;
		public static int iValoresTecnicos = 6;
		public static string sBaseRequerimeintos = "POINT";
		public static string sconexionPoint = @"Password=Bl45o6$9;Persist Security Info=True;User ID= 'AOPILATUÑA';Initial Catalog='POINT';Data Source='192.168.1.165,56399'";
		public static string sconexionInicio = @"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='INICIO';Data Source='192.168.1.165,56399'";

		public static int iConfiguracionTamano = 1;
		public static int iNuevo = 1;
		public static int iVer = 2;
		public static int iEliminar = 3;
		public static int iEditar = 4;

		public static int iConfiguracionDocumento = 2;
		public static int iConfiguracionTasa = 4;
		public static int iConfiguracionGarPor = 5;
		public static int iConfiguracionRegistroAvaluo = 6;
		public static int iConfiguracionCargarDocumentoAvaluo = 7;

		public static int idLineaCreditoTipoBanco = 1;
		public static int idLineaCreditoTipoProveedor = 2;
		
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label lblAsignacionesMias;
		private System.Windows.Forms.CheckBox chkAsignacionesMias;
		public System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Button btnPolitica;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsLineaCredito;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgLineaCredito;
		private System.ComponentModel.IContainer components;

		public frmLineaCredito()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAprobacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaRenovacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MontoTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tasa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeGarantia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmLineaCredito));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEmpresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAprobacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRenovacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MontoTotal");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tasa");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeGarantia");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 0);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnadir", 1);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			this.udsLineaCredito = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.udsResponsable = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsEmpresa = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsAsignado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsRequerimientoTipo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.lblAsignacionesMias = new System.Windows.Forms.Label();
			this.chkAsignacionesMias = new System.Windows.Forms.CheckBox();
			this.label12 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.udgLineaCredito = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnPolitica = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.udsLineaCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAsignado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			this.groupBox3.SuspendLayout();
			this.groupBox4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udgLineaCredito)).BeginInit();
			this.SuspendLayout();
			// 
			// udsLineaCredito
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			this.udsLineaCredito.Band.Columns.AddRange(new object[] {
																																ultraDataColumn1,
																																ultraDataColumn2,
																																ultraDataColumn3,
																																ultraDataColumn4,
																																ultraDataColumn5,
																																ultraDataColumn6,
																																ultraDataColumn7,
																																ultraDataColumn8,
																																ultraDataColumn9});
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
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1255, 65);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(96, 26);
			this.btnSalir.TabIndex = 808;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(1255, 18);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(96, 28);
			this.btnNuevo.TabIndex = 1020;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(1092, 28);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(94, 27);
			this.btnConsultar.TabIndex = 1019;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// udsResponsable
			// 
			ultraDataColumn10.DataType = typeof(System.UInt32);
			this.udsResponsable.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn10,
																															 ultraDataColumn11});
			// 
			// udsEmpresa
			// 
			ultraDataColumn12.DataType = typeof(int);
			this.udsEmpresa.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn12,
																													 ultraDataColumn13});
			// 
			// udsAsignado
			// 
			ultraDataColumn14.DataType = typeof(int);
			this.udsAsignado.Band.Columns.AddRange(new object[] {
																														ultraDataColumn14,
																														ultraDataColumn15});
			// 
			// udsEstado
			// 
			ultraDataColumn16.DataType = typeof(int);
			this.udsEstado.Band.Columns.AddRange(new object[] {
																													ultraDataColumn16,
																													ultraDataColumn17});
			// 
			// udsRequerimientoTipo
			// 
			ultraDataColumn18.DataType = typeof(int);
			this.udsRequerimientoTipo.Band.Columns.AddRange(new object[] {
																																		 ultraDataColumn18,
																																		 ultraDataColumn19});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-710, 166);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(2839, 9);
			this.groupBox1.TabIndex = 1031;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.lblAsignacionesMias);
			this.groupBox2.Controls.Add(this.chkAsignacionesMias);
			this.groupBox2.Controls.Add(this.label12);
			this.groupBox2.Controls.Add(this.label17);
			this.groupBox2.Controls.Add(this.label18);
			this.groupBox2.Controls.Add(this.txtNumero);
			this.groupBox2.Controls.Add(this.dtHasta);
			this.groupBox2.Controls.Add(this.dtDesde);
			this.groupBox2.Controls.Add(this.btnConsultar);
			this.groupBox2.Location = new System.Drawing.Point(19, 83);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1217, 74);
			this.groupBox2.TabIndex = 1038;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Búsqueda";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// lblAsignacionesMias
			// 
			this.lblAsignacionesMias.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblAsignacionesMias.ForeColor = System.Drawing.Color.Red;
			this.lblAsignacionesMias.Location = new System.Drawing.Point(1085, 120);
			this.lblAsignacionesMias.Name = "lblAsignacionesMias";
			this.lblAsignacionesMias.Size = new System.Drawing.Size(230, 27);
			this.lblAsignacionesMias.TabIndex = 1053;
			// 
			// chkAsignacionesMias
			// 
			this.chkAsignacionesMias.Location = new System.Drawing.Point(970, 120);
			this.chkAsignacionesMias.Name = "chkAsignacionesMias";
			this.chkAsignacionesMias.Size = new System.Drawing.Size(96, 28);
			this.chkAsignacionesMias.TabIndex = 1052;
			this.chkAsignacionesMias.Text = "Asignaciones mías";
			this.chkAsignacionesMias.Visible = false;
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(691, 28);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(47, 18);
			this.label12.TabIndex = 1051;
			this.label12.Text = "Buscar";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(355, 28);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(44, 18);
			this.label17.TabIndex = 1047;
			this.label17.Text = "Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(19, 28);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(47, 18);
			this.label18.TabIndex = 1046;
			this.label18.Text = "Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(758, 28);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(192, 24);
			this.txtNumero.TabIndex = 1050;
			this.toolTip1.SetToolTip(this.txtNumero, "Busqueda por número, responsable, asignado, título.");
			// 
			// dtHasta
			// 
			this.dtHasta.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance2;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(422, 28);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(240, 25);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1045;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			this.dtDesde.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance3;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(134, 28);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(192, 25);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1044;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// cmbResponsable
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance4;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResponsable.DataSource = this.udsResponsable;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(125, 28);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.ReadOnly = true;
			this.cmbResponsable.Size = new System.Drawing.Size(297, 25);
			this.cmbResponsable.TabIndex = 1042;
			this.cmbResponsable.ValueMember = "idPersonal";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(10, 28);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(105, 26);
			this.label1.TabIndex = 1038;
			this.label1.Text = "* Usuario";
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Controls.Add(this.label1);
			this.groupBox3.Controls.Add(this.cmbResponsable);
			this.groupBox3.Location = new System.Drawing.Point(19, 9);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1217, 65);
			this.groupBox3.TabIndex = 1050;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Datos principales";
			// 
			// groupBox4
			// 
			this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox4.Controls.Add(this.udgLineaCredito);
			this.groupBox4.Location = new System.Drawing.Point(19, 185);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1332, 258);
			this.groupBox4.TabIndex = 1051;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "Listado líneas de crédito";
			// 
			// udgLineaCredito
			// 
			this.udgLineaCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgLineaCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgLineaCredito.DataSource = this.udsLineaCredito;
			appearance5.BackColor = System.Drawing.Color.White;
			this.udgLineaCredito.DisplayLayout.Appearance = appearance5;
			this.udgLineaCredito.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Número";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 112;
			ultraGridColumn5.Header.Caption = "Nombre";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 250;
			ultraGridColumn6.Format = "dd/MM/yyyy hh:mm:ss";
			ultraGridColumn6.Header.Caption = "F. aprobación";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 168;
			ultraGridColumn7.Format = "dd/MM/yyyy hh:mm:ss";
			ultraGridColumn7.Header.Caption = "F. renovación";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 178;
			appearance6.ImageHAlign = Infragistics.Win.HAlign.Right;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Monto total ($)";
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn8.Width = 162;
			appearance7.ImageHAlign = Infragistics.Win.HAlign.Right;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance7;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Tasa (%)";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 86;
			appearance8.ImageHAlign = Infragistics.Win.HAlign.Right;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance8;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "P. garantía (%)";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 136;
			ultraGridColumn11.Header.Caption = "Tipo";
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Width = 141;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			ultraGridColumn12.CellAppearance = appearance9;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			ultraGridColumn12.CellButtonAppearance = appearance10;
			ultraGridColumn12.Header.Caption = "...";
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn12.Width = 22;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			ultraGridColumn13.CellAppearance = appearance11;
			appearance12.Image = ((object)(resources.GetObject("appearance12.Image")));
			ultraGridColumn13.CellButtonAppearance = appearance12;
			ultraGridColumn13.Header.Caption = "...";
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn13.Width = 21;
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
																										 ultraGridColumn13});
			this.udgLineaCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.udgLineaCredito.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 10F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgLineaCredito.DisplayLayout.Override.HeaderAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgLineaCredito.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgLineaCredito.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.udgLineaCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.udgLineaCredito.Location = new System.Drawing.Point(19, 37);
			this.udgLineaCredito.Name = "udgLineaCredito";
			this.udgLineaCredito.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.udgLineaCredito.Size = new System.Drawing.Size(1275, 203);
			this.udgLineaCredito.TabIndex = 12;
			this.udgLineaCredito.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgSeguimiento_ClickCellButton);
			this.udgLineaCredito.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgLineaCredito_AfterCellUpdate);
			// 
			// btnPolitica
			// 
			this.btnPolitica.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnPolitica.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnPolitica.CausesValidation = false;
			this.btnPolitica.Image = ((System.Drawing.Image)(resources.GetObject("btnPolitica.Image")));
			this.btnPolitica.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnPolitica.Location = new System.Drawing.Point(1255, 452);
			this.btnPolitica.Name = "btnPolitica";
			this.btnPolitica.Size = new System.Drawing.Size(96, 28);
			this.btnPolitica.TabIndex = 1052;
			this.btnPolitica.Text = "&Políticas";
			this.btnPolitica.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnPolitica.Visible = false;
			// 
			// frmLineaCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(1370, 496);
			this.Controls.Add(this.btnPolitica);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnSalir);
			this.Name = "frmLineaCredito";
			this.Text = "Línea de crédito";
			this.Load += new System.EventHandler(this.frmLineaCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.udsLineaCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAsignado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			this.groupBox3.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udgLineaCredito)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmLineaCredito_Load(object sender, System.EventArgs e)
		{
			try
			{
				string ambiente = MenuLatinium.Ambiente;

				if(ambiente == MenuLatinium.AmbientePro)
				{
					if (!Funcion.Permiso("2034", cdsSeteoF))
					{
						MessageBox.Show("No puede ingresar a línea de crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						BeginInvoke(new MethodInvoker(UnloadMe));
						return;
					}
				}

				//Carga id usuario tabla Inicio
				int idUsuarioInicio = Funcion.iEjecucion(sconexionInicio, string.Format("select top 1 idUsuario from Usuarios where NOMBRE = '{0}' Order By idUsuario", MenuLatinium.stUsuario),false);
				//CargarResponsable
				this.cmbResponsable.DataSource = Funcion.dsEjecucion(sconexionInicio, "Exec RequerimientoResponsable");
				//Traer de la tabla de inicio los usuarios y el logeado
				//Cargar el nombre de la empresa que se logeo
				string sempresa = Funcion.sEjecucion(frmRequerimiento.sconexionInicio, string.Format("Select Top 1 e.Empresa From Empresas e Where e.BaseDeDatos = '{0}'", MenuLatinium.stDirFacturacion.Trim()));

				try
				{
					//Seleccionar el personal que ingreso al sistema que es responsable
					this.cmbResponsable.Value = idUsuarioInicio;
					if(idUsuarioInicio == 0)
					{

						MessageBox.Show(string.Format("Su usuario {0}, no se encuentra registrado como usuario de la empresa {1}", 
							MenuLatinium.stUsuario,
							sempresa
							), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
				catch(Exception ex)
				{
					string sms = ex.Message;
				}

				//Ubicar por default las fechas al día de hoy para no estesar la base
				#region Fechas default
				DateTime dtInicio = DateTime.Today.AddMonths(-540);
				this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				this.dtDesde.Value = Convert.ToDateTime("14/02/2024"); // Fecha de inicio de registros
				this.dtHasta.Value = DateTime.Today;
				#endregion Fechas default
				LineaCreditoCon();
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.ToString(), "VALIDADO");
			}
		}

		private void UnloadMe()
		{
			this.Close();
		}


		private void LineaCreditoCon()
		{

			string dtDesde = "";
			string dtHasta = "";
			string sNumero = "";

			try { dtDesde = Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd");} 
			catch(Exception ex){string e = ex.Message;}
			try { dtHasta = Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd");} 
			catch(Exception ex){string e = ex.Message;}
			try { sNumero = this.txtNumero.Text;} 
			catch(Exception ex){string e = ex.Message;}

			this.udgLineaCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoCon '{0}','{1}','{2}'",
				dtDesde,
				dtHasta,
				sNumero
				));
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbEmpresas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			LineaCreditoCon();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void udgSeguimiento_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int idLineaCredito = 0;

			if (e.Cell.Column.ToString() == "btnVer")	
			{
				try
				{
					idLineaCredito = (int)e.Cell.Row.Cells["idLineaCredito"].Value;

					using (frmLineaCreditoCrear frmS = new frmLineaCreditoCrear(
									 frmLineaCreditoCrear.iVer,
									 idLineaCredito
									 ))
					{		
						if (DialogResult.OK == frmS.ShowDialog())
						{
					 
						}
					}
				}
				catch(Exception ex)
				{
					string sms = ex.Message;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "btnAnadir")	
			{
				try
				{
					idLineaCredito = (int)e.Cell.Row.Cells["idLineaCredito"].Value;

					using (frmLineaCreditoCrear frmS = new frmLineaCreditoCrear(
									 frmLineaCreditoCrear.iAnadir,
									 idLineaCredito
									 ))
					{		
						if (DialogResult.OK == frmS.ShowDialog())
						{
							LineaCreditoCon();
						}
					}
				}
				catch(Exception ex)
				{
					string sms = ex.Message;
					return;
				}
			}

		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			#region Añadir garantía
			using (frmLineaCreditoCrear frlcc = new frmLineaCreditoCrear(frmLineaCredito.iNuevo, 0))
			{		
				if (DialogResult.OK == frlcc.ShowDialog())
				{
					LineaCreditoCon();
				}
			}
			#endregion Añadir garantía
		}

		public static bool LineaCreditoConfgVal(C1.Data.C1DataSet ds, int idLineaCreditoConfiguracion, decimal valor, int itipo)
		{
			DataView dvValidacion = Funcion.dvProcedimiento(ds, String.Format("EXEC LineaCreditoConfgVal {0}, {1}, {2}", 
				idLineaCreditoConfiguracion, 
				valor, 
				itipo));
			DataSet dsValidacion = dvValidacion.Table.DataSet;

			bool bResultado = bool.Parse(dsValidacion.Tables[0].Rows[0]["bResultado"].ToString());
			string sMensaje = dsValidacion.Tables[0].Rows[0]["sMensaje"].ToString();

			if(bResultado == false)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}

			return bResultado;
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void udgLineaCredito_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		
		}

	}
}
