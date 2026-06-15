using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCbo_AsignacionCobradores.
	/// </summary>
	public class frmCbo_AsignacionCobradores : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbZona;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridcobradores;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSupervisor;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private System.Windows.Forms.Button btnComandos;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesCopiarUltimoMes;
		private System.Windows.Forms.MenuItem mnuAccionesAgregarNuevoLocal;
		private System.Windows.Forms.Label lblContador;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCbo_AsignacionCobradores()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCbo_AsignacionCobradores));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsignacionCobradores");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idZona");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSupervisor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RangoDesde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RangoHasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Supervisor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSupervisor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Supervisor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idZona");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Canton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idZona");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Canton");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsignacionCobradores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idZona", -1, "cmbZona");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "ultraDropDown1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor", -1, "cmbSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoDesde");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoHasta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEspecialista", 0);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operador");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.label2 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cmbSupervisor = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbZona = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.uGridcobradores = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnVer = new System.Windows.Forms.Button();
			this.cmbCargo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnComandos = new System.Windows.Forms.Button();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesCopiarUltimoMes = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAgregarNuevoLocal = new System.Windows.Forms.MenuItem();
			this.lblContador = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSupervisor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbZona)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridcobradores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(43, 16);
			this.label2.TabIndex = 951;
			this.label2.Text = "Periodo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance1;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 950;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.btnGenerar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGenerar.Enabled = false;
			this.btnGenerar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGenerar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(1096, 8);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(72, 23);
			this.btnGenerar.TabIndex = 963;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1280, 8);
			this.groupBox2.TabIndex = 964;
			this.groupBox2.TabStop = false;
			// 
			// cmbSupervisor
			// 
			this.cmbSupervisor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSupervisor.DataSource = this.ultraDataSource6;
			this.cmbSupervisor.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 88;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 198;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSupervisor.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSupervisor.DisplayMember = "Supervisor";
			this.cmbSupervisor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSupervisor.Location = new System.Drawing.Point(312, 232);
			this.cmbSupervisor.Name = "cmbSupervisor";
			this.cmbSupervisor.Size = new System.Drawing.Size(200, 112);
			this.cmbSupervisor.TabIndex = 965;
			this.cmbSupervisor.ValueMember = "idSupervisor";
			this.cmbSupervisor.Visible = false;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// ultraDataSource4
			// 
			ultraDataColumn15.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// cmbZona
			// 
			this.cmbZona.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbZona.DataSource = this.ultraDataSource3;
			this.cmbZona.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 63;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 95;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 103;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbZona.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbZona.DisplayMember = "Codigo";
			this.cmbZona.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbZona.Location = new System.Drawing.Point(32, 248);
			this.cmbZona.Name = "cmbZona";
			this.cmbZona.Size = new System.Drawing.Size(200, 112);
			this.cmbZona.TabIndex = 967;
			this.cmbZona.ValueMember = "idZona";
			this.cmbZona.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
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
			// uGridcobradores
			// 
			this.uGridcobradores.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridcobradores.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridcobradores.DataSource = this.ultraDataSource2;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridcobradores.DisplayLayout.Appearance = appearance2;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn7.Header.Caption = "Zona";
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn7.Width = 150;
			ultraGridColumn8.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn8.Header.Caption = "Cobrador";
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn8.Width = 280;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 90;
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Width = 90;
			ultraGridColumn11.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.Header.Caption = "Supervisor";
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn11.Width = 200;
			ultraGridColumn12.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn12.Header.Caption = "Cargo";
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn12.Width = 120;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance3;
			ultraGridColumn13.Header.Caption = "Rango Desde";
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Width = 85;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance4;
			ultraGridColumn14.Header.Caption = "Rango Hasta";
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Width = 85;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Width = 110;
			ultraGridColumn16.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			appearance5.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellButtonAppearance = appearance5;
			ultraGridColumn16.Header.Caption = "...";
			ultraGridColumn16.Header.VisiblePosition = 5;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn16.Width = 30;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn16});
			this.uGridcobradores.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridcobradores.DisplayLayout.GroupByBox.Hidden = true;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridcobradores.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.uGridcobradores.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridcobradores.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridcobradores.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.uGridcobradores.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridcobradores.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.uGridcobradores.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridcobradores.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridcobradores.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridcobradores.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.uGridcobradores.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridcobradores.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridcobradores.Location = new System.Drawing.Point(0, 56);
			this.uGridcobradores.Name = "uGridcobradores";
			this.uGridcobradores.Size = new System.Drawing.Size(1248, 384);
			this.uGridcobradores.TabIndex = 971;
			this.uGridcobradores.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridcobradores_AfterRowInsert);
			this.uGridcobradores.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridcobradores_BeforeRowsDeleted);
			this.uGridcobradores.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridcobradores_ClickCellButton);
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DataSource = this.ultraDataSource5;
			this.ultraDropDown1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 139;
			ultraGridColumn18.Header.Caption = "Supervisor/Cobrador";
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 278;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.ultraDropDown1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultraDropDown1.DisplayMember = "Operador";
			this.ultraDropDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown1.Location = new System.Drawing.Point(584, 168);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(280, 104);
			this.ultraDropDown1.TabIndex = 972;
			this.ultraDropDown1.ValueMember = "idPersonal";
			this.ultraDropDown1.Visible = false;
			// 
			// ultraDataSource5
			// 
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(1016, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 973;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cmbCargo
			// 
			this.cmbCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargo.DataSource = this.ultraDataSource7;
			this.cmbCargo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand5.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 100;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 278;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCargo.DisplayMember = "Descripcion";
			this.cmbCargo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargo.Location = new System.Drawing.Point(292, 171);
			this.cmbCargo.Name = "cmbCargo";
			this.cmbCargo.Size = new System.Drawing.Size(280, 104);
			this.cmbCargo.TabIndex = 974;
			this.cmbCargo.ValueMember = "idCargo";
			this.cmbCargo.Visible = false;
			// 
			// ultraDataSource7
			// 
			ultraDataColumn22.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22,
																																 ultraDataColumn23});
			// 
			// btnComandos
			// 
			this.btnComandos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComandos.Enabled = false;
			this.btnComandos.Image = ((System.Drawing.Image)(resources.GetObject("btnComandos.Image")));
			this.btnComandos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComandos.Location = new System.Drawing.Point(1176, 8);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.TabIndex = 976;
			this.btnComandos.Text = "Acciones";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Click += new System.EventHandler(this.btnComandos_Click);
			// 
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesCopiarUltimoMes,
																																								this.mnuAccionesAgregarNuevoLocal});
			// 
			// mnuAccionesCopiarUltimoMes
			// 
			this.mnuAccionesCopiarUltimoMes.Index = 0;
			this.mnuAccionesCopiarUltimoMes.Text = "Copiar Ultimo Mes";
			this.mnuAccionesCopiarUltimoMes.Click += new System.EventHandler(this.mnuAccionesCopiarUltimoMes_Click);
			// 
			// mnuAccionesAgregarNuevoLocal
			// 
			this.mnuAccionesAgregarNuevoLocal.Index = 1;
			this.mnuAccionesAgregarNuevoLocal.Text = "Exportar";
			this.mnuAccionesAgregarNuevoLocal.Click += new System.EventHandler(this.mnuAccionesAgregarNuevoLocal_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(192, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 980;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmCbo_AsignacionCobradores
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1256, 446);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnComandos);
			this.Controls.Add(this.cmbCargo);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.ultraDropDown1);
			this.Controls.Add(this.uGridcobradores);
			this.Controls.Add(this.cmbZona);
			this.Controls.Add(this.cmbSupervisor);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.dtDesde);
			this.MaximizeBox = false;
			this.Name = "frmCbo_AsignacionCobradores";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Asignacion Cobradores";
			this.Load += new System.EventHandler(this.frmCbo_AsignacionCobradores_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSupervisor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbZona)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridcobradores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		int iMes ;
		int iYear ;
		int iMesA ;
		int iYearA ;
		private Acceso miAcceso;
		
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCbo_AsignacionCobradores_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083601");

			if (!Funcion.Permiso("1189", cdsSeteoF))
			{		
				MessageBox.Show("No tiene Acceso a Asignacion de Cobradores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if(miAcceso.BGenerar)this.btnGenerar.Enabled = true;
			if(miAcceso.Nuevo)this.btnComandos.Enabled = true;

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/03/2019");
			this.dtDesde.CalendarInfo.MaxDate = DateTime.Now.AddDays(+5 );		
			this.dtDesde.Value = DateTime.Today ;

			iMesA = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde.Value).Month));
			iYearA = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde.Value).Year));

			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				
			this.ultraDropDown1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaAsignacionCobradores] 1 ");
			this.cmbSupervisor.DataSource = Funcion.dvProcedimiento(cdsSeteoF,  " EXEC [ConsultaAsignacionCobradores] 2 ");
			this.cmbZona.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaAsignacionCobradores] 3 ");
			this.cmbCargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaAsignacionCobradores] 4 ");

			string sSQLPa = string.Format("Exec ConsultaCbo_AsignacionCobradores '{0}', '{1}'", Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"), Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"));
			this.uGridcobradores.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);

		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			#region Valida Comisiones
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridcobradores.Rows.All)
			{
				#region Fechas

				if (dr.Cells["Desde"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione La Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridcobradores.ActiveRow = dr;
					this.uGridcobradores.ActiveRow.Selected = true;

					return;
				}
				if (dr.Cells["Hasta"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione La Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridcobradores.ActiveRow = dr;
					this.uGridcobradores.ActiveRow.Selected = true;

					return;
				}
				#endregion Fechas
				#region Porcentajes Desde C
				if (Convert.ToDateTime(dr.Cells["Desde"].Value) > Convert.ToDateTime(dr.Cells["Hasta"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridcobradores.ActiveRow = dr;
					this.uGridcobradores.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final
				#region Cargo
				if (dr.Cells["idSupervisor"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione el Supervisor ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridcobradores.ActiveRow = dr;
					this.uGridcobradores.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idSupervisor"].Value == 0)
				{
					MessageBox.Show("Seleccione el Supervisor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridcobradores.ActiveRow = dr;
					this.uGridcobradores.ActiveRow.Selected = true;

					return;
				}
				if (dr.Cells["idCargo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione el Cargo del Cobrador/ Verificador ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridcobradores.ActiveRow = dr;
					this.uGridcobradores.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idCargo"].Value == 0)
				{
					MessageBox.Show("Seleccione el Cargo del Cobrador/ Verificador ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridcobradores.ActiveRow = dr;
					this.uGridcobradores.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo
					
				#region Supersisor
				if ((int) dr.Cells["idCargo"].Value == 11 || (int) dr.Cells["idCargo"].Value == 17 || (int) dr.Cells["idCargo"].Value == 18)
				{
					if (dr.Cells["idPersonal"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione el Personal ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridcobradores.ActiveRow = dr;
						this.uGridcobradores.ActiveRow.Selected = true;

						return;
					}

					if ((int)dr.Cells["idPersonal"].Value == 0)
					{
						MessageBox.Show("Seleccione el Personal", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridcobradores.ActiveRow = dr;
						this.uGridcobradores.ActiveRow.Selected = true;

						return;
					}
				}
				#endregion Supersisor

				#region Zona
				if ((int) dr.Cells["idCargo"].Value == 11 || (int) dr.Cells["idCargo"].Value == 17 || (int) dr.Cells["idCargo"].Value == 18)
				{

					if (dr.Cells["idZona"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione la Zona  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridcobradores.ActiveRow = dr;
						this.uGridcobradores.ActiveRow.Selected = true;

						return;
					}

					if ((int)dr.Cells["idZona"].Value == 0)
					{
						MessageBox.Show("Seleccione la Zona ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridcobradores.ActiveRow = dr;
						this.uGridcobradores.ActiveRow.Selected = true;

						return;
					}
				}
				#endregion Zona

				if (dr.Cells["RangoDesde"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el rango Inicial ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridcobradores.ActiveRow = dr;
					this.uGridcobradores.ActiveRow.Selected = true;

					return;
				}
				if ((int)dr.Cells["RangoDesde"].Value == 0)
				{
					MessageBox.Show("El Rango Inicial no Puede ser cero ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridcobradores.ActiveRow = dr;
					this.uGridcobradores.ActiveRow.Selected = true;

					return;
				}
				if (dr.Cells["RangoHasta"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el rango Final ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridcobradores.ActiveRow = dr;
					this.uGridcobradores.ActiveRow.Selected = true;

					return;
				}
				if ((int) dr.Cells["RangoHasta"].Value == 0)
				{
					MessageBox.Show("el Rango Final no Puede ser cero ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridcobradores.ActiveRow = dr;
					this.uGridcobradores.ActiveRow.Selected = true;

					return;
				}
			}
				
			#endregion Valida Comisiones

			#region valida zonas con barrios
				int iContB = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dra in this.uGridcobradores.Rows.All)
			{
				int iBarrios = 0;
				iBarrios = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(1) from dbo.Cre_Barrio where idZona = {0} ", (int)dra.Cells["idZona"].Value ));
				if( iBarrios == 0 )
				{
					iContB = iContB + 1;
					this.uGridcobradores.ActiveRow = dra;
					this.uGridcobradores.ActiveRow.Selected = true;
					//return;
				}
			}

			if( iContB > 0 )
			{
			  MessageBox.Show("Las Zonas seleccionadas no tienen barrios asignados.\n\n Configure en el modulo CrediPoint-> Comisiones -> Asignacioón de barrios.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			  return;
			}
			#endregion valida zonas con barrios

			#region Guarda

			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Guardar los cambios", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{

				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size=75000;"))
				{
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

		
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridcobradores.Rows.All)
						{
					
							string sSQLDetalle = string.Format("Exec GuardaCbo_AsignaCobradores {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}",
								(int)dr.Cells["idAsignacionCobradores"].Value, 
								(int)dr.Cells["idZona"].Value, 
								(int)dr.Cells["idPersonal"].Value, 
								Convert.ToDateTime(dr.Cells["Desde"].Value).ToString("yyyyMMdd"),
								Convert.ToDateTime(dr.Cells["Hasta"].Value).ToString("yyyyMMdd"),
								(int)dr.Cells["idSupervisor"].Value,
								(int)dr.Cells["idCargo"].Value,
								(int)dr.Cells["RangoDesde"].Value,
								(int)dr.Cells["RangoHasta"].Value);

		
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}
			      
				
						oTransaction.Commit();
			
					}
					catch (Exception ex)
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
						return;
					}
					finally
					{
						oConexion.Close();
					}
				}
			}
			#endregion Guarda

			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				
			string sSQLCbo = string.Format("Exec ConsultaCbo_AsignacionCobradores '{0}', '{1}'", Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"), Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"));
			this.uGridcobradores.DataSource = FuncionesProcedimientos.dtGeneral(sSQLCbo);

			int iValidaCo = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Exec ValidaAsignacionCobradores '{0}'", Convert.ToDateTime(dtInicio).ToString("yyyyMMdd")));	

			if (iValidaCo > 0)
			{

				if (DialogResult.Yes == MessageBox.Show("¿Desea Actualizar la Zonas de los clientes ?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					string ssSQLAc = string.Format("Exec ActualizaAsignacionCobradores '{0}' ", Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"));
					Funcion.EjecutaSQL(cdsSeteoF, ssSQLAc);
				}
				if (DialogResult.Yes == MessageBox.Show("¿Desea Actualizar los cobradores en la asignación de clientes ?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					string ssSQLAcP = string.Format("Exec [ActualizaAsignacionCobradoresCuotaCero] '{0}' ", Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"));
					Funcion.EjecutaSQL(cdsSeteoF, ssSQLAcP);

					string ssSQLAcPm = string.Format("Exec [ActualizaAsignacionCobradoresCuotamAYOR] '{0}' ", Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"));
					Funcion.EjecutaSQL(cdsSeteoF, ssSQLAcPm);
				}
			}
			MessageBox.Show("Registro  guardado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void uGridcobradores_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idAsignacionCobradores"].Value = 0;
			e.Row.Cells["idPersonal"].Value = 0;
			e.Row.Cells["idSupervisor"].Value = 0;
			e.Row.Cells["idZona"].Value = 0;
			e.Row.Cells["idCargo"].Value = 0;
				e.Row.Cells["RangoDesde"].Value = 0;
			e.Row.Cells["RangoHasta"].Value = 0;
	
			

		}

		private void uGridcobradores_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idAsignacionCobradores"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idAsignacionCobradores"].Value > 0)
						{
							string sSQL = string.Format("Delete From AsignacionCobradores Where idAsignacionCobradores = {0}", (int)e.Rows[i].Cells["idAsignacionCobradores"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}
					}			
				}
			}
			e.DisplayPromptMsg = false;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione El Periodo ", false, "", cdsSeteoF)) return;
			iMesA = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde.Value).Month));
			iYearA = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde.Value).Year));

			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				this.lblContador.Text = "";
			string sSQLPa = string.Format("Exec ConsultaCbo_AsignacionCobradores '{0}', '{1}'", Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"), Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"));
			this.uGridcobradores.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);

				this.lblContador.Text = this.uGridcobradores.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
					FuncionesProcedimientos.ExportaExcel(this.uGridcobradores);
		}

		private void mnuAccionesCopiarUltimoMes_Click(object sender, System.EventArgs e)
		{
			int iRetorna = 0;
      if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione El Periodo que desea Actualizar  ", false, "", cdsSeteoF)) return;

			string sPeriodo = Funcion.sEscalarSQL(cdsSeteoF, "Select Top 1  convert (varchar ( 20), DatePart ( year, Hasta))+'/'+ convert (varchar ( 20), DatePart ( MONTH, Hasta)) From AsignacionCobradores Order By Hasta Desc");

			if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Copiar La Información del anterior Mes?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{	

				iRetorna = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CopiaAsignacioncobradores '{0}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd")));
//				this.uGridcobradores.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
				if ( iRetorna == 1)
				{

					MessageBox.Show(string.Format("Se copio la Información del Periodo {0} correctamente", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.btnVer_Click(sender, e);
				}
				else
				{
					MessageBox.Show(string.Format("No existen Datos en el Mes anterior {0} o ya se genero una copia en el mes actual ", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.btnVer_Click(sender, e);

				}
			}
		}

		private void btnComandos_Click(object sender, System.EventArgs e)
		{
			this.mnuAcciones.Show(this.btnComandos, new Point(1, 25));
		}

		private void mnuAccionesAgregarNuevoLocal_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.uGridcobradores);
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
		this.btnVer_Click(sender, e);
		}

		private void uGridcobradores_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Operadora

			if (e.Cell.Column.ToString() == "btnEspecialista")
			{
				if (e.Cell.Row.Cells["idCargo"].Value != System.DBNull.Value || e.Cell.Row.Cells["idPersonal"].Value != System.DBNull.Value  )
				{
					if ( (int)e.Cell.Row.Cells["idPersonal"].Value == 0)
					{
						MessageBox.Show(string.Format("Seleccione el Especialista "), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
          
					}
					if ((int)e.Cell.Row.Cells["idCargo"].Value == 0 )
					{
						MessageBox.Show(string.Format("Seleccione la Zona del Especialista  "), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
          
					}
					if ((int)e.Cell.Row.Cells["idCargo"].Value == 13 || (int)e.Cell.Row.Cells["idCargo"].Value == 14  || (int)e.Cell.Row.Cells["idCargo"].Value == 15|| (int)e.Cell.Row.Cells["idCargo"].Value == 16   )
					{

						#region llama
					
					
							using (frmAsignacionEspecialista Gestion = new frmAsignacionEspecialista((int)e.Cell.Row.Cells["idPersonal"].Value, (DateTime )this.dtDesde.Value ))
							{
								if (DialogResult.OK == Gestion.ShowDialog( ))
								{

	
								}
							}
					
							
					
						#endregion llama
						
          
					}
					else
					{
						MessageBox.Show(string.Format("La Asignación es para Especialistas "), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}	
			}
			#endregion Operadora
		}
	}
}
