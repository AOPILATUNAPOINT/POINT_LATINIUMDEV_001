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
	/// Descripción breve de frmCbo_AsignaPorcDias.
	/// </summary>
	public class frmCbo_AsignaPorcDias : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnGenerar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbOperador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtPeriodo;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSupervisor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTramo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Button btnComandos;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesCopiarUltimoMes;
		private System.Windows.Forms.MenuItem mnuAccionesAgregarNuevoLocal;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.MenuItem mnuAccionesActualizar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCbo_AsignaPorcDias()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_PorcentajeDiasCom");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSupervisor");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCbo_AsignaPorcDias));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operador");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Supervisor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSupervisor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Supervisor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_PorcentajeDiasCom");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSupervisor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tramo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_PorcentajeDiasCom");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbOperador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor", -1, "cmbSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tramo", -1, "cmbTramo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tramo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tramo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtPeriodo = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.cmbOperador = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbSupervisor = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnVer = new System.Windows.Forms.Button();
			this.cmbTramo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnComandos = new System.Windows.Forms.Button();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesCopiarUltimoMes = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAgregarNuevoLocal = new System.Windows.Forms.MenuItem();
			this.mnuAccionesActualizar = new System.Windows.Forms.MenuItem();
			this.lblContador = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtPeriodo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOperador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSupervisor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTramo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// dtPeriodo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtPeriodo.Appearance = appearance1;
			this.dtPeriodo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtPeriodo.DateButtons.Add(dateButton1);
			this.dtPeriodo.Format = "MM/yyyy";
			this.dtPeriodo.Location = new System.Drawing.Point(64, 9);
			this.dtPeriodo.Name = "dtPeriodo";
			this.dtPeriodo.NonAutoSizeHeight = 23;
			this.dtPeriodo.Size = new System.Drawing.Size(112, 21);
			this.dtPeriodo.SpinButtonsVisible = true;
			this.dtPeriodo.TabIndex = 947;
			this.dtPeriodo.Value = ((object)(resources.GetObject("dtPeriodo.Value")));
			this.dtPeriodo.ValueChanged += new System.EventHandler(this.dtPeriodo_ValueChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(43, 16);
			this.label2.TabIndex = 949;
			this.label2.Text = "Periodo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGenerar
			// 
			this.btnGenerar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGenerar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGenerar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(664, 8);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(72, 23);
			this.btnGenerar.TabIndex = 950;
			this.btnGenerar.Text = "Guardar";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// cmbOperador
			// 
			this.cmbOperador.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbOperador.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 300;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbOperador.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbOperador.DisplayMember = "Operador";
			this.cmbOperador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbOperador.Location = new System.Drawing.Point(120, 136);
			this.cmbOperador.Name = "cmbOperador";
			this.cmbOperador.Size = new System.Drawing.Size(278, 86);
			this.cmbOperador.TabIndex = 951;
			this.cmbOperador.ValueMember = "idPersonal";
			this.cmbOperador.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(944, 8);
			this.groupBox2.TabIndex = 952;
			this.groupBox2.TabStop = false;
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
			// cmbSupervisor
			// 
			this.cmbSupervisor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSupervisor.DataSource = this.ultraDataSource3;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 200;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 298;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbSupervisor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSupervisor.DisplayMember = "Supervisor";
			this.cmbSupervisor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSupervisor.Location = new System.Drawing.Point(544, 136);
			this.cmbSupervisor.Name = "cmbSupervisor";
			this.cmbSupervisor.Size = new System.Drawing.Size(278, 86);
			this.cmbSupervisor.TabIndex = 953;
			this.cmbSupervisor.ValueMember = "idSupervisor";
			this.cmbSupervisor.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn10.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// ultraDataSource4
			// 
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource4;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance2;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn6.Header.Caption = "Operadora";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn6.Width = 280;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Header.VisiblePosition = 3;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn10.Header.Caption = "Supervisor";
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 280;
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn11.Width = 100;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			ultraGridBand3.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand3.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Hidden = true;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridDocumentos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDocumentos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(8, 56);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(816, 256);
			this.uGridDocumentos.TabIndex = 961;
			this.uGridDocumentos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDocumentos_AfterRowInsert_1);
			this.uGridDocumentos.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDocumentos_BeforeRowsDeleted);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(584, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 974;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cmbTramo
			// 
			this.cmbTramo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTramo.DataSource = this.ultraDataSource5;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbTramo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTramo.DisplayMember = "Descripcion";
			this.cmbTramo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTramo.Location = new System.Drawing.Point(408, 136);
			this.cmbTramo.Name = "cmbTramo";
			this.cmbTramo.Size = new System.Drawing.Size(126, 86);
			this.cmbTramo.TabIndex = 975;
			this.cmbTramo.ValueMember = "Tramo";
			this.cmbTramo.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// btnComandos
			// 
			this.btnComandos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComandos.Enabled = false;
			this.btnComandos.Image = ((System.Drawing.Image)(resources.GetObject("btnComandos.Image")));
			this.btnComandos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComandos.Location = new System.Drawing.Point(744, 8);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.TabIndex = 977;
			this.btnComandos.Text = "Acciones";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Click += new System.EventHandler(this.btnComandos_Click);
			// 
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesCopiarUltimoMes,
																																								this.mnuAccionesAgregarNuevoLocal,
																																								this.mnuAccionesActualizar});
			this.mnuAcciones.Popup += new System.EventHandler(this.mnuAcciones_Popup);
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
			// mnuAccionesActualizar
			// 
			this.mnuAccionesActualizar.Index = 2;
			this.mnuAccionesActualizar.Text = "Actualizar Tramas";
			this.mnuAccionesActualizar.Click += new System.EventHandler(this.mnuAccionesActualizar_Click);
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
			// frmCbo_AsignaPorcDias
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(832, 318);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnComandos);
			this.Controls.Add(this.cmbTramo);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.uGridDocumentos);
			this.Controls.Add(this.cmbSupervisor);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.cmbOperador);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.dtPeriodo);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCbo_AsignaPorcDias";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Asignación Supervisores / Operadoras";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCbo_AsignaPorcDias_KeyDown);
			this.Load += new System.EventHandler(this.frmCbo_AsignaPorcDias_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtPeriodo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOperador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSupervisor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTramo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCbo_AsignaPorcDias_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083602");

			if (!Funcion.Permiso("1190", cdsSeteoF))
			{		
				MessageBox.Show("No tiene Acceso a Asignacion por dias", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if(miAcceso.Nuevo)this.btnComandos.Enabled = true;
			
			this.dtPeriodo.CalendarInfo.MinDate = DateTime.Parse("01/03/2019");
			this.dtPeriodo.CalendarInfo.MaxDate = DateTime.Now.AddDays(+5 );		

			this.dtPeriodo.Value = DateTime.Today;



			int iMes = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Month));
			int iYear = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Year));

	
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(( Convert.ToDateTime(this.dtPeriodo.Value).Year), ( Convert.ToDateTime(this.dtPeriodo.Value).Month)).ToString() + "/" + iMes + "/" + iYear);
			
//			this.cmbOperador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select n.idPersonal, ( ApellidoPaterno + ' '+ ApellidoMaterno +' '+ PrimerNombre + ' '+ SegundoNombre) as Operador from Nomina n Inner Join Usuario u on n.Codigo =  u.Nombre where idGrupo = 19 ");
//			this.cmbOperador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select n.idPersonal,( ApellidoPaterno + ' '+ ApellidoMaterno +' '+ PrimerNombre + ' '+ SegundoNombre) as Operador from Nomina n Inner Join RolCargo c on c.idCargo =  n.idCargo where n.idCargo in (91,120,119)  order by Operador ");
//      this.cmbSupervisor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select n.idPersonal as idSupervisor,( ApellidoPaterno + ' '+ ApellidoMaterno +' '+ PrimerNombre + ' '+ SegundoNombre) as Supervisor from Nomina n Inner Join RolCargo c on c.idCargo =  n.idCargo where n.idCargo in (91,120,119) order by Supervisor ");
//      
			this.cmbOperador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaAsignacionCobradores] 5 ");
			this.cmbSupervisor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaAsignacionCobradores] 6 ");

			this.cmbTramo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select Tramo, Descripcion from TramosOperadoras ");

			string sSQLPa = string.Format("Exec ConsultaCbo_AsignaPorcDias '{0}'", Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"));
			this.uGridDocumentos.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
			

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			int iValida = 0;
			int icont = 0;
			int iRango = 0;

			#region Valida Comisiones
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				#region Porcentajes Desde C
				if ((int)dr.Cells["Desde"].Value == 0)
				{
					MessageBox.Show("El Valor Inicial no puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.ActiveRow = dr;
					this.uGridDocumentos.ActiveRow.Selected = true;

					return;
				}
				if ((int)dr.Cells["Hasta"].Value == 0)
				{
					MessageBox.Show("El Valor Final no puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.ActiveRow = dr;
					this.uGridDocumentos.ActiveRow.Selected = true;

					return;
				}
				if (Convert.ToDecimal(dr.Cells["Desde"].Value) > Convert.ToDecimal(dr.Cells["Hasta"].Value))
				{
					MessageBox.Show("En Dias, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.ActiveRow = dr;
					this.uGridDocumentos.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final
					
				#region Cargo
				if ((int)dr.Cells["idPersonal"].Value == 0)
				{
					MessageBox.Show("Seleccione una Operadora", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.ActiveRow = dr;
					this.uGridDocumentos.ActiveRow.Selected = true;

					return;
				}
				if (dr.Cells["idPersonal"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione una Operadora", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.ActiveRow = dr;
					this.uGridDocumentos.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idSupervisor"].Value == 0)
				{
					MessageBox.Show("Seleccione un Supersivor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.ActiveRow = dr;
					this.uGridDocumentos.ActiveRow.Selected = true;

					return;
				}

				if (dr.Cells["idSupervisor"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Supersivor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.ActiveRow = dr;
					this.uGridDocumentos.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo

				if ((int)dr.Cells["Tramo"].Value == 0)
				{
					MessageBox.Show("Seleccione el Tramo" , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.ActiveRow = dr;
					this.uGridDocumentos.ActiveRow.Selected = true;

					return;
				}
				if (dr.Cells["Tramo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Tramo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.ActiveRow = dr;
					this.uGridDocumentos.ActiveRow.Selected = true;

					return;
				}
			}
			#region Valida Operador 
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drO in this.uGridDocumentos.Rows.All)
			{
				iValida =  (int)	drO.Cells["idPersonal"].Value;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drOp in this.uGridDocumentos.Rows.All)
				{
					if (iValida == (int)drOp.Cells["idPersonal"].Value)
					{
						this.uGridDocumentos.ActiveRow = drOp;
						this.uGridDocumentos.ActiveRow.Selected = true;
						icont = icont + 1 ;
					}
				}
				if ( icont > 1)
				{
					MessageBox.Show("Una Operadora esta asignada mas de dos veces", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				icont = icont = 0;
			}
			#endregion Valida Operador 




			#endregion Valida Comisiones

			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drD in this.uGridDocumentos.Rows.All)
			{
				iRango =  (int)	drD.Cells["idPersonal"].Value;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drH in this.uGridDocumentos.Rows.All)
				{
					if (iValida == (int)drH.Cells["idPersonal"].Value)
					{
						this.uGridDocumentos.ActiveRow = drH;
						this.uGridDocumentos.ActiveRow.Selected = true;
						icont = icont + 1 ;
					}
				}
				if ( icont > 1)
				{
					MessageBox.Show("Una Operadora esta asignada mas de dos veces", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				icont = 0;
			}

			if (this.dtPeriodo.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Del Periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtPeriodo.Focus();
				return;
			}
			int iMes = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Month));
			int iYear = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Year));

	
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(( Convert.ToDateTime(this.dtPeriodo.Value).Year), ( Convert.ToDateTime(this.dtPeriodo.Value).Month)).ToString() + "/" + iMes + "/" + iYear);
			  
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

						#region Detalle
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
						{
					
							string sSQLDetalle = string.Format("Exec GuardaCbo_AsignaPorcDias {0}, {1}, '{2}', {3}, {4}, {5}, {6}, {7}",
								(int)dr.Cells["idCbo_PorcentajeDiasCom"].Value, 
								(int)dr.Cells["idPersonal"].Value, 
								Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"),
								(int)dr.Cells["Desde"].Value, 
								(int)dr.Cells["Hasta"].Value, 1,
								(int)dr.Cells["idSupervisor"].Value,
								(int)dr.Cells["Tramo"].Value);
		
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Detalle	      
				
						oTransaction.Commit();
						string sSQLPa = string.Format("Exec ConsultaCbo_AsignaPorcDias '{0}'", Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"));
						this.uGridDocumentos.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
					}
					catch (Exception ex)
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
					}
					finally
					{
						oConexion.Close();
					}
				}
			}
			this.Close ();
		}

		private void uGridDocumentos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCbo_PorcentajeDiasCom"].Value = 0;
			e.Row.Cells["idPersonal"].Value = 0;
			e.Row.Cells["Hasta"].Value = 0;
			e.Row.Cells["Desde"].Value = 0;
			e.Row.Cells["Origen"].Value = 0;
			e.Row.Cells["idSupervisor"].Value = 0;
			

		}

		private void dtPeriodo_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtPeriodo.Value != System.DBNull.Value)
			{
				int iMes = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Month));
				int iYear = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Year));

	
				DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(( Convert.ToDateTime(this.dtPeriodo.Value).Year), ( Convert.ToDateTime(this.dtPeriodo.Value).Month)).ToString() + "/" + iMes + "/" + iYear);
			  
				string sSQLPa = string.Format("Exec ConsultaCbo_AsignaPorcDias '{0}'", Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"));
				this.uGridDocumentos.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);

			}
		}

		private void uGridDocumentos_AfterRowInsert_1(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCbo_PorcentajeDiasCom"].Value = 0;
			e.Row.Cells["idPersonal"].Value = 0;
			e.Row.Cells["Desde"].Value = 0;
			e.Row.Cells["Hasta"].Value = 0;
			e.Row.Cells["Origen"].Value = 0;
				e.Row.Cells["idSupervisor"].Value = 0;
			e.Row.Cells["Tramo"].Value = 0;

		}

		private void uGridDocumentos_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idCbo_PorcentajeDiasCom"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idCbo_PorcentajeDiasCom"].Value > 0)
						{
							string sSQL = string.Format("Delete From Cbo_AsignaPorcDias Where idCbo_PorcentajeDiasCom = {0}", (int)e.Rows[i].Cells["idCbo_PorcentajeDiasCom"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}
					}			
				}
			}
			e.DisplayPromptMsg = false;
		}

		private void frmCbo_AsignaPorcDias_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			
			if (!Validacion.vbFechaEnDocumentos(this.dtPeriodo, "Seleccione El Periodo ", false, "", cdsSeteoF)) return;
			int iMes = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Month));
			int iYear = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Year));

	 
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(( Convert.ToDateTime(this.dtPeriodo.Value).Year), ( Convert.ToDateTime(this.dtPeriodo.Value).Month)).ToString() + "/" + iMes + "/" + iYear);
			 this.lblContador.Text = "";
			string sSQLPa = string.Format("Exec ConsultaCbo_AsignaPorcDias '{0}'", Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"));
			this.uGridDocumentos.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
						this.lblContador.Text = this.uGridDocumentos.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
		
		}

		private void mnuAccionesAgregarNuevoLocal_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.uGridDocumentos);
		}

		private void mnuAccionesCopiarUltimoMes_Click(object sender, System.EventArgs e)
		{
			int iRetorna = 0;
			if (!Validacion.vbFechaEnDocumentos(this.dtPeriodo, "Seleccione El Periodo que desea Actualizar  ", false, "", cdsSeteoF)) return;

			string sPeriodo = Funcion.sEscalarSQL(cdsSeteoF, "Select Top 1  convert (varchar ( 20), DatePart ( year, Periodo))+'/'+ convert (varchar ( 20), DatePart ( MONTH, Periodo)) From Cbo_AsignaPorcDias Order By Periodo  Desc");

			if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Copiar La Información del anterior Mes?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{	

				iRetorna = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CopiaAsignacionOperadoras '{0}'", Convert.ToDateTime(this.dtPeriodo.Value).ToString("yyyyMMdd")));
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

		private void mnuAccionesActualizar_Click(object sender, System.EventArgs e)
		{
			int iMes = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Month));
			int iYear = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Year));

	
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(( Convert.ToDateTime(this.dtPeriodo.Value).Year), ( Convert.ToDateTime(this.dtPeriodo.Value).Month)).ToString() + "/" + iMes + "/" + iYear);
			  

			int iValidaCo = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Exec ValidaAsignacionOperadoras '{0}'", Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")));	

			if (iValidaCo > 0)
			{
				if (DialogResult.Yes == MessageBox.Show("Esta seguro de Actualizar las tramas de los clientes ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
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

							#region Detalle

					      string sSQLDetalle = string.Format("Exec [AsignacionTramCuotacero] '{0}'", Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"));
								oCmdActualiza.CommandText = sSQLDetalle;
								oCmdActualiza.ExecuteNonQuery();

								string sSQLDetallec = string.Format("Exec [AsignacionTramCuotaMayorCero] '{0}'", Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"));
								oCmdActualiza.CommandText = sSQLDetallec;
								oCmdActualiza.ExecuteNonQuery();

							#endregion Detalle
	      
							oTransaction.Commit();
							MessageBox.Show("Tramas Actualizadas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
						}
						catch (Exception ex)
						{
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
						}
						finally
						{
							oConexion.Close();
						}
					}
				}
        
			}
			else
			{
				MessageBox.Show("La cartera de clientes no esta subida para el periodo seleccionado\n\n Comuniquese con el Administrador Ext 800.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
		}

		private void mnuAcciones_Popup(object sender, System.EventArgs e)
		{
		
		}
	}
}
