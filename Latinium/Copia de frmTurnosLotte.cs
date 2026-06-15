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
	/// Descripción breve de frmTurnosLotte.
	/// </summary>
	public class frmTurnosLotte : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ugTurnos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button bntGrabar;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Button button1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTurno;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFeIni;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFeFin;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargo;
		private System.Windows.Forms.Button btnExporta;
		private System.Windows.Forms.Button btnImprime;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.ComponentModel.IContainer components;

		public frmTurnosLotte()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTurno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Turno", -1, "cmbTurno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaInicia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaFinaliza");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTurno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Turno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaInicia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaFinaliza");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTurnosLotte));
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTurno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorTurno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreTurno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTurno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorTurno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreTurno");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ugTurnos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.bntGrabar = new System.Windows.Forms.Button();
			this.label10 = new System.Windows.Forms.Label();
			this.dtFeIni = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.button1 = new System.Windows.Forms.Button();
			this.btnExporta = new System.Windows.Forms.Button();
			this.btnImprime = new System.Windows.Forms.Button();
			this.cmbTurno = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCargo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFeFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			((System.ComponentModel.ISupportInitialize)(this.ugTurnos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeIni)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTurno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			this.SuspendLayout();
			// 
			// ugTurnos
			// 
			this.ugTurnos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugTurnos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugTurnos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugTurnos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 110;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 300;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 30;
			ultraGridColumn6.Header.Caption = "Cargo";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 200;
			ultraGridColumn7.Header.Caption = "Inicia";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.Header.Caption = "Finaliza";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ugTurnos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugTurnos.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ugTurnos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugTurnos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugTurnos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ugTurnos.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.ugTurnos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugTurnos.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTurnos.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTurnos.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTurnos.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ugTurnos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugTurnos.Location = new System.Drawing.Point(8, 8);
			this.ugTurnos.Name = "ugTurnos";
			this.ugTurnos.Size = new System.Drawing.Size(864, 424);
			this.ugTurnos.TabIndex = 29;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// bntGrabar
			// 
			this.bntGrabar.Image = ((System.Drawing.Image)(resources.GetObject("bntGrabar.Image")));
			this.bntGrabar.Location = new System.Drawing.Point(240, 440);
			this.bntGrabar.Name = "bntGrabar";
			this.bntGrabar.Size = new System.Drawing.Size(75, 40);
			this.bntGrabar.TabIndex = 36;
			this.toolTip1.SetToolTip(this.bntGrabar, "Busca información...!!!");
			this.bntGrabar.Click += new System.EventHandler(this.bntGrabar_Click);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(8, 464);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(66, 16);
			this.label10.TabIndex = 232;
			this.label10.Text = "Fecha Inicia";
			// 
			// dtFeIni
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFeIni.Appearance = appearance8;
			this.dtFeIni.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFeIni.DateButtons.Add(dateButton1);
			this.dtFeIni.Format = "dd/MM/yyyy";
			this.dtFeIni.Location = new System.Drawing.Point(8, 440);
			this.dtFeIni.Name = "dtFeIni";
			this.dtFeIni.NonAutoSizeHeight = 23;
			this.dtFeIni.Size = new System.Drawing.Size(104, 21);
			this.dtFeIni.SpinButtonsVisible = true;
			this.dtFeIni.TabIndex = 231;
			this.dtFeIni.Value = ((object)(resources.GetObject("dtFeIni.Value")));
			// 
			// button1
			// 
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.Location = new System.Drawing.Point(320, 440);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(75, 40);
			this.button1.TabIndex = 233;
			this.toolTip1.SetToolTip(this.button1, "Guardar información");
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// btnExporta
			// 
			this.btnExporta.Image = ((System.Drawing.Image)(resources.GetObject("btnExporta.Image")));
			this.btnExporta.Location = new System.Drawing.Point(400, 440);
			this.btnExporta.Name = "btnExporta";
			this.btnExporta.Size = new System.Drawing.Size(75, 40);
			this.btnExporta.TabIndex = 234;
			this.toolTip1.SetToolTip(this.btnExporta, "Exportar a EXCEL");
			this.btnExporta.Click += new System.EventHandler(this.btnExporta_Click);
			// 
			// btnImprime
			// 
			this.btnImprime.Image = ((System.Drawing.Image)(resources.GetObject("btnImprime.Image")));
			this.btnImprime.Location = new System.Drawing.Point(480, 440);
			this.btnImprime.Name = "btnImprime";
			this.btnImprime.Size = new System.Drawing.Size(75, 40);
			this.btnImprime.TabIndex = 235;
			this.toolTip1.SetToolTip(this.btnImprime, "Imprimir información...!!!");
			this.btnImprime.Click += new System.EventHandler(this.btnImprime_Click);
			// 
			// cmbTurno
			// 
			this.cmbTurno.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTurno.DataSource = this.ultraDataSource2;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.Caption = "#T";
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 35;
			ultraGridColumn11.Header.Caption = "Nombre";
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Width = 150;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbTurno.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTurno.DisplayMember = "ValorTurno";
			this.cmbTurno.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTurno.Location = new System.Drawing.Point(152, 216);
			this.cmbTurno.Name = "cmbTurno";
			this.cmbTurno.Size = new System.Drawing.Size(200, 86);
			this.cmbTurno.TabIndex = 236;
			this.cmbTurno.ValueMember = "idTurno";
			this.cmbTurno.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			this.ultraDataSource2.Band.Key = "CompraNumero";
			// 
			// cmbCargo
			// 
			this.cmbCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargo.DataSource = this.ultraDataSource3;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 300;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCargo.DisplayMember = "Descripcion";
			this.cmbCargo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargo.Location = new System.Drawing.Point(384, 216);
			this.cmbCargo.Name = "cmbCargo";
			this.cmbCargo.Size = new System.Drawing.Size(200, 86);
			this.cmbCargo.TabIndex = 237;
			this.cmbCargo.ValueMember = "idCargo";
			this.cmbCargo.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			this.ultraDataSource3.Band.Key = "CompraNumero";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(120, 464);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(79, 16);
			this.label1.TabIndex = 239;
			this.label1.Text = "Fecha Finaliza";
			// 
			// dtFeFin
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFeFin.Appearance = appearance9;
			this.dtFeFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFeFin.DateButtons.Add(dateButton2);
			this.dtFeFin.Format = "dd/MM/yyyy";
			this.dtFeFin.Location = new System.Drawing.Point(120, 440);
			this.dtFeFin.Name = "dtFeFin";
			this.dtFeFin.NonAutoSizeHeight = 23;
			this.dtFeFin.Size = new System.Drawing.Size(104, 21);
			this.dtFeFin.SpinButtonsVisible = true;
			this.dtFeFin.TabIndex = 238;
			this.dtFeFin.Value = ((object)(resources.GetObject("dtFeFin.Value")));
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// frmTurnosLotte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(880, 493);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFeFin);
			this.Controls.Add(this.cmbCargo);
			this.Controls.Add(this.cmbTurno);
			this.Controls.Add(this.btnImprime);
			this.Controls.Add(this.btnExporta);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.dtFeIni);
			this.Controls.Add(this.bntGrabar);
			this.Controls.Add(this.ugTurnos);
			this.Name = "frmTurnosLotte";
			this.Text = "Genera turnos personal";
			this.Load += new System.EventHandler(this.frmTurnosLotte_Load);
			((System.ComponentModel.ISupportInitialize)(this.ugTurnos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeIni)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTurno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmTurnosLotte_Load(object sender, System.EventArgs e)
		{
			this.dtFeIni.Value = DateTime.Now;
			this.dtFeFin.Value = DateTime.Now;
			string stTurno = string.Format("Select idturno,valorturno,nombreturno from turnos");
			this.cmbTurno.DataSource = Funcion.dvProcedimiento(cdsSeteo, stTurno);
			string stCargo = string.Format("Select idCargo,Descripcion From RolCargo");
			this.cmbCargo.DataSource = Funcion.dvProcedimiento(cdsSeteo, stCargo);
		}

		private void bntGrabar_Click(object sender, System.EventArgs e)
		{
			/*Buscar información*/
			DateTime FIni = (DateTime)this.dtFeIni.Value;
			DateTime FFin = (DateTime)this.dtFeFin.Value;

			string sSQL = string.Format("Exec Visuliza_Personal_Turno '{0}', '{1}'", 
				FIni.ToString("yyyyMMdd"),
				FFin.ToString("yyyyMMdd"));
			this.ugTurnos.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);	
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			/*Grbar información*/
			//SP Graba_Personal_Turno
			if(DialogResult.Yes == MessageBox.Show("¿Procesar INFORMACION..?", "Aviso...!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
				{
					oConexion.Open();

					SqlTransaction oTransaction = null;
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					oCmdActualiza.Connection = oConexion;
					oCmdActualiza.CommandType = CommandType.Text;
					oCmdActualiza.CommandTimeout = 0;
						
					try
					{							
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;

						#region Detalle del Grid
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugTurnos.Rows.All)
						{
							if((int)dr.Cells["Turno"].Value>0)
							{
								#region Fechas
								DateTime FIni = (DateTime)dr.Cells["FechaInicia"].Value;
								DateTime FFin = (DateTime)dr.Cells["FechaFinaliza"].Value;
								#endregion Fechas

								string sSQLDetalle = string.Format("Exec Graba_Personal_Turno {0}, {1}, {2}, {3}, '{4}', '{5}'",
									(int)dr.Cells["idTurno"].Value, (int)dr.Cells["idPersonal"].Value, 
									(int)dr.Cells["Turno"].Value, (int)dr.Cells["idCargo"].Value, 
									FIni.ToString("yyyyMMdd"), FFin.ToString("yyyyMMdd"));
								oCmdActualiza.CommandText = sSQLDetalle;
								oCmdActualiza.ExecuteNonQuery();
							}
						}
						#endregion Detalle del Grid
						
						oTransaction.Commit();
						MessageBox.Show("¿Proceso terminado..!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.button1.Enabled = false;

					}
					catch (Exception ex)
					{
						try
						{
							oTransaction.Rollback();
							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						catch (Exception ex2)
						{
							MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
						}
					}
					finally
					{
						oConexion.Close();
					}
				}
			}
		}

		private void btnExporta_Click(object sender, System.EventArgs e)
		{
			if (this.ugTurnos.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.ugTurnos, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void btnImprime_Click(object sender, System.EventArgs e)
		{
			//Visualizar información
			Imprimir();
		}

		private void Imprimir()
		{
			DateTime FBsk = (DateTime)this.dtFeIni.Value;

			#region Imprime el reporte
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldDesde = new ParameterField ();
			ParameterDiscreteValue discreteValDesde = new ParameterDiscreteValue ();
			paramFieldDesde.ParameterFieldName = "@FeBusk";
			discreteValDesde.Value = FBsk;
			paramFieldDesde.CurrentValues.Add (discreteValDesde);
			paramFields.Add (paramFieldDesde);
				
			//Llama al reporte
			Reporte miRepor = new Reporte("ReporTurnos.rpt", "");
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("Retenciones");
			miRepor.crVista.ParameterFieldInfo = paramFields;
				
			miRepor.Show();
			#endregion Imprime el reporte
		}
	}
}
