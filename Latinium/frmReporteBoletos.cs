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
using System.Threading;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteBoletos.
	/// </summary>
	public class frmReporteBoletos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dthasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtdesde;
		private System.Windows.Forms.Button btnExcel;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbbodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private System.Windows.Forms.Label lblIdentificacion;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugReporteBoletos;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteBoletos()
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
//
		C1.Data.C1DataRow dr;
//		C1.Data.C1DataRow drSeteoF;

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteBoletos));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("N_Boletos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("id_anticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ver", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("N_Boletos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Agencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("id_anticipo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.cmbbodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label6 = new System.Windows.Forms.Label();
			this.ugReporteBoletos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dthasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtdesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnExcel = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cmbbodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugReporteBoletos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dthasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtdesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(72, 16);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 16);
			this.label3.TabIndex = 674;
			this.label3.Text = "Desde";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(256, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 673;
			this.label2.Text = "Hasta";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(328, 40);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(72, 31);
			this.btnConsultar.TabIndex = 672;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// cmbbodega
			// 
			appearance1.FontData.Name = "Tahoma";
			this.cmbbodega.Appearance = appearance1;
			this.cmbbodega.CausesValidation = false;
			this.cmbbodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbbodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbbodega.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbbodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbbodega.DisplayMember = "Nombre";
			this.cmbbodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbbodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbbodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbbodega.Location = new System.Drawing.Point(128, 47);
			this.cmbbodega.MaxDropDownItems = 30;
			this.cmbbodega.Name = "cmbbodega";
			this.cmbbodega.Size = new System.Drawing.Size(180, 22);
			this.cmbbodega.TabIndex = 671;
			this.cmbbodega.ValueMember = "Bodega";
			this.cmbbodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbbodega_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(24, 50);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(92, 16);
			this.label6.TabIndex = 670;
			this.label6.Text = "Centro de Costos";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ugReporteBoletos
			// 
			this.ugReporteBoletos.AlphaBlendMode = Infragistics.Win.AlphaBlendMode.Disabled;
			this.ugReporteBoletos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugReporteBoletos.DataSource = this.ultraDataSource4;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ugReporteBoletos.DisplayLayout.Appearance = appearance2;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance3.BackGradientAlignment = Infragistics.Win.GradientAlignment.Form;
			appearance3.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(224)), ((System.Byte)(192)));
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			ultraGridColumn10.CellAppearance = appearance3;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			appearance4.ImageBackgroundAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			appearance4.ImageBackgroundOrigin = Infragistics.Win.ImageBackgroundOrigin.Container;
			appearance4.ImageBackgroundStyle = Infragistics.Win.ImageBackgroundStyle.Centered;
			appearance4.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.FormulaErrorAppearance = appearance4;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn10.Width = 30;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand2.SummaryFooterCaption = "TOTAL GENERAL";
			this.ugReporteBoletos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ugReporteBoletos.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugReporteBoletos.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.ugReporteBoletos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugReporteBoletos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugReporteBoletos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ugReporteBoletos.DisplayLayout.Override.CardAreaAppearance = appearance6;
			this.ugReporteBoletos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugReporteBoletos.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.ugReporteBoletos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugReporteBoletos.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugReporteBoletos.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ugReporteBoletos.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ugReporteBoletos.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ugReporteBoletos.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ugReporteBoletos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugReporteBoletos.Location = new System.Drawing.Point(24, 120);
			this.ugReporteBoletos.Name = "ugReporteBoletos";
			this.ugReporteBoletos.Size = new System.Drawing.Size(672, 336);
			this.ugReporteBoletos.TabIndex = 669;
			this.ugReporteBoletos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugReporteBoletos_CellChange);
			this.ugReporteBoletos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugReporteBoletos_ClickCellButton);
			this.ugReporteBoletos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugReporteBoletos_InitializeLayout);
			this.ugReporteBoletos.AfterCellActivate += new System.EventHandler(this.ugReporteBoletos_AfterCellActivate);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(System.UInt32);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			this.ultraDataSource4.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource4_CellDataRequested);
			// 
			// dthasta
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dthasta.Appearance = appearance10;
			this.dthasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dthasta.DateButtons.Add(dateButton1);
			this.dthasta.Format = "dd/MM/yyyy";
			this.dthasta.Location = new System.Drawing.Point(296, 11);
			this.dthasta.Name = "dthasta";
			this.dthasta.NonAutoSizeHeight = 23;
			this.dthasta.Size = new System.Drawing.Size(104, 21);
			this.dthasta.SpinButtonsVisible = true;
			this.dthasta.TabIndex = 668;
			this.dthasta.Value = ((object)(resources.GetObject("dthasta.Value")));
			// 
			// dtdesde
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtdesde.Appearance = appearance11;
			this.dtdesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtdesde.DateButtons.Add(dateButton2);
			this.dtdesde.Format = "dd/MM/yyyy";
			this.dtdesde.Location = new System.Drawing.Point(128, 11);
			this.dtdesde.Name = "dtdesde";
			this.dtdesde.NonAutoSizeHeight = 23;
			this.dtdesde.Size = new System.Drawing.Size(104, 21);
			this.dtdesde.SpinButtonsVisible = true;
			this.dtdesde.TabIndex = 667;
			this.dtdesde.Value = ((object)(resources.GetObject("dtdesde.Value")));
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(480, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 958;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
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
			// ultraDataSource2
			// 
			ultraDataColumn10.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			this.ultraDataSource2.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource2_CellDataRequested);
			// 
			// txtRuc
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance12;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(128, 80);
			this.txtRuc.MaxLength = 13;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(176, 21);
			this.txtRuc.TabIndex = 960;
			this.txtRuc.Text = "0";
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(56, 88);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(49, 16);
			this.lblIdentificacion.TabIndex = 961;
			this.lblIdentificacion.Text = "RUC/C.I.";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmReporteBoletos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(712, 462);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.cmbbodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.ugReporteBoletos);
			this.Controls.Add(this.dthasta);
			this.Controls.Add(this.dtdesde);
			this.Name = "frmReporteBoletos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteBoletos_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteBoletos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbbodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugReporteBoletos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dthasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtdesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

	

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			


			if (!Validacion.vbFechaEnDocumentos(this.dtdesde, "Seleccione la Fecha Inicial", false, "", cdsSeteoF)) return;

			if (!Validacion.vbFechaEnDocumentos(this.dthasta, "Seleccione la Fecha Final", false, "", cdsSeteoF)) return;

		
//			if (this.txtRuc.Text.Equals(""))	
//			{  
//				MessageBox.Show("Ingrese Numero de Cedula", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//												this.txtRuc.Focus();
//								return;}	

		if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtdesde.Value), Convert.ToDateTime(this.dthasta.Value), this.dtdesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;
			 
		
			//int  ruc = 0 ;
			int  bodega = 0;
			
			if (this.cmbbodega.ActiveRow != null) bodega = (int)this.cmbbodega.Value;
			//if (this.txtRuc.Text != null) ruc= Convert.ToInt32( this.txtRuc.Text.ToString());
		


			string sSQL = string.Format("Exec [dbo].[ReporteBoleto]'{0}','{1}',{2},{3}", 
				Convert.ToDateTime(this.dtdesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dthasta.Value).ToString("yyyyMMdd"), bodega,0);
				this.ugReporteBoletos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL); 

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmReporteBoletos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();
			
		}

		private void frmReporteBoletos_Load(object sender, System.EventArgs e)

		{
			#region  Acceso
			miAcceso = new Acceso(cdsSeteoF, "200533");

			if (!Funcion.Permiso("1352", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar al reporte ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
											
										
			if (miAcceso.BExportar) this.btnExcel.Enabled =  true;

			#endregion Acceso

			#region  ValidarFecha
			this.dtdesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dthasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");				
								
										
			this.dtdesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dthasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtdesde.Value = DateTime.Today;
			this.dthasta.Value = DateTime.Today; 
			this.cmbbodega.DataSource=Funcion.dvProcedimiento(cdsSeteoF,"select Bodega , Nombre from  Bodega    order by Nombre");
			#endregion ValidarFecha





		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.ugReporteBoletos);
		}

		private void cmbbodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		  
		}

		private void btnImprime_Click(object sender, System.EventArgs e)
		{

		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
			Funcion.CamposNumericos(sender, e);
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
			
		}

		private void ugReporteBoletos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "ruc")
			{
				if (e.Cell.Row.Cells["ruc"].Value != System.DBNull.Value)
				{
					this.ugReporteBoletos.UpdateData();
					
					if ((bool)e.Cell.Row.Cells["ruc"].Value)
					{
//						string sSQL = string.Format("Exec GuardaTemporalImpresionDePrecios {0}, {1}, {2}", 
//							(int)e.Cell.Row.Cells["idListaPreciador"].Value, (int)this.cmbBodega.Value, (int)this.rdgPreciador.EditValue);
//						Funcion.EjecutaSQL(cdsSeteoF, sSQL);
					}
					else
					{
//						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From TemporalListaPreciador Where idListaPreciador = {0}", 
//							(int)e.Cell.Row.Cells["idListaPreciador"].Value));
					}
				}
			}
		}

		private void ultraDataSource2_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void ugReporteBoletos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraDataSource4_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void ugReporteBoletos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{

			  int anticipo;
			 anticipo  =  (int)this.ugReporteBoletos.ActiveRow.Cells["id_anticipo"].Value;
			if (e.Cell.Column.ToString() == "Ver")
			{
				#region Impresion Sorteo
				ParameterFields paramFields = new ParameterFields();

				ParameterField pfIdOrigen2 = new ParameterField();
				ParameterDiscreteValue ValIdOrigen2 = new ParameterDiscreteValue();
				pfIdOrigen2.ParameterFieldName = "@idAnticipo";
				ValIdOrigen2.Value = anticipo;
				pfIdOrigen2.CurrentValues.Add(ValIdOrigen2);
				paramFields.Add(pfIdOrigen2);

				Reporte miRep = new Reporte("sorteo_ticket1.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Cupones de  Sorte");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
				#endregion Impresion Sorteo
//			
			
			
			}
		}

		private void ugReporteBoletos_AfterCellActivate(object sender, System.EventArgs e)
		{
		
		}
	}
}
