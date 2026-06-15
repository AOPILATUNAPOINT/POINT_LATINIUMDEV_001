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
	/// Descripción breve de frmTransferenciasMultiplesLotte.
	/// </summary>
	public class frmTransferenciasMultiplesLotte : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCategoria;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnProcesar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugTransMultiple;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Button btnTfPL;
		private System.Windows.Forms.Button btnTfLf;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTransSal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTransIng;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTransDs;
		private System.Windows.Forms.Button btnTfDs;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTransferenciasMultiplesLotte()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTransferenciasMultiplesLotte));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Categoria");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MinPrin");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MinLOFT");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SalLO");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanTrLO");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MinPLAT");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SalPL");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanTRPL");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoDesign");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MinDESIGN");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SalDS");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanTRDS");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MinPrin");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MinLOFT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SalLO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanTrLO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MinPLAT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SalPL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanTRPL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoDesign");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MinDESIGN");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SalDS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanTRDS");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbCategoria = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnProcesar = new System.Windows.Forms.Button();
			this.ugTransMultiple = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btnTfPL = new System.Windows.Forms.Button();
			this.btnTfLf = new System.Windows.Forms.Button();
			this.txtTransSal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTransIng = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtTransDs = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.btnTfDs = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugTransMultiple)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransSal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransIng)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransDs)).BeginInit();
			this.SuspendLayout();
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
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 35);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(43, 16);
			this.lblBodega.TabIndex = 673;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 215;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 32);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(208, 22);
			this.cmbBodega.TabIndex = 672;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(560, 31);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(80, 24);
			this.btnBuscar.TabIndex = 678;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(280, 35);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(53, 16);
			this.label1.TabIndex = 677;
			this.label1.Text = "Categoria";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCategoria
			// 
			this.cmbCategoria.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCategoria.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 208;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbCategoria.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCategoria.DisplayMember = "Categoria";
			this.cmbCategoria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCategoria.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCategoria.Location = new System.Drawing.Point(344, 33);
			this.cmbCategoria.MaxDropDownItems = 30;
			this.cmbCategoria.Name = "cmbCategoria";
			this.cmbCategoria.Size = new System.Drawing.Size(208, 21);
			this.cmbCategoria.TabIndex = 676;
			this.cmbCategoria.ValueMember = "idCategoria";
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-8, 64);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1660, 8);
			this.groupBox1.TabIndex = 679;
			this.groupBox1.TabStop = false;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1276, 31);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(80, 24);
			this.btnExcel.TabIndex = 683;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnProcesar
			// 
			this.btnProcesar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnProcesar.CausesValidation = false;
			this.btnProcesar.Enabled = false;
			this.btnProcesar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnProcesar.Image = ((System.Drawing.Image)(resources.GetObject("btnProcesar.Image")));
			this.btnProcesar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnProcesar.Location = new System.Drawing.Point(1196, 31);
			this.btnProcesar.Name = "btnProcesar";
			this.btnProcesar.Size = new System.Drawing.Size(80, 24);
			this.btnProcesar.TabIndex = 682;
			this.btnProcesar.Text = "Procesar";
			this.btnProcesar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
			// 
			// ugTransMultiple
			// 
			this.ugTransMultiple.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugTransMultiple.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugTransMultiple.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugTransMultiple.DisplayLayout.Appearance = appearance1;
			this.ugTransMultiple.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 163;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 249;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance2;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Width = 68;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance3;
			ultraGridColumn9.Format = "#,##0";
			ultraGridColumn9.Header.Caption = "Min Principal";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Width = 65;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.BackColor = System.Drawing.Color.Tomato;
			appearance4.BackColor2 = System.Drawing.Color.Tomato;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance4;
			ultraGridColumn10.Format = "#,##0";
			ultraGridColumn10.Header.Caption = "Saldo Loft";
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Width = 63;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance5;
			ultraGridColumn11.Format = "#,##0";
			ultraGridColumn11.Header.Caption = "Min Loft";
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Width = 66;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance6;
			ultraGridColumn12.Format = "#,##0";
			ultraGridColumn12.Header.Caption = "Saldo Loft";
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Width = 66;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(255)), ((System.Byte)(192)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(255)), ((System.Byte)(192)));
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance7;
			ultraGridColumn13.Format = "#,##0";
			ultraGridColumn13.Header.Caption = "Loft";
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Width = 71;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance8;
			ultraGridColumn14.Format = "#,##0";
			ultraGridColumn14.Header.Caption = "Min Platinium";
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Width = 88;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance9;
			ultraGridColumn15.Format = "#,##0";
			ultraGridColumn15.Header.Caption = "Sal Platinium";
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Width = 66;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(255)), ((System.Byte)(255)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(255)), ((System.Byte)(255)));
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance10;
			ultraGridColumn16.Format = "#,##0";
			ultraGridColumn16.Header.Caption = "Platinium";
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn16.Width = 68;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.BackColor = System.Drawing.Color.Tomato;
			appearance11.BackColor2 = System.Drawing.Color.Tomato;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance11;
			ultraGridColumn17.Format = "#,##0";
			ultraGridColumn17.Header.Caption = "Saldo Design";
			ultraGridColumn17.Header.VisiblePosition = 6;
			ultraGridColumn17.Width = 68;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance12;
			ultraGridColumn18.Format = "#,##0";
			ultraGridColumn18.Header.Caption = "Min Design";
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.Width = 69;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance13;
			ultraGridColumn19.Format = "#,##0";
			ultraGridColumn19.Header.Caption = "Sal DS";
			ultraGridColumn19.Header.VisiblePosition = 14;
			ultraGridColumn19.Width = 75;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(224)), ((System.Byte)(192)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(224)), ((System.Byte)(192)));
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance14;
			ultraGridColumn20.Format = "#,##0";
			ultraGridColumn20.Header.Caption = "DESIGN";
			ultraGridColumn20.Header.VisiblePosition = 15;
			ultraGridColumn20.Width = 82;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.ugTransMultiple.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugTransMultiple.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.ugTransMultiple.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugTransMultiple.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugTransMultiple.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.ugTransMultiple.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugTransMultiple.DisplayLayout.Override.CellAppearance = appearance17;
			this.ugTransMultiple.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8.5F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugTransMultiple.DisplayLayout.Override.HeaderAppearance = appearance18;
			appearance19.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance19.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTransMultiple.DisplayLayout.Override.RowAlternateAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTransMultiple.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTransMultiple.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.ugTransMultiple.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugTransMultiple.Location = new System.Drawing.Point(8, 80);
			this.ugTransMultiple.Name = "ugTransMultiple";
			this.ugTransMultiple.Size = new System.Drawing.Size(1348, 424);
			this.ugTransMultiple.TabIndex = 684;
			this.ugTransMultiple.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugTransMultiple_KeyDown);
			this.ugTransMultiple.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugTransMultiple_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(16, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 686;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance22;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 6);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 685;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// btnTfPL
			// 
			this.btnTfPL.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnTfPL.CausesValidation = false;
			this.btnTfPL.Enabled = false;
			this.btnTfPL.Image = ((System.Drawing.Image)(resources.GetObject("btnTfPL.Image")));
			this.btnTfPL.Location = new System.Drawing.Point(940, 4);
			this.btnTfPL.Name = "btnTfPL";
			this.btnTfPL.Size = new System.Drawing.Size(30, 24);
			this.btnTfPL.TabIndex = 688;
			this.btnTfPL.Click += new System.EventHandler(this.btnTfPL_Click);
			// 
			// btnTfLf
			// 
			this.btnTfLf.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnTfLf.CausesValidation = false;
			this.btnTfLf.Enabled = false;
			this.btnTfLf.Image = ((System.Drawing.Image)(resources.GetObject("btnTfLf.Image")));
			this.btnTfLf.Location = new System.Drawing.Point(540, 4);
			this.btnTfLf.Name = "btnTfLf";
			this.btnTfLf.Size = new System.Drawing.Size(30, 24);
			this.btnTfLf.TabIndex = 687;
			this.btnTfLf.Click += new System.EventHandler(this.btnTfLf_Click);
			// 
			// txtTransSal
			// 
			this.txtTransSal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance23.BorderColor = System.Drawing.Color.White;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Center;
			this.txtTransSal.Appearance = appearance23;
			this.txtTransSal.AutoSize = false;
			this.txtTransSal.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.txtTransSal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTransSal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTransSal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTransSal.Location = new System.Drawing.Point(836, 5);
			this.txtTransSal.Name = "txtTransSal";
			this.txtTransSal.ReadOnly = true;
			this.txtTransSal.Size = new System.Drawing.Size(100, 22);
			this.txtTransSal.TabIndex = 692;
			// 
			// txtTransIng
			// 
			this.txtTransIng.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance24.BorderColor = System.Drawing.Color.White;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Center;
			this.txtTransIng.Appearance = appearance24;
			this.txtTransIng.AutoSize = false;
			this.txtTransIng.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.txtTransIng.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTransIng.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTransIng.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTransIng.Location = new System.Drawing.Point(436, 5);
			this.txtTransIng.Name = "txtTransIng";
			this.txtTransIng.ReadOnly = true;
			this.txtTransIng.Size = new System.Drawing.Size(100, 22);
			this.txtTransIng.TabIndex = 691;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Firebrick;
			this.label3.Location = new System.Drawing.Point(580, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(247, 16);
			this.label3.TabIndex = 690;
			this.label3.Text = "TRANSFERENCIA DE ENTRADA PLATINIUM";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.Firebrick;
			this.label2.Location = new System.Drawing.Point(215, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(214, 16);
			this.label2.TabIndex = 689;
			this.label2.Text = "TRANSFERENCIA DE ENTRADA LOFT";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTransDs
			// 
			this.txtTransDs.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance25.BorderColor = System.Drawing.Color.White;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Center;
			this.txtTransDs.Appearance = appearance25;
			this.txtTransDs.AutoSize = false;
			this.txtTransDs.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.txtTransDs.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTransDs.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTransDs.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTransDs.Location = new System.Drawing.Point(1220, 5);
			this.txtTransDs.Name = "txtTransDs";
			this.txtTransDs.ReadOnly = true;
			this.txtTransDs.Size = new System.Drawing.Size(100, 22);
			this.txtTransDs.TabIndex = 694;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.ForeColor = System.Drawing.Color.Firebrick;
			this.label4.Location = new System.Drawing.Point(980, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(229, 16);
			this.label4.TabIndex = 693;
			this.label4.Text = "TRANSFERENCIA DE ENTRADA DESIGN";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnTfDs
			// 
			this.btnTfDs.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnTfDs.CausesValidation = false;
			this.btnTfDs.Enabled = false;
			this.btnTfDs.Image = ((System.Drawing.Image)(resources.GetObject("btnTfDs.Image")));
			this.btnTfDs.Location = new System.Drawing.Point(1325, 4);
			this.btnTfDs.Name = "btnTfDs";
			this.btnTfDs.Size = new System.Drawing.Size(30, 24);
			this.btnTfDs.TabIndex = 695;
			this.btnTfDs.Click += new System.EventHandler(this.btnTfDs_Click);
			// 
			// frmTransferenciasMultiplesLotte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1364, 510);
			this.Controls.Add(this.btnTfDs);
			this.Controls.Add(this.txtTransDs);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtTransSal);
			this.Controls.Add(this.txtTransIng);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btnTfPL);
			this.Controls.Add(this.btnTfLf);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.ugTransMultiple);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnProcesar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbCategoria);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmTransferenciasMultiplesLotte";
			this.Text = "Transferencias de Bodegas";
			this.Load += new System.EventHandler(this.frmTransferenciasMultiplesLotte_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugTransMultiple)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransSal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransIng)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransDs)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		int iBodega = 0;
		
		int idCpaTransfer1 = 0;
		int idCpaTransfer2 = 0;
		int idCpaTransfer3 = 0;

		private void frmTransferenciasMultiplesLotte_Load(object sender, System.EventArgs e)
		{

			miAcceso = new Acceso(cdsSeteoF, "1115");

			if (!Funcion.Permiso("999", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Transferencia de Bodegas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;		
			}

			if (miAcceso.BExportar) this.btnExcel.Enabled = true;
			if (miAcceso.BProcesar) this.btnProcesar.Enabled = true;

			#region Fecha
			this.dtFecha.Value = DateTime.Today;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today.AddDays(1);
			#endregion Fecha

			#region Bodega
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select b.Bodega, b.Nombre From Bodega b Where Bodega in (3, 7) Order by b.Nombre Desc");
			string stExecBodega = string.Format("Select idCategoria, Categoria From ArticuloCategoria Where Activo = 1 Order By Categoria");
			this.cmbCategoria.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExecBodega);			
			#endregion Bodega

			#region Datos Grid
			//this.uGridLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select b.Bodega, b.Nombre, CONVERT(Bit, 0) as Sel FROM Bodega b WHERE Bodega in (4, 5, 6)");
			#endregion Datos Grid

			this.cmbBodega.Focus();
			this.cmbBodega.Value = 3;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void chkTodos_CheckedChanged(object sender, System.EventArgs e)
		{
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
//			{
//				if (this.chkTodos.Checked) dr.Cells["Sel"].Value = 1;
//				else dr.Cells["Sel"].Value = 0;
//			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.ugTransMultiple.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.ugTransMultiple, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void btnProcesar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione la Bodega de Origen")) return;
			string sBodega = "";

			#region Crea Transferencia
			DateTime dtFehaTransferencia = (DateTime)this.dtFecha.Value;			
			int idUsrLot = MenuLatinium.IdUsuario;

			if (DialogResult.Yes == MessageBox.Show("¿Generar TRANSFERENCIA de Inventario?", "Aviso...!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				int idTipoTransporte = 0;
				int idTransporte = 0;
				string NumTransfPL = "";
				string NumTransfLO = "";
				int CanTrnf1 = 0;
				int CanTrnf2 = 0;
				int CanTrnf3 = 0;
				int Proc1 = 0;
				int Proc2 = 0;
				int Proc3 = 0;
				int iBodegaD = 0;
				int iBodegaO = 0;
				string sLeyenda ="";

				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drtr in ugTransMultiple.Rows.All)
						{
							if((int)drtr.Cells["CanTrLO"].Value > 0) CanTrnf1 ++;
							if((int)drtr.Cells["CanTrPL"].Value > 0) CanTrnf2 ++;
							if((int)drtr.Cells["CanTRDS"].Value > 0) CanTrnf3 ++;
						}

						#region Transferencia Bodega Principal a LOFT
						if(CanTrnf1 > 0)
						{
							Proc1 = 1;
							iBodegaD = 5;

							if (true)
							{
								string sSQLNumero = "Exec NumeracionLocales 42, 0, 0";
								oCmdActualiza.CommandText = sSQLNumero;
								NumTransfLO = oCmdActualiza.ExecuteScalar().ToString();
							}

							#region Mensaje
							if((int)this.cmbBodega.Value == 3)
							{
								sLeyenda = "TRANSFERENCIA AUTOMATICA PL a LF";
								iBodegaO = 3;								
							}							

							if((int)this.cmbBodega.Value == 7)
							{
								sLeyenda = "TRANSFERENCIA AUTOMATICA PD a LF";
								iBodegaO = 7;								
							}
						
							#endregion Mensaje					

							string sSQLGrabaMaestro = string.Format("Exec TransferenciasGrabaMaestro {0}, '{1}', '{2}', {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, '{10}','{11}', {12}, {13}", 
								idCpaTransfer1, NumTransfLO, dtFehaTransferencia.ToString("yyyyMMdd HH:mm"), iBodegaO, iBodegaD, 4,
								sLeyenda, "", idTipoTransporte, idTransporte, "","", idUsrLot, idUsrLot);

							oCmdActualiza.CommandText = sSQLGrabaMaestro;
							idCpaTransfer1 = (int)oCmdActualiza.ExecuteScalar();

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ugTransMultiple.Rows.All)
							{
								if((int)dr.Cells["CanTrLO"].Value > 0)
								{
									string sSQLGrabaDetalle = string.Format("Exec TransferenciasGrabaDetalle {0}, {1}, {2}, {3}, {4}",
										0, idCpaTransfer1, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["CanTrLO"].Value, false);
									oCmdActualiza.CommandText = sSQLGrabaDetalle;
									oCmdActualiza.ExecuteNonQuery();
								}
							}
							string stExec = string.Format("Exec TransformaIE {0}", idCpaTransfer1);
							oCmdActualiza.CommandText = stExec;
							oCmdActualiza.ExecuteNonQuery();

							#region Procesa Transferencia SALIDA
							string sSQLProcTSPR = string.Format("Exec Procesa_Transferencia {0}, {1}", idCpaTransfer1, 8);
							oCmdActualiza.CommandText = sSQLProcTSPR;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Procesa Transferencia SALIDA

							#region Procesa Transferencia ENTRADA	
							string sSQLProcTSLO = string.Format("Exec Procesa_Transferencia {0}, {1}", idCpaTransfer1, 9);
							oCmdActualiza.CommandText = sSQLProcTSLO;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Procesa Transferencia ENTRADA
						}
						#endregion Transferencia Bodega Principal a LOFT

						#region Transferencia Bodega Principal a Platimun
						if(CanTrnf2 > 0)
						{
							Proc2 = 1;
							iBodegaD = 6;

							if (true)
							{
								string sSQLNumero = "Exec NumeracionLocales 42, 0, 0";
								oCmdActualiza.CommandText = sSQLNumero;
								NumTransfLO = oCmdActualiza.ExecuteScalar().ToString();
							}

							#region Mensaje
							if((int)this.cmbBodega.Value == 3)
							{
								sLeyenda = "TRANSFERENCIA AUTOMATICA PL a DS";
								iBodegaO = 3;
							}					

							if((int)this.cmbBodega.Value == 7)
							{
								sLeyenda = "TRANSFERENCIA AUTOMATICA PD a DS";
								iBodegaO = 7;
							}							
							#endregion Mensaje					

							string sSQLGrabaMaestroT2 = string.Format("Exec TransferenciasGrabaMaestro {0}, '{1}', '{2}', {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, '{10}','{11}', {12}, {13}", 
								idCpaTransfer2, NumTransfLO, dtFehaTransferencia.ToString("yyyyMMdd HH:mm"), iBodegaO, iBodegaD, 4,
								sLeyenda, "", idTipoTransporte, idTransporte, "","", idUsrLot, idUsrLot);

							oCmdActualiza.CommandText = sSQLGrabaMaestroT2;
							idCpaTransfer2 = (int)oCmdActualiza.ExecuteScalar();

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ugTransMultiple.Rows.All)
							{
								if((int)dr.Cells["CanTrPL"].Value>0)
								{
									string sSQLGrabaDetalleT2 = string.Format("Exec TransferenciasGrabaDetalle {0}, {1}, {2}, {3}, {4}",
										0, idCpaTransfer2, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["CanTrPL"].Value, false);
									oCmdActualiza.CommandText = sSQLGrabaDetalleT2;
									oCmdActualiza.ExecuteNonQuery();
								}
							}
							string stExecT2 = string.Format("Exec TransformaIE {0}", idCpaTransfer2);
							oCmdActualiza.CommandText = stExecT2;
							oCmdActualiza.ExecuteNonQuery();

							#region Procesa Transferencia SALIDA
							string sSQLProcTSPR2 = string.Format("Exec Procesa_Transferencia {0}, {1}", idCpaTransfer2, 8);
							oCmdActualiza.CommandText = sSQLProcTSPR2;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Procesa Transferencia SALIDA

							#region Procesa Transferencia ENTRADA	
							string sSQLProcTSPL = string.Format("Exec Procesa_Transferencia {0}, {1}", idCpaTransfer2, 9);
							oCmdActualiza.CommandText = sSQLProcTSPL;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Procesa Transferencia ENTRADA

						}
						#endregion Transferencia Bodega Principal a Platinum

						#region Transferencia Bodega Principal a Design
						if(CanTrnf3 > 0)
						{
							Proc3 = 1;
							iBodegaD = 4;

							if (true)
							{
								string sSQLNumero = "Exec NumeracionLocales 42, 0, 0";
								oCmdActualiza.CommandText = sSQLNumero;
								NumTransfLO = oCmdActualiza.ExecuteScalar().ToString();
							}

							#region Mensaje
							if((int)this.cmbBodega.Value == 3)
							{
								sLeyenda = "TRANSFERENCIA AUTOMATICA PL a DS";
								iBodegaO = 3;
							}							

							if((int)this.cmbBodega.Value == 7)
							{
								sLeyenda = "TRANSFERENCIA AUTOMATICA PD a DS";
								iBodegaO = 7;
							}
							#endregion Mensaje					

							string sSQLGrabaMaestroT3 = string.Format("Exec TransferenciasGrabaMaestro {0}, '{1}', '{2}', {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, '{10}','{11}', {12}, {13}", 
								idCpaTransfer3, NumTransfLO, dtFehaTransferencia.ToString("yyyyMMdd HH:mm"), iBodegaO, iBodegaD, 4,
								sLeyenda, "", idTipoTransporte, idTransporte, "","", idUsrLot, idUsrLot);

							oCmdActualiza.CommandText = sSQLGrabaMaestroT3;
							idCpaTransfer3 = (int)oCmdActualiza.ExecuteScalar();

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ugTransMultiple.Rows.All)
							{
								if((int)dr.Cells["CanTRDS"].Value>0)
								{
									string sSQLGrabaDetalleT3 = string.Format("Exec TransferenciasGrabaDetalle {0}, {1}, {2}, {3}, {4}",
										0, idCpaTransfer3, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["CanTRDS"].Value, false);
									oCmdActualiza.CommandText = sSQLGrabaDetalleT3;
									oCmdActualiza.ExecuteNonQuery();
								}
							}
							string stExecT3 = string.Format("Exec TransformaIE {0}", idCpaTransfer3);
							oCmdActualiza.CommandText = stExecT3;
							oCmdActualiza.ExecuteNonQuery();

							#region Procesa Transferencia SALIDA
							string sSQLProcTSPR3 = string.Format("Exec Procesa_Transferencia {0}, {1}", idCpaTransfer3, 8);
							oCmdActualiza.CommandText = sSQLProcTSPR3;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Procesa Transferencia SALIDA

							#region Procesa Transferencia ENTRADA	
							string sSQLProcTSPL3 = string.Format("Exec Procesa_Transferencia {0}, {1}", idCpaTransfer3, 9);
							oCmdActualiza.CommandText = sSQLProcTSPL3;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Procesa Transferencia ENTRADA
						}
						#endregion Transferencia Bodega Principal a Design

						oTransaction.Commit();
						
						oConexion.Close();

						this.btnProcesar.Enabled = false;

						if(Proc1 == 1 && Proc2 == 1 && Proc3 == 1) MessageBox.Show("TRANSFERENCIA Procesada..!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
						else if(Proc1 == 1 && Proc2 == 1) MessageBox.Show("TRANSFERENCIA Procesada..!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
						else if(Proc1 == 1 && Proc3 == 1) MessageBox.Show("TRANSFERENCIA Procesada..!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
						else if(Proc2 == 1 && Proc3 == 1) MessageBox.Show("TRANSFERENCIA Procesada..!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
						else if(Proc1 == 1 && Proc2 == 0 && Proc3 == 0) MessageBox.Show("TRANSFERENCIA de LOFT Procesada..!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
						else if(Proc1 == 0 && Proc2 == 1 && Proc3 == 0) MessageBox.Show("TRANSFERENCIA de PLATINIUM Procesada..!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
						else if(Proc1 == 0 && Proc2 == 0 && Proc3 == 1) MessageBox.Show("TRANSFERENCIA de DESIGN Procesada..!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
						else MessageBox.Show("NO se Registra Valores para realizar TRANSFERENCIA.?", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);					
					}
					catch (Exception ex)
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);						
					}
					finally
					{
						oConexion.Close();

						/*Visualiza Números de Transferencia*/
						if(idCpaTransfer1 > 0)
						{
							string lcBskNTrn1 = "SELECT Codigo FROM CompraTransf WHERE idCompraTransf="+idCpaTransfer1.ToString();
							string NumTrn1 = Funcion.sEscalarSQL(cdsSeteoF, lcBskNTrn1);
							this.txtTransIng.Value = NumTrn1.Trim();
							this.btnTfLf.Enabled = true ;
						}

						if(idCpaTransfer2 > 0)
						{
							string lcBskNTrn2 = "SELECT Codigo FROM CompraTransf WHERE idCompraTransf="+idCpaTransfer2.ToString();
							string NumTrn2 = Funcion.sEscalarSQL(cdsSeteoF, lcBskNTrn2);
							this.txtTransSal.Value = NumTrn2.Trim();
							this.btnTfPL.Enabled = true ;
						}

						if(idCpaTransfer3 > 0)
						{
							string lcBskNTrn3 = "SELECT Codigo FROM CompraTransf WHERE idCompraTransf="+idCpaTransfer3.ToString();
							string NumTrn3 = Funcion.sEscalarSQL(cdsSeteoF, lcBskNTrn3);
							this.txtTransDs.Value = NumTrn3.Trim();
							this.btnTfDs.Enabled = true ;
						}
					}
				}
			}
			#endregion Crea Transferencia
		}

		private void ugTransMultiple_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ugTransMultiple);
		}

		private void ugTransMultiple_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Cantidad
			int Cantot = 0;						
			int difSald = 0;
			int iCantidadBodega = 0;
			int SldPrin = 0;

			if((int)this.cmbBodega.Value == 3)
			{
				SldPrin = (int)e.Cell.Row.Cells["SaldoDesign"].Value;
				difSald = (int)e.Cell.Row.Cells["SaldoDesign"].Value - ((int)e.Cell.Row.Cells["CanTrLo"].Value + (int)e.Cell.Row.Cells["CanTrPl"].Value + (int)e.Cell.Row.Cells["CanTRDS"].Value);
			}
			else if ((int)this.cmbBodega.Value == 7)
			{
				SldPrin = (int)e.Cell.Row.Cells["Saldo"].Value;				
				difSald = (int)e.Cell.Row.Cells["Saldo"].Value - ((int)e.Cell.Row.Cells["CanTrLo"].Value + (int)e.Cell.Row.Cells["CanTrPl"].Value + (int)e.Cell.Row.Cells["CanTRDS"].Value);				
			}
				
			if(difSald < 0)
			{
				MessageBox.Show("El SALDO de la Bodega ORIGEN es MENOR a la Cantidad de TRANSFERENCIA.", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cell.Value = 0;
			}
			else if((int)e.Cell.Row.Cells["CanTrLo"].Value < 0)
			{
				MessageBox.Show("Cantidad no puede ser MENOR a CERO.", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cell.Value = 0;
			}
			else if((int)e.Cell.Row.Cells["CanTrPL"].Value < 0)
			{
				MessageBox.Show("Cantidad no puede ser MENOR a CERO.", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cell.Value = 0;
			}
			else if((int)e.Cell.Row.Cells["CanTRDS"].Value < 0)
			{
				MessageBox.Show("Cantidad no puede ser MENOR a CERO.", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cell.Value = 0;
			}
			else
			{
				if (e.Cell.Column.ToString() == "Cantidad")
				{
					if (e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value) e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
					int VerificaCambio = 1; 
					if (false)
					{
						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit = 0 If ((Select idTipoGrupo From Articulo Where idArticulo = {0}) In (1, 6) And (Select idGrupoArticulo From Articulo Where idArticulo = {0}) <> 26) Set @Res = 1 Select @Res", (int)e.Cell.Row.Cells["idArticulo"].Value), true))
						{
							if ((bool)e.Cell.Row.Cells["Principal"].Value) 
							{
								e.Cell.Value = 1;
								MessageBox.Show("Para Productos Principales de un Combo o Promocion la Cantidad es 1", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}

							if ((int)e.Cell.Row.Cells["Tipo"].Value > 0) 
							{
								e.Cell.Value = 1;
								MessageBox.Show("Para Productos que forman parte de un Combo o Promocion la Cantidad es 1", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}

							DateTime dtFecha = (DateTime)this.dtFecha.Value;
							int idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;
							int Existencia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.CalculaExistenciasInventario('{0}', {1}, {2})", dtFecha.ToString("yyyyMMdd"), (int)e.Cell.Row.Cells["idArticulo"].Value, 1));
							int iCantidad = 0;

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugTransMultiple.Rows.All)
							{
								if (idArticulo == (int)dr.Cells["idArticulo"].Value) iCantidad = iCantidad + (int)dr.Cells["Cantidad"].Value;
							}

							if (iCantidad > Existencia)
							{					
								if (Existencia < 1)
								{
									//bEliminaAlValidar = true;
									e.Cell.Row.Delete();
									//bEliminaAlValidar = false;
									MessageBox.Show("Este Articulo No Tiene Existencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								}
								else
								{
									e.Cell.Value = 1;
									MessageBox.Show(string.Format("Existencia Actual : {0}", Existencia), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								}				
							}
						}
					}
				}
			}				
			#endregion Cantidad
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			int idCategoria = 0;
			//string sLocales = "";						

			#region Validacion Bodega
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
//			{
//				if ((bool) dr.Cells["Sel"].Value)
//				{
//					if (sLocales.Length == 0) sLocales = dr.Cells["Bodega"].Value.ToString();
//					else if (sLocales.Length > 0) sLocales = sLocales + ", " + dr.Cells["Bodega"].Value.ToString();
//				}
//			}
//
//			if (sLocales.Length == 0)
//			{
//				MessageBox.Show("Seleccione al menos un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				return;
//			}
			#endregion Validacion Bodega

			if (this.cmbCategoria.ActiveRow != null) idCategoria = (int)this.cmbCategoria.Value;

			string sSQL = string.Format("Exec ExistenciaxBodega {0}", idCategoria);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ugTransMultiple);
		}

		private void uGridLocales_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Check
//			if (e.Cell.Row.Cells["Sel"].Value != System.DBNull.Value)
//			{
//				if (e.Cell.Column.ToString() == "Sel")
//				{
//					//bActualiza = true;
//					this.uGridLocales.UpdateData();
//
//					if ((int)this.uGridLocales.Rows.Count > 1)
//					{
//						int iBodega = 0;
//						int icount = 0;
//						iBodega = (int)e.Cell.Row.Cells["Bodega"].Value;
//						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
//						{								
//							if ((bool)dr.Cells["Sel"].Value)
//							{
//								icount++;
//							}
//						}
//						if (icount > 1)
//						{
//							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
//							{	
//								if ((bool)dr.Cells["Sel"].Value)
//								{
//									if(iBodega != (int)dr.Cells["Bodega"].Value)
//									{
//										dr.Cells["Sel"].Value = false;
//									}
//								}
//							}
//						}
//					}
//				}
//			}
			#endregion Check
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugTransMultiple.Rows.All)
			{
				dr.Cells["CanTrLO"].Value = 0;
				dr.Cells["CanTRPL"].Value = 0;
				dr.Cells["CanTRDS"].Value = 0;
			}
		}

		private void btnTfLf_Click(object sender, System.EventArgs e)
		{
			/*Imprimir TF LF*/
			if(idCpaTransfer1 > 0)
			{
				int idDBEmpresa = MenuLatinium.stIdDB;
				ParameterFields paramFields = new ParameterFields ();

				ParameterField paramFieldTransf = new ParameterField();
				ParameterDiscreteValue discreteValTransf = new ParameterDiscreteValue();
				paramFieldTransf.ParameterFieldName = "@idComprTransf";
				discreteValTransf.Value = idCpaTransfer1;
				paramFieldTransf.CurrentValues.Add (discreteValTransf);
				paramFields.Add (paramFieldTransf);
			
				ParameterField paramFieldEmpresa = new ParameterField ();
				ParameterDiscreteValue discreteValEmpresa = new ParameterDiscreteValue ();
				paramFieldEmpresa.ParameterFieldName = "@idEmpresa";
				discreteValEmpresa.Value = idDBEmpresa; 
				paramFieldEmpresa.CurrentValues.Add (discreteValEmpresa);
				paramFields.Add (paramFieldEmpresa);

				Reporte miRep = new Reporte("Transfer2ConImg.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Transferencia");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
			}
			else MessageBox.Show("NO se ha GENERADO TRANSFERENCIA para LF...!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}

		private void btnTfPL_Click(object sender, System.EventArgs e)
		{
			/*Imprimir TF PL*/
			if(idCpaTransfer2 > 0)
			{
				this.btnTfPL.Enabled = true ;
				int idDBEmpresa = MenuLatinium.stIdDB;
				ParameterFields paramFields = new ParameterFields ();

				ParameterField paramFieldTransf = new ParameterField();
				ParameterDiscreteValue discreteValTransf = new ParameterDiscreteValue();
				paramFieldTransf.ParameterFieldName = "@idComprTransf";
				discreteValTransf.Value = idCpaTransfer2;
				paramFieldTransf.CurrentValues.Add (discreteValTransf);
				paramFields.Add (paramFieldTransf);
			
				ParameterField paramFieldEmpresa = new ParameterField ();
				ParameterDiscreteValue discreteValEmpresa = new ParameterDiscreteValue ();
				paramFieldEmpresa.ParameterFieldName = "@idEmpresa";
				discreteValEmpresa.Value = idDBEmpresa; 
				paramFieldEmpresa.CurrentValues.Add (discreteValEmpresa);
				paramFields.Add (paramFieldEmpresa);

				Reporte miRep = new Reporte("Transfer2ConImg.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Transferencia");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
			}
			else MessageBox.Show("NO se ha GENERADO TRANSFERENCIA para PL...!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}

		private void btnTfDs_Click(object sender, System.EventArgs e)
		{
			/*Imprimir TF DS*/
			if(idCpaTransfer3 > 0)
			{
				this.btnTfPL.Enabled = true ;
				int idDBEmpresa = MenuLatinium.stIdDB;
				ParameterFields paramFields = new ParameterFields ();

				ParameterField paramFieldTransf = new ParameterField();
				ParameterDiscreteValue discreteValTransf = new ParameterDiscreteValue();
				paramFieldTransf.ParameterFieldName = "@idComprTransf";
				discreteValTransf.Value = idCpaTransfer3;
				paramFieldTransf.CurrentValues.Add (discreteValTransf);
				paramFields.Add (paramFieldTransf);
			
				ParameterField paramFieldEmpresa = new ParameterField ();
				ParameterDiscreteValue discreteValEmpresa = new ParameterDiscreteValue ();
				paramFieldEmpresa.ParameterFieldName = "@idEmpresa";
				discreteValEmpresa.Value = idDBEmpresa; 
				paramFieldEmpresa.CurrentValues.Add (discreteValEmpresa);
				paramFields.Add (paramFieldEmpresa);

				Reporte miRep = new Reporte("Transfer2ConImg.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Transferencia");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
			}
			else MessageBox.Show("NO se ha GENERADO TRANSFERENCIA para DS...!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}
	}
}
