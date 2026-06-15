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
	/// Descripción breve de frmTransferLotte.
	/// </summary>
	public class frmTransferLotte : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.ToolTip toolTip1;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugTransMultiple;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCategoria;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTransIng;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTransSal;
		private System.Windows.Forms.Button btnTfLf;
		private System.Windows.Forms.Button btnTfPL;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnProcesar;
		private System.ComponentModel.IContainer components;

		public frmTransferLotte()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTransferLotte));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MinPrin");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MinLOFT");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SalLO");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanTrLO");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MinPLAT");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SalPL");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanTRPL");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MinPrin");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MinLOFT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SalLO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanTrLO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MinPLAT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SalPL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanTRPL");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Categoria");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnProcesar = new System.Windows.Forms.Button();
			this.ugTransMultiple = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btnTfLf = new System.Windows.Forms.Button();
			this.btnTfPL = new System.Windows.Forms.Button();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbCategoria = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtTransIng = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTransSal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugTransMultiple)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransIng)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransSal)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 268;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 267;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			// 
			// btnProcesar
			// 
			this.btnProcesar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnProcesar.CausesValidation = false;
			this.btnProcesar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnProcesar.Image = ((System.Drawing.Image)(resources.GetObject("btnProcesar.Image")));
			this.btnProcesar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnProcesar.Location = new System.Drawing.Point(710, 6);
			this.btnProcesar.Name = "btnProcesar";
			this.btnProcesar.Size = new System.Drawing.Size(80, 24);
			this.btnProcesar.TabIndex = 266;
			this.btnProcesar.Text = "Procesar";
			this.btnProcesar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnProcesar, "Procesa información");
			this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
			// 
			// ugTransMultiple
			// 
			this.ugTransMultiple.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugTransMultiple.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugTransMultiple.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ugTransMultiple.DisplayLayout.Appearance = appearance2;
			this.ugTransMultiple.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 280;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance3;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance4;
			ultraGridColumn5.Format = "#,##0";
			ultraGridColumn5.Header.Caption = "Min Principal";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 80;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.BackColor = System.Drawing.Color.Tomato;
			appearance5.BackColor2 = System.Drawing.Color.Tomato;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance5;
			ultraGridColumn6.Format = "#,##0";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance6;
			ultraGridColumn7.Format = "#,##0";
			ultraGridColumn7.Header.Caption = "Min Loft";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 80;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance7;
			ultraGridColumn8.Format = "#,##0";
			ultraGridColumn8.Header.Caption = "Saldo Loft";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 80;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(255)), ((System.Byte)(192)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(255)), ((System.Byte)(192)));
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance8;
			ultraGridColumn9.Format = "#,##0";
			ultraGridColumn9.Header.Caption = "Loft";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance9;
			ultraGridColumn10.Format = "#,##0";
			ultraGridColumn10.Header.Caption = "Min Platinium";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 85;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance10;
			ultraGridColumn11.Format = "#,##0";
			ultraGridColumn11.Header.Caption = "Sal Platinium";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 80;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(255)), ((System.Byte)(255)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(255)), ((System.Byte)(255)));
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance11;
			ultraGridColumn12.Format = "#,##0";
			ultraGridColumn12.Header.Caption = "Platinium";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 80;
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
																										 ultraGridColumn12});
			this.ugTransMultiple.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugTransMultiple.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.ugTransMultiple.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugTransMultiple.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugTransMultiple.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.ugTransMultiple.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugTransMultiple.DisplayLayout.Override.CellAppearance = appearance14;
			this.ugTransMultiple.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8.5F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugTransMultiple.DisplayLayout.Override.HeaderAppearance = appearance15;
			appearance16.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance16.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTransMultiple.DisplayLayout.Override.RowAlternateAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTransMultiple.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTransMultiple.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.ugTransMultiple.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugTransMultiple.Location = new System.Drawing.Point(8, 88);
			this.ugTransMultiple.Name = "ugTransMultiple";
			this.ugTransMultiple.Size = new System.Drawing.Size(1206, 373);
			this.ugTransMultiple.TabIndex = 271;
			this.ugTransMultiple.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugTransMultiple_KeyDown);
			this.ugTransMultiple.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugTransMultiple_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// btnTfLf
			// 
			this.btnTfLf.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnTfLf.CausesValidation = false;
			this.btnTfLf.Enabled = false;
			this.btnTfLf.Image = ((System.Drawing.Image)(resources.GetObject("btnTfLf.Image")));
			this.btnTfLf.Location = new System.Drawing.Point(1182, 6);
			this.btnTfLf.Name = "btnTfLf";
			this.btnTfLf.Size = new System.Drawing.Size(30, 24);
			this.btnTfLf.TabIndex = 278;
			this.toolTip1.SetToolTip(this.btnTfLf, "Imprime TF LF");
			this.btnTfLf.Click += new System.EventHandler(this.button3_Click);
			// 
			// btnTfPL
			// 
			this.btnTfPL.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnTfPL.CausesValidation = false;
			this.btnTfPL.Enabled = false;
			this.btnTfPL.Image = ((System.Drawing.Image)(resources.GetObject("btnTfPL.Image")));
			this.btnTfPL.Location = new System.Drawing.Point(1182, 38);
			this.btnTfPL.Name = "btnTfPL";
			this.btnTfPL.Size = new System.Drawing.Size(30, 24);
			this.btnTfPL.TabIndex = 279;
			this.toolTip1.SetToolTip(this.btnTfPL, "Imprime TF PL");
			this.btnTfPL.Click += new System.EventHandler(this.button4_Click);
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
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 42);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(53, 16);
			this.label1.TabIndex = 273;
			this.label1.Text = "Categoria";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCategoria
			// 
			this.cmbCategoria.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCategoria.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 180;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbCategoria.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCategoria.DisplayMember = "Categoria";
			this.cmbCategoria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCategoria.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCategoria.Location = new System.Drawing.Point(72, 40);
			this.cmbCategoria.MaxDropDownItems = 30;
			this.cmbCategoria.Name = "cmbCategoria";
			this.cmbCategoria.Size = new System.Drawing.Size(180, 21);
			this.cmbCategoria.TabIndex = 272;
			this.cmbCategoria.ValueMember = "idCategoria";
			this.cmbCategoria.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCategoria_KeyDown);
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.Firebrick;
			this.label2.Location = new System.Drawing.Point(806, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(214, 16);
			this.label2.TabIndex = 274;
			this.label2.Text = "TRANSFERENCIA DE ENTRADA LOFT";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Firebrick;
			this.label3.Location = new System.Drawing.Point(806, 42);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(247, 16);
			this.label3.TabIndex = 275;
			this.label3.Text = "TRANSFERENCIA DE ENTRADA PLATINIUM";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTransIng
			// 
			this.txtTransIng.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance19.BorderColor = System.Drawing.Color.White;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Center;
			this.txtTransIng.Appearance = appearance19;
			this.txtTransIng.AutoSize = false;
			this.txtTransIng.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.txtTransIng.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTransIng.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTransIng.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTransIng.Location = new System.Drawing.Point(1070, 7);
			this.txtTransIng.Name = "txtTransIng";
			this.txtTransIng.ReadOnly = true;
			this.txtTransIng.Size = new System.Drawing.Size(100, 22);
			this.txtTransIng.TabIndex = 276;
			// 
			// txtTransSal
			// 
			this.txtTransSal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance20.BorderColor = System.Drawing.Color.White;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Center;
			this.txtTransSal.Appearance = appearance20;
			this.txtTransSal.AutoSize = false;
			this.txtTransSal.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.txtTransSal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTransSal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTransSal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTransSal.Location = new System.Drawing.Point(1070, 39);
			this.txtTransSal.Name = "txtTransSal";
			this.txtTransSal.ReadOnly = true;
			this.txtTransSal.Size = new System.Drawing.Size(100, 22);
			this.txtTransSal.TabIndex = 277;
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(272, 38);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(80, 24);
			this.btnBuscar.TabIndex = 294;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(710, 38);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(80, 24);
			this.btnExcel.TabIndex = 296;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1222, 8);
			this.groupBox1.TabIndex = 297;
			this.groupBox1.TabStop = false;
			// 
			// frmTransferLotte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1222, 466);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.btnTfPL);
			this.Controls.Add(this.btnTfLf);
			this.Controls.Add(this.txtTransSal);
			this.Controls.Add(this.txtTransIng);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbCategoria);
			this.Controls.Add(this.ugTransMultiple);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btnProcesar);
			this.KeyPreview = true;
			this.Name = "frmTransferLotte";
			this.Text = "Transferencias de Bodegas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmTransferLotte_KeyDown);
			this.Load += new System.EventHandler(this.frmTransferLotte_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugTransMultiple)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransIng)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransSal)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int idCpaTransfer1 = 0;
		int idCpaTransfer2 = 0;

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmTransferLotte_Load(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Today;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today.AddDays(1);

			string stExecBodega = string.Format("Select idCategoria, Categoria From ArticuloCategoria Where Activo = 1 Order By Categoria");
			this.cmbCategoria.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExecBodega);
		}

		private void ugTransMultiple_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ugTransMultiple);
		}

		private void ugTransMultiple_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Cantidad
			int Cantot = 0;
			int SldPrin = (int)e.Cell.Row.Cells["Saldo"].Value;
			int difSald = (int)e.Cell.Row.Cells["Saldo"].Value - ((int)e.Cell.Row.Cells["CanTrLo"].Value + (int)e.Cell.Row.Cells["CanTrPl"].Value);

			if(difSald < 0)
			{
				MessageBox.Show("El SALDO de la Bodega PRINCIPAL es MENOR a la Cantidad de TRANSFERENCIA.?", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cell.Value = 0;
			}
			else if((int)e.Cell.Row.Cells["CanTrLo"].Value<0)
			{
				MessageBox.Show("Cantidad no puede ser MENOR a CERO.?", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cell.Value = 0;
			}
			else if((int)e.Cell.Row.Cells["CanTrPL"].Value<0)
			{
				MessageBox.Show("Cantidad no puede ser MENOR a CERO.?", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
						if (Funcion.bEscalarSQL(cdsSeteo, string.Format("Declare @Res Bit = 0 If ((Select idTipoGrupo From Articulo Where idArticulo = {0}) In (1, 6) And (Select idGrupoArticulo From Articulo Where idArticulo = {0}) <> 26) Set @Res = 1 Select @Res", (int)e.Cell.Row.Cells["idArticulo"].Value), true))
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
							int Existencia = Funcion.iEscalarSQL(cdsSeteo, string.Format("Select dbo.CalculaExistenciasInventario('{0}', {1}, {2})", dtFecha.ToString("yyyyMMdd"), (int)e.Cell.Row.Cells["idArticulo"].Value, 1));
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
		
		private void button3_Click(object sender, System.EventArgs e)
		{
			/*Imprimir TF LF*/
			if(idCpaTransfer1>0)
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

		private void button4_Click(object sender, System.EventArgs e)
		{
			/*Imprimir TF PL*/
			if(idCpaTransfer2>0)
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

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			int idCategoria = 0;
			if (this.cmbCategoria.ActiveRow != null) idCategoria = (int)this.cmbCategoria.Value;

			string sSQL = string.Format("Exec ExistenciaxBodega {0}", idCategoria);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ugTransMultiple);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ugTransMultiple);
		}

		private void btnProcesar_Click(object sender, System.EventArgs e)
		{
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
				int Proc1 = 0;
				int Proc2 = 0;

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
							if((int)drtr.Cells["CanTrLO"].Value>0) CanTrnf1 ++;
							if((int)drtr.Cells["CanTrPL"].Value>0) CanTrnf2 ++;
						}

						#region Transferencia Bodega Principal a LOFT
						if(CanTrnf1 > 0)
						{
							Proc1 = 1;
							if (true)
							{
								string sSQLNumero = "Exec NumeracionLocales 42, 0, 0";
								oCmdActualiza.CommandText = sSQLNumero;
								NumTransfLO = oCmdActualiza.ExecuteScalar().ToString();
							}	
							string sSQLGrabaMaestro = string.Format("Exec TransferenciasGrabaMaestro {0}, '{1}', '{2}', {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, '{10}','{11}', {12}, {13}", 
								idCpaTransfer1, NumTransfLO, dtFehaTransferencia.ToString("yyyyMMdd HH:mm"), 3, 1, 4, 
								"TRANSFERENCIA AUTOMATICA PR a LF", "", idTipoTransporte, idTransporte, "","", idUsrLot, idUsrLot);

							oCmdActualiza.CommandText = sSQLGrabaMaestro;
							idCpaTransfer1 = (int)oCmdActualiza.ExecuteScalar();

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ugTransMultiple.Rows.All)
							{
								if((int)dr.Cells["CanTrLO"].Value>0)
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

						#region Transferencia Bodega Principal a PLATINIUM
						if(CanTrnf2 > 0)
						{
							Proc2 = 1;
							if (true)
							{
								string sSQLNumero = "Exec NumeracionLocales 42, 0, 0";
								oCmdActualiza.CommandText = sSQLNumero;
								NumTransfPL = oCmdActualiza.ExecuteScalar().ToString();
							}	

							string sSQLGrabaMaestroT2 = string.Format("Exec TransferenciasGrabaMaestro {0}, '{1}', '{2}', {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, '{10}','{11}', {12}, {13}", 
								idCpaTransfer2, NumTransfPL, dtFehaTransferencia.ToString("yyyyMMdd HH:mm"), 3, 2, 4, 
								"TRANSFERENCIA AUTOMATICA PR a PL", "", idTipoTransporte, idTransporte, "","", idUsrLot, idUsrLot);

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
						#endregion Transferencia Bod Principal a PLATINIUM

						oTransaction.Commit();
						
						oConexion.Close();

						this.btnProcesar.Enabled = false;

						if(Proc1 == 1 && Proc2 == 1) MessageBox.Show("TRANSFERENCIA Procesada..!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
						else if(Proc1 == 1 && Proc2 == 0) MessageBox.Show("TRANSFERENCIA de LOFT Procesada..!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
						else if(Proc1 == 0 && Proc2 == 1) MessageBox.Show("TRANSFERENCIA de PLATINIUM Procesada..!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
						if(idCpaTransfer1>0)
						{
							string lcBskNTrn1 = "SELECT Codigo FROM CompraTransf WHERE idCompraTransf="+idCpaTransfer1.ToString();
							string NumTrn1 = Funcion.sEscalarSQL(cdsSeteo, lcBskNTrn1);
							this.txtTransIng.Value = NumTrn1.Trim();
							this.btnTfLf.Enabled = true ;
						}
						if(idCpaTransfer2>0)
						{
							string lcBskNTrn2 = "SELECT Codigo FROM CompraTransf WHERE idCompraTransf="+idCpaTransfer2.ToString();
							string NumTrn2 = Funcion.sEscalarSQL(cdsSeteo, lcBskNTrn2);
							this.txtTransSal.Value = NumTrn2.Trim();
							this.btnTfPL.Enabled = true ;
						}
					}
				}
			}
			#endregion Crea Transferencia
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCategoria.Focus();
		}

		private void cmbCategoria_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnBuscar_Click(sender, e);
		}

		private void frmTransferLotte_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
