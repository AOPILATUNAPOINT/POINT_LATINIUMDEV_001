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
	/// Descripción breve de frmHabitaciones.
	/// </summary>
	public class frmHabitaciones : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnGrabar;
		private System.Windows.Forms.Button btnExportar;
		private C1.Data.C1DataSet cdsBodega;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugHabitaciones;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTematica;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSucursal;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.ComponentModel.IContainer components;

		public frmHabitaciones()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmHabitaciones));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTematica", -1, "cmbTematica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroHabitacion");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal", -1, "cmbSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HabitacionNormal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HabitacionYacuzzi");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HabitacionLoft");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acc_Silla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acc_Tubo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acc_Columbio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acc_Chimenea");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioNormal");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioReserva");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTematica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSucursal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HabitacionNormal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HabitacionYacuzzi");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HabitacionLoft");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Acc_Silla");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Acc_Tubo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Acc_Columbio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Acc_Chimenea");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioNormal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioReserva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTematica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTematica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.btnGrabar = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.cdsBodega = new C1.Data.C1DataSet();
			this.ugHabitaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTematica = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSucursal = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugHabitaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTematica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			this.SuspendLayout();
			// 
			// btnGrabar
			// 
			this.btnGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGrabar.CausesValidation = false;
			this.btnGrabar.Image = ((System.Drawing.Image)(resources.GetObject("btnGrabar.Image")));
			this.btnGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGrabar.Location = new System.Drawing.Point(16, 408);
			this.btnGrabar.Name = "btnGrabar";
			this.btnGrabar.Size = new System.Drawing.Size(72, 23);
			this.btnGrabar.TabIndex = 16;
			this.btnGrabar.Text = "Grabar";
			this.btnGrabar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGrabar.Click += new System.EventHandler(this.btnGrabar_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(88, 408);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 23);
			this.btnExportar.TabIndex = 17;
			this.btnExportar.Text = "Excel";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// cdsBodega
			// 
			this.cdsBodega.BindingContextCtrl = this;
			this.cdsBodega.DataLibrary = "LibFacturacion";
			this.cdsBodega.DataLibraryUrl = "";
			this.cdsBodega.DataSetDef = "dsBodega";
			this.cdsBodega.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsBodega.Name = "cdsBodega";
			this.cdsBodega.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsBodega.SchemaDef = null;
			this.cdsBodega.BeforeFill += new C1.Data.FillEventHandler(this.cdsBodega_BeforeFill);
			// 
			// ugHabitaciones
			// 
			this.ugHabitaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugHabitaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugHabitaciones.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugHabitaciones.DisplayLayout.Appearance = appearance1;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn1.CellAppearance = appearance2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "TEMATICA";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn3.CellAppearance = appearance3;
			ultraGridColumn3.Header.Caption = "# HABITACION";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 30;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.Header.Caption = "NOMBRE";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 180;
			ultraGridColumn5.Header.Caption = "SUCURSAL";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 200;
			ultraGridColumn6.Header.Caption = "N";
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn6.Width = 20;
			ultraGridColumn7.Header.Caption = "Y";
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn7.Width = 20;
			ultraGridColumn8.Header.Caption = "L";
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn8.Width = 20;
			ultraGridColumn9.Header.Caption = "S";
			ultraGridColumn9.Header.VisiblePosition = 10;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn9.Width = 20;
			ultraGridColumn10.Header.Caption = "T";
			ultraGridColumn10.Header.VisiblePosition = 11;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn10.Width = 20;
			ultraGridColumn11.Header.Caption = "C";
			ultraGridColumn11.Header.VisiblePosition = 12;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn11.Width = 20;
			ultraGridColumn12.Header.Caption = "H";
			ultraGridColumn12.Header.VisiblePosition = 13;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn12.Width = 20;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance4;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "PRECIO";
			ultraGridColumn13.Header.VisiblePosition = 5;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance5;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "RESERVA";
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn15.Header.Caption = "ACTIVO";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn15.Width = 20;
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
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ugHabitaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugHabitaciones.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.ugHabitaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugHabitaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugHabitaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ugHabitaciones.DisplayLayout.Override.CardAreaAppearance = appearance7;
			this.ugHabitaciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugHabitaciones.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.ugHabitaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugHabitaciones.Location = new System.Drawing.Point(0, 0);
			this.ugHabitaciones.Name = "ugHabitaciones";
			this.ugHabitaciones.Size = new System.Drawing.Size(840, 400);
			this.ugHabitaciones.TabIndex = 18;
			this.ugHabitaciones.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ugHabitaciones_AfterRowInsert);
			this.ugHabitaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugHabitaciones_KeyDown);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn8.DataType = typeof(bool);
			ultraDataColumn9.DataType = typeof(bool);
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn11.DataType = typeof(bool);
			ultraDataColumn12.DataType = typeof(bool);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(bool);
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
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15});
			// 
			// cmbTematica
			// 
			this.cmbTematica.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTematica.DataSource = this.ultraDataSource2;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 177;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbTematica.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTematica.DisplayMember = "Nombre";
			this.cmbTematica.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTematica.Location = new System.Drawing.Point(40, 192);
			this.cmbTematica.Name = "cmbTematica";
			this.cmbTematica.Size = new System.Drawing.Size(200, 86);
			this.cmbTematica.TabIndex = 19;
			this.cmbTematica.ValueMember = "idTematica";
			this.cmbTematica.Visible = false;
			// 
			// ultraDataSource2
			// 
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			this.ultraDataSource2.Band.Key = "CompraNumero";
			// 
			// cmbSucursal
			// 
			this.cmbSucursal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSucursal.DataSource = this.ultraDataSource3;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 177;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbSucursal.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbSucursal.DisplayMember = "Nombre";
			this.cmbSucursal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSucursal.Location = new System.Drawing.Point(248, 192);
			this.cmbSucursal.Name = "cmbSucursal";
			this.cmbSucursal.Size = new System.Drawing.Size(200, 86);
			this.cmbSucursal.TabIndex = 20;
			this.cmbSucursal.ValueMember = "idBodega";
			this.cmbSucursal.Visible = false;
			// 
			// ultraDataSource3
			// 
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			this.ultraDataSource3.Band.Key = "CompraNumero";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(520, 416);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(62, 16);
			this.label1.TabIndex = 21;
			this.label1.Text = "N = Normal";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(520, 432);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(64, 16);
			this.label2.TabIndex = 22;
			this.label2.Text = "Y = Yacuzzi";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(592, 416);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(42, 16);
			this.label3.TabIndex = 23;
			this.label3.Text = "L = Loft";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(592, 432);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(46, 16);
			this.label4.TabIndex = 24;
			this.label4.Text = "S = Silla";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(640, 416);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(50, 16);
			this.label5.TabIndex = 25;
			this.label5.Text = "T = Tubo";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(640, 432);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(73, 16);
			this.label6.TabIndex = 26;
			this.label6.Text = "C = Columbio";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(720, 416);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(77, 16);
			this.label7.TabIndex = 27;
			this.label7.Text = "H = Chimenea";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(720, 432);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(55, 16);
			this.label8.TabIndex = 28;
			this.label8.Text = "A = Activo";
			// 
			// frmHabitaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(840, 453);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbSucursal);
			this.Controls.Add(this.cmbTematica);
			this.Controls.Add(this.ugHabitaciones);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnGrabar);
			this.Name = "frmHabitaciones";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Registro de Habitaciones TANTRA";
			this.Load += new System.EventHandler(this.frmHabitaciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugHabitaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTematica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsBodega_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsBodega.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void VisulizaHabitaciones()
		{
			string lcQry = "Select idHabitacion,idTematica,NumeroHabitacion,Nombre,idSucursal,PrecioNormal,PrecioReserva,HabitacionNormal,";
			lcQry += "HabitacionYacuzzi,HabitacionLoft,Acc_Silla,Acc_Tubo,Acc_Columbio,Acc_Chimenea,Estado ";
			lcQry += "From Habitaciones order By idSucursal,NumeroHabitacion";
		  this.ugHabitaciones.DataSource = Funcion.dvProcedimiento(cdsBodega, lcQry);
		}
		private void frmHabitaciones_Load(object sender, System.EventArgs e)
		{
			this.cmbTematica.DataSource = Funcion.dvProcedimiento(cdsBodega, "Select idTematica, Nombre From Tematica");
			this.cmbSucursal.DataSource = Funcion.dvProcedimiento(cdsBodega, "Select idBodega, Nombre, Codigo From Bodega Where Activo=1 and Almacen=1");
			VisulizaHabitaciones();
		}

		private void ugHabitaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ugHabitaciones.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{
							#region Enter
						case (int) Keys.Enter:
							ugHabitaciones.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ugHabitaciones.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							ugHabitaciones.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
							#endregion Enter
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error del sistema.");
				}		
			}
		}

		private void ugHabitaciones_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idHabitacion"].Value = 0;
			e.Row.Cells["idTematica"].Value = 0; 
			e.Row.Cells["NumeroHabitacion"].Value = 0; 
			e.Row.Cells["Nombre"].Value = "";
			e.Row.Cells["idSucursal"].Value = 0;
			e.Row.Cells["HabitacionNormal"].Value = false;
			e.Row.Cells["HabitacionYacuzzi"].Value = false;
			e.Row.Cells["HabitacionLoft"].Value = false;
			e.Row.Cells["Acc_Silla"].Value = false;
			e.Row.Cells["Acc_Tubo"].Value = false;
			e.Row.Cells["Acc_Columbio"].Value = false;
			e.Row.Cells["Acc_Chimenea"].Value = false;
			e.Row.Cells["Estado"].Value = false;
			e.Row.Cells["PrecioNormal"].Value = 0;
			e.Row.Cells["PrecioReserva"].Value = 0;
		}

		private void btnGrabar_Click(object sender, System.EventArgs e)
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

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugHabitaciones.Rows.All)
					{
						string sSQLDetalle = string.Format("Exec GrabaHabitaciones {0}, {1}, {2}, '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}",
							(int)dr.Cells["idHabitacion"].Value, 
							(int)dr.Cells["idTematica"].Value, 
							(int)dr.Cells["NumeroHabitacion"].Value, 
							dr.Cells["Nombre"].Value.ToString(),
							(int)dr.Cells["idSucursal"].Value,
							(bool)dr.Cells["HabitacionNormal"].Value, 
							(bool)dr.Cells["HabitacionYacuzzi"].Value, 
							(bool)dr.Cells["HabitacionLoft"].Value, 
							(bool)dr.Cells["Acc_Silla"].Value,
							(bool)dr.Cells["Acc_Tubo"].Value,
							(bool)dr.Cells["Acc_Columbio"].Value,
							(bool)dr.Cells["Acc_Chimenea"].Value,
							(decimal)dr.Cells["PrecioNormal"].Value, 
							(decimal)dr.Cells["PrecioReserva"].Value,
							(bool)dr.Cells["Estado"].Value);
						oCmdActualiza.CommandText = sSQLDetalle;
						oCmdActualiza.ExecuteNonQuery();
					}
          oTransaction.Commit();
				}
				catch (Exception ex)
				{
					try
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
			// Visuliza información
      MessageBox.Show("Información grabada correctamente...!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			VisulizaHabitaciones();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.ugHabitaciones.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.ugHabitaciones, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}
	}
}
