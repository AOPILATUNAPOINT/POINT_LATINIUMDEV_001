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
	/// Descripción breve de frmConfiguracionGEXyCompraProtegida.
	/// </summary>
	public class frmConfiguracionGEXyCompraProtegida : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridComisiones;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSubGrupo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoGrupo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnGuardar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConfiguracionGEXyCompraProtegida()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idConfiguracionGEXyCompraProtegida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoProducto", -1, "cmbTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubgrupo", -1, "cmbSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoInicial");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoFinal");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tamaño_Precio");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioCredito");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioContado");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idConfiguracionGEXyCompraProtegida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoProducto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubgrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RangoInicial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RangoFinal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tamaño_Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioContado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubgrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subgrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubgrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subgrupo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoProducto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoProducto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConfiguracionGEXyCompraProtegida));
			this.uGridComisiones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipoGrupo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.uGridComisiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridComisiones
			// 
			this.uGridComisiones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridComisiones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridComisiones.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridComisiones.DisplayLayout.Appearance = appearance1;
			this.uGridComisiones.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "Serie";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.Caption = "Producto";
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn2.Width = 160;
			ultraGridColumn3.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.Header.Caption = "Grupo De Artículos";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn3.Width = 150;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "Rango Inicial";
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 80;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "Rango Final";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 80;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Tamaño/Precio";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 65;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Precio Credito";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 90;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Precio Contado";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 90;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance7;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellButtonAppearance = appearance8;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 80;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridComisiones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridComisiones.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.uGridComisiones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridComisiones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridComisiones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridComisiones.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 7.75F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridComisiones.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.uGridComisiones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridComisiones.DisplayLayout.Override.RowAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridComisiones.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			this.uGridComisiones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridComisiones.Location = new System.Drawing.Point(8, 56);
			this.uGridComisiones.Name = "uGridComisiones";
			this.uGridComisiones.Size = new System.Drawing.Size(816, 404);
			this.uGridComisiones.TabIndex = 738;
			this.uGridComisiones.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridComisiones_AfterRowInsert);
			this.uGridComisiones.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridComisiones_BeforeRowsDeleted);
			this.uGridComisiones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridComisiones_KeyDown);
			this.uGridComisiones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridComisiones_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn6.DefaultValue = false;
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
			// cmbSubGrupo
			// 
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataSource = this.ultraDataSource1;
			this.cmbSubGrupo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 119;
			ultraGridColumn11.Header.Caption = "Grupo";
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 278;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSubGrupo.DisplayMember = "Subgrupo";
			this.cmbSubGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(136, 168);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(280, 112);
			this.cmbSubGrupo.TabIndex = 739;
			this.cmbSubGrupo.ValueMember = "idSubgrupo";
			this.cmbSubGrupo.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn10.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// cmbTipoGrupo
			// 
			this.cmbTipoGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoGrupo.DataSource = this.ultraDataSource3;
			this.cmbTipoGrupo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 127;
			ultraGridColumn13.Header.Caption = "Tipo Producto";
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 278;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbTipoGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipoGrupo.DisplayMember = "Descripcion";
			this.cmbTipoGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoGrupo.Location = new System.Drawing.Point(432, 168);
			this.cmbTipoGrupo.Name = "cmbTipoGrupo";
			this.cmbTipoGrupo.Size = new System.Drawing.Size(280, 112);
			this.cmbTipoGrupo.TabIndex = 740;
			this.cmbTipoGrupo.ValueMember = "TipoProducto";
			this.cmbTipoGrupo.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(736, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(90, 24);
			this.btnExcel.TabIndex = 957;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(640, 8);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(90, 24);
			this.btnGuardar.TabIndex = 958;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
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
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(832, 8);
			this.groupBox1.TabIndex = 959;
			this.groupBox1.TabStop = false;
			// 
			// frmConfiguracionGEXyCompraProtegida
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(832, 466);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.cmbTipoGrupo);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.uGridComisiones);
			this.KeyPreview = true;
			this.Name = "frmConfiguracionGEXyCompraProtegida";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Configuración Garantía Extendida - Compra Protegida";
			this.Load += new System.EventHandler(this.frmConfiguracionGEXyCompraProtegida_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridComisiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmConfiguracionGEXyCompraProtegida_Load(object sender, System.EventArgs e)
		{

			this.cmbSubGrupo.DataSource = FuncionesProcedimientos.dtGeneral("Select idSubgrupo, SubGrupo From ArticuloSubGrupo Order By SubGrupo");
			this.cmbTipoGrupo.DataSource = FuncionesProcedimientos.dtGeneral("Select TipoProducto, Descripcion From TipoProducto ");
			string sSQLPa = string.Format("Exec ConsultaConfiguracionGEXyCompraProtegida ");
			this.uGridComisiones.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridComisiones);
		}

		private void uGridComisiones_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idConfiguracionGEXyCompraProtegida"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idConfiguracionGEXyCompraProtegida"].Value > 0)
						{
							string sSQL = string.Format("Delete From ConfiguracionGEXyCompraProtegida Where idConfiguracionGEXyCompraProtegida = {0}", (int)e.Rows[i].Cells["idConfiguracionGEXyCompraProtegida"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
							e.DisplayPromptMsg = false;
						}
						else e.DisplayPromptMsg = false;
					}		
				}
			}		
			else e.Cancel = true;
		}

		private void uGridComisiones_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			
			e.Row.Cells["idConfiguracionGEXyCompraProtegida"].Value = 0;
			e.Row.Cells["TipoProducto"].Value = 0;
			e.Row.Cells["idSubgrupo"].Value = 0;

			e.Row.Cells["RangoInicial"].Value = 0.00m;  
			e.Row.Cells["RangoFinal"].Value = 0.00m; 	
			e.Row.Cells["Tamaño_Precio"].Value = false; 

			e.Row.Cells["PrecioCredito"].Value = 0.00m; 
			e.Row.Cells["PrecioContado"].Value = 0.00m; 	
				e.Row.Cells["Costo"].Value = 0.00m; 	

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void uGridComisiones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.uGridComisiones);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Valida Comisiones
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridComisiones.Rows.All)
			{
				#region Porcentajes Desde C
				if (Convert.ToDecimal(dr.Cells["RangoInicial"].Value) > Convert.ToDecimal(dr.Cells["RangoFinal"].Value))
				{
					MessageBox.Show("En Porcentajes, el Rango Inicial no puede ser mayor al Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridComisiones.ActiveRow = dr;
					this.uGridComisiones.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final
					
				#region Cargo
				if (dr.Cells["TipoProducto"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Tipo de Producto ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridComisiones.ActiveRow = dr;
					this.uGridComisiones.ActiveRow.Selected = true;

					return;
				}
				if ((int)dr.Cells["TipoProducto"].Value == 0)
				{
					MessageBox.Show("Seleccione un Tipo de Producto ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridComisiones.ActiveRow = dr;
					this.uGridComisiones.ActiveRow.Selected = true;

					return;
				}

				if (dr.Cells["idSubgrupo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione el Grupo ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridComisiones.ActiveRow = dr;
					this.uGridComisiones.ActiveRow.Selected = true;

					return;
				}
				if ((int)dr.Cells["idSubgrupo"].Value == 0)
				{
					MessageBox.Show("Seleccione el Grupo ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridComisiones.ActiveRow = dr;
					this.uGridComisiones.ActiveRow.Selected = true;

					return;
				}

				#endregion Cargo


				#region Porcentajes repetidos
//				string sSQL = string.Format("Exec Cob_ValidaPorcentajes {0}, {1}, {2}, {3}", 
//					(int)dr.Cells["idConfiguracionGEXyCompraProtegida"].Value, (int)dr.Cells["idSubgrupo"].Value,
//					 Convert.ToDecimal(dr.Cells["RangoInicial"].Value), 
//					 Convert.ToDecimal(dr.Cells["RangoFinal"].Value));
//
//				if (Funcion.bEscalarSQL(cdsSeteoF, sSQL, true))
//				{
//					MessageBox.Show("Esta ingresando valores repetidos en Porcentajes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.uGridComisiones.ActiveRow = dr;
//					this.uGridComisiones.ActiveRow.Selected = true;
//
//					return;
//				}
				#endregion Porcentajes repetidos				
		
			}
				
			#endregion Valida Comisiones

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
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridComisiones.Rows.All)
						{																											 
							string sSQLDetalle = string.Format("Exec GuardaConfiguracionGEXyCompraProtegida {0}, {1}, {2}, {3}, {4}, {5}, {6} , {7}, {8}",
								(int)dr.Cells["idConfiguracionGEXyCompraProtegida"].Value, 
								(int)dr.Cells["TipoProducto"].Value, 
								(int)dr.Cells["idSubgrupo"].Value, 
								Convert.ToDecimal (dr.Cells["RangoInicial"].Value), 
								Convert.ToDecimal (dr.Cells["RangoFinal"].Value), 
								Convert.ToBoolean (dr.Cells["Tamaño_Precio"].Value), 
								Convert.ToDecimal (dr.Cells["PrecioCredito"].Value), 
								Convert.ToDecimal (dr.Cells["PrecioContado"].Value ),
								Convert.ToDecimal (dr.Cells["Costo"].Value ));
		
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Detalle	      
				
						oTransaction.Commit();

						string sSQLPa = string.Format("Exec ConsultaConfiguracionGEXyCompraProtegida ");
						this.uGridComisiones.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
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

		private void uGridComisiones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "RangoInicial", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "RangoFinal", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "PrecioCredito", 2);	
			FuncionesProcedimientos.FormatoGrid(e, "PrecioContado", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "Costo", 2);	
		}
	}
}
