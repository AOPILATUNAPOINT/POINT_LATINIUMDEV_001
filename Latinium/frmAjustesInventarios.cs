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
	/// Descripción breve de frmAjustesInventarios.
	/// </summary>
	public class frmAjustesInventarios : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnGrabar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugMinimos;
		public DevExpress.XtraEditors.SimpleButton btVer;
		public DevExpress.XtraEditors.SimpleButton btExportar;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCategoria;

		private C1DataRow drSeteoF;

		public frmAjustesInventarios()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAjustesInventarios));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Minimo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Minimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Categoria");
			this.btnGrabar = new System.Windows.Forms.Button();
			this.ugMinimos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btVer = new DevExpress.XtraEditors.SimpleButton();
			this.btExportar = new DevExpress.XtraEditors.SimpleButton();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.label1 = new System.Windows.Forms.Label();
			this.cmbCategoria = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.ugMinimos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoria)).BeginInit();
			this.SuspendLayout();
			// 
			// btnGrabar
			// 
			this.btnGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGrabar.CausesValidation = false;
			this.btnGrabar.Image = ((System.Drawing.Image)(resources.GetObject("btnGrabar.Image")));
			this.btnGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGrabar.Location = new System.Drawing.Point(168, 432);
			this.btnGrabar.Name = "btnGrabar";
			this.btnGrabar.Size = new System.Drawing.Size(72, 32);
			this.btnGrabar.TabIndex = 261;
			this.btnGrabar.Text = "Grabar";
			this.btnGrabar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGrabar.Click += new System.EventHandler(this.btnGrabar_Click);
			// 
			// ugMinimos
			// 
			this.ugMinimos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugMinimos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugMinimos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugMinimos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 200;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 540;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Format = "#,##0";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Format = "#,##0";
			ultraGridColumn6.Header.VisiblePosition = 5;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Format = "#,##0";
			ultraGridColumn7.Header.Caption = "C.AJST";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.ugMinimos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugMinimos.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.ugMinimos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugMinimos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugMinimos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ugMinimos.DisplayLayout.Override.CardAreaAppearance = appearance6;
			this.ugMinimos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugMinimos.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugMinimos.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugMinimos.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugMinimos.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ugMinimos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugMinimos.Location = new System.Drawing.Point(8, 56);
			this.ugMinimos.Name = "ugMinimos";
			this.ugMinimos.Size = new System.Drawing.Size(960, 368);
			this.ugMinimos.TabIndex = 260;
			this.ugMinimos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugMinimos_KeyDown);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Decimal);
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
			// btVer
			// 
			this.btVer.Appearance.Options.UseTextOptions = true;
			this.btVer.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btVer.Image = ((System.Drawing.Image)(resources.GetObject("btVer.Image")));
			this.btVer.ImageIndex = 4;
			this.btVer.Location = new System.Drawing.Point(312, 8);
			this.btVer.Name = "btVer";
			this.btVer.Size = new System.Drawing.Size(64, 32);
			this.btVer.TabIndex = 259;
			this.btVer.Text = "&Ver";
			this.btVer.Click += new System.EventHandler(this.btVer_Click);
			// 
			// btExportar
			// 
			this.btExportar.Appearance.Options.UseTextOptions = true;
			this.btExportar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btExportar.Image = ((System.Drawing.Image)(resources.GetObject("btExportar.Image")));
			this.btExportar.ImageIndex = 2;
			this.btExportar.Location = new System.Drawing.Point(240, 432);
			this.btExportar.Name = "btExportar";
			this.btExportar.Size = new System.Drawing.Size(72, 32);
			this.btExportar.TabIndex = 258;
			this.btExportar.Text = "&Exportar";
			this.btExportar.Click += new System.EventHandler(this.btExportar_Click);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBodega.Location = new System.Drawing.Point(136, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(42, 16);
			this.lblBodega.TabIndex = 257;
			this.lblBodega.Text = "LOCAL";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 30;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 30;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Width = 180;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Width = 30;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Codigo";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(136, 24);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(160, 21);
			this.cmbBodega.TabIndex = 256;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// label6
			// 
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(8, 432);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(40, 24);
			this.label6.TabIndex = 264;
			this.label6.Text = "Fecha PROC";
			// 
			// dtFecha
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance11;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(48, 434);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 263;
			this.dtFecha.Value = new System.DateTime(2013, 7, 1, 0, 0, 0, 0);
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
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(16, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(73, 16);
			this.label1.TabIndex = 275;
			this.label1.Text = "CATEGORIA";
			this.label1.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmbCategoria
			// 
			this.cmbCategoria.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbCategoria.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCategoria.DisplayMember = "Categoria";
			this.cmbCategoria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCategoria.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCategoria.Location = new System.Drawing.Point(16, 24);
			this.cmbCategoria.MaxDropDownItems = 30;
			this.cmbCategoria.Name = "cmbCategoria";
			this.cmbCategoria.Size = new System.Drawing.Size(112, 21);
			this.cmbCategoria.TabIndex = 274;
			this.cmbCategoria.ValueMember = "idCategoria";
			// 
			// frmAjustesInventarios
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(976, 469);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbCategoria);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btnGrabar);
			this.Controls.Add(this.ugMinimos);
			this.Controls.Add(this.btVer);
			this.Controls.Add(this.btExportar);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Name = "frmAjustesInventarios";
			this.Text = "Ajustar INVENTARIOS";
			this.Load += new System.EventHandler(this.frmAjustesInventarios_Load);
			((System.ComponentModel.ISupportInitialize)(this.ugMinimos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoria)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmAjustesInventarios_Load(object sender, System.EventArgs e)
		{
			string stExecBodega = string.Format("Select idBodega, Bodega, Codigo, Nombre From Bodega Where Activo=1 and Almacen=1");			
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExecBodega);
			this.dtFecha.Value = DateTime.Today;
			string stCateg = string.Format("Select idCategoria, Categoria From ArticuloCategoria Where idCategoria in (214,215)");
			this.cmbCategoria.DataSource = Funcion.dvProcedimiento(cdsSeteo, stCateg);
		}

		private void btVer_Click(object sender, System.EventArgs e)
		{
			int idCat = 0;
			if (this.cmbCategoria.ActiveRow != null) idCat = (int)this.cmbCategoria.Value;
			if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
			}
			else
			{
				string sSQL = string.Format("Exec ExistenciaArtAjuste {0},{1}", (int)this.cmbBodega.Value, idCat);
				this.ugMinimos.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);	
			}
		}

		private void ugMinimos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ugMinimos.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{
							#region Flechas
						case (int) Keys.Up:
							//							if (ultraGrid1.ActiveCell.DroppedDown) return;

							ugMinimos.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ugMinimos.PerformAction(UltraGridAction.AboveCell, false, false);
							e.Handled = true;
							ugMinimos.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Left:
							ugMinimos.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ugMinimos.PerformAction(UltraGridAction.PrevCellByTab, false, false);
							e.Handled = true;
							ugMinimos.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Down:
							//							if (ultraGrid1.ActiveCell.DroppedDown) return;
			
							ugMinimos.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ugMinimos.PerformAction(UltraGridAction.BelowCell, false, false);
							e.Handled = true;
							ugMinimos.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
							#endregion Flechas
							#region Enter
						case (int) Keys.Enter:
							ugMinimos.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ugMinimos.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							ugMinimos.PerformAction(UltraGridAction.EnterEditMode, false, false);
							//							if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
							//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
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

		private void btnGrabar_Click(object sender, System.EventArgs e)
		{
			/*Proceso para ajustar inventario*/
			if (DialogResult.Yes == MessageBox.Show(string.Format("¿Generar AJUSTE de Inventario de la Bodega '{0}'?", this.cmbBodega.Text.ToString()), "Aviso...!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				int idIngresoBodega = 0;
				int idEgresoBodega = 0;
				int ProcAjuIngr = 0;
				int ProcAjuEgrs = 0;
				int ProcesaAjInf = 0;
				int ProcesaInf = 0;

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

						#region Crea Ajuste de Ingreso
						DateTime dtFehaIngreso = (DateTime)this.dtFecha.Value;
						string sSQL = string.Format("Exec CICreaAjusteIngreso {0}, '{1}'", 
							(int)this.cmbBodega.Value, dtFehaIngreso.ToString("yyyyMMdd"));
						oCmdActualiza.CommandText = sSQL;
						idIngresoBodega = (int)oCmdActualiza.ExecuteScalar();

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugMinimos.Rows.All)
						{
							#region guarda datos ajuste
							if((int)dr.Cells["Cantidad"].Value>0)
							{
								decimal subt = 0.00m;
								subt = (int)dr.Cells["Cantidad"].Value*(decimal)dr.Cells["Costo"].Value;
								DateTime FeProceIng = (DateTime)this.dtFecha.Value;

								string sSQLDetalleIngreso = string.Format("Exec RegistraDetalleCpa_Lotte {0}, {1}, '{2}', '{3}', {4}, {5}, {6}, {7}", 
									idIngresoBodega, 
									(int)dr.Cells["idArticulo"].Value, 
									dr.Cells["Codigo"].Value, 
									dr.Cells["Articulo"].Value, 
									(int)dr.Cells["Cantidad"].Value, 
									(decimal)dr.Cells["Costo"].Value, 
									subt,
									(int)this.cmbBodega.Value);
								oCmdActualiza.CommandText = sSQLDetalleIngreso;
								oCmdActualiza.ExecuteNonQuery();

								int idart = (int)dr.Cells["idArticulo"].Value;
								int idBo = (int)dr.Cells["idBodega"].Value;
								int mini = (int)dr.Cells["Minimo"].Value;
								int sald = (int)dr.Cells["Saldo"].Value;
								int cant = (int)dr.Cells["Cantidad"].Value;

								string sSQLAjst = string.Format("Exec RegistraAjustesLotte '{0}', {1}, {2}, {3}, {4}, {5}, {6}",
									FeProceIng.ToString("yyyyMMdd"), idart, idBo, mini, sald, cant,	idIngresoBodega);
								oCmdActualiza.CommandText = sSQLAjst;
								oCmdActualiza.ExecuteNonQuery();
								ProcAjuIngr ++;
							}
							#endregion guarda datos ajuste
						}
						if (ProcAjuIngr==0)
						{
							string sSQLDelIn = string.Format("Delete From Compra Where idCompra="+idIngresoBodega.ToString());
							oCmdActualiza.CommandText = sSQLDelIn;
							oCmdActualiza.ExecuteNonQuery();
						}
						else ProcesaAjInf = 1;
						if(ProcesaAjInf == 1)
						{
							string sQlTtAjst = string.Format("Exec CITotalizaAjustes {0}", idIngresoBodega);
							oCmdActualiza.CommandText = sQlTtAjst;
							oCmdActualiza.ExecuteNonQuery();
							string sQlAsiAjst = string.Format("Exec CreaAsiIngreso {0}", idIngresoBodega);
							oCmdActualiza.CommandText = sQlAsiAjst;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Crea Ajuste de Ingreso

						#region Crea Ajuste de Egreso
						DateTime dtFehaEgreso = (DateTime)this.dtFecha.Value;
						string sQlSalida = string.Format("Exec CICreaAjusteEgreso {0}, '{1}'", (int)this.cmbBodega.Value, dtFehaIngreso.ToString("yyyyMMdd"));
						oCmdActualiza.CommandText = sQlSalida;
						idEgresoBodega = (int)oCmdActualiza.ExecuteScalar();

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dreg in this.ugMinimos.Rows.All)
						{
							#region guarda datos ajuste
							if((int)dreg.Cells["Cantidad"].Value<0)
							{
								decimal subt = 0.00m;
								subt = ((int)dreg.Cells["Cantidad"].Value*-1)*(decimal)dreg.Cells["Costo"].Value;
								DateTime FeProceEgr = (DateTime)this.dtFecha.Value;

								string sSQLDetalleEgreso = string.Format("Exec RegistraDetalleCpa_Lotte {0}, {1}, '{2}', '{3}', {4}, {5}, {6}, {7}", 
									idEgresoBodega, 
									(int)dreg.Cells["idArticulo"].Value, 
									dreg.Cells["Codigo"].Value, 
									dreg.Cells["Articulo"].Value, 
									(int)dreg.Cells["Cantidad"].Value*-1, 
									(decimal)dreg.Cells["Costo"].Value, 
									subt,
									(int)this.cmbBodega.Value);
								oCmdActualiza.CommandText = sSQLDetalleEgreso;
								oCmdActualiza.ExecuteNonQuery();

								string sSQLAjEg = string.Format("Exec RegistraAjustesLotte '{0}', {1}, {2}, {3}, {4}, {5}, {6}",
									FeProceEgr.ToString("yyyyMMdd"),
									(int)dreg.Cells["idArticulo"].Value, 
									(int)dreg.Cells["idBodega"].Value, 
									(int)dreg.Cells["Minimo"].Value, 
									(int)dreg.Cells["Saldo"].Value, 
									(int)dreg.Cells["Cantidad"].Value,
									idEgresoBodega);
								oCmdActualiza.CommandText = sSQLAjEg;
								oCmdActualiza.ExecuteNonQuery();
								ProcAjuEgrs ++;
							}
							#endregion guarda datos ajuste
						}

						/*Sino hay procesos realizados eliminar ajuste de COMPRA*/
						if (ProcAjuEgrs==0)
						{
							string sSQLDelEg = string.Format("Delete From Compra Where idCompra="+idEgresoBodega.ToString());
							oCmdActualiza.CommandText = sSQLDelEg;
							oCmdActualiza.ExecuteNonQuery();
						}
						else ProcesaInf = 1;
						if(ProcesaInf == 1)
						{
							string sQlTtlAjSal = string.Format("Exec CITotalizaAjustes {0}", idEgresoBodega);
							oCmdActualiza.CommandText = sQlTtlAjSal;
							oCmdActualiza.ExecuteNonQuery();
							string sQlAsiSal = string.Format("Exec CreaAsiEgreso {0}", idEgresoBodega);
							oCmdActualiza.CommandText = sQlAsiSal;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Crea Ajuste de Egreso
						// Visuliza información
						MessageBox.Show("Información PROCESADA correctamente...!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.btnGrabar.Enabled = false;
						oTransaction.Commit();
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
							oTransaction.Rollback();

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

		private void btExportar_Click(object sender, System.EventArgs e)
		{
			if (this.ugMinimos.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.ugMinimos, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}
	}
}
