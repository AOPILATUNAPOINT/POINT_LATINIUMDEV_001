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
	/// Descripción breve de frmReporteInventarioDetallado.
	/// </summary>
	public class frmReporteInventarioDetallado : System.Windows.Forms.Form
	{
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteInventarioDetallado()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteInventarioDetallado));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Inventario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_I");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadJustifica");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Egreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Inventario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado_I");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadJustifica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Egreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			this.cboFecha = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label55 = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.lblSeguimiento = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExcel = new System.Windows.Forms.Button();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btnImprimir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// cboFecha
			// 
			this.cboFecha.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboFecha.Location = new System.Drawing.Point(864, 9);
			this.cboFecha.MaxDropDownItems = 15;
			this.cboFecha.Name = "cboFecha";
			this.cboFecha.Size = new System.Drawing.Size(112, 21);
			this.cboFecha.TabIndex = 398;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(808, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 16);
			this.label1.TabIndex = 397;
			this.label1.Text = "Fecha";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.BackColor = System.Drawing.Color.Transparent;
			this.label55.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label55.Location = new System.Drawing.Point(24, 11);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(41, 17);
			this.label55.TabIndex = 393;
			this.label55.Text = "Bodega";
			this.label55.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label53.Location = new System.Drawing.Point(376, 11);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 17);
			this.label53.TabIndex = 392;
			this.label53.Text = "Fecha";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblSeguimiento
			// 
			this.lblSeguimiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblSeguimiento.AutoSize = true;
			this.lblSeguimiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSeguimiento.Location = new System.Drawing.Point(32, -200);
			this.lblSeguimiento.Name = "lblSeguimiento";
			this.lblSeguimiento.Size = new System.Drawing.Size(0, 17);
			this.lblSeguimiento.TabIndex = 396;
			this.lblSeguimiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(24, 44);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 395;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnVer
			// 
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(568, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 391;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(88, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(264, 22);
			this.cmbBodega.TabIndex = 388;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// dtFecha
			// 
			appearance1.FontData.Name = "Tahoma";
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(432, 9);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 389;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(-144, 80);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1296, 8);
			this.groupBox1.TabIndex = 639;
			this.groupBox1.TabStop = false;
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(896, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 640;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance2;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Agencia";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 200;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Fecha Inventario";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 110;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Estado del Conteo";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 120;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Articulo";
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 200;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance3;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.Header.Appearance = appearance4;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Estado";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Width = 120;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance5;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.Header.Appearance = appearance6;
			ultraGridColumn10.Header.Caption = "Cantidad. Cruce";
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Width = 200;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Width = 120;
			ultraGridColumn13.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Width = 120;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Width = 120;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Observaciones";
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Width = 250;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 11;
			ultraGridColumn16.Width = 130;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Hidden = true;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance7;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance10.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.uGridDocumentos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDocumentos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(8, 104);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(968, 240);
			this.uGridDocumentos.TabIndex = 641;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
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
																																 ultraDataColumn13});
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
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(808, 40);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 782;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Visible = false;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// frmReporteInventarioDetallado
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(992, 350);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.uGridDocumentos);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cboFecha);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.lblSeguimiento);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.Name = "frmReporteInventarioDetallado";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Inventario Detallado";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteInventarioDetallado_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteInventarioDetallado_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private System.Windows.Forms.ComboBox cboFecha;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label55;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label lblSeguimiento;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Button btnImprimir;

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmReporteInventarioDetallado_Load(object sender, System.EventArgs e)
		{
			string carpeta_1 ="";
			carpeta_1 = @"C:\Latinium\Reportes\ReporteEncabezadoAuditoria.rpt";
			if (!File.Exists(carpeta_1))
			{
				Process.Start("cmd.exe", "/C" + "net use \\192.168.1.10  /user:latinium 10dgpv15.p");
				Directory.GetDirectories(@"\\192.168.1.10\Proyectos");
				System.IO.File.Copy  (  @"\\192.168.1.10\Proyectos\Latinium\Reportes\ReporteEncabezadoAuditoria.rpt" ,@"C:\Latinium\Reportes\ReporteEncabezadoAuditoria.rpt", false);
			}

			miAcceso = new Acceso(cdsSeteoF, "200120");

			if (!Funcion.Permiso("2013", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Inventario Detallado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

      if (miAcceso.BExportar) this.btnExcel.Visible =  true;
			if (miAcceso.Imprimir) this.btnImprimir.Visible =  true;
			string sSQL = string.Format("Exec ListaBodegas {0}, 47", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			
			int iBodega = 0;

			sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodega > 0)
			{
				this.cmbBodega.Value = iBodega;
				this.cmbBodega.Enabled = false;
			}
			else this.cmbBodega.Value = 1;
									
			this.dtFecha.Value = DateTime.Today;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today.AddDays(2);
		}

		private void frmReporteInventarioDetallado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null) return;

			string sSQL = string.Format("Select Fecha From dbo.CierreConteoInventario Where Bodega = {0} And Estado = 2 Order By Fecha", (int)this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
			this.cboFecha.Items.Clear();

			DateTime Fecha = DateTime.Today;
			
			while (dr.Read())
			{
				Fecha = dr.GetDateTime(0);
				this.cboFecha.Items.Add(Fecha.ToString("dd/MM/yyyy"));
			}
			dr.Close();

			this.cboFecha.Text = Fecha.ToString("dd/MM/yyyy");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.uGridDocumentos);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Bodega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = {0} And CONVERT(DATE, Fecha) = '{1}'", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"))) == 0)
			{
				MessageBox.Show("No existe Conteo de esta Bodega en esta Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
		
			
			int iBodega = 0;				
						
			if (this.cmbBodega.ActiveRow != null) iBodega = (int) this.cmbBodega.Value;
		
			string sSQL = string.Format("Exec ReporteInventarioDetallado {0}, '{1}'", 
				iBodega, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDocumentos);				
	
			

			this.lblContador.Text = this.uGridDocumentos.Rows.Count + " ARTICULOS ENCONTRADOS";
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Bodega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From dbo.CierreConteoInventario Where Bodega = {0} And CONVERT(DATE, Fecha) = '{1}'", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"))) == 0)
			{
				MessageBox.Show("No existe Conteo de esta Bodega en esta Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
		
			DateTime dtFecha = (DateTime)this.dtFecha.Value;

			int iBodega = 0;				
						
			if (this.cmbBodega.ActiveRow != null) iBodega = (int) this.cmbBodega.Value;


			ParameterFields paramFields = new ParameterFields ();

					
			ParameterField pfCuotas = new ParameterField ();
			ParameterDiscreteValue ValCuotas = new ParameterDiscreteValue ();
			pfCuotas.ParameterFieldName = "@Bodega";
			ValCuotas.Value = iBodega;
			pfCuotas.CurrentValues.Add (ValCuotas);
			paramFields.Add (pfCuotas);
				
			ParameterField pfCapital = new ParameterField ();
			ParameterDiscreteValue ValCapital = new ParameterDiscreteValue ();
			pfCapital.ParameterFieldName = "@Fecha";
			ValCapital.Value = dtFecha;
			pfCapital.CurrentValues.Add (ValCapital);
			paramFields.Add (pfCapital);
		
			
			Reporte miRepor = new Reporte("ReporteEncabezadoAuditoria.rpt", "");
			miRepor.MdiParent = this.MdiParent;			
			miRepor.crVista.ParameterFieldInfo = paramFields;
			miRepor.Show();	

		}
	}
}
