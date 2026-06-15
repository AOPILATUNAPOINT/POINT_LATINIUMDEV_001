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
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteAutorizacionTransporte.
	/// </summary>
	public class frmReporteAutorizacionTransporte : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago_1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnExportar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteAutorizacionTransporte()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Chofer");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbFormaPago_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRegistroTransporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aceptar", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteAutorizacionTransporte));
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anular", 1);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Imprimir", 2);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Chofer");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRegistroTransporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlNube");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jefe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jefe");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbFormaPago_1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnExportar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago_1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Width = 70;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Width = 110;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 5;
			ultraGridColumn3.Width = 150;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 7;
			ultraGridColumn4.Width = 180;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 110;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 9;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Width = 100;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 13;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn9.Header.Caption = "Jefe / SubJefe";
			ultraGridColumn9.Header.VisiblePosition = 12;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 250;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			appearance3.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellButtonAppearance = appearance3;
			ultraGridColumn12.Header.Caption = "Aprobar";
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn12.Width = 30;
			ultraGridColumn13.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			appearance4.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellButtonAppearance = appearance4;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn13.Width = 45;
			ultraGridColumn14.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			appearance5.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellButtonAppearance = appearance5;
			ultraGridColumn14.Header.Caption = "Ver";
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn14.Width = 30;
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
																										 ultraGridColumn14});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance6.ForeColor = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.CaptionAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 152);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1184, 256);
			this.ultraGrid1.TabIndex = 336;
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
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
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1224, 8);
			this.groupBox1.TabIndex = 335;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(464, 104);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 334;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(360, 104);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 332;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(24, 107);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(39, 16);
			this.lblCliente.TabIndex = 333;
			this.lblCliente.Text = "Buscar";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(208, 8);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 329;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(24, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 328;
			this.label6.Text = "Desde";
			// 
			// txtBusqueda
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance12;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(80, 105);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(272, 21);
			this.txtBusqueda.TabIndex = 331;
			// 
			// dtHasta
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance13;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 327;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.ValueChanged += new System.EventHandler(this.dtHasta_ValueChanged);
			// 
			// dtDesde
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance14;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(80, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 326;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1120, 104);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1011;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// cmbFormaPago
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance15;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource1;
			ultraGridColumn15.Header.Caption = "EstadoA";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 8;
			ultraGridColumn16.Header.Caption = "Estado";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 174;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormaPago.DisplayMember = "Nombre";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(80, 72);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(176, 21);
			this.cmbFormaPago.TabIndex = 1012;
			this.cmbFormaPago.ValueMember = "Estado";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 1013;
			this.label1.Text = "Estado";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(24, 40);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 1015;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			appearance16.FontData.Name = "Tahoma";
			this.cmbBodega.Appearance = appearance16;
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 180;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(80, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 1014;
			this.cmbBodega.ValueMember = "Bodega";
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
			// cmbFormaPago_1
			// 
			this.cmbFormaPago_1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago_1.DataSource = this.ultraDataSource3;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Width = 250;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbFormaPago_1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbFormaPago_1.DisplayMember = "Jefe";
			this.cmbFormaPago_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago_1.Location = new System.Drawing.Point(360, 195);
			this.cmbFormaPago_1.Name = "cmbFormaPago_1";
			this.cmbFormaPago_1.Size = new System.Drawing.Size(272, 88);
			this.cmbFormaPago_1.TabIndex = 1016;
			this.cmbFormaPago_1.ValueMember = "idPersonal";
			this.cmbFormaPago_1.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn14.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn14,
																																 ultraDataColumn15});
			// 
			// dtFecha
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance17;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton3);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(384, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(136, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1017;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.Visible = false;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1120, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 24);
			this.btnExportar.TabIndex = 1018;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Visible = false;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// frmReporteAutorizacionTransporte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1200, 414);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbFormaPago_1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.KeyPreview = true;
			this.Name = "frmReporteAutorizacionTransporte";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Autorización Voucher Transporte";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteAutorizacionTransporte_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteAutorizacionTransporte_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago_1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close ();
		}
		int txtBodegaPredefinida = 0;
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void frmReporteAutorizacionTransporte_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "201103");

			if (!Funcion.Permiso("1822", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a autorización de Transporte", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExportar.Visible = true;
			if (!miAcceso.Autorizacion) this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
			if (!miAcceso.Buscar) this.ultraGrid1.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = true;
			if (!miAcceso.Anular) this.ultraGrid1.DisplayLayout.Bands[0].Columns["Anular"].Hidden = true;


			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("20/10/2022");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("20/10/2022");			

			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			txtBodegaPredefinida = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (txtBodegaPredefinida == 0) this.cmbBodega.Value = 1;
			else this.cmbBodega.Value = txtBodegaPredefinida;

			if (txtBodegaPredefinida == 0) this.cmbBodega.Enabled = true;
			else this.cmbBodega.Enabled = false;

			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format(" select Estado, Nombre from dbo.EstadoRegistroTransporte"));
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);

		}

		private void dtHasta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtHasta.Value != System.DBNull.Value)
			{
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 1, (DateTime)this.dtHasta.Value, false, cdsSeteoF);

				this.cmbFormaPago_1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format(" Exec [ConsultaJefeTransporte] '{0}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")));
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}
			#endregion Valida Fechas

			int iLocal = 0;
			int iEstado = 0;
			if (this.cmbBodega.ActiveRow != null) iLocal = (int)this.cmbBodega.Value;
			if (this.cmbFormaPago.ActiveRow != null) iEstado = (int)this.cmbFormaPago.Value;

			this.lblContador.Text =  "";

				string sSQL = string.Format("Exec [ConsultaAutorizacionTransporte] '{0}', '{1}', {2}, {3}, '{4}'",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),	Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), iLocal, iEstado,  this.txtBusqueda.Text ) ;			
				FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);

			
			this.lblContador.Text =  this.ultraGrid1.Rows.Count + " Registros Encontrados";

		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Imprimir")	
			{

				if (!miAcceso.Buscar)
				{
					MessageBox.Show("No tiene Acceso a Visualizar el Archivo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				#region Bube
				if ((int)e.Cell.Row.Cells["iEstado"].Value < 2 )
				{
					MessageBox.Show("El Archivo PDF no esta subido.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				
					return;
				}
				if (Convert.ToString( e.Cell.Row.Cells["UrlNube"].Value).Length < 10 )
				{
					MessageBox.Show("El Archivo PDF no esta subido.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					
					return;
				}
				System.Diagnostics.Process.Start("chrome.exe", Convert.ToString( e.Cell.Row.Cells["UrlNube"].Value));
				#endregion Bube
			}
			if (e.Cell.Column.ToString() == "Anular")	
			{

				if (!miAcceso.Anular)
				{
					MessageBox.Show("No tiene Acceso Anular.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#region Bube
				if ((int)e.Cell.Row.Cells["iEstado"].Value == 3 )
				{
					MessageBox.Show("El voucher esta anulado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					
					return;
				}
				if ((int)e.Cell.Row.Cells["iEstado"].Value >=4 )
				{
					MessageBox.Show("El voucher ya esta Autorizado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					
					return;
				}
				if (DialogResult.Yes == MessageBox.Show("Desea Anular el Deposito ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					using (frmNotas Notas  = new frmNotas())
					{
						if (DialogResult.OK == Notas.ShowDialog())
						{
							string sSQL = string.Format("Update RegistroTransporte set NotaAnula = '{0}', Estado = 3, FechaAnula = GETDATE (), UsuarioAnula = suser_sname() where idRegistroTransporte = {1}", Notas.txtNotas.Text, 
								(int)e.Cell.Row.Cells["idRegistroTransporte"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL);
							this.btnBuscar_Click(sender,e);
						}
					}
				}
				#endregion Bube
			}

			if (e.Cell.Column.ToString() == "Aceptar")	
			{
				if (!miAcceso.Autorizacion)
				{
					MessageBox.Show("No tiene Acceso Aprobar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
	
				#region Bube
				if ((int)e.Cell.Row.Cells["iEstado"].Value < 2 )
				{
					MessageBox.Show("El voucher esta en estado Pendiente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				
					return;
				}
				if ((int)e.Cell.Row.Cells["iEstado"].Value == 3 )
				{
					MessageBox.Show("El voucher esta anulado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			
					return;
				}
				if ((int)e.Cell.Row.Cells["iEstado"].Value > 3 )
				{
					MessageBox.Show("El voucher ya esta Autorizado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
		
					return;
				}
				#endregion Bube

				if ((int)e.Cell.Row.Cells["iEstado"].Value == 2 )
				{
					if ((int)e.Cell.Row.Cells["idPersonal"].Value == 0 )
					{
						MessageBox.Show("Seleccione el Jefe o SubJefe de Almacen.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						
						return;
					}
					if (e.Cell.Row.Cells["idPersonal"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione el Jefe o SubJefe de Almacen.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						
						return;
					}	
					if (DialogResult.Yes == MessageBox.Show("Desea Autorizar el Voucher ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					{
						#region Guardar
						using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
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

								#region Autorizacion de pago

								string sSQLAut = string.Format("Exec CreaAutorizacionesDePago {0}, {1}, '{2}', '{3}', {4}, {5}, '1 DE 1', '{6}', {7}, 0", 
									26, (int)e.Cell.Row.Cells["idRegistroTransporte"].Value, 
									Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
									Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
									Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value), 7, 
									"VOUCHER TRANSPORTE" + " " + Convert.ToString(e.Cell.Row.Cells["Numero"].Value), 0);			
								oCmdActualiza.CommandText = sSQLAut;
								oCmdActualiza.ExecuteNonQuery();		

								string sSQLAsiEs = string.Format("update RegistroTransporte set Estado = 4, idPersonal = {1} where idRegistroTransporte = {0}", 
									(int)e.Cell.Row.Cells["idRegistroTransporte"].Value, (int)e.Cell.Row.Cells["idPersonal"].Value);			
								oCmdActualiza.CommandText = sSQLAsiEs;
								oCmdActualiza.ExecuteNonQuery();
				
								#endregion Autorizacion de pago

								oTransaction.Commit();

								Cursor = Cursors.Default;	

								MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.btnBuscar_Click (sender, e);
					
							}
							catch(Exception ex)
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
				
						#endregion Guardar
					}
				}
			}	
		}

		private void frmReporteAutorizacionTransporte_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();	
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}
	}
}
