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

namespace Latinium
{
	/// <summary>
	/// Summary description for frmDptoTecnico.
	/// </summary>
	public class frmDptoTecnico : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPersonal;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.Misc.UltraButton btTransferir;
		private Infragistics.Win.Misc.UltraButton btSalir;
		private Infragistics.Win.Misc.UltraButton btnConsultar;
		private System.Windows.Forms.ComboBox cmbEstados;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmDptoTecnico()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonalOrden");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenProd");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Inicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Terminado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonalOrden");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrdenProd");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Inicio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Terminado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDptoTecnico));
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProduccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonalOrden");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenProd");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstCase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SopProc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SopMemoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SopLecTar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SopCDDVD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SopHDD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BusDeDatos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistOpe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ofimatica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Otros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DetEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnboundColumn1", 0);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnboundColumn2", 1);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProduccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonalOrden");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrdenProd");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstCase");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SopProc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SopMemoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SopLecTar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SopCDDVD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SopHDD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BusDeDatos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SistOpe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ofimatica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Otros");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DetEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			this.lblBodega = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btSalir = new Infragistics.Win.Misc.UltraButton();
			this.btTransferir = new Infragistics.Win.Misc.UltraButton();
			this.btnConsultar = new Infragistics.Win.Misc.UltraButton();
			this.cmbEstados = new System.Windows.Forms.ComboBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			this.SuspendLayout();
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(176, 11);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(42, 16);
			this.lblBodega.TabIndex = 152;
			this.lblBodega.Text = "Técnico";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 97;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 91;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 87;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 137;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 70;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 67;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Número";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 98;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 99;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Código";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 180;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Artículo";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 300;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance3;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 67;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 125;
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
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(928, 152);
			this.ultraGrid1.TabIndex = 4;
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn6.DataType = typeof(System.Double);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
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
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// cmbPersonal
			// 
			this.cmbPersonal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonal.DataSource = this.ultraDataSource1;
			this.cmbPersonal.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 301;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbPersonal.DisplayMember = "Nombre";
			this.cmbPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonal.Location = new System.Drawing.Point(240, 8);
			this.cmbPersonal.MaxDropDownItems = 30;
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.Size = new System.Drawing.Size(320, 21);
			this.cmbPersonal.TabIndex = 1;
			this.cmbPersonal.ValueMember = "idPersonal";
			this.cmbPersonal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbPersonal_KeyDown);
			this.cmbPersonal.ValueChanged += new System.EventHandler(this.cmbPersonal_ValueChanged);
			this.cmbPersonal.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbPersonal_InitializeLayout);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(584, 11);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 157;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(640, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 2;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
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
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource3;
			appearance10.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance10;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 77;
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 4;
			ultraGridColumn19.Width = 40;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.Caption = "Código";
			ultraGridColumn20.Header.VisiblePosition = 5;
			ultraGridColumn20.Width = 127;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.Caption = "Artículo";
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 166;
			ultraGridColumn22.Header.Caption = "Case";
			ultraGridColumn22.Header.VisiblePosition = 7;
			ultraGridColumn22.Width = 38;
			ultraGridColumn23.Header.Caption = "CPU";
			ultraGridColumn23.Header.VisiblePosition = 8;
			ultraGridColumn23.Width = 38;
			ultraGridColumn24.Header.Caption = "RAM";
			ultraGridColumn24.Header.VisiblePosition = 9;
			ultraGridColumn24.Width = 38;
			ultraGridColumn25.Header.Caption = "Lector";
			ultraGridColumn25.Header.VisiblePosition = 10;
			ultraGridColumn25.Width = 45;
			ultraGridColumn26.Header.Caption = "CD/DVD";
			ultraGridColumn26.Header.VisiblePosition = 11;
			ultraGridColumn26.Width = 50;
			ultraGridColumn27.Header.Caption = "HDD";
			ultraGridColumn27.Header.VisiblePosition = 12;
			ultraGridColumn27.Width = 38;
			ultraGridColumn28.Header.Caption = "Bus de datos";
			ultraGridColumn28.Header.VisiblePosition = 13;
			ultraGridColumn28.Width = 80;
			ultraGridColumn29.Header.Caption = "SO";
			ultraGridColumn29.Header.VisiblePosition = 14;
			ultraGridColumn29.Width = 27;
			ultraGridColumn30.Header.Caption = "Ofimática";
			ultraGridColumn30.Header.VisiblePosition = 15;
			ultraGridColumn30.Width = 60;
			ultraGridColumn31.Header.VisiblePosition = 16;
			ultraGridColumn31.Width = 38;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 17;
			ultraGridColumn32.Width = 126;
			ultraGridColumn33.Header.VisiblePosition = 18;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn34.Header.Caption = "Estado";
			ultraGridColumn34.Header.VisiblePosition = 19;
			ultraGridColumn34.Width = 120;
			ultraGridColumn35.Header.VisiblePosition = 20;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn36.CellAppearance = appearance11;
			ultraGridColumn36.Header.Caption = "";
			ultraGridColumn36.Header.VisiblePosition = 21;
			ultraGridColumn36.NullText = "CREAR SERIAL";
			ultraGridColumn36.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn36.Width = 85;
			ultraGridColumn37.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn37.CellAppearance = appearance12;
			ultraGridColumn37.Header.Caption = "";
			ultraGridColumn37.Header.VisiblePosition = 22;
			ultraGridColumn37.NullText = "IMPRIMIR";
			ultraGridColumn37.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn37.Width = 60;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance13;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance14;
			this.ultraGrid2.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance15;
			appearance16.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance16.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(8, 216);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(1031, 296);
			this.ultraGrid2.TabIndex = 5;
			this.ultraGrid2.DoubleClick += new System.EventHandler(this.ultraGrid2_DoubleClick);
			this.ultraGrid2.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_CellChange);
			this.ultraGrid2.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid2_BeforeCellUpdate);
			this.ultraGrid2.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_ClickCellButton);
			this.ultraGrid2.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_AfterCellUpdate);
			this.ultraGrid2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid2_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(bool);
			ultraDataColumn22.DefaultValue = false;
			ultraDataColumn23.DataType = typeof(bool);
			ultraDataColumn23.DefaultValue = false;
			ultraDataColumn24.DataType = typeof(bool);
			ultraDataColumn24.DefaultValue = false;
			ultraDataColumn25.DataType = typeof(bool);
			ultraDataColumn25.DefaultValue = false;
			ultraDataColumn26.DataType = typeof(bool);
			ultraDataColumn26.DefaultValue = false;
			ultraDataColumn27.DataType = typeof(bool);
			ultraDataColumn27.DefaultValue = false;
			ultraDataColumn28.DataType = typeof(bool);
			ultraDataColumn28.DefaultValue = false;
			ultraDataColumn29.DataType = typeof(bool);
			ultraDataColumn29.DefaultValue = false;
			ultraDataColumn30.DataType = typeof(bool);
			ultraDataColumn30.DefaultValue = false;
			ultraDataColumn31.DataType = typeof(bool);
			ultraDataColumn31.DefaultValue = false;
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35});
			// 
			// btSalir
			// 
			this.btSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance19.Image = 14;
			this.btSalir.Appearance = appearance19;
			this.btSalir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSalir.Location = new System.Drawing.Point(968, 8);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(72, 21);
			this.btSalir.TabIndex = 6;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// btTransferir
			// 
			this.btTransferir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance20.Image = 14;
			this.btTransferir.Appearance = appearance20;
			this.btTransferir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btTransferir.Enabled = false;
			this.btTransferir.Location = new System.Drawing.Point(968, 56);
			this.btTransferir.Name = "btTransferir";
			this.btTransferir.Size = new System.Drawing.Size(72, 21);
			this.btTransferir.TabIndex = 7;
			this.btTransferir.Text = "&Transferir";
			this.btTransferir.Visible = false;
			this.btTransferir.Click += new System.EventHandler(this.btTransferir_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnConsultar.Location = new System.Drawing.Point(784, 8);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(72, 21);
			this.btnConsultar.TabIndex = 3;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// cmbEstados
			// 
			this.cmbEstados.Location = new System.Drawing.Point(16, 8);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(144, 21);
			this.cmbEstados.TabIndex = 0;
			this.cmbEstados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstados_KeyDown);
			this.cmbEstados.Validated += new System.EventHandler(this.cmbEstados_Validated);
			this.cmbEstados.SelectedIndexChanged += new System.EventHandler(this.cmbEstados_SelectedIndexChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1048, 8);
			this.groupBox1.TabIndex = 158;
			this.groupBox1.TabStop = false;
			// 
			// frmDptoTecnico
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1048, 518);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btTransferir);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbPersonal);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "frmDptoTecnico";
			this.Text = "Pendientes de Producción";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDptoTecnico_KeyDown);
			this.Load += new System.EventHandler(this.frmDptoTecnico_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bLoad = false;

		private void frmDptoTecnico_Load(object sender, System.EventArgs e)
		{
			bLoad = true;

			this.cmbEstados.Text = "PENDIENTE";
			this.cmbEstados.Items.Add("PENDIENTE");
			this.cmbEstados.Items.Add("TERMINADO");
			this.cmbEstados.Items.Add("TRANSFERIDO");			
			
			this.dtFecha.Value = DateTime.Today;

			bLoad = false;

			this.cmbPersonal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select p.idPersonal, p.Nombre From Personal p Inner Join Usuario u On u.Nombre = p.Codigo Where u.idGrupo = 14 Order By p.Nombre Asc");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cmbPersonal_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void Consulta()
		{
			if (this.cmbEstados.Text.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione un Estado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbEstados.Focus();
				return;
			}

			if (this.cmbPersonal.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Técnico", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbPersonal.Focus();
				return;
			}

			int iEstado = 0;
			
			if ((string)this.cmbEstados.Text == "PENDIENTE") iEstado = 0;
			if ((string)this.cmbEstados.Text == "TERMINADO") iEstado = 1;
			if ((string)this.cmbEstados.Text == "TRANSFERIDO") iEstado = 2;

			this.ultraGrid2.DataSource = this.ultraDataSource3;
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select * From ConsultaProduccion(0, 0)");

			DateTime dtFecha = (DateTime)this.dtFecha.Value;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			string sSQL = string.Format("Exec DepartamentoTecnico {0}, '{1}', {2}", (int)this.cmbPersonal.Value, dtFecha.ToString("yyyyMMdd"), iEstado);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			if (this.ultraGrid1.Rows.Count == 0) this.btTransferir.Enabled = false;
			else  this.btTransferir.Enabled = true;
		}

		private void ConsultaProduccion()
		{
			if (this.ultraGrid1.ActiveRow.Cells["idOrdenProd"].Value == System.DBNull.Value) return;

			string sSQL = string.Format("Select * From ConsultaProduccion({0}, {1})", (int)this.ultraGrid1.ActiveRow.Cells["idPersonal"].Value, (int)this.ultraGrid1.ActiveRow.Cells["idOrdenProd"].Value);
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void cmbPersonal_ValueChanged(object sender, System.EventArgs e)
		{
		//	Consulta();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			ConsultaProduccion();
		}

		private void ultraGrid2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid2_DoubleClick(object sender, System.EventArgs e)
		{
			if (this.ultraGrid2.ActiveRow.Cells["idProduccion"].Value != System.DBNull.Value)
			{
				//			if (this.ultraGrid2.ActiveRow.Cells["Codigo"].IsActiveCell || this.ultraGrid2.ActiveRow.Cells["Articulo"].IsActiveCell)
				//			{
				using(frmSerialesProduccion miSerialDeProduccion = new frmSerialesProduccion((int)this.ultraGrid1.ActiveRow.Cells["idOrdenProd"].Value, (int)this.ultraGrid1.ActiveRow.Cells["idPersonal"].Value, (int)this.ultraGrid2.ActiveRow.Cells["idProduccion"].Value, this.ultraGrid2.ActiveRow.Cells["Articulo"].Value.ToString().Trim(), (int)this.ultraGrid2.ActiveRow.Cells["Estado"].Value))
				{
					if (DialogResult.OK == miSerialDeProduccion.ShowDialog())
					{						
						ConsultaProduccion();
					}
				}
				//}
			}
		}

		private void ultraGrid2_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			//			if (e.Cell.Column.ToString() == "EstCase")
			//			{
			if ((int)e.Cell.Row.Cells["Estado"].Value == 0)
			{
				e.Cancel = true; 
				MessageBox.Show("No ha registrado las seriales.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				//e.Cell.Row.Cells["EstCase"].Value = (bool)e.Cell.Row.Cells["EstCase"].OriginalValue;
				return;
			}
				
			if ((int)e.Cell.Row.Cells["Estado"].Value == 4)
			{
				e.Cancel = true; 
				MessageBox.Show("Partes dañadas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				//e.Cell.Row.Cells["EstCase"].Value = (bool)e.Cell.Row.Cells["EstCase"].OriginalValue;
				return;
			}

			if ((int)e.Cell.Row.Cells["Estado"].Value == 6)
			{
				e.Cancel = true; 
				MessageBox.Show("No puede modificar ya registro la serial.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				//e.Cell.Row.Cells["EstCase"].Value = (bool)e.Cell.Row.Cells["EstCase"].OriginalValue;
				return;
			}

			if ((int)e.Cell.Row.Cells["Estado"].Value == 3)
			{
				e.Cancel = true; 
				MessageBox.Show("No puede modificar ya imprimió la serial.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				//e.Cell.Row.Cells["EstCase"].Value = (bool)e.Cell.Row.Cells["EstCase"].OriginalValue;
				return;
			}

			if ((int)e.Cell.Row.Cells["Estado"].Value == 7)
			{
				e.Cancel = true; 
				MessageBox.Show("No puede modificar ya Transfirió", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				//e.Cell.Row.Cells["EstCase"].Value = (bool)e.Cell.Row.Cells["EstCase"].OriginalValue;
				return;
			}
			//			}
		}

		private void ultraGrid2_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void ActualizaEstado(Infragistics.Win.UltraWinGrid.CellEventArgs e, int IdProduccion)
		{
			int iCont = 0;
				
			for (int j = 7; j < 17; j++)
			{
				if ((bool)e.Cell.Row.Cells[j].Value)
				{
					iCont++;
				}
			}
			
			if (iCont == 10)
			{
				string sSQL = string.Format("Update Produccion Set Estado = 5 Where idProduccion = {0}", IdProduccion);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

				e.Cell.Row.Cells["DetEstado"].Value = "REVISADO";
			}
			else
			{
				string sSQL = string.Format("Update Produccion Set Estado = 2 Where idProduccion = {0}", IdProduccion);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

				e.Cell.Row.Cells["DetEstado"].Value = "SERIALIZADO";
			}

			ConsultaProduccion();
		}

		private void ultraGrid2_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "UnboundColumn1")
			{	
				#region Genera Serial
				if ((int)e.Cell.Row.Cells["Estado"].Value == 0)
				{
					MessageBox.Show("Registre las seriales de los componentes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				if ((int)e.Cell.Row.Cells["Estado"].Value == 2)
				{
					MessageBox.Show("Registre el estado de los componentes para crear la serial.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				if ((int)e.Cell.Row.Cells["Estado"].Value == 3 || (int)e.Cell.Row.Cells["Estado"].Value == 6 || (int)e.Cell.Row.Cells["Estado"].Value == 7)
				{
					MessageBox.Show("La producción seleccionada ya esta terminada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				if ((int)e.Cell.Row.Cells["Estado"].Value == 4)
				{
					MessageBox.Show("No puede crear Serial para una Producción con Partes Dañadas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				string sSQL = string.Format("Exec SerialProdTerminado {0}, {1}", (int)e.Cell.Row.Cells["idOrdenProd"].Value, (int)e.Cell.Row.Cells["Numero"].Value);				
				string sSerial = Funcion.sEscalarSQL(cdsSeteoF, sSQL);

				string sSQLActSerial = string.Format("Update Produccion Set Serial = '{0}', Estado = 6 Where idProduccion = {1}", sSerial, (int)e.Cell.Row.Cells["idProduccion"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLActSerial);

				MessageBox.Show("Serial " + sSerial + " creada con exito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				ConsultaProduccion();
				#endregion Genera Serial
			}

			if (e.Cell.Column.ToString() == "UnboundColumn2")
			{	
				#region Imprime Etiquetas
				if ((int)e.Cell.Row.Cells["Estado"].Value == 0)
				{
					MessageBox.Show("Registre las Seriales de los Componentes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				if ((int)e.Cell.Row.Cells["Estado"].Value == 2)
				{
					MessageBox.Show("Registre el Estado de los componentes para Crear la Serial e Imprimir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				if ((int)e.Cell.Row.Cells["Estado"].Value == 4)
				{
					MessageBox.Show("No puede Imprimir una Producción con Partes Dañadas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				if ((int)e.Cell.Row.Cells["Estado"].Value == 5)
				{
					MessageBox.Show("Para Imprimir Debe Generar la Serial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				int IdOP = (int) e.Cell.Row.Cells["idOrdenProd"].Value;
				int IdProduccion = (int) e.Cell.Row.Cells["idProduccion"].Value;

				#region Guarda Movimiento de la Serial En Ingreso de Bodega
				int iEstado = (int)e.Cell.Row.Cells["Estado"].Value;
				int idSerial = (int)e.Cell.Row.Cells["idSerial"].Value;
				if (iEstado == 6)
				{
					string sSQLRetId = string.Format("Exec RetornaIdDetalleIB {0}", IdOP);
					int idDetCompra = Funcion.iEscalarSQL(cdsSeteoF, sSQLRetId);
	
					string sSQLGSerial = string.Format("Exec GuardaSeriales {0}, {1}, '{2}', 1, 20", 
						idSerial, (int)e.Cell.Row.Cells["idArticulo"].Value, e.Cell.Row.Cells["Serial"].Value.ToString());
					idSerial = Funcion.iEscalarSQL(cdsSeteoF, sSQLGSerial);

					string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, 0, 20, '', 1, 'PRODUCTO TERMINADO ORDEN DE PRODUCCIÓN N. {2}', 0", 
						idSerial, idDetCompra, this.ultraGrid1.ActiveRow.Cells["Factura"].Value.ToString());
					Funcion.EjecutaSQL(cdsSeteoF, sSQLGDetalleSeriales, true);

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Produccion Set idSerial = {0} Where idProduccion = {1}", idSerial, IdProduccion));
				}
				#endregion Guarda Movimiento de la Serial En Ingreso de Bodega
			
				string sSQLImprime = string.Format("Exec ReporteProdTerminado {0}, {1}", IdOP, IdProduccion);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLImprime);

				string sDir = Funcion.sEscalarSQL(cdsSeteoF, "Select ADLabelJoyOP From SeteoF");
				string sDir1 = Funcion.sEscalarSQL(cdsSeteoF, "Select ADLabelJoyOP1 From SeteoF");

				Process.Start(sDir, "");
				Process.Start(sDir1, "");
				
				string sSQLActEstado = string.Format("Exec ActualizaEstadosProduccion {0}, {1}, {2}, {3}", 
					(int)e.Cell.Row.Cells["Estado"].Value, (int)e.Cell.Row.Cells["idProduccion"].Value, (int)e.Cell.Row.Cells["idPersonalOrden"].Value, (int)e.Cell.Row.Cells["idOrdenProd"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLActEstado);				

				ConsultaProduccion();
				#endregion Imprime Etiquetas
			}
		}

		private void btTransferir_Click(object sender, System.EventArgs e)
		{			
			double dCantidad = 0;

			if (this.ultraGrid1.Rows.Count == 0) 
			{
				MessageBox.Show("No hay ordenes de producción pendientes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (this.ultraGrid2.Rows.Count == 0) 
			{
				MessageBox.Show("No hay productos terminados para transferir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if ((int)dr.Cells["Estado"].Value == 3)	dCantidad++;
			}
			
			if (dCantidad == 0) 
			{
				MessageBox.Show("No hay productos terminados para transferir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			int idTransferencia = Funcion.iEscalarSQL(cdsSeteoF, "Exec CreaTransferenciaDeProductoTerminado");

			string sSQL = string.Format("Exec CreaDetalleTransfProduccion {0}, {1}, '{2}', '{3}', {4}", idTransferencia, (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value, this.ultraGrid1.ActiveRow.Cells["Articulo"].Value.ToString(), this.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString(), dCantidad); 
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			
			sSQL = string.Format("Exec TransformaIE {0}", idTransferencia);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

			int idDetCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ConsultaidTransf {0}, {1}", idTransferencia, (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value));

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if ((int)dr.Cells["Estado"].Value == 6 || (int)dr.Cells["Estado"].Value == 3)
				{
					sSQL = string.Format("Exec GuardaSeriales 0, {0}, '{1}', 3", (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value, dr.Cells["Serial"].Value.ToString());
					int idSerial = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

					string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, 0, 20, '', 1, 'Ordén de Producción {2}', 0", idSerial, idDetCompra, this.ultraGrid1.ActiveRow.Cells["Factura"].Value.ToString());
					Funcion.EjecutaSQL(cdsSeteoF, sSQLGDetalleSeriales, true);

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Produccion Set Estado = 7 Where idProduccion = {0}", (int)dr.Cells["idProduccion"].Value));
				}
			}
			
			int IdCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ProcesaProduccion {0}", idTransferencia));			
			string sSQLAjuste = string.Format("Exec ActualizaEnvioRecepcionTransferencia {0}, 8, {1}, ''", idDetCompra, dCantidad);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLAjuste);

			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update OrdenProduccion Set Transferido = Transferido + {0} Where idOrdenProd = {1}", dCantidad, (int)this.ultraGrid1.ActiveRow.Cells["idOrdenProd"].Value));				
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update OrdenProduccion Set Estado = Case When Cantidad = Transferido Then 4 Else 2 End Where idOrdenProd = {0}", (int)this.ultraGrid1.ActiveRow.Cells["idOrdenProd"].Value));				

			MessageBox.Show(string.Format("{0} Producto(s) Transferido(s) de la Ordén de Producción N. '{1}'", dCantidad, this.ultraGrid1.ActiveRow.Cells["Factura"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.Cursor = Cursors.WaitCursor;

			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldCompra = new ParameterField ();
			ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
			paramFieldCompra.ParameterFieldName = "@idCompra";
			ValIdCompra.Value = IdCompra;
			paramFieldCompra.CurrentValues.Add (ValIdCompra);
			paramFields.Add (paramFieldCompra);

			Reporte miRep = new Reporte("RepTransfSalida.rpt", "");
			miRep.MdiParent = this.MdiParent;			
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();

			this.Cursor = Cursors.Default;

			ConsultaProduccion();
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
//			if (!bLoad) Consulta();
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value) this.dtFecha.Value = DateTime.Today;
		}

		private void cmbEstados_Validated(object sender, System.EventArgs e)
		{
			if (this.cmbEstados.Text.ToString().Length == 0) this.cmbEstados.Text = "PENDIENTE";			
		}

		private void cmbEstados_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbPersonal.Focus();
		}

		private void cmbPersonal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFecha.Focus();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnConsultar_Click(sender, e);
		}

		private void cmbEstados_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			//if (!bLoad)Consulta();
		}

		private void frmDptoTecnico_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void ultraGrid2_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idProduccion"].Value != System.DBNull.Value) 
			{
				int idProduccion = (int)e.Cell.Row.Cells["idProduccion"].Value;
			
				int iSel = 0;
				string sSQL = "";

				if (e.Cell.Column.ToString() == "EstCase")
				{
					this.ultraGrid2.UpdateData();

					if ((bool)e.Cell.Row.Cells["EstCase"].Value) iSel = 1; else iSel = 0;

					sSQL = string.Format("Update Produccion Set EstCase = {0} Where idProduccion = {1}", iSel, idProduccion);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					ActualizaEstado(e, idProduccion);
				}

				if (e.Cell.Column.ToString() == "SopProc")
				{
					this.ultraGrid2.UpdateData();

					if ((bool)e.Cell.Row.Cells["SopProc"].Value) iSel = 1; else iSel = 0;				

					sSQL = string.Format("Update Produccion Set SopProc = {0} Where idProduccion = {1}", iSel, idProduccion);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					ActualizaEstado(e, idProduccion);
				}

				if (e.Cell.Column.ToString() == "SopMemoria")
				{
					this.ultraGrid2.UpdateData();

					if ((bool)e.Cell.Row.Cells["SopMemoria"].Value) iSel = 1; else iSel = 0;				

					sSQL = string.Format("Update Produccion Set SopMemoria = {0} Where idProduccion = {1}", iSel, idProduccion);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					ActualizaEstado(e, idProduccion);
				}

				if (e.Cell.Column.ToString() == "SopLecTar")
				{
					this.ultraGrid2.UpdateData();

					if ((bool)e.Cell.Row.Cells["SopLecTar"].Value) iSel = 1; else iSel = 0;				

					sSQL = string.Format("Update Produccion Set SopLecTar = {0} Where idProduccion = {1}", iSel, idProduccion);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					ActualizaEstado(e, idProduccion);
				}

				if (e.Cell.Column.ToString() == "SopCDDVD")
				{
					this.ultraGrid2.UpdateData();

					if ((bool)e.Cell.Row.Cells["SopCDDVD"].Value) iSel = 1; else iSel = 0;				

					sSQL = string.Format("Update Produccion Set SopCDDVD = {0} Where idProduccion = {1}", iSel, idProduccion);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					ActualizaEstado(e, idProduccion);
				}

				if (e.Cell.Column.ToString() == "SopHDD")
				{
					this.ultraGrid2.UpdateData();

					if ((bool)e.Cell.Row.Cells["SopHDD"].Value) iSel = 1; else iSel = 0;				

					sSQL = string.Format("Update Produccion Set SopHDD = {0} Where idProduccion = {1}", iSel, idProduccion);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					ActualizaEstado(e, idProduccion);
				}

				if (e.Cell.Column.ToString() == "BusDeDatos")
				{
					this.ultraGrid2.UpdateData();

					if ((bool)e.Cell.Row.Cells["BusDeDatos"].Value) iSel = 1; else iSel = 0;				

					sSQL = string.Format("Update Produccion Set BusDeDatos = {0} Where idProduccion = {1}", iSel, idProduccion);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					ActualizaEstado(e, idProduccion);
				}

				if (e.Cell.Column.ToString() == "SistOpe")
				{
					this.ultraGrid2.UpdateData();

					if ((bool)e.Cell.Row.Cells["SistOpe"].Value) iSel = 1; else iSel = 0;				

					sSQL = string.Format("Update Produccion Set SistOpe = {0} Where idProduccion = {1}", iSel, idProduccion);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					ActualizaEstado(e, idProduccion);
				}

				if (e.Cell.Column.ToString() == "Ofimatica")
				{
					this.ultraGrid2.UpdateData();

					if ((bool)e.Cell.Row.Cells["Ofimatica"].Value) iSel = 1; else iSel = 0;				

					sSQL = string.Format("Update Produccion Set Ofimatica = {0} Where idProduccion = {1}", iSel, idProduccion);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					ActualizaEstado(e, idProduccion);
				}

				if (e.Cell.Column.ToString() == "Otros")
				{
					this.ultraGrid2.UpdateData();

					if ((bool)e.Cell.Row.Cells["Otros"].Value) iSel = 1; else iSel = 0;				

					sSQL = string.Format("Update Produccion Set Otros = {0} Where idProduccion = {1}", iSel, idProduccion);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					ActualizaEstado(e, idProduccion);
				}
			}
		}
	}
}

