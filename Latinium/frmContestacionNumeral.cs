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
	/// Descripción breve de frmContestacionNumeral.
	/// </summary>
	public class frmContestacionNumeral : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridReporte;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmContestacionNumeral()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idContestacionNumeral");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Oficio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImposibilidadCobro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCobranzaEfectiva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCorreoNegacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCartaNotificacionNegacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MotivoNegacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuantiaSobreVersaReclamoAsciende");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MontoTotalCompraFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntregaAseguradora");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaInformeRiesgroCrediticio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorFacturaAeguradoraSur");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroFacturaPoint");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo31Mayo2022");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Canton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnNumeral", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmContestacionNumeral));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idContestacionNumeral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Oficio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ImposibilidadCobro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCobranzaEfectiva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCorreoNegacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCartaNotificacionNegacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MotivoNegacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuantiaSobreVersaReclamoAsciende");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MontoTotalCompraFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEntregaAseguradora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaInformeRiesgroCrediticio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorFacturaAeguradoraSur");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroFacturaPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo31Mayo2022");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Agencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Canton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Region");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.uGridReporte = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExcel = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
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
			// uGridReporte
			// 
			this.uGridReporte.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridReporte.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridReporte.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 127;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 134;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 210;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Imposibilidad Cobro";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 158;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Fecha Cobranza Efectiva";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 158;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Fecha Correo Negacion";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 143;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Fecha Carta Notificacion Negacion";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 202;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Motivo Negacion";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 107;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Cuantia Sobre Versa Reclamo Asciende";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 228;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Monto Total Compra Factura";
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Width = 174;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Fecha Entrega Aseguradora";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Width = 162;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "Fecha Informe Riesgro Crediticio";
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Width = 191;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.Caption = "Valor Factura Aeguradora Sur";
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Numero Factura Point";
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn15.Width = 139;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.Caption = "Saldo 31 Mayo 2022";
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn16.Width = 121;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 18;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 19;
			ultraGridColumn20.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.ForegroundAlpha = Infragistics.Win.Alpha.Transparent;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			appearance2.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn20.CellButtonAppearance = appearance2;
			ultraGridColumn20.Header.Caption = "CN";
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn20.Width = 42;
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
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.uGridReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridReporte.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridReporte.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridReporte.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridReporte.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridReporte.DisplayLayout.Override.CellAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8.5F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridReporte.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.uGridReporte.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridReporte.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridReporte.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridReporte.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridReporte.Location = new System.Drawing.Point(8, 56);
			this.uGridReporte.Name = "uGridReporte";
			this.uGridReporte.Size = new System.Drawing.Size(1088, 384);
			this.uGridReporte.TabIndex = 1043;
			this.uGridReporte.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridReporte_ClickCellButton);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.DateTime);
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(-568, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(2296, 8);
			this.groupBox1.TabIndex = 1042;
			this.groupBox1.TabStop = false;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1008, 5);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(88, 24);
			this.btnExcel.TabIndex = 1041;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(392, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 1040;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel1.Location = new System.Drawing.Point(8, 10);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 15);
			this.ultraLabel1.TabIndex = 1039;
			this.ultraLabel1.Text = "Buscar";
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(304, 6);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 1038;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// txtBusqueda
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance8;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(64, 7);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(232, 21);
			this.txtBusqueda.TabIndex = 1037;
			// 
			// frmContestacionNumeral
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1104, 446);
			this.Controls.Add(this.uGridReporte);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.txtBusqueda);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmContestacionNumeral";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Contestacion Numeral";
			this.Load += new System.EventHandler(this.frmContestacionNumeral_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			string sSQLReporte = string.Format("Exec ConsultaContestacionNumeral '{0}'", this.txtBusqueda.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQLReporte, this.uGridReporte);

			this.lblContador.Text = this.uGridReporte.Rows.Count + " Registros Encontrados";
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.uGridReporte.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			FuncionesProcedimientos.ExportaExcel(this.uGridReporte);
		}

		private void frmContestacionNumeral_Load(object sender, System.EventArgs e)
		{
				miAcceso = new Acceso(cdsSeteoF, "023305");
				
				if (!Funcion.Permiso("1982", cdsSeteoF))
				{				
					MessageBox.Show("No puede ingresar al modulo Contestacion Numeral", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					BeginInvoke(new MethodInvoker(UnloadMe));				
	
					return;
				}
	
				if (miAcceso.BExportar) this.btnExcel.Enabled = true;
		}

		private void uGridReporte_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Contestacion Numeral
			if (e.Cell.Column.ToString() == "btnNumeral")
			{
				int idRecursosApelacion = 0;
				idRecursosApelacion = (int)e.Cell.Row.Cells["idContestacionNumeral"].Value;
				if (DialogResult.Yes == MessageBox.Show(" ¿Esta Seguro de Imprimir el Documento?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					ParameterFields paramFields = new ParameterFields ();
				
					ParameterField pfRet = new ParameterField ();
					ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
					pfRet.ParameterFieldName = "@idContestacionNumeral";
					idCompraRet.Value = idRecursosApelacion;
					pfRet.CurrentValues.Add (idCompraRet);
					paramFields.Add (pfRet);
				
					Reporte miRepor = new Reporte("CartaContestacionNumeral.rpt", "");
					miRepor.MdiParent = this.MdiParent;
					miRepor.Titulo("");
					miRepor.crVista.ParameterFieldInfo = paramFields;

					miRepor.Show();
				}
			}
			#endregion Contestacion Numeral
		}
	}
}
