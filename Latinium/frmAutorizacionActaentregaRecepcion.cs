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
	/// Descripción breve de frmAutorizacionActaentregaRecepcion.
	/// </summary>
	public class frmAutorizacionActaentregaRecepcion : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTransporte;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinGrid.UltraGrid grdFacturas;
		private Infragistics.Win.Misc.UltraLabel lblLocal;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button button1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAutorizacionActaentregaRecepcion()
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

		int IBodega = 0;
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActaEntregaRecepcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega_Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jefe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotaAnulacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioAnula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotaAprueba");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioAprueba");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotaEntrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuarioentrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Entrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAceptar", 0);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAutorizacionActaentregaRecepcion));
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnular", 1);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnImprimir", 2);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 3);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEntrega", 4);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActaEntregaRecepcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega_Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jefe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotaAnulacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioAnula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sector");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotaAprueba");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioAprueba");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotaEntrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuarioentrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.optTransporte = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.grdFacturas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblLocal = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.lblContador = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnExportar = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.optTransporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdFacturas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// optTransporte
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTransporte.Appearance = appearance1;
			this.optTransporte.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTransporte.CheckedIndex = 0;
			this.optTransporte.ItemAppearance = appearance2;
			this.optTransporte.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendiente";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Procesado";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Anulado";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Entregado";
			this.optTransporte.Items.Add(valueListItem1);
			this.optTransporte.Items.Add(valueListItem2);
			this.optTransporte.Items.Add(valueListItem3);
			this.optTransporte.Items.Add(valueListItem4);
			this.optTransporte.Items.Add(valueListItem5);
			this.optTransporte.ItemSpacingVertical = 10;
			this.optTransporte.Location = new System.Drawing.Point(68, 72);
			this.optTransporte.Name = "optTransporte";
			this.optTransporte.Size = new System.Drawing.Size(396, 24);
			this.optTransporte.TabIndex = 1034;
			this.optTransporte.Text = "Todos";
			this.optTransporte.ValueChanged += new System.EventHandler(this.optTransporte_ValueChanged);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Location = new System.Drawing.Point(16, 352);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 1033;
			this.btnCancelar.Text = "&Salir";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// grdFacturas
			// 
			this.grdFacturas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdFacturas.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdFacturas.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdFacturas.DisplayLayout.Appearance = appearance3;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 135;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Fecha Procesa";
			ultraGridColumn2.Header.VisiblePosition = 8;
			ultraGridColumn2.Width = 71;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 7;
			ultraGridColumn3.Width = 80;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 9;
			ultraGridColumn4.Width = 89;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Bodega Origen";
			ultraGridColumn5.Header.VisiblePosition = 10;
			ultraGridColumn5.Width = 150;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Bodega Realiza";
			ultraGridColumn6.Header.VisiblePosition = 11;
			ultraGridColumn6.Width = 150;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 12;
			ultraGridColumn7.Width = 225;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 13;
			ultraGridColumn8.Width = 225;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 17;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Nota Anulacion";
			ultraGridColumn10.Header.VisiblePosition = 22;
			ultraGridColumn10.Width = 150;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Usuario Anula";
			ultraGridColumn11.Header.VisiblePosition = 23;
			ultraGridColumn11.Width = 150;
			ultraGridColumn12.Header.VisiblePosition = 20;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 14;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 15;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Nota Aprueba";
			ultraGridColumn15.Header.VisiblePosition = 18;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.Caption = "Usuario Aprueba";
			ultraGridColumn16.Header.VisiblePosition = 19;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.Caption = "Nota Entrega";
			ultraGridColumn17.Header.VisiblePosition = 21;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.Caption = "Usuario Entrega";
			ultraGridColumn18.Header.VisiblePosition = 24;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.Caption = "Fecha Entrega";
			ultraGridColumn19.Header.VisiblePosition = 16;
			ultraGridColumn20.Header.Caption = "...";
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn21.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			appearance4.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn21.CellButtonAppearance = appearance4;
			ultraGridColumn21.Header.Caption = "...";
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn21.Width = 30;
			ultraGridColumn22.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			appearance5.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn22.CellButtonAppearance = appearance5;
			ultraGridColumn22.Header.Caption = "...";
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn22.Width = 30;
			ultraGridColumn23.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance6.Image = ((object)(resources.GetObject("appearance6.Image")));
			appearance6.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn23.CellButtonAppearance = appearance6;
			ultraGridColumn23.Header.Caption = "...";
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn23.Width = 30;
			ultraGridColumn24.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance7.Image = ((object)(resources.GetObject("appearance7.Image")));
			appearance7.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellButtonAppearance = appearance7;
			ultraGridColumn24.Header.Caption = "...";
			ultraGridColumn24.Header.VisiblePosition = 6;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn24.Width = 30;
			ultraGridColumn25.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			appearance8.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellButtonAppearance = appearance8;
			ultraGridColumn25.Header.Caption = "..";
			ultraGridColumn25.Header.VisiblePosition = 5;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn25.Width = 40;
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
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			ultraGridBand1.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.grdFacturas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdFacturas.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.grdFacturas.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 9F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdFacturas.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.grdFacturas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdFacturas.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdFacturas.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdFacturas.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.grdFacturas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdFacturas.Location = new System.Drawing.Point(12, 144);
			this.grdFacturas.Name = "grdFacturas";
			this.grdFacturas.Size = new System.Drawing.Size(840, 200);
			this.grdFacturas.TabIndex = 1031;
			this.grdFacturas.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdFacturas_ClickCellButton);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.DateTime);
			ultraDataColumn20.DataType = typeof(bool);
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
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// lblLocal
			// 
			this.lblLocal.AutoSize = true;
			this.lblLocal.BackColor = System.Drawing.Color.Transparent;
			this.lblLocal.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblLocal.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.lblLocal.Location = new System.Drawing.Point(12, 40);
			this.lblLocal.Name = "lblLocal";
			this.lblLocal.Size = new System.Drawing.Size(30, 15);
			this.lblLocal.TabIndex = 1030;
			this.lblLocal.Text = "Local";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel1.Location = new System.Drawing.Point(12, 108);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 15);
			this.ultraLabel1.TabIndex = 1029;
			this.ultraLabel1.Text = "Buscar";
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(372, 106);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 1028;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 180;
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(68, 37);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(280, 21);
			this.cmbBodega.TabIndex = 1026;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(276, 104);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 1025;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel4.Location = new System.Drawing.Point(12, 80);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(38, 15);
			this.ultraLabel4.TabIndex = 1027;
			this.ultraLabel4.Text = "Estado";
			// 
			// txtBusqueda
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance15;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(68, 105);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(192, 21);
			this.txtBusqueda.TabIndex = 1024;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(196, 10);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 1023;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(12, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 1022;
			this.label6.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance16;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(244, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1021;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.ValueChanged += new System.EventHandler(this.dtHasta_ValueChanged);
			// 
			// dtDesde
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance17;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(68, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1020;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
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
			// btnExportar
			// 
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(776, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 24);
			this.btnExportar.TabIndex = 1035;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// button1
			// 
			this.button1.CausesValidation = false;
			this.button1.Location = new System.Drawing.Point(728, 112);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(104, 23);
			this.button1.TabIndex = 1036;
			this.button1.Text = "&Procesar varios";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// frmAutorizacionActaentregaRecepcion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(864, 382);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.optTransporte);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.grdFacturas);
			this.Controls.Add(this.lblLocal);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Name = "frmAutorizacionActaentregaRecepcion";
			this.Text = "AUTORIZACIÓN ACTA DE ENTREGA-RECPECIÓN";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAutorizacionActaentregaRecepcion_KeyDown);
			this.Load += new System.EventHandler(this.frmAutorizacionActaentregaRecepcion_Load);
			((System.ComponentModel.ISupportInitialize)(this.optTransporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdFacturas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void frmAutorizacionActaentregaRecepcion_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200111");

			if (!Funcion.Permiso("1887", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar  Autorización Acta Entrega Recepción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			DateTime dtInicio = DateTime.Today.AddMonths(-3500);
			

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;
		  DateTime dtIniciod = (DateTime) this.dtHasta.Value;
			
			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtHasta.Value = DateTime.Today;

			FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 1, dtIniciod, false, cdsSeteoF);
		
			IBodega = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec BodegaAsignadaPorUsuarioBuscaFacturas {0}", MenuLatinium.IdUsuario));

			if (IBodega > 0)
			{
					this.cmbBodega.Value = IBodega;
				this.cmbBodega.Enabled = false;
			}
			else
			{
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 1, dtIniciod, false, cdsSeteoF);
				
				this.cmbBodega.Enabled = true;
			}
			this.grdFacturas.DisplayLayout.Bands[0].Columns["btnAceptar"].Hidden = true;
			this.grdFacturas.DisplayLayout.Bands[0].Columns["btnAnular"].Hidden = true;
			this.grdFacturas.DisplayLayout.Bands[0].Columns["btnImprimir"].Hidden = true;
			this.grdFacturas.DisplayLayout.Bands[0].Columns["btnEntrega"].Hidden = true;
			this.grdFacturas.DisplayLayout.Bands[0].Columns["Sel"].Hidden = true;
			this.grdFacturas.DisplayLayout.Bands[0].Columns["btnVer"].Hidden = false;
			this.txtBusqueda.Focus();


			string carpeta_1 ="";
			carpeta_1 = @"C:\Latinium\Reportes\ReporteEntregayRecepcion3.rpt";

		}

		private void dtHasta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtHasta.Value != System.DBNull.Value)
			{
			
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 1, (DateTime)this.dtHasta.Value, false, cdsSeteoF);
			
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la fecha inicial", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la fecha final", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;
			
			if (this.cmbBodega.ActiveRow != null) IBodega = (int)this.cmbBodega.Value;

			string sSQL = "";

			sSQL = string.Format("Exec ConsultaGeneralAprobacionActaEntregaRecepcion '{0}', '{1}', {2}, {3}, '{4}'", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 
				IBodega, (int) this.optTransporte.Value,  this.txtBusqueda.Text.ToString());

			this.grdFacturas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			if (this.grdFacturas.Rows.Count > 0) this.grdFacturas.Focus();

			this.lblContador.Text = this.grdFacturas.Rows.Count + " REGISTROS ENCONTRADOS";

			//if (this.cmbBodega.ActiveRow != null) IBodega = (int)this.cmbBodega.Value;

		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
		this.Close();	
		}

		private void frmAutorizacionActaentregaRecepcion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();	
		}

		private void optTransporte_ValueChanged(object sender, System.EventArgs e)
		{
			if((int) this.optTransporte.Value == 0)
			{
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnAceptar"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnAnular"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnImprimir"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnVer"].Hidden = false;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnEntrega"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["Sel"].Hidden = true;
			}

			if((int) this.optTransporte.Value == 1)
			{
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnAceptar"].Hidden = false;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnAnular"].Hidden = false;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnImprimir"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnVer"].Hidden = false;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnEntrega"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["Sel"].Hidden = false;
			}
			if((int) this.optTransporte.Value == 2)
			{

				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnAceptar"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnAnular"].Hidden = false;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnImprimir"].Hidden = false;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnVer"].Hidden = false;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnEntrega"].Hidden = false;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["Sel"].Hidden = true;
			}
			
			if((int) this.optTransporte.Value == 3)
			{
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnAceptar"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnAnular"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnImprimir"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnVer"].Hidden = false;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnEntrega"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["Sel"].Hidden = true;
			}
			if((int) this.optTransporte.Value == 4)
			{
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnAceptar"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnAnular"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnImprimir"].Hidden = false;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnVer"].Hidden = false;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["btnEntrega"].Hidden = true;
				this.grdFacturas.DisplayLayout.Bands[0].Columns["Sel"].Hidden = true;
			}
	   	this.btnBuscar_Click (sender, e);
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.grdFacturas);
		}

		private void grdFacturas_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnAceptar")
			{	
				if( (int)e.Cell.Row.Cells["iEstado"].Value == 1)
				{
					if (!miAcceso.BProcesar)
					{
						MessageBox.Show("No tiene Acceso a Procesar el documento Acta Entrega Recepción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
					int icont=0;
					icont= Funcion.iEscalarSQL(cdsSeteoF, string.Format(" select count(1) from dbo.DetalleActaEntregaRecepcion where idActaEntregaRecepcion =  {0}",
						(int)e.Cell.Row.Cells["idActaEntregaRecepcion"].Value));

					if (icont == 0 )
					{
						MessageBox.Show("El documento Acta Entrega Recepción, no tiene productos seleccionados.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Aprobar el documento. ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					{
						using (frmNotas Notas  = new frmNotas())
						{
							if (DialogResult.OK == Notas.ShowDialog())
							{
								string sSQL = string.Format("Update ActaEntregaRecepcion set iEstado = 2, FechaAprueba = Getdate (), Fecha = Getdate (), UsuarioAprueba = suser_sname(),  NotaAprueba = '{1}'  where idActaEntregaRecepcion = {0}",
									(int)e.Cell.Row.Cells["idActaEntregaRecepcion"].Value, Notas.txtNotas.Text);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL);

								this.btnBuscar_Click(sender, e);
							}
						}
					}
				}
			}

			if (e.Cell.Column.ToString() == "btnAnular")
			{	
				if ((int)e.Cell.Row.Cells["iEstado"].Value == 1 )
				{
					if (!miAcceso.Anular)
					{
						MessageBox.Show("No tiene Acceso, Anular el documento Acta Entrega Recepción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
				if ((int)e.Cell.Row.Cells["iEstado"].Value == 2 )
				{
					if (!miAcceso.AnularCrediPoint)
					{
						MessageBox.Show("No tiene Acceso Anular Documentos ya procesados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
				if ((int)e.Cell.Row.Cells["iEstado"].Value == 3 )
				{
					return;
					
				}
				if (DialogResult.Yes == MessageBox.Show("Esta seguro de Anular el documento. ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					using (frmNotas Notas  = new frmNotas())
					{
						if (DialogResult.OK == Notas.ShowDialog())
						{
							string sSQL = string.Format("Update ActaEntregaRecepcion set iEstado = 3, FechaAnula = Getdate (), UsuarioAnula = suser_sname(),  NotaAnulacion = '{1}'  where idActaEntregaRecepcion = {0}",
								(int)e.Cell.Row.Cells["idActaEntregaRecepcion"].Value, Notas.txtNotas.Text);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL);

							this.btnBuscar_Click(sender, e);
						}
					}
				}
			}

			#region Entrega
			if (e.Cell.Column.ToString() == "btnEntrega")
			{	
				if ((int)e.Cell.Row.Cells["iEstado"].Value == 2 )
				{
					if (!miAcceso.BAutorizarPedido)
					{
						MessageBox.Show("No tiene Acceso a registrar la entrega.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
				if ((int)e.Cell.Row.Cells["iEstado"].Value >= 3 )
				{
					return;
					
				}
				if (DialogResult.Yes == MessageBox.Show("Esta seguro de Registar la Entrega. ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					using (frmNotas Notas  = new frmNotas())
					{
						if (DialogResult.OK == Notas.ShowDialog())
						{
							string sSQL = string.Format("Update ActaEntregaRecepcion set iEstado = 4, FechaEntrega = Getdate (), UsuarioEntrega = suser_sname(),  NotaEntrega = '{1}'  where idActaEntregaRecepcion = {0}",
								(int)e.Cell.Row.Cells["idActaEntregaRecepcion"].Value, Notas.txtNotas.Text);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL);

							this.btnBuscar_Click(sender, e);
						}
					}
				}
			}
			#endregion Entrega
			#region Imprimir
			if (e.Cell.Column.ToString() == "btnImprimir")
			{
				if ((int)e.Cell.Row.Cells["iEstado"].Value != 3 )
				{
					if (!miAcceso.Imprimir)
					{
						MessageBox.Show("No tiene Acceso Imprimir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
					else
					{
						ParameterFields paramFields = new ParameterFields();

						ParameterField paramFieldCedula = new ParameterField ();
						ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
						paramFieldCedula.ParameterFieldName = "@idActaEntregaRecepcion";
						ValCedula.Value = Convert.ToInt32( e.Cell.Row.Cells["idActaEntregaRecepcion"].Value);
						paramFieldCedula.CurrentValues.Add (ValCedula);
						paramFields.Add (paramFieldCedula);

						Reporte miRep = new Reporte("ReporteEntregayRecepcion3.rpt", "");
						miRep.MdiParent = this.MdiParent;
						miRep.Titulo("Voucher Transporte");
						miRep.crVista.ParameterFieldInfo = paramFields;
						miRep.Show();
					}
				}
				else
				{
					return;
				}
			}

			#endregion Imprimir

			#region Ver
			if (e.Cell.Column.ToString() == "btnVer")
			{
				using (frmActaEntregaRecepcion Notas  = new frmActaEntregaRecepcion(true,(int)e.Cell.Row.Cells["idCompra"].Value,(int)e.Cell.Row.Cells["idActaEntregaRecepcion"].Value))
				{
					if (DialogResult.OK == Notas.ShowDialog( ))
					{

	
					}
				}
			}
			#endregion Ver
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			int icont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdFacturas.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value == true)
				{
					icont = icont +1;
				}	
			}

			if ( icont == 0 )
			{
				MessageBox.Show(string.Format("Seleccione al menos un Documento"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Aprobar los documentos Seleccionados. ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				using (frmNotas Notas  = new frmNotas())
				{
					if (DialogResult.OK == Notas.ShowDialog())
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdFacturas.Rows.All)
						{
							if ((bool)dr.Cells["Sel"].Value == true)
							{
								string sSQL = string.Format("Update ActaEntregaRecepcion set iEstado = 2, FechaAprueba = Getdate (), Fecha = Getdate (), UsuarioAprueba = suser_sname(),  NotaAprueba = '{1}'  where idActaEntregaRecepcion = {0}",
									Convert.ToInt32 ( dr.Cells["idActaEntregaRecepcion"].Value), Notas.txtNotas.Text);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL);
							}
						}
						this.btnBuscar_Click(sender,e);
					}
				}
			}

		}
	}
}
