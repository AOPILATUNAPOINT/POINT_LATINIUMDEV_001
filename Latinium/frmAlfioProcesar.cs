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
	/// Descripción breve de frmAlfioProcesar.
	/// </summary>
	public class frmAlfioProcesar : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label lblBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridReporte;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTransporte;
		private System.Windows.Forms.Button btnDeposito;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCargos;
		private System.Windows.Forms.Button btnSalir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAlfioProcesar()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAlfioProcesar));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormularioSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormularioCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConfirmacionAPP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Formulario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NroCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NroCedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acreditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("comision");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Voucher");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFioCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abono");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 0);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDepositar", 1);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEfectivo", 2);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormularioSolicitud");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormularioCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ConfirmacionAPP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Formulario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NroCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NroCedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Acreditado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Email");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Voucher");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFioCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Abono");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			this.btnExcel = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.lblBuscar = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridReporte = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.optTransporte = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnDeposito = new System.Windows.Forms.Button();
			this.chkCargos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnSalir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTransporte)).BeginInit();
			this.SuspendLayout();
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1146, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(72, 24);
			this.btnExcel.TabIndex = 882;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(456, 71);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 881;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBuscar
			// 
			this.lblBuscar.AutoSize = true;
			this.lblBuscar.Location = new System.Drawing.Point(16, 73);
			this.lblBuscar.Name = "lblBuscar";
			this.lblBuscar.Size = new System.Drawing.Size(39, 16);
			this.lblBuscar.TabIndex = 879;
			this.lblBuscar.Text = "Buscar";
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(72, 72);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(288, 19);
			this.txtBuscar.TabIndex = 878;
			// 
			// btnConsultar
			// 
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(368, 70);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 877;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(192, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(34, 16);
			this.label3.TabIndex = 873;
			this.label3.Text = "Hasta";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 872;
			this.label1.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.AutoSelect = false;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 871;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.AutoSelect = false;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(72, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 870;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// uGridReporte
			// 
			this.uGridReporte.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridReporte.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridReporte.DisplayLayout.Appearance = appearance3;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 5;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "# Solicitud";
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.Width = 70;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Cedula";
			ultraGridColumn6.Header.VisiblePosition = 9;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Codigo Nuovo";
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Cliente";
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "# Formulario";
			ultraGridColumn9.Header.VisiblePosition = 15;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance4;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 11;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Banco Beneficiario";
			ultraGridColumn11.Header.VisiblePosition = 18;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Tipo Cuenta Beneficiario";
			ultraGridColumn12.Header.VisiblePosition = 19;
			ultraGridColumn13.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "# Cuenta Beneficiario";
			ultraGridColumn13.Header.VisiblePosition = 20;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.Caption = "# Cedula Beneficiario";
			ultraGridColumn14.Header.VisiblePosition = 17;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Beneficiario Acreditar";
			ultraGridColumn15.Header.VisiblePosition = 16;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance5;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Comision";
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance6;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 13;
			ultraGridColumn18.Header.VisiblePosition = 21;
			ultraGridColumn19.Header.VisiblePosition = 23;
			ultraGridColumn20.Header.VisiblePosition = 24;
			ultraGridColumn21.Header.VisiblePosition = 22;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance7;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 14;
			ultraGridColumn23.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			appearance8.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn23.CellButtonAppearance = appearance8;
			ultraGridColumn23.Header.Caption = "Ir";
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn23.Width = 30;
			ultraGridColumn24.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			appearance9.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellButtonAppearance = appearance9;
			ultraGridColumn24.Header.Caption = "Deposito";
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn24.Width = 30;
			ultraGridColumn25.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance10.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellAppearance = appearance10;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			appearance11.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellButtonAppearance = appearance11;
			ultraGridColumn25.Header.Caption = "Pago";
			ultraGridColumn25.Header.VisiblePosition = 3;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn25.Width = 30;
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
			this.uGridReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridReporte.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridReporte.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.uGridReporte.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8.5F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridReporte.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.uGridReporte.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridReporte.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridReporte.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.uGridReporte.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridReporte.Location = new System.Drawing.Point(8, 104);
			this.uGridReporte.Name = "uGridReporte";
			this.uGridReporte.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.uGridReporte.Size = new System.Drawing.Size(1216, 272);
			this.uGridReporte.TabIndex = 883;
			this.uGridReporte.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridReporte_ClickCellButton);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22});
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
			// optTransporte
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTransporte.Appearance = appearance16;
			this.optTransporte.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTransporte.CheckedIndex = 0;
			this.optTransporte.ItemAppearance = appearance17;
			this.optTransporte.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = -1;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendiente";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Acreditado";
			valueListItem4.DataValue = 4;
			valueListItem4.DisplayText = "Mensaje Enviado";
			valueListItem5.DataValue = 5;
			valueListItem5.DisplayText = "Cancelado";
			this.optTransporte.Items.Add(valueListItem1);
			this.optTransporte.Items.Add(valueListItem2);
			this.optTransporte.Items.Add(valueListItem3);
			this.optTransporte.Items.Add(valueListItem4);
			this.optTransporte.Items.Add(valueListItem5);
			this.optTransporte.ItemSpacingVertical = 10;
			this.optTransporte.Location = new System.Drawing.Point(72, 40);
			this.optTransporte.Name = "optTransporte";
			this.optTransporte.Size = new System.Drawing.Size(352, 24);
			this.optTransporte.TabIndex = 1019;
			this.optTransporte.Text = "Todos";
			this.optTransporte.ValueChanged += new System.EventHandler(this.optTransporte_ValueChanged);
			// 
			// btnDeposito
			// 
			this.btnDeposito.CausesValidation = false;
			this.btnDeposito.Image = ((System.Drawing.Image)(resources.GetObject("btnDeposito.Image")));
			this.btnDeposito.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDeposito.Location = new System.Drawing.Point(1144, 72);
			this.btnDeposito.Name = "btnDeposito";
			this.btnDeposito.Size = new System.Drawing.Size(76, 24);
			this.btnDeposito.TabIndex = 1020;
			this.btnDeposito.Text = "&Procesar";
			this.btnDeposito.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDeposito.Visible = false;
			this.btnDeposito.Click += new System.EventHandler(this.btnDeposito_Click);
			// 
			// chkCargos
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkCargos.Appearance = appearance18;
			this.chkCargos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCargos.Location = new System.Drawing.Point(64, 104);
			this.chkCargos.Name = "chkCargos";
			this.chkCargos.Size = new System.Drawing.Size(13, 13);
			this.chkCargos.TabIndex = 1095;
			this.chkCargos.CheckedChanged += new System.EventHandler(this.chkCargos_CheckedChanged);
			// 
			// btnSalir
			// 
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(1144, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 1096;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// frmAlfioProcesar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1232, 382);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.chkCargos);
			this.Controls.Add(this.btnDeposito);
			this.Controls.Add(this.optTransporte);
			this.Controls.Add(this.uGridReporte);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.lblBuscar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.KeyPreview = true;
			this.Name = "frmAlfioProcesar";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "ALFIO Transferencia";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAlfioProcesar_KeyDown);
			this.Load += new System.EventHandler(this.frmAlfioProcesar_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTransporte)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}


		private void frmAlfioProcesar_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "20120102");

			if (!Funcion.Permiso("1918", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Reporte Pedidos ALFIO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar) this.btnExcel.Visible = true;
			if (miAcceso.Nuevo) this.btnDeposito.Visible = true;


			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2023");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2023");

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtDesde.Value=DateTime.Today.AddMonths(-1);

			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.Value=DateTime.Today;

			this.uGridReporte.DisplayLayout.Bands[0].Columns["btnVer"].Hidden = true;
			this.uGridReporte.DisplayLayout.Bands[0].Columns["btnDepositar"].Hidden = true;
			
				this.uGridReporte.DisplayLayout.Bands[0].Columns["btnEfectivo"].Hidden = true;

		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			#region valida Fechas
				if(!Validacion.vbFechaEnDocumentos(this.dtDesde,"Seleccione la fecha inicial",false,"",cdsSeteoF))return;
				if(!Validacion.vbFechaEnDocumentos(this.dtHasta,"Seleccione la fecha final",false,"",cdsSeteoF))return;
				if(!Validacion.vbComparaFechas((DateTime)this.dtDesde.Value,(DateTime)this.dtHasta.Value,this.dtHasta,"La fecha incial no puede ser posterior a la fecha final",false))return;
			#endregion valida Fechas

			string sSQL =string.Format("Exec [ReporteAprobacionFio] '{0}','{1}','{2}',{3}",
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyMMdd"),Convert.ToDateTime(this.dtHasta.Value).ToString("yyyMMdd"),
				this.txtBuscar.Text.ToString(), (int) this.optTransporte.Value );
			this.uGridReporte.DataSource=Funcion.dvProcedimiento(cdsSeteoF,sSQL);

			this.lblContador.Text = this.uGridReporte.Rows.Count + " REGISTROS ENCONTRADOS";

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}

		private void frmAlfioProcesar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void btnDeposito_Click(object sender, System.EventArgs e)
		{
			#region Validación
			if ((int) this.optTransporte.Value != 1 )
			{
				MessageBox.Show("Seleccione el estado Pendiente para Procesar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			int idPersonal = 0;
			if ((int) this.uGridReporte.Rows.Count == 0)
			{
				MessageBox.Show("No existen registros a procesar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			int icont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridReporte.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value == true)
				{
					icont = icont + 1;
				}
			}
			if ( icont  == 0 )
			{
				MessageBox.Show("Seleccione al menos un registro.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			int idCuponesWeb = 0;
			int iFormulario = 0;
			string snombre = "";
			string scodigo ="";
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridReporte.Rows.All)
			{
				idCuponesWeb = 0;
				snombre ="";
				iFormulario = 0;
				if ((bool)dr.Cells["Sel"].Value == true)
				{
					idCuponesWeb = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(1) from dbo.FormularioSolicitud where borrar = 0 and iestado = 3 and ConfirmacionAPP = '{0}' and cedula != '{1}'",
						Convert.ToString( dr.Cells["ConfirmacionAPP"].Value) , Convert.ToString( dr.Cells["cedula"].Value)));	

					iFormulario = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(1) from dbo.FormularioSolicitud where borrar = 0 and iestado = 3 and idFormularioSolicitud = {0}",
						Convert.ToInt32( dr.Cells["idFormularioSolicitud"].Value) ));	


					if( idCuponesWeb > 0 )
					{
						snombre = Funcion.sEscalarSQL(cdsSeteoF, string.Format("select top 1 cl.Nombre from dbo.FormularioSolicitud f Inner Join cliente cl on cl.Ruc = f.cedula where f.borrar = 0 and iestado = 3 and ConfirmacionAPP = '{0}' and cedula != '{1}' and proveedor = 0",
							Convert.ToString( dr.Cells["ConfirmacionAPP"].Value) , Convert.ToString( dr.Cells["cedula"].Value)));	
						scodigo = Funcion.sEscalarSQL(cdsSeteoF, string.Format("select top 1 f.ConfirmacionAPP from dbo.FormularioSolicitud f Inner Join cliente cl on cl.Ruc = f.cedula where f.borrar = 0 and iestado = 3 and ConfirmacionAPP = '{0}' and cedula != '{1}' and proveedor = 0",
							Convert.ToString( dr.Cells["ConfirmacionAPP"].Value) , Convert.ToString( dr.Cells["cedula"].Value)));
						this.uGridReporte.ActiveRow = dr;
						this.uGridReporte.ActiveRow.Selected = true;
						MessageBox.Show("El codigo Nuovo: " + scodigo + ", ya fue registrado a "+ snombre +" ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;	
					}
					if(iFormulario > 0)
					{
						MessageBox.Show("El item seleccionado ya fue procesado. Actualiza la consulta. ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;	
					}
				}
			}
			#endregion Validación
			#region Cliente
			int idCliente = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridReporte.Rows.All)
			{
				idCliente = 0;
				if ((bool)dr.Cells["Sel"].Value == true)
				{
					 
					idCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) from  dbo.Cliente where Proveedor = 0 and Ruc = '{0}'",
						Convert.ToString( dr.Cells["cedula"].Value)));	

					if( idCliente == 0)
					{
						this.uGridReporte.ActiveRow = dr;
						this.uGridReporte.ActiveRow.Selected = true;
						MessageBox.Show("El Cliente no existe : " + Convert.ToString( dr.Cells["cedula"].Value) + " ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;	
					}
				}
			}
			#endregion Cliente

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridReporte.Rows.All)
			{
				int bNuovovl = 0;
				if ((bool)dr.Cells["Sel"].Value == true)
				{
					bNuovovl = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(1) from dbo.FioCredito where Nuovo = '{0}' and Estado > 0 and Borrar = 0 ",
						Convert.ToString( dr.Cells["ConfirmacionAPP"].Value)));	
					if ( bNuovovl > 0 )
					{
						string scliente = Funcion.sEscalarSQL(cdsSeteoF, string.Format("select top 1 cl.Ruc from dbo.FioCredito f Inner Join dbo.Cliente cl on cl.idCliente = f.idCliente where Nuovo = '{0}' and f.Estado > 0 and f.Borrar = 0 order by idFioCredito desc ",
							Convert.ToString( dr.Cells["ConfirmacionAPP"].Value)));
              
						if ( scliente != Convert.ToString( dr.Cells["cedula"].Value))
						{
							MessageBox.Show("El codigo Nuovo: " + Convert.ToString( dr.Cells["ConfirmacionAPP"].Value) + ", ya fue registrado por "+ Convert.ToString( dr.Cells["nombre"].Value) +" ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;	
						}
					}
				}
			}

			icont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridReporte.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value == true)
				{
					icont = icont + 1;
				}
			}

			if (DialogResult.Yes ==	MessageBox.Show(string.Format("¿Esta Seguro de procesar los siguientes registro:  {0} ", icont), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				#region g
				//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridReporte.Rows.All)
				//				{
				//					if ((bool)dr.Cells["Sel"].Value == true)
				//					{
				//						string ssSQL = string.Format("Update FormularioSolicitud Set iEstado = 3 Where idFormularioSolicitud = {0}", Convert.ToInt32( dr.Cells["idFormularioSolicitud"].Value));
				//						Funcion.EjecutaSQL(cdsSeteoF, ssSQL);
				//
				//						string ssSQLc = string.Format("Update FormularioCuenta Set iEstado = 3, idFormularioSolicitud = {1} Where idFormularioCuenta = {0}", 
				//							Convert.ToInt32( dr.Cells["idFormularioCuenta"].Value), Convert.ToInt32( dr.Cells["idFormularioSolicitud"].Value));
				//						Funcion.EjecutaSQL(cdsSeteoF, ssSQLc);
				//					}
				//				}
				//				MessageBox.Show("Transferencia Exitosa.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//				this.btnConsultar_Click(sender,e);
				//				return;
				#endregion g

				#region Guardar

				//			DateTime dtInicia = DateTime.Today;
				//			DateTime dtRetorna = DateTime.Today;
				//				dtInicia = (DateTime)this.dtInicia.Value;
				//				dtRetorna = (DateTime)this.dtRetorna.Value;

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

					

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridReporte.Rows.All)
						{
							if ((bool)dr.Cells["Sel"].Value == true)
							{
								string ssSQL = string.Format("Update FormularioSolicitud Set iEstado = 3 Where idFormularioSolicitud = {0}", Convert.ToInt32( dr.Cells["idFormularioSolicitud"].Value));
								oCmdActualiza.CommandText = ssSQL;
								oCmdActualiza.ExecuteNonQuery();	

								string ssSQLc = string.Format("Update FormularioCuenta Set iEstado = 3, idFormularioSolicitud = {1} Where idFormularioCuenta = {0}", 
									Convert.ToInt32( dr.Cells["idFormularioCuenta"].Value), Convert.ToInt32( dr.Cells["idFormularioSolicitud"].Value));
								oCmdActualiza.CommandText = ssSQLc;
								oCmdActualiza.ExecuteNonQuery();	

								string  sSQLDeta = string.Format("Exec [GuardaFioCreditos] {0}, {1}" ,
									Convert.ToInt32 ( dr.Cells["idFormularioSolicitud"].Value), 
									Convert.ToInt32 ( dr.Cells["idFormularioCuenta"].Value)
									);
								oCmdActualiza.CommandText = sSQLDeta;
								oCmdActualiza.ExecuteNonQuery();	
							}
						}
			          
						oTransaction.Commit();

						Cursor = Cursors.Default;	

						MessageBox.Show("Transferencia Exitosa.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.btnConsultar_Click(sender,e);
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

		private void optTransporte_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int) this.optTransporte.Value >= 3 )
			{
				this.uGridReporte.DisplayLayout.Bands[0].Columns["btnVer"].Hidden = false;
			}
			else
			{
				this.uGridReporte.DisplayLayout.Bands[0].Columns["btnVer"].Hidden = true;
					this.uGridReporte.DisplayLayout.Bands[0].Columns["btnEfectivo"].Hidden = true;
			}
			if ((int) this.optTransporte.Value >= 4 )
			{
				this.uGridReporte.DisplayLayout.Bands[0].Columns["btnDepositar"].Hidden = false;
					this.uGridReporte.DisplayLayout.Bands[0].Columns["btnEfectivo"].Hidden = false;
			}
			else
			{
				this.uGridReporte.DisplayLayout.Bands[0].Columns["btnDepositar"].Hidden = true;
					this.uGridReporte.DisplayLayout.Bands[0].Columns["btnEfectivo"].Hidden = true;

			}


			this.btnConsultar_Click(sender, e);
		}

		private void chkCargos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridReporte.Rows.All)
				dr.Cells["Sel"].Value = this.chkCargos.Checked;
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridReporte);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();		
		}

		private void uGridReporte_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnVer")
			{	
				using (frmFioCliente Notas  = new frmFioCliente( (int)e.Cell.Row.Cells["idFormularioSolicitud"].Value))
				{
					if (DialogResult.OK == Notas.ShowDialog())
					{
						
					}
				}		
			}

			if (e.Cell.Column.ToString() == "btnDepositar")
			{	
				if (Convert.ToString( e.Cell.Row.Cells["UrlNube"].Value).Length > 5 &&  Convert.ToString( e.Cell.Row.Cells["Voucher"].Value) == "ARCHIVO SUBIDO" )
				{
					System.Diagnostics.Process.Start("chrome.exe", Convert.ToString( e.Cell.Row.Cells["UrlNube"].Value));
				}
				else
				{
					#region Deposito
					try
					{
						MessageBox.Show("Acontinuación seleccione el voucher de transporte en formato JPG.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
						string txtDirectorioarchivo ="";
						string txtURL ="";
						string SArhivoNuevo = Convert.ToString( e.Cell.Row.Cells["idFormularioSolicitud"].Value)+"_"+ Convert.ToString( e.Cell.Row.Cells["idFormularioCuenta"].Value)+ "_"+ Convert.ToString( e.Cell.Row.Cells["numero"].Value);
						OpenFileDialog buscar = new OpenFileDialog ();
						buscar.Filter = "Text files (*.jpg)|*.jpg";
						if (buscar.ShowDialog () == DialogResult.OK)
						{

							txtDirectorioarchivo = buscar.FileName;//nombre del archivo
							FileInfo fi= new FileInfo(txtDirectorioarchivo);
							string nombre = fi.Name;//nombre.jpg
							string Directorio = txtDirectorioarchivo.Replace(nombre, "")+SArhivoNuevo+".jpg";//archivo nuevo
							File.Delete(Directorio); // First delete the new file if it exists, because it wont be overwritten by Move
							File.Move(txtDirectorioarchivo, Directorio); 
						
							string carpeta_1="";
							string carpeta_2="";
							string carpeta_3="";
							string carpeta ="";
							//string txtDirectorioarchivo ="";
							//	string txtURL ="";
							string sUrlNube = "";
							string ScARPETAUR = @"192.168.1.10";
							string sCarpeta ="";
							//
							carpeta_1 = @"\\"+ScARPETAUR+@"\Proyectos\VOUCHERFIO";
							//@"\\"+
							carpeta_2 = @"\\"+ScARPETAUR+@"\Proyectos\VOUCHERFIO\"+MenuLatinium.stDirFacturacion;
							// @"\\192.168.1.10\Proyectos\Transporte\"+MenuLatinium.stDirFacturacion+@"\" + this.txtNumero.Text.Replace(" ", "_") +"_"+this.txtRuc.Text+"_"+ Convert.ToString (this.txtidRegistroTransporte.Value);
							carpeta = @"\\"+ScARPETAUR+@"\Proyectos\VOUCHERFIO\"+MenuLatinium.stDirFacturacion+@"\" + Convert.ToString( e.Cell.Row.Cells["idFormularioSolicitud"].Value)+"_"+ Convert.ToString( e.Cell.Row.Cells["idFormularioCuenta"].Value)+ "_"+ Convert.ToString( e.Cell.Row.Cells["numero"].Value);
							carpeta_3 = @""+MenuLatinium.stDirFacturacion+@"\" + Convert.ToString( e.Cell.Row.Cells["idFormularioSolicitud"].Value)+"_"+ Convert.ToString( e.Cell.Row.Cells["idFormularioCuenta"].Value)+ "_"+ Convert.ToString( e.Cell.Row.Cells["numero"].Value);
							Process.Start("cmd.exe", "/C" + "net use \\192.168.1.10  /user:latinium 10dgpv15.p");
							Directory.GetDirectories(@"\\192.168.1.10\Proyectos");

							if (!Directory.Exists(carpeta_1))
							{
								Directory.CreateDirectory(carpeta);
					
							}
							if (!Directory.Exists(carpeta_2))
							{
								Directory.CreateDirectory(carpeta);
					
							}

							if (!Directory.Exists(carpeta))
							{
								Directory.CreateDirectory(carpeta);
					
							}
							txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",Directorio,@"\"));
							carpeta  = (carpeta + @"\" + txtURL.Replace(" ", "_"));
							if (System.IO.File.Exists(carpeta))
							{
								//entonces que me abra el archivo PDF

								MessageBox.Show(string.Format("El archivo '{0}' ya existe. ", carpeta), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
								return;

							}
							else
							{
								System.IO.File.Copy (  Directorio, carpeta, false);
								sCarpeta = carpeta;
								sUrlNube = carpeta_3+ @"\" + txtURL;
								string ssSQL = string.Format("Exec [GrabaUrlNubeALFIO] {0}, '{1}', '{2}'", Convert.ToInt32( e.Cell.Row.Cells["idFormularioSolicitud"].Value), sCarpeta, sUrlNube );
								Funcion.EjecutaSQL(cdsSeteoF, ssSQL);
								MessageBox.Show(string.Format(" Archivo '{0}' subido . ", sUrlNube), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.btnConsultar_Click(sender,e);
							}

						}
					}
					catch (Exception ex)
					{
						MessageBox.Show ("Error: "+ ex.Message);
					}
					#endregion Deposito
				}
			}

			if (e.Cell.Column.ToString() == "btnEfectivo")
			{	
				if(Convert.ToDecimal( e.Cell.Row.Cells["Abono"].Value ) > 0 )
				{
					if (DialogResult.Yes == MessageBox.Show("¿Desea Registrar un cobro?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
					{
						using (frmPagoFioCredito Notas  = new frmPagoFioCredito( (int)e.Cell.Row.Cells["idFioCredito"].Value))
						{
							if (DialogResult.OK == Notas.ShowDialog())
							{
								#region Guardar

								//			DateTime dtInicia = DateTime.Today;
								//			DateTime dtRetorna = DateTime.Today;
								//				dtInicia = (DateTime)this.dtInicia.Value;
								//				dtRetorna = (DateTime)this.dtRetorna.Value;
								string sNumeracion ="";
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

										#region Numeracion
								
										string ssSQL = string.Format("Exec NumeracionLocales 87, 0, 0");
										oCmdActualiza.CommandText = ssSQL;
										sNumeracion = oCmdActualiza.ExecuteScalar().ToString();						
					
										#endregion Numeracion

										#region Graba Maestro
										int iEstado = 2;
										if ( Convert.ToDecimal( Notas.txtValor.Value) < Convert.ToDecimal( Notas.txtvalorPendiente.Value) ) iEstado = 1;
								
										string  sSQL1 = string.Format("Exec [GrabaPagosAlFioCredito] {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', {7},'{8}', '{9}'" ,
											0 , Convert.ToInt32( e.Cell.Row.Cells["idFioCredito"].Value), iEstado, 
											Convert.ToDecimal(  Notas.txtValor.Value),
											Notas.txtOrigen.Text,  	Convert.ToDateTime(Notas.dtFecha.Value).ToString("yyyyMMdd"), sNumeracion, Convert.ToInt32(  Notas.optTransporte.Value),
											Notas.sCarpeta, Notas.sUrlNube
											);
										oCmdActualiza.CommandText = sSQL1;
										oCmdActualiza.ExecuteNonQuery();
										#endregion Graba Maestro

										if ( iEstado == 2 )
										{
											string sSQLAsi = string.Format("update FioCredito set Estado = 2 where idFioCredito = {0} ", 
												Convert.ToInt32( e.Cell.Row.Cells["idFioCredito"].Value) );			
											oCmdActualiza.CommandText = sSQLAsi;
											oCmdActualiza.ExecuteNonQuery();
										}
          
										oTransaction.Commit();

										Cursor = Cursors.Default;	

										MessageBox.Show("Cobro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

				
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
								this.btnConsultar_Click (sender, e);
							}
						}		
					}
					else
					{
						using (frmHistorialPagoFio Notas  = new frmHistorialPagoFio( (int)e.Cell.Row.Cells["idFioCredito"].Value))
						{
							if (DialogResult.OK == Notas.ShowDialog())
							{
									this.btnConsultar_Click(sender, e);
							}
						}
					}
				}
				else
				{
					using (frmPagoFioCredito Notas  = new frmPagoFioCredito( (int)e.Cell.Row.Cells["idFioCredito"].Value))
					{
						if (DialogResult.OK == Notas.ShowDialog())
						{
							#region Guardar

							//			DateTime dtInicia = DateTime.Today;
							//			DateTime dtRetorna = DateTime.Today;
							//				dtInicia = (DateTime)this.dtInicia.Value;
							//				dtRetorna = (DateTime)this.dtRetorna.Value;
							string sNumeracion ="";
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

									#region Numeracion
								
									string ssSQL = string.Format("Exec NumeracionLocales 87, 0, 0");
									oCmdActualiza.CommandText = ssSQL;
									sNumeracion = oCmdActualiza.ExecuteScalar().ToString();						
					
									#endregion Numeracion

									#region Graba Maestro
									int iEstado = 2;
									if ( Convert.ToDecimal( Notas.txtValor.Value) < Convert.ToDecimal( Notas.txtvalorPendiente.Value) ) iEstado = 1;
								
									string  sSQL1 = string.Format("Exec [GrabaPagosAlFioCredito] {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', {7},'{8}', '{9}'" ,
										0 , Convert.ToInt32( e.Cell.Row.Cells["idFioCredito"].Value), iEstado, 
										Convert.ToDecimal(  Notas.txtValor.Value),
										Notas.txtOrigen.Text,  	Convert.ToDateTime(Notas.dtFecha.Value).ToString("yyyyMMdd"), sNumeracion, Convert.ToInt32(  Notas.optTransporte.Value),
										Notas.sCarpeta, Notas.sUrlNube
										);
									oCmdActualiza.CommandText = sSQL1;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Graba Maestro

									if ( iEstado == 2 )
									{
										string sSQLAsi = string.Format("update FioCredito set Estado = 2 where idFioCredito = {0} ", 
											Convert.ToInt32( e.Cell.Row.Cells["idFioCredito"].Value) );			
										oCmdActualiza.CommandText = sSQLAsi;
										oCmdActualiza.ExecuteNonQuery();
									}
          
									oTransaction.Commit();

									Cursor = Cursors.Default;	

									MessageBox.Show("Cobro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

				
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
								this.btnConsultar_Click(sender, e);
						}
					}	
				}
			}
		}
	}
}

