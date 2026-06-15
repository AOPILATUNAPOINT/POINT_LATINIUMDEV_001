using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;
using C1.Data;
using System.Globalization;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmGastosPorServicio.
	/// </summary>
	public class frmGastosPorServicio : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtEntrega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDevolucion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtInicio;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTiposDeGastos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdLiquidacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblEstado;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private System.Windows.Forms.Button btnActualizar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnGrabar;
		private System.Windows.Forms.Button btnGastos;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnImprimir;

		private C1DataRow drSeteoF;

		public frmGastosPorServicio()
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
		
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmGastosPorServicio));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetLiquidacionGastos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLiquidacionGastos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGasto", -1, "cmbTiposDeGastos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Adicional");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 13, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetLiquidacionGastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLiquidacionGastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Iva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoGasto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Adicional");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGasto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripción");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoGasto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripción");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyectoZoom");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyectoZoom");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.label9 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.dtEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDevolucion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTiposDeGastos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtIdLiquidacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnGrabar = new System.Windows.Forms.Button();
			this.btnGastos = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnImprimir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dtEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDevolucion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiposDeGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdLiquidacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			this.SuspendLayout();
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(208, 40);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(43, 16);
			this.label9.TabIndex = 268;
			this.label9.Text = "Entrega";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label9.Click += new System.EventHandler(this.label9_Click);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(408, 40);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(59, 16);
			this.label8.TabIndex = 267;
			this.label8.Text = "Devolución";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Location = new System.Drawing.Point(8, 40);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(31, 16);
			this.label20.TabIndex = 266;
			this.label20.Text = "Inicio";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Font = new System.Drawing.Font("Tahoma", 11F, System.Drawing.FontStyle.Bold);
			this.label7.Location = new System.Drawing.Point(808, 466);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(60, 21);
			this.label7.TabIndex = 264;
			this.label7.Text = "TOTAL:";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label7.Visible = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(47, 16);
			this.label1.TabIndex = 263;
			this.label1.Text = "Proyecto";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtEntrega
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtEntrega.Appearance = appearance1;
			this.dtEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtEntrega.DateButtons.Add(dateButton1);
			this.dtEntrega.Format = "dd/MM/yyyy";
			this.dtEntrega.Location = new System.Drawing.Point(272, 38);
			this.dtEntrega.Name = "dtEntrega";
			this.dtEntrega.NonAutoSizeHeight = 23;
			this.dtEntrega.Size = new System.Drawing.Size(112, 21);
			this.dtEntrega.SpinButtonsVisible = true;
			this.dtEntrega.TabIndex = 256;
			this.dtEntrega.Value = ((object)(resources.GetObject("dtEntrega.Value")));
			// 
			// dtDevolucion
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDevolucion.Appearance = appearance2;
			this.dtDevolucion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDevolucion.DateButtons.Add(dateButton2);
			this.dtDevolucion.Format = "dd/MM/yyyy";
			this.dtDevolucion.Location = new System.Drawing.Point(488, 38);
			this.dtDevolucion.Name = "dtDevolucion";
			this.dtDevolucion.NonAutoSizeHeight = 23;
			this.dtDevolucion.Size = new System.Drawing.Size(112, 21);
			this.dtDevolucion.SpinButtonsVisible = true;
			this.dtDevolucion.TabIndex = 257;
			this.dtDevolucion.Value = ((object)(resources.GetObject("dtDevolucion.Value")));
			// 
			// dtInicio
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtInicio.Appearance = appearance3;
			this.dtInicio.AutoSelect = false;
			this.dtInicio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtInicio.DateButtons.Add(dateButton3);
			this.dtInicio.Format = "dd/MM/yyyy";
			this.dtInicio.Location = new System.Drawing.Point(72, 38);
			this.dtInicio.Name = "dtInicio";
			this.dtInicio.NonAutoSizeHeight = 23;
			this.dtInicio.Size = new System.Drawing.Size(112, 21);
			this.dtInicio.SpinButtonsVisible = true;
			this.dtInicio.TabIndex = 255;
			this.dtInicio.Value = ((object)(resources.GetObject("dtInicio.Value")));
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance4;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Tahoma", 11F);
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(944, 463);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(96, 27);
			this.txtTotal.TabIndex = 265;
			this.txtTotal.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource3;
			appearance5.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance5;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 109;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 113;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 74;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 69;
			ultraGridColumn5.Header.Caption = "R.U.C.";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Razón Social";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 200;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 70;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "N. Factura";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 90;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance6;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 75;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance7;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "I.V.A.";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 75;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance8;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Valor";
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Width = 75;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.Caption = "Gasto";
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Width = 180;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance9;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Proformado";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.PromptChar = ' ';
			ultraGridColumn13.Width = 75;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance10;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 75;
			ultraGridColumn15.Header.Caption = "...";
			ultraGridColumn15.Header.VisiblePosition = 14;
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
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance11;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance15;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1056, 368);
			this.ultraGrid1.TabIndex = 258;
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(System.Double);
			ultraDataColumn10.DataType = typeof(System.Double);
			ultraDataColumn11.DataType = typeof(System.Double);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.Double);
			ultraDataColumn14.DataType = typeof(System.Double);
			ultraDataColumn15.DataType = typeof(bool);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
			// cmbTiposDeGastos
			// 
			this.cmbTiposDeGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTiposDeGastos.DataSource = this.ultraDataSource2;
			ultraGridBand2.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 320;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbTiposDeGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTiposDeGastos.DisplayMember = "Descripción";
			this.cmbTiposDeGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTiposDeGastos.Location = new System.Drawing.Point(24, 192);
			this.cmbTiposDeGastos.Name = "cmbTiposDeGastos";
			this.cmbTiposDeGastos.Size = new System.Drawing.Size(312, 104);
			this.cmbTiposDeGastos.TabIndex = 270;
			this.cmbTiposDeGastos.ValueMember = "idTipoGasto";
			this.cmbTiposDeGastos.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn16.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.BackColor = System.Drawing.Color.Transparent;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(696, 37);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 271;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// ultraDataSource1
			// 
			ultraDataColumn18.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// txtIdLiquidacion
			// 
			this.txtIdLiquidacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdLiquidacion.Appearance = appearance18;
			this.txtIdLiquidacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdLiquidacion.Enabled = false;
			this.txtIdLiquidacion.FormatString = "";
			this.txtIdLiquidacion.Location = new System.Drawing.Point(744, 464);
			this.txtIdLiquidacion.Name = "txtIdLiquidacion";
			this.txtIdLiquidacion.PromptChar = ' ';
			this.txtIdLiquidacion.Size = new System.Drawing.Size(40, 22);
			this.txtIdLiquidacion.TabIndex = 288;
			this.txtIdLiquidacion.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstado.Appearance = appearance19;
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.FormatString = "";
			this.txtEstado.Location = new System.Drawing.Point(696, 464);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(40, 22);
			this.txtEstado.TabIndex = 289;
			this.txtEstado.Visible = false;
			// 
			// cmbProyecto
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance20;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataSource = this.ultraDataSource1;
			ultraGridBand3.AddButtonCaption = "SubProyecto";
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 320;
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridColumn20.Width = 250;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(72, 8);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(312, 21);
			this.cmbProyecto.TabIndex = 291;
			this.cmbProyecto.ValueMember = "idProyectoZoom";
			this.cmbProyecto.ValueChanged += new System.EventHandler(this.cmbProyecto_ValueChanged);
			this.cmbProyecto.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbProyecto_InitializeLayout);
			// 
			// btnActualizar
			// 
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizar.Location = new System.Drawing.Point(171, 464);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(75, 24);
			this.btnActualizar.TabIndex = 314;
			this.btnActualizar.Text = "&Actualizar";
			this.btnActualizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(335, 464);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 24);
			this.btnSalir.TabIndex = 315;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnGrabar
			// 
			this.btnGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGrabar.Image = ((System.Drawing.Image)(resources.GetObject("btnGrabar.Image")));
			this.btnGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGrabar.Location = new System.Drawing.Point(8, 464);
			this.btnGrabar.Name = "btnGrabar";
			this.btnGrabar.Size = new System.Drawing.Size(72, 24);
			this.btnGrabar.TabIndex = 316;
			this.btnGrabar.Text = "&Grabar";
			this.btnGrabar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGrabar.Click += new System.EventHandler(this.btnGrabar_Click);
			// 
			// btnGastos
			// 
			this.btnGastos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGastos.Image = ((System.Drawing.Image)(resources.GetObject("btnGastos.Image")));
			this.btnGastos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGastos.Location = new System.Drawing.Point(253, 464);
			this.btnGastos.Name = "btnGastos";
			this.btnGastos.Size = new System.Drawing.Size(75, 24);
			this.btnGastos.TabIndex = 317;
			this.btnGastos.Text = "Gastos";
			this.btnGastos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGastos.Click += new System.EventHandler(this.btnGastos_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1072, 8);
			this.groupBox1.TabIndex = 318;
			this.groupBox1.TabStop = false;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(87, 465);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 319;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// frmGastosPorServicio
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1072, 494);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGastos);
			this.Controls.Add(this.btnGrabar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdLiquidacion);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.cmbTiposDeGastos);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtEntrega);
			this.Controls.Add(this.dtDevolucion);
			this.Controls.Add(this.dtInicio);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "frmGastosPorServicio";
			this.Text = "Liquidación de Proyectos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGastosPorServicio_KeyDown);
			this.Load += new System.EventHandler(this.frmGastosPorServicio_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDevolucion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiposDeGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdLiquidacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void label9_Click(object sender, System.EventArgs e)
		{
		
		}

		private void frmGastosPorServicio_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0626");

//			if (!Funcion.Permiso("991", cdsSeteoF))
//			{				
//				MessageBox.Show("No tiene Acceso LIQUIDACIÓN DE GASTOS", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				BeginInvoke(new MethodInvoker(UnloadMe));
//				return;
//			}
			this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyectoZoom, Numero, Nombre From ProyectoZoom Order By Nombre");
			this.cmbTiposDeGastos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoGasto, Descripción From tiposdegasto Where Activo = 1 Order By idTipoGasto");
			this.cmbProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, "Select Top 1 idProyectoZoom From ProyectoZoom Order By idProyectoZoom Desc");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
		
		private void Calculo(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{				
			e.Cell.Row.Cells["Subtotal"].Value = (double)e.Cell.Row.Cells["Adicional"].Value;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				this.txtTotal.Value = (double) this.txtTotal.Value + (double)dr.Cells["Subtotal"].Value;		 
			}
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Adicional")
			{
				Calculo(e);
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);

			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes Busqueda = new BuscaClientes("", 4, 1, 1, 1, false))
				{
					if (DialogResult.OK == Busqueda.ShowDialog())
					{
						this.ultraGrid1.ActiveRow.Cells["idCliente"].Value = (int)Busqueda.ultraGrid1.ActiveRow.Cells["idCliente"].Value;
						this.ultraGrid1.ActiveRow.Cells["Ruc"].Value = Busqueda.ultraGrid1.ActiveRow.Cells["Ruc"].Value.ToString();
						this.ultraGrid1.ActiveRow.Cells["Nombre"].Value = Busqueda.ultraGrid1.ActiveRow.Cells["Nombre"].Value.ToString();
					}
				}
			}
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.ultraGrid1.ActiveRow.Cells["idCompra"].Value != System.DBNull.Value) e.Handled = true;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Adicional", 2);
		}

		private void frmGastosPorServicio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}


		private void cmbProyecto_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtEstado.Value = 0;
			this.dtInicio.Value = System.DBNull.Value;
			this.dtEntrega.Value = System.DBNull.Value;
			this.dtDevolucion.Value = System.DBNull.Value;

			string sSQL = string.Format("Exec ConsultaLiquidacionGasto {0}", (int)this.cmbProyecto.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtIdLiquidacion.Value = dr.GetInt32(0);
				if (dr.GetValue(1) != System.DBNull.Value) this.dtInicio.Value = dr.GetDateTime(1);
				if (dr.GetValue(2) != System.DBNull.Value) this.dtEntrega.Value = dr.GetDateTime(2);
				if (dr.GetValue(3) != System.DBNull.Value) this.dtDevolucion.Value = dr.GetDateTime(3);
				this.txtEstado.Value = dr.GetInt32(4);
			}
			dr.Close();

			if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "POR LIQUIDAR";
			if ((int)this.txtEstado.Value == 2) this.lblEstado.Text = "LIQUIDADO";

			if ((int)this.txtEstado.Value == 1) sSQL = string.Format("Exec DetalleLiquidacionDeGastosPorLiquidar {0}", (int)this.cmbProyecto.Value);
			if ((int)this.txtEstado.Value == 2) sSQL = string.Format("Exec ConsultaDetalleLiquidacionGastos {0}", (int)this.cmbProyecto.Value);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			this.cmbProyecto_ValueChanged(sender, e);
		}

		private void btnGastos_Click(object sender, System.EventArgs e)
		{
			using (frmTiposDeGastos TipoDeGasto = new frmTiposDeGastos())
			{
				if (DialogResult.OK == TipoDeGasto.ShowDialog())
				{
					if (this.cmbProyecto.ActiveRow == null) return;

					this.cmbProyecto_ValueChanged(sender, e);
				}
			}
		}

		private void btnGrabar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtEntrega, "Seleccione la fecha de entrega", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtEntrega, "Seleccione la fecha de devolución", false, "", cdsSeteoF)) return;

			#region Estado
			int iFacturado = 0;
			int iSel = 0;
			string sMensaje = "Proyecto Por Liquidar.";
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in ultraGrid1.Rows.All)
			{
				if (dr1.Cells["idCompra"].Value != System.DBNull.Value) iFacturado ++;
				if ((bool)dr1.Cells["Sel"].Value) iSel ++;
			}
					
			if (iFacturado == iSel)
			{
				this.lblEstado.Text = "LIQUIDADO";
				this.txtEstado.Value = 2;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr2 in ultraGrid1.Rows.All)
				{
					if (dr2.Cells["idCompra"].Value == System.DBNull.Value) 
					{
						dr2.Cells["Adicional"].Value = 0;
						dr2.Cells["Subtotal"].Value = 0;
					}
				}
				sMensaje = "Proyecto Liquidado.";
			}
			else 
			{
				this.lblEstado.Text = "POR LIQUIDAR";
				this.txtEstado.Value = 1;
			}
			#endregion Estado	

			#region Guarda Liquidacion
			string sSQL = "";
			sSQL = string.Format("Exec ActualizaLiqGastos {0}, {1}, {2}, '{3}', '{4}', '{5}'", 
				(int)this.txtIdLiquidacion.Value, (double)this.txtTotal.Value, (int)this.txtEstado.Value,
				Convert.ToDateTime(this.dtInicio.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtEntrega.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtDevolucion.Value).ToString("yyyyMMdd"));
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idCompra"].Value == System.DBNull.Value)
				{
					sSQL = string.Format("Exec ActualizaDetLiqGastos {0}, {1}, {2}, {3}, {4}", 
						(int)dr.Cells["idDetLiquidacionGastos"].Value, (int)dr.Cells["idCliente"].Value, (double)dr.Cells["Adicional"].Value, (double)dr.Cells["Subtotal"].Value, (bool)dr.Cells["Sel"].Value); 
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}
			}
			#endregion Guarda Liquidacion

			MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.cmbProyecto_ValueChanged(sender, e);
		}

		private void cmbProyecto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldCompra = new ParameterField ();
			ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
			paramFieldCompra.ParameterFieldName = "@idLiquidacionGastos";
			ValIdCompra.Value = int.Parse(this.txtIdLiquidacion.Value.ToString());
			paramFieldCompra.CurrentValues.Add (ValIdCompra);
			paramFields.Add (paramFieldCompra);
				
			Reporte Rep = new Reporte("LiquidacionCotizaciones.rpt");
			Rep.MdiParent = this.MdiParent;			
			Rep.crVista.ParameterFieldInfo = paramFields;
			Rep.Show();
		}
	}
}

