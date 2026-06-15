using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Shared;
using System.Diagnostics;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de PagosPendientes.
	/// </summary>
	public class PagosPendientes : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.Label lblPagos;
		private System.Windows.Forms.Label lblDescuento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaPago;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnGenerarPago;
		private System.Windows.Forms.Label lblContador;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public PagosPendientes()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(PagosPendientes));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RUC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioAut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPago");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel", 0);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Autorizado", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Autorizado", 19, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 20, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RUC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioAut");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAut");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.label2 = new System.Windows.Forms.Label();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label3 = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblPagos = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.dtFechaPago = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnGenerarPago = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaPago)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 17);
			this.label2.TabIndex = 53;
			this.label2.Text = "Fecha";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFecha
			// 
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Format = "dd/MM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(64, 8);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(120, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 52;
			this.txtFecha.Value = ((object)(resources.GetObject("txtFecha.Value")));
			this.txtFecha.ValueChanged += new System.EventHandler(this.txtDesde_ValueChanged);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 80;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 88;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.Caption = "Tipo";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 140;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 79;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Documento";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 80;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "R.U.C.";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 220;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 220;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 22;
			ultraGridColumn9.Width = 220;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.Caption = "Solicitado Por";
			ultraGridColumn10.Header.VisiblePosition = 17;
			ultraGridColumn10.Width = 90;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.Caption = "Autorizado Por";
			ultraGridColumn11.Header.VisiblePosition = 19;
			ultraGridColumn11.Width = 90;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.Caption = "Fecha Solicitud";
			ultraGridColumn12.Header.VisiblePosition = 18;
			ultraGridColumn12.Width = 90;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.Caption = "Fecha Autorizado";
			ultraGridColumn13.Header.VisiblePosition = 20;
			ultraGridColumn13.Width = 100;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.Caption = "Vence";
			ultraGridColumn14.Header.VisiblePosition = 21;
			ultraGridColumn14.Width = 70;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 9;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 98;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.Caption = "Forma de Pago";
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Width = 100;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.VisiblePosition = 11;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 51;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.Caption = "Cuenta";
			ultraGridColumn18.Header.VisiblePosition = 12;
			ultraGridColumn18.Width = 180;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn19.CellAppearance = appearance2;
			ultraGridColumn19.Header.Caption = "Número";
			ultraGridColumn19.Header.VisiblePosition = 13;
			ultraGridColumn19.Width = 50;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance3.BackColor2 = System.Drawing.Color.LightSteelBlue;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance3;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 14;
			ultraGridColumn20.Width = 75;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.BackColor = System.Drawing.Color.LightBlue;
			appearance4.BackColor2 = System.Drawing.Color.LightBlue;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance4;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Saldo";
			ultraGridColumn21.Header.VisiblePosition = 15;
			ultraGridColumn21.Width = 75;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.DataType = typeof(bool);
			ultraGridColumn23.DefaultCellValue = false;
			ultraGridColumn23.Header.Caption = "...";
			ultraGridColumn23.Header.VisiblePosition = 16;
			ultraGridColumn23.Width = 20;
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
																										 ultraGridColumn23});
			ultraGridBand1.Header.Caption = "";
			appearance5.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance5;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.LightBlue;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance6;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(5, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1302, 224);
			this.ultraGrid1.TabIndex = 54;
			this.ultraGrid1.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.DateTime);
			ultraDataColumn12.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn13.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn14.DataType = typeof(System.DateTime);
			ultraDataColumn14.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn20.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn21.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn22.DataType = typeof(int);
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
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 17);
			this.label3.TabIndex = 131;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBusqueda
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance13;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(64, 40);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(220, 22);
			this.txtBusqueda.TabIndex = 130;
			this.txtBusqueda.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBusqueda_KeyPress);
			this.txtBusqueda.ValueChanged += new System.EventHandler(this.txtBusqueda_ValueChanged);
			// 
			// lblPagos
			// 
			this.lblPagos.AutoSize = true;
			this.lblPagos.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPagos.ForeColor = System.Drawing.Color.Firebrick;
			this.lblPagos.Location = new System.Drawing.Point(400, 42);
			this.lblPagos.Name = "lblPagos";
			this.lblPagos.Size = new System.Drawing.Size(0, 18);
			this.lblPagos.TabIndex = 631;
			this.lblPagos.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(1040, 43);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(29, 16);
			this.lblDescuento.TabIndex = 633;
			this.lblDescuento.Text = "Total";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtValor
			// 
			this.txtValor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance14.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtValor.Appearance = appearance14;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Underline);
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(1088, 40);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(100, 22);
			this.txtValor.TabIndex = 634;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1312, 8);
			this.groupBox1.TabIndex = 635;
			this.groupBox1.TabStop = false;
			// 
			// dtFechaPago
			// 
			this.dtFechaPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaPago.DateButtons.Add(dateButton2);
			this.dtFechaPago.Format = "dd/MM/yyyy";
			this.dtFechaPago.Location = new System.Drawing.Point(1008, 144);
			this.dtFechaPago.Name = "dtFechaPago";
			this.dtFechaPago.NonAutoSizeHeight = 23;
			this.dtFechaPago.Size = new System.Drawing.Size(120, 21);
			this.dtFechaPago.SpinButtonsVisible = true;
			this.dtFechaPago.TabIndex = 636;
			this.dtFechaPago.Value = ((object)(resources.GetObject("dtFechaPago.Value")));
			this.dtFechaPago.Visible = false;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(304, 39);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 637;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnGenerarPago
			// 
			this.btnGenerarPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerarPago.CausesValidation = false;
			this.btnGenerarPago.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerarPago.Image")));
			this.btnGenerarPago.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerarPago.Location = new System.Drawing.Point(1213, 40);
			this.btnGenerarPago.Name = "btnGenerarPago";
			this.btnGenerarPago.Size = new System.Drawing.Size(95, 23);
			this.btnGenerarPago.TabIndex = 638;
			this.btnGenerarPago.Text = "&Generar Pago";
			this.btnGenerarPago.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerarPago.Click += new System.EventHandler(this.btnGenerarPago_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(400, 9);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 18);
			this.lblContador.TabIndex = 639;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// PagosPendientes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1312, 318);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btnGenerarPago);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtFechaPago);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.lblPagos);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.txtFecha);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "PagosPendientes";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "Lista de Pagos Autorizados";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.PagosPendientes_KeyDown);
			this.Load += new System.EventHandler(this.PagosPendientes_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaPago)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		private void SQL()
		{
			this.txtValor.Value = 0;
			this.lblPagos.Text = "";
			
			string sSQL = string.Format("Exec ListaPagosAutorizados '{0}', '{1}'", 
				Convert.ToDateTime(this.txtFecha.Value).ToString("yyyyMMdd"), this.txtBusqueda.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);

			this.lblContador.Text = this.ultraGrid1.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void txtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			SQL();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
		
		private void Suma()
		{
			this.txtValor.Value = 0.00m;
			this.lblPagos.Text = "";
			int iCont = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					this.txtValor.Value = Convert.ToDecimal(this.txtValor.Value) + Convert.ToDecimal(dr.Cells["Autorizado"].Value);
					iCont ++;
				}
			}

			this.lblPagos.Text = iCont.ToString() + " PAGOS SELECCIONADOS";
		}

		private void ImpresionEgreso(int idAsiento, int idTipo)
		{
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfidAsiento = new ParameterField ();
			ParameterDiscreteValue ValIdAsiento = new ParameterDiscreteValue ();
			pfidAsiento.ParameterFieldName = "@idAsiento";
			ValIdAsiento.Value = idAsiento;
			pfidAsiento.CurrentValues.Add (ValIdAsiento);
			paramFields.Add (pfidAsiento);
				
			ParameterField pfidTipo = new ParameterField ();
			ParameterDiscreteValue ValidTipo = new ParameterDiscreteValue ();
			pfidTipo.ParameterFieldName = "@idTipo";
			ValidTipo.Value = idTipo;
			pfidTipo.CurrentValues.Add (ValidTipo);
			paramFields.Add (pfidTipo);
				
			ParameterField pfidAsientoDet = new ParameterField ();
			ParameterDiscreteValue ValIdAsientoDet = new ParameterDiscreteValue ();
			pfidAsientoDet.ParameterFieldName = "@idAsientoDet";
			ValIdAsientoDet.Value = idAsiento;
			pfidAsientoDet.CurrentValues.Add (ValIdAsientoDet);
			paramFields.Add (pfidAsientoDet);
			
			string strReporte = "ComprobanteDeEgreso.rpt";

			if (MenuLatinium.stDirFacturacion == "CROMATIKA")	strReporte = "ComprobanteDeEgresoCromatika.rpt";
			else if (MenuLatinium.stDirFacturacion == "ZOOMMATIKA")	strReporte = "ComprobanteDeEgresoCromatika.rpt";
			else if (MenuLatinium.stDirFacturacion == "LOTESERVICIO") strReporte = "ComprobanteDeEgreso.rpt";
			else if (MenuLatinium.stDirFacturacion == "GRETASANTORINI") strReporte = "ComprobanteDeEgreso.rpt";
			else if (MenuLatinium.stDirFacturacion == "IWANTYOU") strReporte = "ComprobanteDeEgreso.rpt";
			else if (MenuLatinium.stDirFacturacion == "POINT") strReporte = "ComprobanteDeEgresoA4.rpt";
			else strReporte = "ComprobanteDeEgresoA4.rpt";

			Reporte miRepor = new Reporte(strReporte, "");
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("Egreso");
			miRepor.crVista.ParameterFieldInfo = paramFields;
			miRepor.Show();	
		}

		private void ImpresionCheque(decimal fValorCheque, string sNombre, DateTime dtFechaCheque, int idAsiento, int idCuenta)
		{			
			string strFiltro = "{Asiento.idAsiento} = " + idAsiento;
			
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
						
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "ValorCheque";
			discreteVal.Value = fValorCheque;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
			
			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "Nombre";
			discreteVal1.Value = sNombre;
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);

			ParameterField paramField2 = new ParameterField ();
			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			paramField2.ParameterFieldName = "Fecha";
			discreteVal2.Value = dtFechaCheque;
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);

			string sSQL = string.Format("Select Reporte From CuentaCorriente Where idCuenta = {0}", idCuenta);
			string sReporte = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
			
			Reporte miRep = new Reporte(string.Format("{0}.rpt", sReporte), strFiltro);
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.MdiParent = MdiParent;
			miRep.Show();
		}

		private int Egreso(Infragistics.Win.UltraWinGrid.CellEventArgs e, string sBeneficiario, DateTime Fecha, string sNumeroCheque, string sNumEgreso)
		{			
			int idAsiento = 0;
			int idPago = 0;

			if ((int)e.Cell.Row.Cells["idTipo"].Value == 4) /* COMPRA */
			{
				/* CREAR EL PAGO DE LA AUTORIZACION SELECCIONADA */

				string sSQL = string.Format("Exec InsertaPagosAutorizados {0}, '{1}', '{2}', '{3}', '{4}'",  (int)e.Cell.Row.Cells["idAutorizacionPago"].Value, sBeneficiario, Fecha.ToString("yyyyMMdd"), sNumeroCheque, sNumEgreso);
				idPago = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
	
				string stSelect = string.Format("Exec PagoAsientoE {0}, '{1}'", idPago, sNumEgreso);
				Funcion.EjecutaSQL(cdsSeteoF, stSelect, true);
          
				idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idAsiento, 0) From Pago Where idPago = {0}", idPago));	
																					
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Asiento Set Cheque = {0} Where idAsiento = {1}", sNumeroCheque, idAsiento));

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update CuentaCorriente Set NumCheque = {0} Where idCuenta = {1}", sNumeroCheque, (int)e.Cell.Row.Cells["idCuenta"].Value));


				return idAsiento;
			}
			else if ((int)e.Cell.Row.Cells["idTipo"].Value == 5) /* ABONO */
			{
				string sSQL = string.Format("Exec InsertaPagosAutorizados {0}, '{1}', '{2}', '{3}', '{4}'",  (int)e.Cell.Row.Cells["idAutorizacionPago"].Value, sBeneficiario, Fecha.ToString("yyyyMMdd"), sNumeroCheque, sNumEgreso);
				idPago = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaPagoPendiente {0}, 5", (int)e.Cell.Row.Cells["idAutorizacionPago"].Value));

				string stSelect = string.Format("Exec PagoAsiento {0}", idPago); //
				Funcion.EjecutaSQL(cdsSeteoF, stSelect, true);
          
				idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idAsiento, 0) From Pago Where idPago = {0}", idPago));					

				return idAsiento;
			}
			else /* PRESTAMO Y ANTICIPO EMPLEADOS */
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec PrestamoAsiento {0}", (int)e.Cell.Row.Cells["idOrigen"].Value));

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update RolPrestamo Set Estado = 1 Where idRolPrestamo = {0}", (int)e.Cell.Row.Cells["idOrigen"].Value));
					
				idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idAsiento, 0) From RolPrestamo Where idRolPrestamo = {0}", (int)e.Cell.Row.Cells["idOrigen"].Value));
				
				return idAsiento;
			}			
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "UnboundColumn1")
			{							
				int idAsiento = 0;

				if (e.Cell.Row.Cells["FormaPago"].Value.ToString() == "CHEQUES")
				{
					string sSQL = string.Format("Select dbo.NumeroCheque({0})", (int)e.Cell.Row.Cells["idCuenta"].Value);
					string NumCheque = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
					string sNombre = e.Cell.Row.Cells["Nombre"].Value.ToString();

					using (frmNumeracionCheque miCheque = new frmNumeracionCheque(NumCheque, sNombre, (int)e.Cell.Row.Cells["idCuenta"].Value, (int)e.Cell.Row.Cells["idFormaPago"].Value, (DateTime)this.dtFechaPago.Value))
					{
						if (DialogResult.OK == miCheque.ShowDialog())
						{								
							this.Cursor = Cursors.WaitCursor;

							idAsiento = Egreso(e, miCheque.txtNombre.Text.ToString(), (DateTime)miCheque.dtFecha.Value, miCheque.txtNumCheque.Text.ToString(), miCheque.txtNumEgreso.Text.ToString());
							
							ImpresionEgreso(idAsiento, (int)e.Cell.Row.Cells["idTipo"].Value);

							ImpresionCheque((decimal)e.Cell.Row.Cells["Autorizado"].Value, miCheque.txtNombre.Text.ToString(), (DateTime)miCheque.dtFecha.Value, idAsiento, (int)e.Cell.Row.Cells["idCuenta"].Value);

							SQL();

							this.Cursor = Cursors.Default;
						}
					}
				}
				else if (e.Cell.Row.Cells["FormaPago"].Value.ToString() == "TRANSFERENCIA")
				{
					string sNombre = e.Cell.Row.Cells["Nombre"].Value.ToString();

					using (frmNumeracionCheque miCheque = new frmNumeracionCheque("", sNombre, (int)e.Cell.Row.Cells["idCuenta"].Value, (int)e.Cell.Row.Cells["idFormaPago"].Value, (DateTime)this.dtFechaPago.Value)) /* CONFIRMAR PARAMETROS FALTA ENVIAR SI ES TRANSFERENCIA O NBTA DE CREDITO */
					{
						if (DialogResult.OK == miCheque.ShowDialog())
						{
							this.Cursor = Cursors.WaitCursor;

							idAsiento = Egreso(e, miCheque.txtNombre.Text.ToString(), (DateTime)miCheque.dtFecha.Value, miCheque.txtNumCheque.Text.ToString(), miCheque.txtNumEgreso.Text.ToString());
					
							//							if ((int)e.Cell.Row.Cells["idTipo"].Value == 4 || (int)e.Cell.Row.Cells["idTipo"].Value == 5) idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idAsiento, 0) From Pago Where idPago = {0}", (int)e.Cell.Row.Cells["idOrigen"].Value));
							//							else idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idAsiento, 0) From RolPrestamo Where idRolPrestamo = {0}", (int)e.Cell.Row.Cells["idOrigen"].Value));

							ImpresionEgreso(idAsiento, (int)e.Cell.Row.Cells["idTipo"].Value);

							SQL();

							this.Cursor = Cursors.Default;
						}
					}
				}
			}			
		}

		private void PagosPendientes_Load(object sender, System.EventArgs e)
		{
			Funcion.Resolucion(this, this.Height, this.Width);
			this.txtFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.txtFecha.Value = DateTime.Today;
		}

		private void PagosPendientes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void txtBusqueda_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtBusqueda_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnVer_Click(sender, e);
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idAutorizacionPago"].Value == System.DBNull.Value) return;

			if (e.Cell.Column.ToString() == "Sel")
			{
				this.ultraGrid1.UpdateData();
				Suma();
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{			
			SQL();
		}

		private void btnGenerarPago_Click(object sender, System.EventArgs e)
		{
			int idAsiento = 0;
			string sSQL = "";
			DateTime Fecha = DateTime.Today;
			string sNotas = "";
			string sDocumento = "";
			
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay pagos pendientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBusqueda.Focus();
				return;
			}

			#region Valida Seleccion
			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				if ((bool)dr.Cells["Sel"].Value) iCont++;

			if (iCont == 0)
			{
				MessageBox.Show("Seleccione los pagos a generar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			#endregion Valida Seleccion
			
			#region Valida Valor a Pagar
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value) 
				{
					decimal dValor = (decimal)dr.Cells["Autorizado"].Value;
					decimal dSaldo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec RetornaSaldoAutorizacionPago {0}", (int)dr.Cells["idAutorizacionPago"].Value));

					if (Math.Round(dValor, 2) > Math.Round(dSaldo, 2))
					{
						MessageBox.Show("El Valor a Pagar es Mayor al Saldo del Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}
			}
			#endregion Valida Valor a Pagar

			#region Valida Proveedor
			string sRUC = "";
			int idTipo = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					if (sRUC.Length == 0) sRUC = dr.Cells["RUC"].Value.ToString();
					else if (sRUC != dr.Cells["RUC"].Value.ToString() && (int)dr.Cells["idTipo"].Value != 9)
					{
						MessageBox.Show("No puede generar egreso para proveedores distintos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}
			}		
			#endregion Valida Proveedor

			#region Valida Forma de pago
			int idFormaPago = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					if (idFormaPago == 0) idFormaPago = (int)dr.Cells["idFormaPago"].Value;
					else if ((idFormaPago != (int)dr.Cells["idFormaPago"].Value) && (int)dr.Cells["idFormaPago"].Value != 4)
					{
						MessageBox.Show("No puede generar egreso con formas de pago distintas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}
			}		
			#endregion Valida Forma de pago

			#region Valida Cuenta
			int idCuenta = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					if (idCuenta == 0) idCuenta = (int)dr.Cells["idCuenta"].Value;
					else if (idCuenta != (int)dr.Cells["idCuenta"].Value)
					{
						MessageBox.Show("No puede generar egreso con cuentas distintas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}
			}		
			#endregion Valida Cuenta

			#region Tipo - Nombre
			string sNombre = "";

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					if (idTipo == 0) 
					{
						idTipo = (int)dr.Cells["idTipo"].Value;
						sNombre = dr.Cells["Nombre"].Value.ToString();
						sDocumento = dr.Cells["Numero"].Value.ToString();

						continue;
					}
				}
			}
			#endregion Tipo - Nombre

			#region Valida Fecha
			this.dtFechaPago.Value = System.DBNull.Value;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					if (this.dtFechaPago.Value == System.DBNull.Value) this.dtFechaPago.Value = (DateTime)dr.Cells["FechaSol"].Value;
					else if ((DateTime)this.dtFechaPago.Value > (DateTime)dr.Cells["FechaSol"].Value)
					{
						this.dtFechaPago.Value = (DateTime)dr.Cells["FechaSol"].Value;
					}
				}
			}		
			#endregion Valida Fecha
			
			#region Totaliza valor a Pagar
			decimal dSuma = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				if ((bool)dr.Cells["Sel"].Value)
					if ((int)dr.Cells["idFormapago"].Value != 4) 
						dSuma = dSuma + (decimal)dr.Cells["Autorizado"].Value;			
			#endregion Totaliza valor a Pagar
			
			#region Retorna Numero de Cheque
			sSQL = string.Format("Select dbo.NumeroCheque({0})", idCuenta);
			string NumCheque = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
			#endregion Retorna Numero de Cheque	

			#region Crea Egreso
			using (frmNumeracionCheque miCheque = new frmNumeracionCheque(NumCheque, sNombre, idCuenta, idFormaPago, (DateTime)this.dtFechaPago.Value))
			{
				if (DialogResult.OK == miCheque.ShowDialog())
				{
					this.Cursor = Cursors.WaitCursor;
					Fecha = (DateTime)miCheque.dtFecha.Value;

					#region Notas
					if (idTipo == 7) sNotas = "Caja Chica";
					if (idTipo == 8) sNotas = "Cruce Anticipos Clientes Banco Solidario";
					if (idTipo == 9) sNotas = "Finiquito Empleados";
					if (idTipo == 10) sNotas = "Anticipos Por Viaje";
					#endregion Notas

					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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

							#region Egreso
							DateTime dtFechaEgreso = (DateTime)miCheque.dtFecha.Value;
							DateTime dtFechaCobro = (DateTime)miCheque.dtFechaCheque.Value;					

							string sSQLEncEgreso = string.Format("Exec GeneraEncabezadoEgreso {0}, '{1}', '{2}', {3}, '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', {10}", 
								0, miCheque.txtNumEgreso.Text.ToString(), miCheque.txtNumCheque.Text.ToString(), 19, miCheque.txtNombre.Text.ToString(), sNotas, sDocumento, 
								dtFechaEgreso.ToString("yyyyMMdd"), dtFechaCobro.ToString("yyyyMMdd"), sRUC, idTipo);
							oCmdActualiza.CommandText = sSQLEncEgreso;									
							idAsiento = (int)oCmdActualiza.ExecuteScalar();
												
							#region Detalle Egreso
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
							{
								if ((bool) dr.Cells["Sel"].Value)
								{
									idTipo = Convert.ToInt32(dr.Cells["idTipo"].Value);

									string sSQLDetEgreso = string.Format("Exec GeneraDetalleEgreso {0}, {1}, {2}, {3}, {4}, '{5}', {6}, '{7}'", 
										idAsiento, (int)dr.Cells["idAutorizacionPago"].Value, (int)dr.Cells["idOrigen"].Value, (int)dr.Cells["idTipo"].Value, (int)dr.Cells["idProyecto"].Value, 
										dr.Cells["Ruc"].Value.ToString(), Convert.ToDecimal(dr.Cells["Autorizado"].Value), dr.Cells["Numero"].Value.ToString());
									oCmdActualiza.CommandText = sSQLDetEgreso;
									oCmdActualiza.ExecuteNonQuery();

									string sSQLAutorizacion = string.Format("Exec ActualizaSaldosCXC {0}, {1}, 1", (int)dr.Cells["idAutorizacionPago"].Value, idAsiento);
									oCmdActualiza.CommandText = sSQLAutorizacion;
									oCmdActualiza.ExecuteNonQuery();
								
									string sSQLEstados = string.Format("Exec PagosActualizaEstadoDocumentos {0}, {1}, {2}, {3}, '{4}', '{5}'",
										(int)dr.Cells["idAutorizacionPago"].Value, (int)dr.Cells["idOrigen"].Value, (int)dr.Cells["idTipo"].Value, idAsiento, 
										dr.Cells["Concepto"].Value.ToString(), miCheque.txtNumEgreso.Text.ToString());
									oCmdActualiza.CommandText = sSQLEstados;
									oCmdActualiza.ExecuteNonQuery();									
								}
							}							
							#endregion Detalle Egreso

							#region Bancos
							string sSQLBancosCaja = string.Format("Exec EgresoBancosCaja {0}, {1}, {2}, {3}", idAsiento, idCuenta, dSuma, idFormaPago);
							oCmdActualiza.CommandText = sSQLBancosCaja;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Bancos

							#region Actualiza numero de cheque en cuenta bancaria
							if (idFormaPago == 2)
							{
								oCmdActualiza.CommandText = string.Format("Update CuentaCorriente Set NumCheque = {0} Where idCuenta = {1}", miCheque.txtNumCheque.Text.ToString(), idCuenta);
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Actualiza numero de cheque en cuenta bancaria

							oTransaction.Commit();							
					
							oConexion.Close();

							ImpresionEgreso(idAsiento, idTipo);

							if (idFormaPago == 2) 
								ImpresionCheque(dSuma, miCheque.txtNombre.Text.ToString(), (DateTime)miCheque.dtFechaCheque.Value, idAsiento, idCuenta);
							#endregion Egreso
												
							SQL();

							this.Cursor = Cursors.Default;
						}
						catch (Exception ex)
						{
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);												

							this.Cursor = Cursors.Default;
						}
						finally
						{
							oConexion.Close();
						}
					}
				}
			}
			#endregion Crea Egreso
		}			
	}
}
