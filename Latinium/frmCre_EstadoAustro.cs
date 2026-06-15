using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_EstadoAustro.
	/// </summary>
	public class frmCre_EstadoAustro : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optSeleccionados;
		private System.Windows.Forms.Label lblSeleccion;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugrTemporal;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridOperaciones;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_EstadoAustro()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_EstadoAustro));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("REFERENCIA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IDENTIFICACION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FECHA_INICIO_CUOTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FECHA_VENCIMIENTO_CUOTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FECHA_PAGO_CUOTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VALOR_PAGADO_CAPITAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VALOR_PAGADO_INTERES");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VALOR_PAGADO_MORA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NUMERO_DE_CUOTA_PAGADA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CODIGO_DE_LA_TRANSACCION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NUMERO_TOTAL_DE_CUOTAS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VALOR_TOTAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VALOR_PAGADO_ADICIONAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TIPO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DIAS_VENCIMIENTO");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Identificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaInicioCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVencimientoCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPagoCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPagadoCapital");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPagadoInteres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPagadoMora");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroCuotaPagada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoTransaccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroTotalCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPagadoAdicional");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasVencimiento");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Identificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaInicioCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVencimientoCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPagoCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorPagadoCapital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorPagadoInteres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorPagadoMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroCuotaPagada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoTransaccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroTotalCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorPagadoAdicional");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasVencimiento");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("REFERENCIA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IDENTIFICACION");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FECHA_INICIO_CUOTA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FECHA_VENCIMIENTO_CUOTA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FECHA_PAGO_CUOTA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VALOR_PAGADO_CAPITAL");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VALOR_PAGADO_INTERES");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VALOR_PAGADO_MORA");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NUMERO_DE_CUOTA_PAGADA");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CODIGO_DE_LA_TRANSACCION");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NUMERO_TOTAL_DE_CUOTAS");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VALOR_TOTAL");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VALOR_PAGADO_ADICIONAL");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TIPO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DIAS_VENCIMIENTO");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "VALOR_PAGADO_CAPITAL", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "VALOR_PAGADO_CAPITAL", 7, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "VALOR_PAGADO_INTERES", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "VALOR_PAGADO_INTERES", 8, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "VALOR_PAGADO_MORA", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "VALOR_PAGADO_MORA", 9, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "VALOR_TOTAL", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "VALOR_TOTAL", 13, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "VALOR_PAGADO_ADICIONAL", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "VALOR_PAGADO_ADICIONAL", 14, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.optSeleccionados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblSeleccion = new System.Windows.Forms.Label();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.ugrTemporal = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridOperaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optSeleccionados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugrTemporal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOperaciones)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 755;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 6);
			this.dtFecha.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 756;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// optEstado
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance2;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.CheckedIndex = 1;
			this.optEstado.ItemAppearance = appearance3;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendiente";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Enviado";
			this.optEstado.Items.Add(valueListItem1);
			this.optEstado.Items.Add(valueListItem2);
			this.optEstado.Items.Add(valueListItem3);
			this.optEstado.ItemSpacingVertical = 10;
			this.optEstado.Location = new System.Drawing.Point(184, 3);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(224, 26);
			this.optEstado.TabIndex = 761;
			this.optEstado.Text = "Pendiente";
			this.optEstado.ValueChanged += new System.EventHandler(this.optEstado_ValueChanged);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(376, 40);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 770;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 41);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(39, 16);
			this.label7.TabIndex = 769;
			this.label7.Text = "Buscar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(56, 40);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(216, 19);
			this.txtBuscar.TabIndex = 768;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(280, 37);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 767;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1056, 5);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 771;
			this.btnExportar.Text = "Generar XLS";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(int);
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
																																 ultraDataColumn17});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 64);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1304, 8);
			this.groupBox1.TabIndex = 773;
			this.groupBox1.TabStop = false;
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
			// optSeleccionados
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.optSeleccionados.Appearance = appearance4;
			this.optSeleccionados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optSeleccionados.CheckedIndex = 0;
			this.optSeleccionados.ItemAppearance = appearance5;
			this.optSeleccionados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = false;
			valueListItem4.DisplayText = "No Seleccionados";
			valueListItem5.DataValue = true;
			valueListItem5.DisplayText = "Seleccionados";
			this.optSeleccionados.Items.Add(valueListItem4);
			this.optSeleccionados.Items.Add(valueListItem5);
			this.optSeleccionados.ItemSpacingVertical = 10;
			this.optSeleccionados.Location = new System.Drawing.Point(424, 3);
			this.optSeleccionados.Name = "optSeleccionados";
			this.optSeleccionados.Size = new System.Drawing.Size(232, 26);
			this.optSeleccionados.TabIndex = 774;
			this.optSeleccionados.Text = "No Seleccionados";
			this.optSeleccionados.ValueChanged += new System.EventHandler(this.optSeleccionados_ValueChanged);
			// 
			// lblSeleccion
			// 
			this.lblSeleccion.AutoSize = true;
			this.lblSeleccion.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSeleccion.ForeColor = System.Drawing.Color.Firebrick;
			this.lblSeleccion.Location = new System.Drawing.Point(672, 8);
			this.lblSeleccion.Name = "lblSeleccion";
			this.lblSeleccion.Size = new System.Drawing.Size(0, 17);
			this.lblSeleccion.TabIndex = 775;
			this.lblSeleccion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtValor
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance6;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(1096, 38);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(48, 21);
			this.txtValor.TabIndex = 776;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(1056, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(29, 16);
			this.label1.TabIndex = 777;
			this.label1.Text = "Total";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ugrTemporal
			// 
			this.ugrTemporal.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugrTemporal.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugrTemporal.DataSource = this.ultraDataSource2;
			appearance7.BackColor = System.Drawing.Color.White;
			this.ugrTemporal.DisplayLayout.Appearance = appearance7;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 12;
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
																										 ultraGridColumn16});
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.ugrTemporal.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ugrTemporal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugrTemporal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugrTemporal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.ugrTemporal.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugrTemporal.DisplayLayout.Override.CellAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8.5F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugrTemporal.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.ugrTemporal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugrTemporal.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugrTemporal.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ugrTemporal.Enabled = false;
			this.ugrTemporal.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugrTemporal.Location = new System.Drawing.Point(8, 80);
			this.ugrTemporal.Name = "ugrTemporal";
			this.ugrTemporal.Size = new System.Drawing.Size(1136, 320);
			this.ugrTemporal.TabIndex = 778;
			this.ugrTemporal.Visible = false;
			this.ugrTemporal.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugrTemporal_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(System.DateTime);
			ultraDataColumn22.DataType = typeof(System.DateTime);
			ultraDataColumn23.DataType = typeof(System.DateTime);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn33});
			// 
			// uGridOperaciones
			// 
			this.uGridOperaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridOperaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridOperaciones.DataSource = this.ultraDataSource1;
			appearance13.BackColor = System.Drawing.Color.White;
			this.uGridOperaciones.DisplayLayout.Appearance = appearance13;
			ultraGridColumn17.Header.Caption = "...";
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Width = 28;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Width = 113;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Width = 112;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.Caption = "FECHA INICIO CUOTA";
			ultraGridColumn21.Header.VisiblePosition = 4;
			ultraGridColumn21.Width = 135;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.Caption = "FECHA VENCIMIENTO CUOTA";
			ultraGridColumn22.Header.VisiblePosition = 5;
			ultraGridColumn22.Width = 177;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Format = "";
			ultraGridColumn23.Header.Caption = "FECHA PAGO CUOTA";
			ultraGridColumn23.Header.VisiblePosition = 6;
			ultraGridColumn23.Width = 131;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance14;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "VALOR PAGADO CAPITAL";
			ultraGridColumn24.Header.VisiblePosition = 7;
			ultraGridColumn24.Width = 155;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance15;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.Caption = "VALOR PAGADO INTERES";
			ultraGridColumn25.Header.VisiblePosition = 8;
			ultraGridColumn25.Width = 159;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance16;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.Caption = "VALOR PAGADO MORA";
			ultraGridColumn26.Header.VisiblePosition = 9;
			ultraGridColumn26.Width = 147;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance17;
			ultraGridColumn27.Header.Caption = "NUMERO DE CUOTA PAGADA";
			ultraGridColumn27.Header.VisiblePosition = 10;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.Caption = "CODIGO DE LA TRANSACCION";
			ultraGridColumn28.Header.VisiblePosition = 12;
			ultraGridColumn28.Width = 186;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance18;
			ultraGridColumn29.Header.Caption = "NUMERO TOTAL DE CUOTAS";
			ultraGridColumn29.Header.VisiblePosition = 13;
			ultraGridColumn29.Width = 172;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance19;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.Caption = "VALOR TOTAL";
			ultraGridColumn30.Header.VisiblePosition = 14;
			ultraGridColumn30.Width = 102;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance20;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.Caption = "VALOR PAGADO ADICIONAL";
			ultraGridColumn31.Header.VisiblePosition = 15;
			ultraGridColumn31.Width = 171;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 16;
			ultraGridColumn32.Width = 122;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance21;
			ultraGridColumn33.Header.Caption = "DIAS VENCIMIENTO";
			ultraGridColumn33.Header.VisiblePosition = 11;
			ultraGridColumn33.Width = 125;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn33});
			ultraGridBand2.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance22;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance23;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance24;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance25;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance26;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5});
			ultraGridBand2.SummaryFooterCaption = "Totales";
			this.uGridOperaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridOperaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridOperaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridOperaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance27.BackColor = System.Drawing.Color.Transparent;
			this.uGridOperaciones.DisplayLayout.Override.CardAreaAppearance = appearance27;
			appearance28.ForeColor = System.Drawing.Color.Black;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOperaciones.DisplayLayout.Override.CellAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance29.FontData.BoldAsString = "True";
			appearance29.FontData.Name = "Arial";
			appearance29.FontData.SizeInPoints = 8.5F;
			appearance29.ForeColor = System.Drawing.Color.White;
			appearance29.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridOperaciones.DisplayLayout.Override.HeaderAppearance = appearance29;
			this.uGridOperaciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOperaciones.DisplayLayout.Override.RowSelectorAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOperaciones.DisplayLayout.Override.SelectedRowAppearance = appearance31;
			this.uGridOperaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridOperaciones.Location = new System.Drawing.Point(8, 80);
			this.uGridOperaciones.Name = "uGridOperaciones";
			this.uGridOperaciones.Size = new System.Drawing.Size(1136, 320);
			this.uGridOperaciones.TabIndex = 779;
			this.uGridOperaciones.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridOperaciones_CellChange_1);
			// 
			// frmCre_EstadoAustro
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1152, 406);
			this.Controls.Add(this.uGridOperaciones);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.lblSeleccion);
			this.Controls.Add(this.optSeleccionados);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.optEstado);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.ugrTemporal);
			this.Name = "frmCre_EstadoAustro";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Estado Austro";
			this.Load += new System.EventHandler(this.frmCre_EstadoAustro_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optSeleccionados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugrTemporal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOperaciones)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variable
		decimal dTotalSeleccion = 0.00m;
		#endregion Variable

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

		private void frmCre_EstadoAustro_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "08320701");

			if (!Funcion.Permiso("1803", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar Pagos Fideicomiso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;		
			}

			this.dtFecha.CalendarInfo.MinDate = DateTime.Parse("02/08/2022");
			this.dtFecha.Value = DateTime.Today;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today.AddDays(30);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}

			string sSQL = string.Format("Exec Cre_ListaPagosFideicomisoAustro '{0}', {1}, '{2}', {3}",
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.optEstado.Value, this.txtBuscar.Text.ToString(), (bool)this.optSeleccionados.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridOperaciones);

			this.lblContador.Text = this.uGridOperaciones.Rows.Count + " Registros Encontrados";
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			bool bSeleccionados = (bool)this.optSeleccionados.Value;
			int iCont = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridOperaciones.Rows.All)
				if ((bool) dr.Cells["Sel"].Value) iCont ++;
			
			if (iCont == 0)
			{
				MessageBox.Show("Seleccione los documentos para generar el pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridOperaciones.Focus();				
				return;
			}

			if (bSeleccionados)
			{
				MessageBox.Show("No puede Exportar filas si no esta Seleccionadas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.optSeleccionados.Focus();
				return;
			}
			else
			{
				if (DialogResult.Yes == MessageBox.Show("Esta Seguro de Generar el proceso de Seleccion \n\nEste proceso no puede ser revertido", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridOperaciones.Rows.All)
					{
						bool bCheck = false;

						bCheck = (bool)dr.Cells["Sel"].Value;

						if (bCheck)
						{
							#region Variables
							int iTablaAmortizacion = 0;
							int iNumeroCuotaPagada = 0;
							int iNumeroTotalCuotas = 0;
							int iDiasVencimiento = 0;
							string sReferencia = "";
							string sIdentificacion = "";
							string sTipo = "";
							string sCodigoTransaccion = "";
							DateTime dtFechaInicioCuota = DateTime.Now;
							DateTime dtFechaVencimientoCuota = DateTime.Now;
							DateTime dtFechaPagoCuota = DateTime.Now;
							decimal dValorPagadoCapital = 0.00m;
							decimal dValorPagadoInteres = 0.00m;
							decimal dValorPagadoMora = 0.00m;
							decimal dValorTotal = 0.00m;
							decimal dValorPagadoAdicional = 0.00m;
							#endregion Variables

							#region Datos
							iTablaAmortizacion = (int)dr.Cells["idCre_TablaDeAmortizacion"].Value;
							sReferencia = (string)dr.Cells["REFERENCIA"].Value;
							sIdentificacion = (string)dr.Cells["IDENTIFICACION"].Value;
							dtFechaInicioCuota = Convert.ToDateTime(dr.Cells["FECHA_INICIO_CUOTA"].Value);
							dtFechaVencimientoCuota = Convert.ToDateTime(dr.Cells["FECHA_VENCIMIENTO_CUOTA"].Value);
							dtFechaPagoCuota = Convert.ToDateTime(dr.Cells["FECHA_PAGO_CUOTA"].Value);
							dValorPagadoCapital = (decimal)dr.Cells["VALOR_PAGADO_CAPITAL"].Value;
							dValorPagadoInteres = (decimal)dr.Cells["VALOR_PAGADO_INTERES"].Value;
							dValorPagadoMora = (decimal)dr.Cells["VALOR_PAGADO_MORA"].Value;
							iNumeroCuotaPagada = (int)dr.Cells["NUMERO_DE_CUOTA_PAGADA"].Value;
							iDiasVencimiento = (int)dr.Cells["DIAS_VENCIMIENTO"].Value;
							sCodigoTransaccion = (string)dr.Cells["CODIGO_DE_LA_TRANSACCION"].Value;
							iNumeroTotalCuotas = (int)dr.Cells["NUMERO_TOTAL_DE_CUOTAS"].Value;
							dValorTotal = (decimal)dr.Cells["VALOR_TOTAL"].Value;
							dValorPagadoAdicional = (decimal)dr.Cells["VALOR_PAGADO_ADICIONAL"].Value;
							sTipo = (string)dr.Cells["TIPO"].Value;

							#endregion Datos

							#region Inserta Datos temporales
							string sSQL = string.Format("Exec Cre_InsertaFideicomisoTemporal {0}, '{1}', '{2}', '{3}', '{4}', '{5}', {6}, {7}, {8}, {9}, '{10}', {11}, {12}, {13}, {14}, '{15}'",
								iTablaAmortizacion, sReferencia, sIdentificacion, dtFechaInicioCuota.ToString("yyyyMMdd"), dtFechaVencimientoCuota.ToString("yyyyMMdd"), dtFechaPagoCuota.ToString("yyyyMMdd"), 
								dValorPagadoCapital, dValorPagadoInteres, dValorPagadoMora, iNumeroCuotaPagada, sCodigoTransaccion, 
								iDiasVencimiento, iNumeroTotalCuotas, dValorTotal, dValorPagadoAdicional, sTipo);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
							#endregion Inserta Datos temporales

							#region Modifica Estado
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaEstadoFideicomiso {0}", iTablaAmortizacion));
							#endregion Modifica Estado
						}

						#region Muestra Datos Temporales
						ConsultaTemporal();
						#endregion Muestra Datos Temporales
					}

					#region Asiento
					string sSQLAS = string.Format("Exec Cre_AsientoCuotasFideicomisoAustro {0}", dTotalSeleccion);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAS, true);
					#endregion Asiento

					#region Exportar
					FuncionesProcedimientos.ExportaExcel(this.ugrTemporal);
					#endregion Exportar

					#region Elimina Datos Temporales
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete from FideicomisoAustroTemporal"));
					ConsultaTemporal();

					this.txtValor.Value = 0.00m;
					dTotalSeleccion = (decimal)this.txtValor.Value;
					#endregion Elimina Datos Temporales

					#region Consulta
					this.btnVer_Click(sender, e);
					#endregion Consulta

					MessageBox.Show("Proceso Terminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
		}

		private void ConsultaTemporal()
		{
			string sSQLCO = string.Format("Exec Cre_ConsultaFideicomisoTemporal");
			FuncionesProcedimientos.dsGeneral(sSQLCO, this.ugrTemporal);
		}

		private void suma()
		{			
			this.txtValor.Value = 0.00m;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridOperaciones.Rows.All)
				if ((bool)dr.Cells["Sel"].Value)
					this.txtValor.Value = Convert.ToDecimal(this.txtValor.Value) + Convert.ToDecimal(dr.Cells["VALOR_TOTAL"].Value);
					dTotalSeleccion = (decimal)this.txtValor.Value;

		}

		private void CampSeleccionados()
		{
			this.lblSeleccion.Text = "";
			int iCont = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridOperaciones.Rows.All)
				if ((bool)dr.Cells["Sel"].Value) iCont ++;

			this.lblSeleccion.Text = iCont.ToString() + " DOCUMENTOS SELECCIONADOS   ";
		}

		private void uGridOperaciones_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridOperaciones.UpdateData();

				CampSeleccionados();

				suma();
			}
		}

		private void uGridOperaciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void optEstado_ValueChanged(object sender, System.EventArgs e)
		{
			int iEstados = 0;
			iEstados = (int)this.optEstado.Value;

			if (iEstados == 2 || iEstados == 0)
			{
				this.btnExportar.Enabled = false;
				this.uGridOperaciones.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			}
			else
			{
				this.btnExportar.Enabled = true;
				this.uGridOperaciones.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}

			this.btnVer_Click(sender, e);
		}

		private void optSeleccionados_ValueChanged(object sender, System.EventArgs e)
		{
			bool bSeleccion = false;
			bSeleccion = (bool)this.optSeleccionados.Value;

			if (bSeleccion)
			{
				this.uGridOperaciones.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			}
			else
			{
				this.uGridOperaciones.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}
			
			this.btnVer_Click(sender, e);
		}

		private void ugrTemporal_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridOperaciones_CellChange_1(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridOperaciones.UpdateData();

				CampSeleccionados();

				suma();
			}
		}
	}
}
