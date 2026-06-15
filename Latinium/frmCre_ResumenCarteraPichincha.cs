using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_ResumenCarteraPichincha.
	/// </summary>
	public class frmCre_ResumenCarteraPichincha : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnSubirArchivos;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDatos;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFecha;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_ResumenCarteraPichincha()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idResumenCarteraVencidaPichincha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSubida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaProceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVenceCuota");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto_Financiado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapitalVencido");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mora");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorAPagar");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroCuota");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasMora");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasVencidas");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SCapital");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SInteres");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SMora");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SValorAPagar");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SNumeroCuota");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SDiasMora");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDoc");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Documento");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoPago");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoPagoM");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPago");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldosPorCruzar");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorAPagar", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorAPagar", 13, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SValorAPagar", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SValorAPagar", 20, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Diferencia", 32, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Diferencia", 32, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idResumenCarteraVencidaPichincha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSubida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaProceso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVenceCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto_Financiado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapitalVencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorAPagar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasVencidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SCapital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SInteres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SValorAPagar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SNumeroCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SDiasMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDoc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Documento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoPagoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldosPorCruzar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencia");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_ResumenCarteraPichincha));
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.uGridDatos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnSubirArchivos = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.dtpFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.uGridDatos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridDatos
			// 
			this.uGridDatos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDatos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDatos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDatos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.Caption = "Subida";
			ultraGridColumn2.Header.VisiblePosition = 5;
			ultraGridColumn2.Width = 75;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.Caption = "Proceso";
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 75;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 100;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Width = 250;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.BackColor = System.Drawing.Color.CornflowerBlue;
			appearance2.BackColor2 = System.Drawing.Color.CornflowerBlue;
			ultraGridColumn9.CellAppearance = appearance2;
			ultraGridColumn9.Header.Caption = "Vence";
			ultraGridColumn9.Header.VisiblePosition = 12;
			ultraGridColumn9.Width = 75;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance3;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Financiado";
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 100;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.BackColor = System.Drawing.Color.CornflowerBlue;
			appearance4.BackColor2 = System.Drawing.Color.CornflowerBlue;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance4;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Capital";
			ultraGridColumn11.Header.VisiblePosition = 13;
			ultraGridColumn11.Width = 100;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.BackColor = System.Drawing.Color.CornflowerBlue;
			appearance5.BackColor2 = System.Drawing.Color.CornflowerBlue;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance5;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 14;
			ultraGridColumn12.Width = 100;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.BackColor = System.Drawing.Color.CornflowerBlue;
			appearance6.BackColor2 = System.Drawing.Color.CornflowerBlue;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance6;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 15;
			ultraGridColumn13.Width = 100;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.BackColor = System.Drawing.Color.CornflowerBlue;
			appearance7.BackColor2 = System.Drawing.Color.CornflowerBlue;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance7;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "A Pagar";
			ultraGridColumn14.Header.VisiblePosition = 16;
			ultraGridColumn14.Width = 100;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.BackColor = System.Drawing.Color.CornflowerBlue;
			appearance8.BackColor2 = System.Drawing.Color.CornflowerBlue;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance8;
			ultraGridColumn15.Format = "#,##0";
			ultraGridColumn15.Header.Caption = "N. Cuota";
			ultraGridColumn15.Header.VisiblePosition = 9;
			ultraGridColumn15.Width = 90;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.BackColor = System.Drawing.Color.CornflowerBlue;
			appearance9.BackColor2 = System.Drawing.Color.CornflowerBlue;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance9;
			ultraGridColumn16.Format = "#,##0";
			ultraGridColumn16.Header.Caption = "Dias Mora";
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Width = 90;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.BackColor = System.Drawing.Color.CornflowerBlue;
			appearance10.BackColor2 = System.Drawing.Color.CornflowerBlue;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance10;
			ultraGridColumn17.Format = "#,##0";
			ultraGridColumn17.Header.Caption = "C Vencidas";
			ultraGridColumn17.Header.VisiblePosition = 11;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 90;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.BackColor = System.Drawing.Color.LightSlateGray;
			appearance11.BackColor2 = System.Drawing.Color.LightSlateGray;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance11;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Capital";
			ultraGridColumn18.Header.VisiblePosition = 19;
			ultraGridColumn18.Width = 100;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.BackColor = System.Drawing.Color.LightSlateGray;
			appearance12.BackColor2 = System.Drawing.Color.LightSlateGray;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance12;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Interes";
			ultraGridColumn19.Header.VisiblePosition = 20;
			ultraGridColumn19.Width = 100;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.BackColor = System.Drawing.Color.LightSlateGray;
			appearance13.BackColor2 = System.Drawing.Color.LightSlateGray;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance13;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "Mora";
			ultraGridColumn20.Header.VisiblePosition = 21;
			ultraGridColumn20.Width = 100;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.BackColor = System.Drawing.Color.LightSlateGray;
			appearance14.BackColor2 = System.Drawing.Color.LightSlateGray;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance14;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "A Pagar";
			ultraGridColumn21.Header.VisiblePosition = 22;
			ultraGridColumn21.Width = 100;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.BackColor = System.Drawing.Color.LightSlateGray;
			appearance15.BackColor2 = System.Drawing.Color.LightSlateGray;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance15;
			ultraGridColumn22.Format = "#,##0";
			ultraGridColumn22.Header.Caption = "N Cuota";
			ultraGridColumn22.Header.VisiblePosition = 17;
			ultraGridColumn22.Width = 90;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.BackColor = System.Drawing.Color.LightSlateGray;
			appearance16.BackColor2 = System.Drawing.Color.LightSlateGray;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance16;
			ultraGridColumn23.Format = "#,##0";
			ultraGridColumn23.Header.Caption = "Dias Mora";
			ultraGridColumn23.Header.VisiblePosition = 18;
			ultraGridColumn23.Width = 90;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.BackColor = System.Drawing.Color.RoyalBlue;
			appearance17.BackColor2 = System.Drawing.Color.RoyalBlue;
			ultraGridColumn25.CellAppearance = appearance17;
			ultraGridColumn25.Header.Caption = "Doc";
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Width = 50;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.BackColor = System.Drawing.Color.RoyalBlue;
			appearance18.BackColor2 = System.Drawing.Color.RoyalBlue;
			ultraGridColumn26.CellAppearance = appearance18;
			ultraGridColumn26.Header.Caption = "N. Documento";
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 100;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.BackColor = System.Drawing.Color.RoyalBlue;
			appearance19.BackColor2 = System.Drawing.Color.RoyalBlue;
			ultraGridColumn27.CellAppearance = appearance19;
			ultraGridColumn27.Header.Caption = "Estado Cuota";
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 100;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.BackColor = System.Drawing.Color.RoyalBlue;
			appearance20.BackColor2 = System.Drawing.Color.RoyalBlue;
			ultraGridColumn28.CellAppearance = appearance20;
			ultraGridColumn28.Header.Caption = "Estado Cuota";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Width = 100;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.BackColor = System.Drawing.Color.RoyalBlue;
			appearance21.BackColor2 = System.Drawing.Color.RoyalBlue;
			ultraGridColumn29.CellAppearance = appearance21;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.BackColor = System.Drawing.Color.RoyalBlue;
			appearance22.BackColor2 = System.Drawing.Color.RoyalBlue;
			ultraGridColumn30.CellAppearance = appearance22;
			ultraGridColumn30.Header.Caption = "Estado Pago";
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Width = 100;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.BackColor = System.Drawing.Color.RoyalBlue;
			appearance23.BackColor2 = System.Drawing.Color.RoyalBlue;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance23;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.Caption = "Valor Cobro";
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Width = 100;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance24;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.Caption = "Por Cruzar";
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Width = 100;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance25;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Width = 100;
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
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33});
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance26;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance27;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance28;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridDatos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDatos.DisplayLayout.Override.ActiveRowAppearance = appearance29;
			this.uGridDatos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDatos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDatos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance30.BackColor = System.Drawing.Color.Transparent;
			this.uGridDatos.DisplayLayout.Override.CardAreaAppearance = appearance30;
			appearance31.ForeColor = System.Drawing.Color.Black;
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDatos.DisplayLayout.Override.CellAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance32.FontData.BoldAsString = "True";
			appearance32.FontData.Name = "Arial";
			appearance32.FontData.SizeInPoints = 10F;
			appearance32.ForeColor = System.Drawing.Color.White;
			appearance32.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDatos.DisplayLayout.Override.HeaderAppearance = appearance32;
			this.uGridDatos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance33.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance33.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDatos.DisplayLayout.Override.RowAlternateAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDatos.DisplayLayout.Override.RowSelectorAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDatos.DisplayLayout.Override.SelectedRowAppearance = appearance35;
			this.uGridDatos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDatos.Location = new System.Drawing.Point(8, 88);
			this.uGridDatos.Name = "uGridDatos";
			this.uGridDatos.Size = new System.Drawing.Size(1314, 312);
			this.uGridDatos.TabIndex = 243;
			this.uGridDatos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDatos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn2.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn9.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn12.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn13.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn14.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn18.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn19.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn31.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn32.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn33.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			// btnSubirArchivos
			// 
			this.btnSubirArchivos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSubirArchivos.CausesValidation = false;
			this.btnSubirArchivos.Enabled = false;
			this.btnSubirArchivos.Image = ((System.Drawing.Image)(resources.GetObject("btnSubirArchivos.Image")));
			this.btnSubirArchivos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSubirArchivos.Location = new System.Drawing.Point(1234, 8);
			this.btnSubirArchivos.Name = "btnSubirArchivos";
			this.btnSubirArchivos.Size = new System.Drawing.Size(80, 23);
			this.btnSubirArchivos.TabIndex = 731;
			this.btnSubirArchivos.Text = "Importar";
			this.btnSubirArchivos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSubirArchivos.Click += new System.EventHandler(this.btnSubirArchivos_Click);
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
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(368, 37);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 736;
			this.btnVer.Text = "&Ver";
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 735;
			this.label1.Text = "Buscar";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance36;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(64, 38);
			this.txtBuscar.MaxLength = 15;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(288, 21);
			this.txtBuscar.TabIndex = 734;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 10);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(36, 16);
			this.label5.TabIndex = 733;
			this.label5.Text = "Fecha";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpFecha
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpFecha.Appearance = appearance37;
			this.dtpFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtpFecha.DateButtons.Add(dateButton1);
			this.dtpFecha.Format = "dd/MM/yyyy";
			this.dtpFecha.Location = new System.Drawing.Point(64, 8);
			this.dtpFecha.Name = "dtpFecha";
			this.dtpFecha.NonAutoSizeHeight = 23;
			this.dtpFecha.Size = new System.Drawing.Size(104, 21);
			this.dtpFecha.SpinButtonsVisible = true;
			this.dtpFecha.TabIndex = 732;
			this.dtpFecha.Value = ((object)(resources.GetObject("dtpFecha.Value")));
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1330, 8);
			this.groupBox1.TabIndex = 737;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(418, 40);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 738;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.CornflowerBlue;
			this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label2.Location = new System.Drawing.Point(808, 9);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(101, 19);
			this.label2.TabIndex = 739;
			this.label2.Text = "Información Banco";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.LightSlateGray;
			this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label3.Location = new System.Drawing.Point(912, 9);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(109, 19);
			this.label3.TabIndex = 740;
			this.label3.Text = "Información Sistema";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.RoyalBlue;
			this.label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label4.Location = new System.Drawing.Point(1024, 9);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(99, 19);
			this.label4.TabIndex = 741;
			this.label4.Text = "Pagos de la Cuota";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmCre_ResumenCarteraPichincha
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1330, 406);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.dtpFecha);
			this.Controls.Add(this.btnSubirArchivos);
			this.Controls.Add(this.uGridDatos);
			this.KeyPreview = true;
			this.Name = "frmCre_ResumenCarteraPichincha";
			this.Text = "Resumen Cartera Pichincha";
			this.Load += new System.EventHandler(this.frmCre_ResumenCarteraPichincha_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDatos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).EndInit();
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

		private void frmCre_ResumenCarteraPichincha_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "082908");

			if (!Funcion.Permiso("949", cdsSeteoF))
			{		
				MessageBox.Show("No tiene Acceso RESUMEN CARTERA PICHINCHA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if(miAcceso.BImportar)this.btnSubirArchivos.Enabled = true;

		}

		private void btnSubirArchivos_Click(object sender, System.EventArgs e)
		{
			try
			{
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;                                  

					int iCont = 0;

					#region Sube archivo
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [Cartera_Vencida$]", "Cartera_Vencida").Tables[0].Rows)
					{
						#region Variables Archivo
						string sCedula = "";
						if (row["CEDULA"] != System.DBNull.Value) sCedula = row["CEDULA"].ToString();										
						if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
						int iFila = int.Parse(row["NROFILA"].ToString());
						string sOperacion =	row["OPERACION"].ToString();							
						string sCliente =	row["CLIENTE"].ToString();
						string dtFechaInicioCont = row["FECHAINICIOCONT"].ToString();
						string dtFechaVenceCont = row["FECHAVCTOCONT"].ToString();
						string dtFechaVenceCuota = row["FECHAVCTOCUOTA"].ToString();
						decimal dMonto = decimal.Parse(row["MONTOFINANCIADO"].ToString());
						decimal dCapital = decimal.Parse(row["CAPITALAMORTIZACION"].ToString());
						decimal dInteres = decimal.Parse(row["INTERES"].ToString());
						decimal dMora = decimal.Parse(row["MORA"].ToString());
						decimal dValorAPagar = decimal.Parse(row["VALORAPAGAR"].ToString());
						int iNroCuota = int.Parse(row["NROCUOTA"].ToString());
						int iDiasMora = int.Parse(row["DIASMORA"].ToString());
						int iNCV = int.Parse(row["NROCUOTASVENCIDAS"].ToString());
						#endregion Variables Archivo	

						#region Variables Consulta
						int idCre_TablaDeAmortizacion = 0;
						int isNCuota = 0;
						decimal dSCapital = 0.00m;
						decimal dSInteres = 0.00m;
						decimal dSMora = 0.00m;
						int isDiasMora = 0;
						int iEstado = 0;
						int idCompra = 0;
						int iSEstadoPago = 0;
						decimal dSValorDoc = 0.00m;
						string sTipoDoc = "";
						string sDocumento = "";
						decimal dSValorAPagar = 0.00m;
						decimal dDiferencias = 0.00m;
						int idPago = 0;
						int idCliente = 0;
						decimal dSaldosPorCruzar = 0.00m;
						#endregion Variables Consulta

						#region Datos Cuotas
						string sSQL = string.Format("Exec Cre_CuotasResumenCartera '{0}', {1}, '{2}'", sOperacion, iNroCuota, DateTime.Today.ToString("yyyyMMdd"));
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);						
						dr.Read();
						if (dr.HasRows)
						{
							idCre_TablaDeAmortizacion = dr.GetInt32(0);
							isNCuota = dr.GetInt32(1);
							dSCapital = dr.GetDecimal(2);
							dSInteres = dr.GetDecimal(3);
							dSMora = dr.GetDecimal(4);
							isDiasMora = dr.GetInt32(5);
							iEstado = dr.GetInt32(6);
							idCompra = dr.GetInt32(7);
							iSEstadoPago = dr.GetInt32(8);
							dSValorDoc = dr.GetDecimal(9);
							sTipoDoc = dr.GetString(10);
							sDocumento = dr.GetString(11);
							idPago = dr.GetInt32(12);
							idCliente = dr.GetInt32(13);
							dSaldosPorCruzar = dr.GetDecimal(14);
						}
						dr.Close();
						#endregion Datos Cuotas

						dSValorAPagar = dSCapital + dSInteres + dSMora;
						dDiferencias = dValorAPagar - dSValorAPagar;

						#region Guarda Informacion
						sSQL = string.Format("Exec Cre_GuardaResumenCarteraPichincha {0}, '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, '{23}', {24}, '{25}', '{26}', {27}, {28}, {29}, {30}, {31}, {32}",
							iFila, DateTime.Today.ToString("yyyyMMdd"), sOperacion, sCedula, sCliente.Trim(), //4
							dtFechaInicioCont, dtFechaVenceCont, dtFechaVenceCuota, //7
							dMonto, dCapital, dInteres, dMora, dValorAPagar, iNroCuota, iDiasMora, iNCV, //15
							dSCapital, dSInteres, dSMora, dSValorAPagar, isNCuota, isDiasMora, iEstado, //22
							iSEstadoPago, dSValorDoc, sTipoDoc, sDocumento, idCompra, idCre_TablaDeAmortizacion, dDiferencias, idPago, idCliente, dSaldosPorCruzar);//32
						Funcion.EjecutaSQL(cdsSeteoF, sSQL);
						#endregion Guarda Informacion

						iCont++;
					}				
					#endregion Sube archivo

					MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		
					//this.btnVer_Click(sender, e);
					this.Cursor = Cursors.Default;	
				}
			}
			catch (Exception ex)
			{
				try
				{
					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
					this.Cursor = Cursors.Default;	
					return;
				}
				catch (Exception ex2)
				{
					MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					this.Cursor = Cursors.Default;	
					return;
				}						
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec Cre_ResumenCarteraVencidaPichincha '{0}'", Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"));
			this.uGridDatos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.lblContador.Text = this.uGridDatos.Rows.Count + " Registros encontrados"; 
		}

		private void uGridDatos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}