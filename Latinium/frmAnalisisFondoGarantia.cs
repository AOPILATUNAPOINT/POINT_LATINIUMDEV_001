using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmAnalisisFondoGarantia.
	/// </summary>
	public class frmAnalisisFondoGarantia : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnConsultar;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGridPrincipal;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridPendientes;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		//private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridClientesSube;
		private System.Windows.Forms.Button btnGenerar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuota;
		private System.Windows.Forms.GroupBox groupBox1;
		
		/// <summary>
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAnalisisFondoGarantia()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OPERACION");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CEDULA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CLIENTE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MONTO_ORIGINAL");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SALDOFDG");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CUOTAS_RESC", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CUOTAS_RESC_COB");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CUOTAS_PENDIENTES");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VALOR_RESCATE");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SALDOFDG", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SALDOFDG", 4, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "VALOR_RESCATE", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "VALOR_RESCATE", 8, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OPERACION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CEDULA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CLIENTE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MONTO_ORIGINAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SALDOFDG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CUOTAS_RESC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CUOTAS_RESC_COB");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CUOTAS_PENDIENTES");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VALOR_RESCATE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CUOTAS_PENDIENTES");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorFDG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Rescate");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAnalisisFondoGarantia));
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CUOTAS_PENDIENTES");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operaciones");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorFDG");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Rescate");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Operaciones", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Operaciones", 1, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorFDG", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorFDG", 2, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Rescate", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Rescate", 3, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CEDULA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OPERACION");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CLIENTE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FECHA_VCTO_CUOTA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MONTO_FINANCIADO");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CAPITAL_AMORTIZACION");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("INTERES");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MORA");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VALOR_A_PAGAR");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NUMERO_CUOTA");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DIAS_MORA");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NUMERO_CUOTAS_VENCIDAS");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CAPITAL_AMORTIZACION", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CAPITAL_AMORTIZACION", 5, true);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "INTERES", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "INTERES", 6, true);
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MORA", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MORA", 7, true);
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "VALOR_A_PAGAR", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "VALOR_A_PAGAR", 8, true);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CEDULA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OPERACION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CLIENTE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FECHA_VCTO_CUOTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MONTO_FINANCIADO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CAPITAL_AMORTIZACION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("INTERES");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MORA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VALOR_A_PAGAR");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NUMERO_CUOTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DIAS_MORA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NUMERO_CUOTAS_VENCIDAS");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			this.ultraGridPrincipal = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.uGridPendientes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.btnExcel = new System.Windows.Forms.Button();
			this.uGridClientesSube = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtCuota = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.ultraGridPrincipal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPendientes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridClientesSube)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuota)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGridPrincipal
			// 
			this.ultraGridPrincipal.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGridPrincipal.DataSource = this.ultraDataSource4;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGridPrincipal.DisplayLayout.Appearance = appearance1;
			this.ultraGridPrincipal.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 88;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 89;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 180;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "MONTO";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 79;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "SALDO F D G";
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Width = 90;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Header.Caption = "CUOTAS RESC";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Header.Caption = "CUOTAS PEND";
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Width = 101;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance7;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "VALOR RESCATE";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 101;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand1.Header.Caption = "DETALLE DE CLIENTES A RESCATAR";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance8;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance9;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGridPrincipal.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGridPrincipal.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.ultraGridPrincipal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGridPrincipal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGridPrincipal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ultraGridPrincipal.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 7.5F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGridPrincipal.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.ultraGridPrincipal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGridPrincipal.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGridPrincipal.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGridPrincipal.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.ultraGridPrincipal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGridPrincipal.Location = new System.Drawing.Point(8, 64);
			this.ultraGridPrincipal.Name = "ultraGridPrincipal";
			this.ultraGridPrincipal.Size = new System.Drawing.Size(968, 208);
			this.ultraGridPrincipal.TabIndex = 776;
			this.ultraGridPrincipal.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.UInt32);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(208, 12);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 780;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// uGridPendientes
			// 
			this.uGridPendientes.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPendientes.DataSource = this.ultraDataSource1;
			appearance16.BackColor = System.Drawing.Color.White;
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPendientes.DisplayLayout.Appearance = appearance16;
			this.uGridPendientes.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance17;
			ultraGridColumn10.Header.Caption = "CUOTAS PEND";
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Width = 91;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn11.CellAppearance = appearance18;
			ultraGridColumn11.Header.Caption = "N. OPER";
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 60;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance19;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "SALDO F D G";
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Width = 101;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance20;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "VALOR RESCATE";
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 104;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			ultraGridBand2.Header.Caption = "RESUMEN POR CUOTAS PENDIENTES DE COBRO";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Center;
			summarySettings3.Appearance = appearance21;
			summarySettings3.DisplayFormat = "{0: #,##0}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance22;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance23;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridPendientes.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPendientes.DisplayLayout.Override.ActiveRowAppearance = appearance24;
			this.uGridPendientes.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPendientes.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPendientes.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance25.BackColor = System.Drawing.Color.Transparent;
			this.uGridPendientes.DisplayLayout.Override.CardAreaAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 7.5F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPendientes.DisplayLayout.Override.HeaderAppearance = appearance26;
			this.uGridPendientes.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance27.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance27.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPendientes.DisplayLayout.Override.RowAlternateAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPendientes.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPendientes.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.uGridPendientes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPendientes.Location = new System.Drawing.Point(984, 64);
			this.uGridPendientes.Name = "uGridPendientes";
			this.uGridPendientes.Size = new System.Drawing.Size(377, 208);
			this.uGridPendientes.TabIndex = 781;
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel4.Location = new System.Drawing.Point(16, 16);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(48, 15);
			this.ultraLabel4.TabIndex = 786;
			this.ultraLabel4.Text = "Cuota N.";
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(301, 12);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(76, 23);
			this.btnExcel.TabIndex = 787;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// uGridClientesSube
			// 
			this.uGridClientesSube.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridClientesSube.DataSource = this.ultraDataSource2;
			appearance30.BackColor = System.Drawing.Color.White;
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridClientesSube.DisplayLayout.Appearance = appearance30;
			this.uGridClientesSube.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Width = 90;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 93;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Width = 264;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.Caption = "FECHA VCTO";
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn17.Width = 79;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance31;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "MONTO";
			ultraGridColumn18.Header.VisiblePosition = 4;
			ultraGridColumn18.Width = 100;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance32;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "CAPITAL";
			ultraGridColumn19.Header.VisiblePosition = 5;
			ultraGridColumn19.Width = 100;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance33;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 6;
			ultraGridColumn20.Width = 100;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance34;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 7;
			ultraGridColumn21.Width = 101;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance35;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "A PAGAR";
			ultraGridColumn22.Header.VisiblePosition = 8;
			ultraGridColumn22.Width = 101;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance36;
			ultraGridColumn23.Header.Caption = "CUOTA N.";
			ultraGridColumn23.Header.VisiblePosition = 9;
			ultraGridColumn23.Width = 101;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance37;
			ultraGridColumn24.Header.Caption = "DIAS MORA";
			ultraGridColumn24.Header.VisiblePosition = 10;
			ultraGridColumn24.Width = 101;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance38;
			ultraGridColumn25.Header.Caption = "CUOTAS VENCIDAS";
			ultraGridColumn25.Header.VisiblePosition = 11;
			ultraGridColumn25.Width = 101;
			ultraGridBand3.Columns.AddRange(new object[] {
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
			ultraGridBand3.Header.Caption = "LISTADO DE CLIENTES A RESCATAR POR CUOTA";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance39;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance40;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance41;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance42;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8,
																																															summarySettings9});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridClientesSube.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance43.ForeColor = System.Drawing.Color.Black;
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridClientesSube.DisplayLayout.Override.ActiveRowAppearance = appearance43;
			this.uGridClientesSube.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridClientesSube.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridClientesSube.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance44.BackColor = System.Drawing.Color.Transparent;
			this.uGridClientesSube.DisplayLayout.Override.CardAreaAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance45.FontData.BoldAsString = "True";
			appearance45.FontData.Name = "Arial";
			appearance45.FontData.SizeInPoints = 7.5F;
			appearance45.ForeColor = System.Drawing.Color.White;
			appearance45.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridClientesSube.DisplayLayout.Override.HeaderAppearance = appearance45;
			this.uGridClientesSube.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance46.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance46.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridClientesSube.DisplayLayout.Override.RowAlternateAppearance = appearance46;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridClientesSube.DisplayLayout.Override.RowSelectorAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridClientesSube.DisplayLayout.Override.SelectedRowAppearance = appearance48;
			this.uGridClientesSube.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridClientesSube.Location = new System.Drawing.Point(8, 280);
			this.uGridClientesSube.Name = "uGridClientesSube";
			this.uGridClientesSube.Size = new System.Drawing.Size(1352, 208);
			this.uGridClientesSube.TabIndex = 788;
			// 
			// btnGenerar
			// 
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(1240, 12);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(120, 23);
			this.btnGenerar.TabIndex = 789;
			this.btnGenerar.Text = "Generar Rescates";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn17.DataType = typeof(System.DateTime);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn25});
			// 
			// txtCuota
			// 
			appearance49.ForeColor = System.Drawing.Color.Black;
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuota.Appearance = appearance49;
			this.txtCuota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuota.FormatString = "###0";
			this.txtCuota.Location = new System.Drawing.Point(88, 13);
			this.txtCuota.MaxValue = 18;
			this.txtCuota.MinValue = 0;
			this.txtCuota.Name = "txtCuota";
			this.txtCuota.PromptChar = ' ';
			this.txtCuota.Size = new System.Drawing.Size(100, 22);
			this.txtCuota.TabIndex = 790;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 48);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1368, 8);
			this.groupBox1.TabIndex = 793;
			this.groupBox1.TabStop = false;
			// 
			// frmAnalisisFondoGarantia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1368, 494);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtCuota);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.uGridClientesSube);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.uGridPendientes);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.ultraGridPrincipal);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmAnalisisFondoGarantia";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Rescates";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAnalisisFondoGarantia_KeyDown);
			this.Load += new System.EventHandler(this.frmAnalisisFondoGarantia_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGridPrincipal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPendientes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridClientesSube)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuota)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmAnalisisFondoGarantia_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0719");
			
			if (!Funcion.Permiso("649", cdsSeteoF))
			{				
				MessageBox.Show("Analisis Fondo de Garantia ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				

				return;
			}

			this.txtCuota.Value = 18;

			Consulta();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void Consulta()
		{
			string sSQL = string.Format("Exec AnalisisFondoGarantia {0}, {1}", 1, (int)this.txtCuota.Value);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.ultraGridPrincipal);

			sSQL = string.Format("Exec AnalisisFondoGarantia {0}, {1}", 2, (int)this.txtCuota.Value);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridPendientes);

			sSQL = string.Format("Exec AnalisisGeneralFondoGarantia {0}", (int)this.txtCuota.Value);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridClientesSube);			
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{       
			Consulta();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}		

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGridPrincipal);
			FuncionesProcedimientos.ExportaExcel(this.uGridPendientes);
			FuncionesProcedimientos.ExportaExcel(this.uGridClientesSube);
		}

		private void frmAnalisisFondoGarantia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.OK;
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.OK == MessageBox.Show(string.Format("Esta seguro de Subir {0} Rescates", this.uGridClientesSube.Rows.Count), "Point Technology", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				int iCont = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridClientesSube.Rows.All)
				{				
					int iCuoC_CuoV = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select NumCuotas - Cuotas_Vendidas From Compra Where NumeroDeOperacion = '{0}' And idEntidadFinanciera = 4", dr.Cells["OPERACION"].Value.ToString()));

					int iNUMERO_CUOTA = 0;

					if (iCuoC_CuoV > 0) iNUMERO_CUOTA = Convert.ToInt32(dr.Cells["NUMERO_CUOTA"].Value) - iCuoC_CuoV;

					int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Cre_ValidaCuotaPagadaRescate '{0}', {1}", dr.Cells["OPERACION"].Value.ToString(), iNUMERO_CUOTA));

					int iDiasVencidos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT DATEDIFF(DAY, '{0}', GETDATE())", Convert.ToDateTime(dr.Cells["FECHA_VCTO_CUOTA"].Value).ToString("yyyyMMdd")));
						
					if (iEstado != 2)
					{				
						if (Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ValidaSubeCarteraRescatada '{0}'", dr.Cells["OPERACION"].Value.ToString())) == 0)
						{
							string sSQLAct = string.Format("Exec GrabaRegFondo '{0}', '{1}', '{2}', '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}, {11}", 
								dr.Cells["CEDULA"].Value.ToString(), dr.Cells["OPERACION"].Value.ToString(), dr.Cells["CLIENTE"].Value.ToString(), Convert.ToDateTime(dr.Cells["FECHA_VCTO_CUOTA"].Value).ToString("yyyyMMdd"), 
								Convert.ToDecimal(dr.Cells["MONTO_FINANCIADO"].Value), Convert.ToDecimal(dr.Cells["CAPITAL_AMORTIZACION"].Value), Convert.ToDecimal(dr.Cells["INTERES"].Value), Convert.ToDecimal(dr.Cells["MORA"].Value), Convert.ToDecimal(dr.Cells["VALOR_A_PAGAR"].Value), 
								iNUMERO_CUOTA, Convert.ToInt32(dr.Cells["DIAS_MORA"].Value), Convert.ToInt32(dr.Cells["NUMERO_CUOTAS_VENCIDAS"].Value));
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
						
							iCont++;
						}			
					}					
				}

				MessageBox.Show(string.Format("{0} registros de rescates subidos correctamente", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information); 
			}
		}
	}
}
