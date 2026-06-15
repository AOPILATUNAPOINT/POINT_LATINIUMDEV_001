using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmProyeccionRescates.
	/// </summary>
	public class frmProyeccionRescates : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCartera;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmProyeccionRescates()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyeccionRescates");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("COLOCACION");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CEDULA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OPERACION");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CLIENTE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VENCIMIENTO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MONTO");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CAPITAL_CUOTA");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("INTERES_CUOTA");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MORA_CUOTA");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("A_PAGAR_CUOTA");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CUOTA");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DIAS_MORA_AL_02_12_2019");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DIAS_MORA_ACTUAL");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NUMERO_CUOTAS_VENCIDAS");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CUOTAS_CANCELADAS");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CUOTAS_VENCIDAS");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SALDO_CAPITAL");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("INTERES_X_PAGAR");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MORA_X_PAGAR");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("V_CUOTAS_VENCIDAS");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENVIADO_A_PAGAR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("COBRADO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gestion_Operador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gestion_Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Ultima_Cancelacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ultimo_Pago");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MONTO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MONTO", 6, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CAPITAL_CUOTA", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CAPITAL_CUOTA", 7, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "INTERES_CUOTA", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "INTERES_CUOTA", 8, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MORA_CUOTA", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MORA_CUOTA", 9, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "A_PAGAR_CUOTA", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "A_PAGAR_CUOTA", 10, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SALDO_CAPITAL", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SALDO_CAPITAL", 17, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "INTERES_X_PAGAR", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "INTERES_X_PAGAR", 18, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MORA_X_PAGAR", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MORA_X_PAGAR", 19, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "V_CUOTAS_VENCIDAS", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "V_CUOTAS_VENCIDAS", 20, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Ultimo_Pago", 27, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Ultimo_Pago", 27, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyeccionRescates");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("COLOCACION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CEDULA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OPERACION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CLIENTE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VENCIMIENTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MONTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CAPITAL_CUOTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("INTERES_CUOTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MORA_CUOTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_PAGAR_CUOTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CUOTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DIAS_MORA_AL_02_12_2019");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DIAS_MORA_ACTUAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NUMERO_CUOTAS_VENCIDAS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CUOTAS_CANCELADAS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CUOTAS_VENCIDAS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SALDO_CAPITAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("INTERES_X_PAGAR");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MORA_X_PAGAR");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("V_CUOTAS_VENCIDAS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENVIADO_A_PAGAR");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("COBRADO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gestion_Operador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gestion_Cobrador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Ultima_Cancelacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ultimo_Pago");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmProyeccionRescates));
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pagos");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OPERACION");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorEnviado");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrado");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapitalE");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorEnviado", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorEnviado", 3, true);
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cobrado", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cobrado", 4, true);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 5, true);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CapitalE", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CapitalE", 6, true);
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pagos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OPERACION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorEnviado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapitalE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.uGridCartera = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnVer = new System.Windows.Forms.Button();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.uGridCartera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridCartera
			// 
			this.uGridCartera.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCartera.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCartera.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridCartera.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Colocación";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Cédula";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Operación";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Cliente";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 200;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Vencimiento";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance2;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Monto";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance3;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Capital Cuota";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance4;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Interes Cuota";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance5;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Mora Cuota";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance6;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "A pagar Cuota";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance7;
			ultraGridColumn12.Header.Caption = "Cuota";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance8;
			ultraGridColumn13.Header.Caption = "Dias Mora";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance9;
			ultraGridColumn14.Header.Caption = "Dias Mora Actual";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance10;
			ultraGridColumn15.Header.Caption = "# Cuotas Vencidas";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance11;
			ultraGridColumn16.Header.Caption = "Cuotas Canceladas";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn17.CellAppearance = appearance12;
			ultraGridColumn17.Header.Caption = "Cuotas Vencidas";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance13;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Saldo Capital";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance14;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Interes x Pagar";
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance15;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "Mora X Pagar";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance16;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "V. Cuotas Vencidas";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.Caption = "Enviado a Pagar";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.Caption = "Cobrado";
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 25;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.Caption = "Gestión Operador";
			ultraGridColumn25.Header.VisiblePosition = 26;
			ultraGridColumn25.Width = 350;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.Caption = "Gestión Cobrador";
			ultraGridColumn26.Header.VisiblePosition = 27;
			ultraGridColumn26.Width = 350;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.Caption = "Fecha U. Cancelación";
			ultraGridColumn27.Header.VisiblePosition = 24;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance17;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.Caption = "Ultimo Pago";
			ultraGridColumn28.Header.VisiblePosition = 23;
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
																										 ultraGridColumn28});
			appearance18.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance18;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance19;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance20;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance21;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance22;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance23;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance24;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance25;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance26;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance27;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance28;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10});
			this.uGridCartera.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCartera.DisplayLayout.Override.ActiveRowAppearance = appearance29;
			this.uGridCartera.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance30.BackColor = System.Drawing.Color.Transparent;
			this.uGridCartera.DisplayLayout.Override.CardAreaAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance31.FontData.BoldAsString = "True";
			appearance31.FontData.Name = "Arial";
			appearance31.FontData.SizeInPoints = 8F;
			appearance31.ForeColor = System.Drawing.Color.White;
			appearance31.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCartera.DisplayLayout.Override.HeaderAppearance = appearance31;
			this.uGridCartera.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCartera.DisplayLayout.Override.RowSelectorAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCartera.DisplayLayout.Override.SelectedRowAppearance = appearance33;
			this.uGridCartera.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridCartera.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCartera.Location = new System.Drawing.Point(8, 64);
			this.uGridCartera.Name = "uGridCartera";
			this.uGridCartera.Size = new System.Drawing.Size(1184, 184);
			this.uGridCartera.TabIndex = 201;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.DateTime);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(System.DateTime);
			ultraDataColumn28.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn28});
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(312, 16);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 1016;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// txtBusqueda
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance34;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(112, 16);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(192, 21);
			this.txtBusqueda.TabIndex = 1015;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel1.Location = new System.Drawing.Point(8, 16);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(96, 15);
			this.ultraLabel1.TabIndex = 1014;
			this.ultraLabel1.Text = "Operación/ Cedula";
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
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(408, 16);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 1017;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1128, 16);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(64, 24);
			this.btnExcel.TabIndex = 1018;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-176, 48);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1576, 8);
			this.groupBox1.TabIndex = 1019;
			this.groupBox1.TabStop = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance35.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance35;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn29.CellAppearance = appearance36;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.Caption = "Operación";
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 120;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.Caption = "Cédula";
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn31.Width = 120;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance37;
			ultraGridColumn32.Header.Caption = "Valor Enviado";
			ultraGridColumn32.Header.VisiblePosition = 4;
			ultraGridColumn32.Width = 100;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance38;
			ultraGridColumn33.Header.VisiblePosition = 5;
			ultraGridColumn33.Width = 100;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance39;
			ultraGridColumn34.Header.VisiblePosition = 6;
			ultraGridColumn34.Width = 100;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance40;
			ultraGridColumn35.Header.Caption = "Capital E";
			ultraGridColumn35.Header.VisiblePosition = 7;
			ultraGridColumn35.Width = 100;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn36.Header.Caption = "Cliente";
			ultraGridColumn36.Header.VisiblePosition = 3;
			ultraGridColumn36.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36});
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance41;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance42;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance43;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance44;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13,
																																															summarySettings14});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance45.ForeColor = System.Drawing.Color.Black;
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance45;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance46.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance46;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance47.FontData.BoldAsString = "True";
			appearance47.FontData.Name = "Arial";
			appearance47.FontData.SizeInPoints = 8F;
			appearance47.ForeColor = System.Drawing.Color.White;
			appearance47.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance47;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance49;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 288);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1184, 216);
			this.ultraGrid1.TabIndex = 1020;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(-176, 248);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1576, 8);
			this.groupBox2.TabIndex = 1021;
			this.groupBox2.TabStop = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(160, 264);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(0, 17);
			this.label1.TabIndex = 1022;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn35.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36});
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.Black;
			this.label2.Location = new System.Drawing.Point(8, 264);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(122, 17);
			this.label2.TabIndex = 1023;
			this.label2.Text = "Clientes Rescatados";
			// 
			// frmProyeccionRescates
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1200, 510);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.uGridCartera);
			this.Name = "frmProyeccionRescates";
			this.Text = "Rescates";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmProyeccionRescates_KeyDown);
			this.Load += new System.EventHandler(this.frmProyeccionRescates_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridCartera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.lblContador.Text = "";
			this.label1.Text = "";

			string sSQLPa = string.Format("Exec ReporteProyeccionRescates '{0}' ", this.txtBusqueda.Text);
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridCartera);
			
			this.lblContador.Text = this.uGridCartera.Rows.Count + " REGISTROS ENCONTRADOS";



			string sSQLP = string.Format("Exec ReporteClientesRescatados '{0}' ", this.txtBusqueda.Text);
			FuncionesProcedimientos.dsGeneral(sSQLP , this.ultraGrid1);
			
			this.label1.Text = this.ultraGrid1.Rows.Count + " REGISTROS ENCONTRADOS";
		}
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmProyeccionRescates_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200607");

			if (!Funcion.Permiso("1317", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar al Reporte Seleccionado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar) this.btnExcel.Enabled = true;
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridCartera);
		}

		private void frmProyeccionRescates_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
