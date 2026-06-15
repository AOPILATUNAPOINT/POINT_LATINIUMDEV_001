using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmProyeccionDePagosQuincenal.
	/// </summary>
	public class frmProyeccionDePagosQuincenal : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSaldosBancarios;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPrestamos;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdProvision;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridProvisiones;
		private System.Windows.Forms.Button btnActualizar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.ComboBox cboFecha;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmProyeccionDePagosQuincenal()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvisionesSaldosBancarios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoBancario");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ChequesPorEntregar");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Provision");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoAFavor");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ChequesPorEntregar", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ChequesPorEntregar", 4, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoBancario", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoBancario", 3, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Provision", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Provision", 5, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoAFavor", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoAFavor", 6, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvisionesSaldosBancarios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoBancario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ChequesPorEntregar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Provision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoAFavor");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvisionesPrestamos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvisiones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiaEmision");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiaPago");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Parcial");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Parcial", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Parcial", 6, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Reserva", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Reserva", 7, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvisionesPrestamos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvisiones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiaEmision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Parcial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Reserva");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleProvisiones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvisiones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPlanificacionPagos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiaEmision");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiaPago");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaContable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaContable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Parcial");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAsignar", 0);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmProyeccionDePagosQuincenal));
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Parcial", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Parcial", 11, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Reserva", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Reserva", 12, true);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleProvisiones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvisiones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPlanificacionPagos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiaEmision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaContable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuentaContable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Parcial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Reserva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			this.uGridSaldosBancarios = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridPrestamos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridProvisiones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtIdProvision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.cboFecha = new System.Windows.Forms.ComboBox();
			((System.ComponentModel.ISupportInitialize)(this.uGridSaldosBancarios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPrestamos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridProvisiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProvision)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridSaldosBancarios
			// 
			this.uGridSaldosBancarios.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSaldosBancarios.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridSaldosBancarios.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 318;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 223;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "(+) Saldo Bancario";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 110;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "(-) Ch/Entregar";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 110;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "(-) Provision";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 110;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "(=) Saldo A Favor";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 110;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand1.Header.Caption = "Saldo Bancario";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance6;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance7;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance8;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance9;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridSaldosBancarios.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSaldosBancarios.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.uGridSaldosBancarios.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSaldosBancarios.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridSaldosBancarios.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSaldosBancarios.DisplayLayout.Override.CellAppearance = appearance12;
			this.uGridSaldosBancarios.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8.5F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSaldosBancarios.DisplayLayout.Override.HeaderAppearance = appearance13;
			appearance14.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance14.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSaldosBancarios.DisplayLayout.Override.RowAlternateAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSaldosBancarios.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSaldosBancarios.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.uGridSaldosBancarios.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSaldosBancarios.Location = new System.Drawing.Point(8, 48);
			this.uGridSaldosBancarios.Name = "uGridSaldosBancarios";
			this.uGridSaldosBancarios.Size = new System.Drawing.Size(704, 152);
			this.uGridSaldosBancarios.TabIndex = 445;
			this.uGridSaldosBancarios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSaldosBancarios_KeyDown);
			this.uGridSaldosBancarios.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSaldosBancarios_AfterCellUpdate);
			this.uGridSaldosBancarios.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSaldosBancarios_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn6.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn7.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// uGridPrestamos
			// 
			this.uGridPrestamos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPrestamos.DataSource = this.ultraDataSource2;
			appearance17.BackColor = System.Drawing.Color.White;
			this.uGridPrestamos.DisplayLayout.Appearance = appearance17;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 219;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 131;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 98;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Cuenta Bancaria";
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Width = 217;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance18;
			ultraGridColumn12.Header.Caption = "Dia De Emision";
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 100;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance19;
			ultraGridColumn13.Header.Caption = "Dia De Pago";
			ultraGridColumn13.Header.VisiblePosition = 5;
			ultraGridColumn13.Width = 98;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance20;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Width = 99;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance21;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Width = 101;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			ultraGridBand2.Header.Caption = "Prestamos";
			ultraGridBand2.HeaderVisible = true;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance22;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance23;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings5,
																																															summarySettings6});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridPrestamos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPrestamos.DisplayLayout.Override.ActiveRowAppearance = appearance24;
			this.uGridPrestamos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPrestamos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance25.BackColor = System.Drawing.Color.Transparent;
			this.uGridPrestamos.DisplayLayout.Override.CardAreaAppearance = appearance25;
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPrestamos.DisplayLayout.Override.CellAppearance = appearance26;
			this.uGridPrestamos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance27.FontData.BoldAsString = "True";
			appearance27.FontData.Name = "Arial";
			appearance27.FontData.SizeInPoints = 8.5F;
			appearance27.ForeColor = System.Drawing.Color.White;
			appearance27.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPrestamos.DisplayLayout.Override.HeaderAppearance = appearance27;
			appearance28.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance28.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPrestamos.DisplayLayout.Override.RowAlternateAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPrestamos.DisplayLayout.Override.RowSelectorAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPrestamos.DisplayLayout.Override.SelectedRowAppearance = appearance30;
			this.uGridPrestamos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPrestamos.Location = new System.Drawing.Point(896, 48);
			this.uGridPrestamos.Name = "uGridPrestamos";
			this.uGridPrestamos.Size = new System.Drawing.Size(200, 152);
			this.uGridPrestamos.TabIndex = 446;
			this.uGridPrestamos.Visible = false;
			this.uGridPrestamos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridPrestamos_KeyDown);
			this.uGridPrestamos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPrestamos_AfterCellUpdate);
			this.uGridPrestamos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridPrestamos_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15});
			// 
			// uGridProvisiones
			// 
			this.uGridProvisiones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridProvisiones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridProvisiones.DataSource = this.ultraDataSource3;
			appearance31.BackColor = System.Drawing.Color.White;
			this.uGridProvisiones.DisplayLayout.Appearance = appearance31;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 148;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 100;
			ultraGridColumn18.Header.Caption = "Pagado";
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Width = 30;
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 171;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Width = 160;
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 74;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.Caption = "Cuenta Bancaria";
			ultraGridColumn22.Header.VisiblePosition = 6;
			ultraGridColumn22.Width = 220;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance32;
			ultraGridColumn23.Header.Caption = "Desde";
			ultraGridColumn23.Header.VisiblePosition = 9;
			ultraGridColumn23.Width = 45;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance33;
			ultraGridColumn24.Header.Caption = "Hasta";
			ultraGridColumn24.Header.VisiblePosition = 10;
			ultraGridColumn24.Width = 45;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.VisiblePosition = 7;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.VisiblePosition = 8;
			ultraGridColumn26.Width = 220;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance34;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 13;
			ultraGridColumn27.Width = 100;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance35;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 11;
			ultraGridColumn28.Width = 100;
			ultraGridColumn29.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 14;
			ultraGridColumn29.Width = 400;
			ultraGridColumn30.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			//appearance36.Image = ((object)(resources.GetObject("appearance36.Image")));
			appearance36.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance36.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn30.CellButtonAppearance = appearance36;
			ultraGridColumn30.Header.Caption = "...";
			ultraGridColumn30.Header.VisiblePosition = 12;
			ultraGridColumn30.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn30.Width = 25;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn30});
			ultraGridBand3.Header.Caption = "Provisiones";
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance37;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance38;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings7,
																																															summarySettings8});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridProvisiones.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridProvisiones.DisplayLayout.GroupByBox.Hidden = true;
			appearance39.ForeColor = System.Drawing.Color.Black;
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridProvisiones.DisplayLayout.Override.ActiveRowAppearance = appearance39;
			this.uGridProvisiones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridProvisiones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance40.BackColor = System.Drawing.Color.Transparent;
			this.uGridProvisiones.DisplayLayout.Override.CardAreaAppearance = appearance40;
			appearance41.ForeColor = System.Drawing.Color.Black;
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridProvisiones.DisplayLayout.Override.CellAppearance = appearance41;
			this.uGridProvisiones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance42.FontData.BoldAsString = "True";
			appearance42.FontData.Name = "Arial";
			appearance42.FontData.SizeInPoints = 8.5F;
			appearance42.ForeColor = System.Drawing.Color.White;
			appearance42.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridProvisiones.DisplayLayout.Override.HeaderAppearance = appearance42;
			appearance43.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance43.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridProvisiones.DisplayLayout.Override.RowAlternateAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridProvisiones.DisplayLayout.Override.RowSelectorAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridProvisiones.DisplayLayout.Override.SelectedRowAppearance = appearance45;
			this.uGridProvisiones.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridProvisiones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridProvisiones.Location = new System.Drawing.Point(8, 208);
			this.uGridProvisiones.Name = "uGridProvisiones";
			this.uGridProvisiones.Size = new System.Drawing.Size(1085, 296);
			this.uGridProvisiones.TabIndex = 447;
			this.uGridProvisiones.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridProvisiones_BeforeRowsDeleted);
			this.uGridProvisiones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridProvisiones_KeyDown);
			this.uGridProvisiones.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridProvisiones_ClickCellButton);
			this.uGridProvisiones.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridProvisiones_AfterCellUpdate);
			this.uGridProvisiones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridProvisiones_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(bool);
			ultraDataColumn18.DefaultValue = false;
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn29});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 16);
			this.label1.TabIndex = 449;
			this.label1.Text = "Fecha";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGenerar
			// 
			this.btnGenerar.CausesValidation = false;
			//this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(632, 7);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(80, 23);
			this.btnGenerar.TabIndex = 711;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1120, 8);
			this.groupBox1.TabIndex = 712;
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
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(320, 7);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 713;
			// 
			// txtIdProvision
			// 
			this.txtIdProvision.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdProvision.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdProvision.Enabled = false;
			this.txtIdProvision.Location = new System.Drawing.Point(1074, 8);
			this.txtIdProvision.Name = "txtIdProvision";
			this.txtIdProvision.PromptChar = ' ';
			this.txtIdProvision.Size = new System.Drawing.Size(16, 21);
			this.txtIdProvision.TabIndex = 714;
			this.txtIdProvision.Visible = false;
			// 
			// btnActualizar
			// 
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Enabled = false;
			//this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizar.Location = new System.Drawing.Point(544, 7);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(80, 23);
			this.btnActualizar.TabIndex = 715;
			this.btnActualizar.Text = "&Actualizar";
			this.btnActualizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Enabled = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			//this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(192, 7);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(80, 23);
			this.btnVer.TabIndex = 739;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cboFecha
			// 
			this.cboFecha.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboFecha.Enabled = false;
			this.cboFecha.Location = new System.Drawing.Point(64, 8);
			this.cboFecha.MaxDropDownItems = 15;
			this.cboFecha.Name = "cboFecha";
			this.cboFecha.Size = new System.Drawing.Size(112, 21);
			this.cboFecha.TabIndex = 757;
			this.cboFecha.SelectedIndexChanged += new System.EventHandler(this.cboFecha_SelectedIndexChanged);
			// 
			// frmProyeccionDePagosQuincenal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1098, 510);
			this.Controls.Add(this.cboFecha);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.txtIdProvision);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.uGridProvisiones);
			this.Controls.Add(this.uGridPrestamos);
			this.Controls.Add(this.uGridSaldosBancarios);
			this.KeyPreview = true;
			this.Name = "frmProyeccionDePagosQuincenal";
			this.Text = "Provisión De Pagos Quincenal";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmProyeccionDePagosQuincenal_KeyDown);
			this.Load += new System.EventHandler(this.frmProyeccionDePagosQuincenal_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridSaldosBancarios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPrestamos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridProvisiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProvision)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		int iEstado = 0;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void Fechas()
		{
			DateTime Fecha = DateTime.Today;

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select Fecha From Provisiones Order By Fecha", true);						
			
			while (dr.Read())
			{
				Fecha = dr.GetDateTime(0);
				this.cboFecha.Items.Add(Fecha.ToString("dd/MM/yyyy"));
			}
			dr.Close();

			this.cboFecha.Text = Fecha.ToString("dd/MM/yyyy");
		}

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}


		private void frmProyeccionDePagosQuincenal_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "044402");
			
			if (!Funcion.Permiso("915", cdsSeteoF))
			{				
				
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			Fechas();

			if (this.cboFecha.Items.Count > 0)
			{
				this.cboFecha.Enabled = true; 
				this.btnVer.Enabled = true;
			}
		}

		private void Consultas()
		{
			FuncionesProcedimientos.dsGeneral(string.Format("Exec ConsultaDetalleProvisiones {0}, 1", (int)this.txtIdProvision.Value), this.uGridSaldosBancarios);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec ConsultaDetalleProvisiones {0}, 2", (int)this.txtIdProvision.Value), this.uGridPrestamos);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec ConsultaDetalleProvisiones {0}, 3", (int)this.txtIdProvision.Value), this.uGridProvisiones);

			this.uGridProvisiones.DisplayLayout.Bands[0].SortedColumns.Add("Cuenta", false, true);			
			this.uGridProvisiones.Rows.ExpandAll(true);
		}

		private void SumatoriaProvisiones(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{			
//			this.uGridProvisiones.DisplayLayout.Bands[0].SortedColumns.Add("Cuenta", false, false);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridSaldosBancarios.Rows.All)
			{
				if ((int)dr.Cells["idCuenta"].Value == (int)e.Cell.Row.Cells["idCuenta"].Value)
				{
					dr.Cells["Provision"].Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select SUM(Reserva) From DetalleProvisiones Where idProvisiones = {0} And idCuenta = {1}",
						(int)this.txtIdProvision.Value, (int)e.Cell.Row.Cells["idCuenta"].Value));

					dr.Cells["SaldoAFavor"].Value = 
						Convert.ToDecimal(dr.Cells["SaldoBancario"].Value) - 
						Convert.ToDecimal(dr.Cells["ChequesPorEntregar"].Value) -
						Convert.ToDecimal(dr.Cells["Provision"].Value); 

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ProvisionesSaldosBancarios Set SaldoFinal = SaldoAFavor - PagoProveedores Where idProvisiones = {0} And idCuenta = {1}",	
						(int)this.txtIdProvision.Value, (int)e.Cell.Row.Cells["idCuenta"].Value));
				}
			}

//			this.uGridProvisiones.DisplayLayout.Bands[0].SortedColumns.Add("Cuenta", false, true);
			this.uGridProvisiones.Rows.ExpandAll(true);
		}

		private void SumatoriaPrestamos(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			decimal dProvisiones = 0.00m;
			
			this.uGridProvisiones.DisplayLayout.Bands[0].SortedColumns.Add("Cuenta", false, false);

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridProvisiones.Rows.All)
				if ((int)dr.Cells["idCuenta"].Value == (int)e.Cell.Row.Cells["idCuenta"].Value)
					dProvisiones = dProvisiones + Convert.ToDecimal(dr.Cells["Reserva"].Value);

			this.uGridProvisiones.DisplayLayout.Bands[0].SortedColumns.Add("Cuenta", false, true);

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridSaldosBancarios.Rows.All)
			{
				if ((int)dr.Cells["idCuenta"].Value == (int)e.Cell.Row.Cells["idCuenta"].Value)
				{
					dr.Cells["SaldoAFavor"].Value = 
						Convert.ToDecimal(dr.Cells["SaldoBancario"].Value) - 
						Convert.ToDecimal(dr.Cells["ChequesPorEntregar"].Value) -
						Convert.ToDecimal(e.Cell.Row.Cells["Reserva"].Value) - 
						dProvisiones;
				}
			}
		}

		private void SumatoriaSaldosBancarios(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{			
			decimal dProvisiones = 0.00m;
			decimal dPrestamos = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPrestamos.Rows.All)
				if ((int)dr.Cells["idCuenta"].Value == (int)e.Cell.Row.Cells["idCuenta"].Value)
					dPrestamos = dPrestamos + Convert.ToDecimal(dr.Cells["Reserva"].Value);

			this.uGridProvisiones.DisplayLayout.Bands[0].SortedColumns.Add("Cuenta", false, false);

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridProvisiones.Rows.All)
				if ((int)dr.Cells["idCuenta"].Value == (int)e.Cell.Row.Cells["idCuenta"].Value)
					dProvisiones = dProvisiones + Convert.ToDecimal(dr.Cells["Reserva"].Value);

			this.uGridProvisiones.DisplayLayout.Bands[0].SortedColumns.Add("Cuenta", false, true);

			e.Cell.Row.Cells["SaldoAFavor"].Value = 
				Convert.ToDecimal(e.Cell.Row.Cells["SaldoBancario"].Value) - 
				Convert.ToDecimal(e.Cell.Row.Cells["ChequesPorEntregar"].Value) -
				dPrestamos - dProvisiones;
		}
		
		private void uGridSaldosBancarios_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "ChequesPorEntregar", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SaldoBancario", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SaldoAFavor", 2);
		}

		private void uGridPrestamos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Parcial", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Reserva", 2);
		}

		private void uGridProvisiones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Parcial", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Reserva", 2);
		}

		private void uGridSaldosBancarios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridSaldosBancarios);
		}

		private void uGridPrestamos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridPrestamos);
		}

		private void uGridProvisiones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridProvisiones);
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
//			if (Funcion.iEscalarSQL(cdsSeteoF, "Select COUNT(*) From PlanificacionPagos Where Activo = 1") == 0)
//			{
//				MessageBox.Show("Antes de continuar ingrese la Planificacion de Pagos a provisionar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				return;
//			}
//			
//			if (Funcion.iEscalarSQL(cdsSeteoF, "Select COUNT(*) From CuentaCorriente Where Flujo = 1") == 0)
//			{
//				MessageBox.Show("Antes de continuar seleccione las cuentas bancarias de las cuales va a generar la provision", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				return;
//			}
//			
			if (Funcion.iEscalarSQL(cdsSeteoF, "Select COUNT(*) From Provisiones Where Estado = 1") > 0)
			{
				MessageBox.Show("Cierre las provisiones abiertas para continuar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de generar una nueva provision?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				DateTime dtFecha = DateTime.Today;
				if (this.cboFecha.Text.ToString().Length > 0) dtFecha = Convert.ToDateTime(this.cboFecha.Text);

				string sSQL = string.Format("Exec GeneraProvisiones 0, '{0}', ''", dtFecha.ToString("yyyyMMdd"));
				this.txtIdProvision.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				MessageBox.Show("Provisión Generada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				Consultas();

				Fechas();

				this.cboFecha.Enabled = true; 
				this.btnVer.Enabled = true;
			}
		}

		private void frmProyeccionDePagosQuincenal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridSaldosBancarios_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "SaldoBancario")
			{
				this.SumatoriaSaldosBancarios(sender, e);
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ProvisionesSaldosBancarios Set SaldoBancario = {0} Where idProvisionesSaldosBancarios = {1}",
					Convert.ToDecimal(e.Cell.Row.Cells["SaldoBancario"].Value), (int)e.Cell.Row.Cells["idProvisionesSaldosBancarios"].Value));
			}

			if (e.Cell.Column.ToString() == "Provision")
			{
				this.SumatoriaSaldosBancarios(sender, e);
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ProvisionesSaldosBancarios Set Provision = {0} Where idProvisionesSaldosBancarios = {1}",
					Convert.ToDecimal(e.Cell.Row.Cells["Provision"].Value), (int)e.Cell.Row.Cells["idProvisionesSaldosBancarios"].Value));
			}

			if (e.Cell.Column.ToString() == "SaldoAFavor")
			{
				this.SumatoriaSaldosBancarios(sender, e);
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ProvisionesSaldosBancarios Set SaldoAFavor = {0} Where idProvisionesSaldosBancarios = {1}",
					Convert.ToDecimal(e.Cell.Row.Cells["SaldoAFavor"].Value), (int)e.Cell.Row.Cells["idProvisionesSaldosBancarios"].Value));
			}
		}

		private void uGridPrestamos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Reserva")
			{
				this.SumatoriaPrestamos(sender, e);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ProvisionesPrestamos Set Reserva = {0} Where idProvisionesPrestamos = {1}",
					Convert.ToDecimal(e.Cell.Row.Cells["Reserva"].Value), (int)e.Cell.Row.Cells["idProvisionesPrestamos"].Value));
			}
		}

		private void uGridProvisiones_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Reserva")
			{				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleProvisiones Set Reserva = {0} Where idDetalleProvisiones = {1}",
					Convert.ToDecimal(e.Cell.Row.Cells["Reserva"].Value), (int)e.Cell.Row.Cells["idDetalleProvisiones"].Value));

				this.SumatoriaProvisiones(sender, e);
			}
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec GeneraProvisiones {0}, '{1}', ''", 
				(int)this.txtIdProvision.Value, Convert.ToDateTime(this.cboFecha.Text).ToString("yyyyMMdd"));
			this.txtIdProvision.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			MessageBox.Show("Provisión Actualizada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			Consultas();
		}

		private void uGridProvisiones_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (iEstado == 2)
			{
				MessageBox.Show("No puede eliminar lineas de Provisiones Cerradas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
			
			for (int i = 0; i < e.Rows.Length; i++)
			{
				if (Convert.ToDecimal(e.Rows[i].Cells["Reserva"].Value) > 0.00m)
				{
					MessageBox.Show("La linea seleccionada ya tiene valor reservado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}

				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar las lineas seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					string sSQL = string.Format("Delete From DetalleProvisiones Where idDetalleProvisiones = {0}", (int)e.Rows[i].Cells["idDetalleProvisiones"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					e.DisplayPromptMsg = false;
				}
				else e.Cancel = true;
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			Consultas();
		}

		private void cboFecha_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.cboFecha.Text.ToString().Length == 0) return;

			this.txtIdProvision.Value = 0;
			iEstado = 0;
			this.lblEstado.Text = "";
			
			this.btnVer.Enabled = false;
			this.btnActualizar.Enabled = false;

			this.uGridSaldosBancarios.DataSource = this.ultraDataSource1;
			this.uGridPrestamos.DataSource = this.ultraDataSource2;
			this.uGridProvisiones.DataSource = this.ultraDataSource3;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridSaldosBancarios);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridPrestamos);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridProvisiones);

			string sSQL = string.Format("Exec ConsultaProvisiones '{0}'", Convert.ToDateTime(this.cboFecha.Text).ToString("yyyyMMdd"));
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{				
				this.txtIdProvision.Value = dr.GetInt32(0);
				iEstado = dr.GetInt32(3);
				this.lblEstado.Text = dr.GetString(4);
			}
			dr.Close();

			if (iEstado == 0)
			{				
				FuncionesProcedimientos.EstadoGrids(true, this.uGridSaldosBancarios);
				FuncionesProcedimientos.EstadoGrids(true, this.uGridPrestamos);
				FuncionesProcedimientos.EstadoGrids(true, this.uGridProvisiones);
			}
			else if (iEstado == 1)
			{
				this.btnVer.Enabled = true;
				this.btnActualizar.Enabled = true;

				Consultas();

				FuncionesProcedimientos.EstadoGrids(true, this.uGridSaldosBancarios);
				FuncionesProcedimientos.EstadoGrids(true, this.uGridPrestamos);
				FuncionesProcedimientos.EstadoGrids(true, this.uGridProvisiones);
			}
			else if (iEstado == 2) 
			{
				this.btnVer.Enabled = true;
				Consultas();
			}
		}

		private void uGridProvisiones_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (iEstado == 2)
			{
				MessageBox.Show("La provisión esta cerrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Convert.ToDecimal(e.Cell.Row.Cells["Parcial"].Value) == 0.00m)
			{
				MessageBox.Show("EL valor parcial esta en Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			e.Cell.Row.Cells["Reserva"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Parcial"].Value);	
		}
	}
}
