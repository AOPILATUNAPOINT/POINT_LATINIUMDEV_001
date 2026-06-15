using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmAprobacionGastosDeViaje.
	/// </summary>
	public class frmAprobacionGastosDeViaje : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaDePago;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbcuenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAprobacionGastosDeViaje()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAprobacionGastosDeViaje));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Personal_Solicitante", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRetorna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Destino");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorSolicitado");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaDePago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbcuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorAutorizado");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGastosViaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aceptar", 0);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anular", 1);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Imprimir", 2);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorSolicitado", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorSolicitado", 6, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorAutorizado", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorAutorizado", 10, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Personal_Solicitante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaRetorna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Destino");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorSolicitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorAutorizado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGastosViaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDeposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("N_Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngresoReembolsoViaticos", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGastosViajes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReembolsoViaticos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Definido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AdministrativoSolicitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GerentesSolicitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Administrativo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gerentes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotalAcreditado");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbcuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbFormaDePago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbcuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaDePago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			this.SuspendLayout();
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(200, 16);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 318;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 317;
			this.label6.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 16);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 316;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(72, 16);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 315;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// optEstados
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance3;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance4;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendientes";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Aprobado";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Anulados";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(72, 45);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(312, 29);
			this.optEstados.TabIndex = 319;
			this.optEstados.Text = "Todos";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(456, 85);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 323;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(352, 85);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 321;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(16, 88);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(39, 16);
			this.lblCliente.TabIndex = 322;
			this.lblCliente.Text = "Buscar";
			// 
			// txtBusqueda
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance5;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(72, 86);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(272, 21);
			this.txtBusqueda.TabIndex = 320;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 120);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1184, 8);
			this.groupBox1.TabIndex = 324;
			this.groupBox1.TabStop = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource4;
			appearance6.BackColor = System.Drawing.Color.White;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Width = 70;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Personal Solicitante";
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Width = 230;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 5;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Fecha Salida";
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.Width = 95;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Fecha Retorna";
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Width = 95;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Width = 190;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance7;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Valor Solicitado";
			ultraGridColumn7.Header.VisiblePosition = 9;
			ultraGridColumn7.Width = 90;
			ultraGridColumn8.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn8.Header.Caption = "Forma De Pago";
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn8.Width = 150;
			ultraGridColumn9.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 11;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.Header.VisiblePosition = 12;
			ultraGridColumn10.Width = 190;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance8;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Valor Autorizado";
			ultraGridColumn11.Header.VisiblePosition = 16;
			ultraGridColumn11.Width = 95;
			ultraGridColumn12.Header.VisiblePosition = 13;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 15;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			appearance9.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellButtonAppearance = appearance9;
			ultraGridColumn15.Header.Caption = "Aprobar";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn15.Width = 45;
			ultraGridColumn16.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			appearance10.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellButtonAppearance = appearance10;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn16.Width = 43;
			ultraGridColumn17.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			appearance11.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn17.CellButtonAppearance = appearance11;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn17.Width = 50;
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
																										 ultraGridColumn17});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance12;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance13;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance14.ForeColor = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.CaptionAppearance = appearance14;
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 8F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance17;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(0, 136);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1184, 232);
			this.ultraGrid1.TabIndex = 325;
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(System.DateTime);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn14});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(System.DateTime);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// cmbcuenta
			// 
			this.cmbcuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbcuenta.DataSource = this.ultraDataSource3;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 150;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.cmbcuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbcuenta.DisplayMember = "Descripcion";
			this.cmbcuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbcuenta.Location = new System.Drawing.Point(1152, 8);
			this.cmbcuenta.Name = "cmbcuenta";
			this.cmbcuenta.Size = new System.Drawing.Size(24, 32);
			this.cmbcuenta.TabIndex = 986;
			this.cmbcuenta.ValueMember = "idCuenta";
			this.cmbcuenta.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// cmbFormaDePago
			// 
			this.cmbFormaDePago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaDePago.DataSource = this.ultraDataSource2;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Width = 150;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbFormaDePago.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbFormaDePago.DisplayMember = "FormaPago";
			this.cmbFormaDePago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaDePago.Location = new System.Drawing.Point(1072, 8);
			this.cmbFormaDePago.Name = "cmbFormaDePago";
			this.cmbFormaDePago.Size = new System.Drawing.Size(24, 32);
			this.cmbFormaDePago.TabIndex = 987;
			this.cmbFormaDePago.ValueMember = "idFormaPago";
			this.cmbFormaDePago.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn25.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26});
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
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			appearance20.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance20;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn26.Header.VisiblePosition = 3;
			ultraGridColumn27.Header.VisiblePosition = 5;
			ultraGridColumn28.Header.VisiblePosition = 6;
			ultraGridColumn29.Header.VisiblePosition = 7;
			ultraGridColumn30.Header.VisiblePosition = 8;
			ultraGridColumn31.Header.VisiblePosition = 10;
			ultraGridColumn32.Header.VisiblePosition = 11;
			ultraGridColumn33.Header.VisiblePosition = 9;
			ultraGridColumn34.Header.VisiblePosition = 12;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			ultraGridBand4.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultraGrid2.DisplayLayout.GroupByBox.Hidden = true;
			appearance21.ForeColor = System.Drawing.Color.Black;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance21;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance22;
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.CellAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance24.FontData.BoldAsString = "True";
			appearance24.FontData.Name = "Arial";
			appearance24.FontData.SizeInPoints = 8F;
			appearance24.ForeColor = System.Drawing.Color.White;
			appearance24.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance24;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance25.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance25.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance27;
			this.ultraGrid2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(1104, 8);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(40, 40);
			this.ultraGrid2.TabIndex = 1009;
			this.ultraGrid2.Visible = false;
			// 
			// frmAprobacionGastosDeViaje
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1192, 374);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.cmbFormaDePago);
			this.Controls.Add(this.cmbcuenta);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.MaximizeBox = false;
			this.Name = "frmAprobacionGastosDeViaje";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Aprobación Gastos De Viaje";
			this.Load += new System.EventHandler(this.frmAprobacionGastosDeViaje_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbcuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaDePago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmAprobacionGastosDeViaje_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0224");

			if (!Funcion.Permiso("1281", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar  Solicitud De Gastos De Viaje", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
//			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
//			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = DateTime.Today.AddDays(-30);
			this.dtHasta.Value = DateTime.Today;

			this.cmbcuenta.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("EXEC ConsultaGastosViajes 2 "));
			this.cmbFormaDePago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ConsultaGastosViajes 3 "));
			this.Seguridad ();
		}

		private void Seguridad ()
		{
			if ( (int) this.optEstados.Value == 0)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Anular"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = true;
			}
			if ( (int) this.optEstados.Value == 1)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = false;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Anular"].Hidden = false;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = true;
			}
			if ( (int) this.optEstados.Value == 2)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Anular"].Hidden = false;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = false;
			}
			if ( (int) this.optEstados.Value == 3)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Anular"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = true;
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
			this.lblContador.Text = "";
			string sSQL = string.Format("Exec ConsultaAprobacionGastosViajes '{0}', '{1}', {2}", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),(int)this.optEstados.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);

			this.lblContador.Text = this.ultraGrid1.Rows.Count + " Registros Encontrados";
			this.Seguridad ();
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Aprobacion
			if (e.Cell.Column.ToString() == "Aceptar")
			{	
				if (!miAcceso.Autorizacion) 
				{
					MessageBox.Show("No tiene acceso AUTORIZAR ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#region Valida Observacion

				if (e.Cell.Row.Cells["Observacion"].Value.ToString().Trim().Length == 0 )
				{
					MessageBox.Show("Ingrese un Observación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
				if (e.Cell.Row.Cells["Observacion"].Value.ToString().Trim().Length < 10 )
				{
					MessageBox.Show("La Observación debe tener minimo 10 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}

				if (e.Cell.Row.Cells["Observacion"].Value.ToString().Trim().Length > 100 )
				{
					MessageBox.Show("La Observación debe tener maximo 100 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}

				if (e.Cell.Row.Cells["idFormaPago"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la forma de pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}		
				
				if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 0)
				{
					MessageBox.Show("Seleccione la forma de pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}	

       	#endregion Valida Observacion

				using (frmReembolsoViaticos Viaticos = new frmReembolsoViaticos( true , (int)e.Cell.Row.Cells["idPersonal"].Value, Convert.ToDateTime ( e.Cell.Row.Cells["FechaSalida"].Value), Convert.ToDateTime ( e.Cell.Row.Cells["FechaRetorna"].Value), (int) e.Cell.Row.Cells["idGastosViaje"].Value, (int) e.Cell.Row.Cells["Tipo"].Value))
				{
					if (DialogResult.OK == Viaticos.ShowDialog( ))
					{
						decimal DSolicitado = 0.00m;
						if (Viaticos.UgridDetalle.Rows.Count > 0)
						{
							//decimal DSolicitado = 0;
							string sSQLPa = string.Format(" EXEC ConsultaDetalleIngresoReembolsoViaticos   {0}, {1}, {2} ", 0, 0, 0);
							FuncionesProcedimientos.dsGeneral(sSQLPa, this.ultraGrid2);
							e.Cell.Row.Cells["ValorAutorizado"].Value = 0;
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Viaticos.UgridDetalle.Rows.All)
							{
								this.ultraGrid2.Rows.Band.AddNew();
								this.ultraGrid2.ActiveRow.Cells["idIngresoReembolsoViaticos"].Value = (int) dr.Cells["idIngresoReembolsoViaticos"].Value;
								this.ultraGrid2.ActiveRow.Cells["idReembolsoViaticos"].Value = (int) dr.Cells["idReembolsoViaticos"].Value;
								this.ultraGrid2.ActiveRow.Cells["Concepto"].Value = Convert.ToString ( dr.Cells["Concepto"].Value);
								this.ultraGrid2.ActiveRow.Cells["AdministrativoSolicitado"].Value = Convert.ToDecimal( dr.Cells["AdministrativoSolicitado"].Value);
								this.ultraGrid2.ActiveRow.Cells["GerentesSolicitado"].Value = Convert.ToDecimal( dr.Cells["GerentesSolicitado"].Value);
								this.ultraGrid2.ActiveRow.Cells["Administrativo"].Value = Convert.ToDecimal( dr.Cells["Administrativo"].Value);
								this.ultraGrid2.ActiveRow.Cells["Gerentes"].Value = Convert.ToDecimal( dr.Cells["Gerentes"].Value);
								this.ultraGrid2.ActiveRow.Cells["SubTotal"].Value = Convert.ToDecimal( dr.Cells["SubTotal"].Value);
								this.ultraGrid2.ActiveRow.Cells["SubTotalAcreditado"].Value = Convert.ToDecimal( dr.Cells["SubTotalAcreditado"].Value);
								this.ultraGrid2.ActiveRow.Cells["idGastosViajes"].Value = (int) dr.Cells["idGastosViajes"].Value;
							//	DSolicitado = Convert.ToDecimal( e.Cell.Row.Cells["ValorAutorizado"].Value )  + Convert.ToDecimal( dr.Cells["SubTotalAcreditado"].Value);
								//e.Cell.Row.Cells["ValorAutorizado"].Value = (((Convert.ToDecimal( e.Cell.Row.Cells["ValorAutorizado"].Value)  + Convert.ToDecimal( dr.Cells["SubTotalAcreditado"].Value )) *90 )/100);
								e.Cell.Row.Cells["ValorAutorizado"].Value = (((Convert.ToDecimal( e.Cell.Row.Cells["ValorAutorizado"].Value)  + Convert.ToDecimal( dr.Cells["SubTotalAcreditado"].Value ))  ));
							   
							}
						}
						if(  (decimal) e.Cell.Row.Cells["ValorAutorizado"].Value == 0)
						{
							MessageBox.Show("El valor autorizado no puede ser 0 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
             DSolicitado = ((decimal)e.Cell.Row.Cells["ValorAutorizado"].Value  );
             e.Cell.Row.Cells["ValorAutorizado"].Value  = (DSolicitado * 90 )/100;
						if	(DialogResult.Yes == MessageBox.Show( " Valor Solcitado es: " + DSolicitado + "\n\n el valor acreditado sera " + Convert.ToDecimal(e.Cell.Row.Cells["ValorAutorizado"].Value) + "\n\n ¿Esta seguro de Continuar?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{

							#region Guardar      
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

									#region Graba Ingreso Gastos
									if((int) this.ultraGrid2.Rows.Count > 0)
									{						
										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
										{
											string  sSQLDeta = string.Format("Exec GuardaDettaleIngresoReembolsoViaticos {0}, {1}, '{2}', {3}, {4}, {5}, {6}, {7}, {8},{9}" ,
												Convert.ToInt32 ( dr.Cells["idIngresoReembolsoViaticos"].Value), 
												Convert.ToInt32 ( dr.Cells["idReembolsoViaticos"].Value), 
												Convert.ToString  ( dr.Cells["Concepto"].Value),
												Convert.ToDecimal ( dr.Cells["AdministrativoSolicitado"].Value),
												Convert.ToDecimal ( dr.Cells["GerentesSolicitado"].Value)
												,Convert.ToDecimal ( dr.Cells["Administrativo"].Value)
												,Convert.ToDecimal ( dr.Cells["Gerentes"].Value)
												, (int) e.Cell.Row.Cells["idGastosViaje"].Value,
												Convert.ToDecimal  ( dr.Cells["SubTotal"].Value),
												Convert.ToDecimal ( dr.Cells["SubTotalAcreditado"].Value));
											//							Funcion.EjecutaSQL(cdsSeteoF, sSQLDeta);

											oCmdActualiza.CommandText = sSQLDeta;
											oCmdActualiza.ExecuteNonQuery();	
										}						
									}
									#endregion Graba Ingreso Gastos
					      
									string sSQLAutEs = string.Format("Update GastosDeViaje set Estado = 2 , Observacion = '{1}', ValorAutorizado = {2}, idFormaPago = {3} where idGastosViaje = {0}",  
										(int) e.Cell.Row.Cells["idGastosViaje"].Value, 
										Convert.ToString ( e.Cell.Row.Cells["Observacion"].Value), (decimal) e.Cell.Row.Cells["ValorAutorizado"].Value , (int) e.Cell.Row.Cells["idFormaPago"].Value);			
									oCmdActualiza.CommandText = sSQLAutEs;
									oCmdActualiza.ExecuteNonQuery();		

									string sSQLAut = string.Format("Exec CreaAutorizacionesDePago {0}, {1}, '{2}', '{3}', {4}, {5}, '1 DE 1', '{6}', 0, 0", 
										10, (int) e.Cell.Row.Cells["idGastosViaje"].Value, 
										Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd"), Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd"), 
										Convert.ToDecimal(e.Cell.Row.Cells["ValorAutorizado"].Value ), (int) e.Cell.Row.Cells["idFormaPago"].Value,
										0, 0);			
									oCmdActualiza.CommandText = sSQLAut;
									oCmdActualiza.ExecuteNonQuery();	

									oTransaction.Commit();

									Cursor = Cursors.Default;	

									#region impresion 
									//								ParameterFields paramFields = new ParameterFields();
									//								ParameterField paramFieldCedula = new ParameterField ();
									//								ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
									//								paramFieldCedula.ParameterFieldName = "@idGastosViaje";
									//								ValCedula.Value = (int) e.Cell.Row.Cells["idGastosViaje"].Value;
									//								paramFieldCedula.CurrentValues.Add (ValCedula);
									//								paramFields.Add (paramFieldCedula);
									//								Reporte miRep = new Reporte("ReporteGastosViaje.rpt", "");
									//								miRep.MdiParent = this.MdiParent;
									//								miRep.Titulo("Planificacion De Carpa");
									//								miRep.crVista.ParameterFieldInfo = paramFields;
									//								miRep.Show();
									#endregion impresion 

									MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
									this.btnBuscar_Click (sender, e);
				
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
						else
						{

							this.DialogResult = DialogResult.No;
						
						}
					}
				}
			}
			#endregion Aprobacion

			#region impresiones

			#region impresion 
								
			if (e.Cell.Column.ToString() == "Imprimir")
			{	
				ParameterFields paramFields = new ParameterFields();
				ParameterField paramFieldCedula = new ParameterField ();
				ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
				paramFieldCedula.ParameterFieldName = "@idGastosViaje";
				ValCedula.Value = (int) e.Cell.Row.Cells["idGastosViaje"].Value;
				paramFieldCedula.CurrentValues.Add (ValCedula);
				paramFields.Add (paramFieldCedula);
				Reporte miRep = new Reporte("ReporteGastosViaje.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Planificacion De Carpa");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
			}
			#endregion impresion 

			#endregion impresiones

			if (e.Cell.Column.ToString() == "Anular")
			{	
				if (!miAcceso.Anular) 
				{
					MessageBox.Show("No tiene acceso ANULAR ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro de Anular el registro ?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
				int estado =	Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CambiaEstadoGestionesViajes {0}, 3 ", (int) e.Cell.Row.Cells["idGastosViaje"].Value));
			 this.btnBuscar_Click (sender, e);
				}
			}
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnBuscar_Click (sender, e );
		}
	}
}
