using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Threading;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_PagosCarteraPichincha.
	/// </summary>
	public class frmCre_PagosCarteraPichincha : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipos;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPagos;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFMora;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIMora;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnContabilizar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnImportarCartera;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optSeleccionados;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado_Cuota;
		private System.Windows.Forms.Button btnActualizaPagos;
		private System.Windows.Forms.Button btnGenerarPagos;
		private System.Windows.Forms.Button btnRatio;
		private System.Windows.Forms.Button btnProyeccionPagos;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnSeleccionar;
		private System.Windows.Forms.Button btnRescates;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIMoraP;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFMoraP;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_PagosCarteraPichincha()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_PagosCarteraPichincha));
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_Lista_Pagos_Cartera_Pichincha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo_Cobro");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operacion");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuota");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vence");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Cuota");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Cancelacion");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AbInteres");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AbCapital");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AbMora");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AbDiferimiento");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abono");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Cuota");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rango_Morosidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Pago");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Envio");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 0);
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Abono", 25, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Abono", 25, true);
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_Lista_Pagos_Cartera_Pichincha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo_Cobro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vence");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Cuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Cancelacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AbInteres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AbCapital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AbMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AbDiferimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Abono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado_Cuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rango_Morosidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Envio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estacion");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			this.optTipos = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label7 = new System.Windows.Forms.Label();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.uGridPagos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtFMora = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIMora = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.optSeleccionados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnContabilizar = new System.Windows.Forms.Button();
			this.btnActualizaPagos = new System.Windows.Forms.Button();
			this.btnImportarCartera = new System.Windows.Forms.Button();
			this.label3 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.optEstado_Cuota = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnGenerarPagos = new System.Windows.Forms.Button();
			this.btnRatio = new System.Windows.Forms.Button();
			this.btnProyeccionPagos = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnSeleccionar = new System.Windows.Forms.Button();
			this.btnRescates = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.txtIMoraP = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtFMoraP = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.optTipos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFMora)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIMora)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optSeleccionados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado_Cuota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIMoraP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFMoraP)).BeginInit();
			this.SuspendLayout();
			// 
			// optTipos
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipos.Appearance = appearance1;
			this.optTipos.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTipos.CheckedIndex = 0;
			this.optTipos.Enabled = false;
			this.optTipos.ItemAppearance = appearance2;
			this.optTipos.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Cobros";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Pre cancelaciones";
			this.optTipos.Items.Add(valueListItem1);
			this.optTipos.Items.Add(valueListItem2);
			this.optTipos.Items.Add(valueListItem3);
			this.optTipos.ItemSpacingVertical = 10;
			this.optTipos.Location = new System.Drawing.Point(8, 39);
			this.optTipos.Name = "optTipos";
			this.optTipos.Size = new System.Drawing.Size(304, 26);
			this.optTipos.TabIndex = 971;
			this.optTipos.Text = "Todos";
			this.optTipos.ValueChanged += new System.EventHandler(this.optTipos_ValueChanged);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 108);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(37, 17);
			this.label7.TabIndex = 969;
			this.label7.Text = "Buscar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optEstado
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance3;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.CheckedIndex = 1;
			this.optEstado.Enabled = false;
			this.optEstado.ItemAppearance = appearance4;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Todos";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Pendiente";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Enviado";
			valueListItem7.DataValue = 3;
			valueListItem7.DisplayText = "Liquidado";
			this.optEstado.Items.Add(valueListItem4);
			this.optEstado.Items.Add(valueListItem5);
			this.optEstado.Items.Add(valueListItem6);
			this.optEstado.Items.Add(valueListItem7);
			this.optEstado.ItemSpacingVertical = 10;
			this.optEstado.Location = new System.Drawing.Point(320, 72);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(304, 26);
			this.optEstado.TabIndex = 970;
			this.optEstado.Text = "Pendiente";
			this.optEstado.ValueChanged += new System.EventHandler(this.optEstado_ValueChanged);
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Enabled = false;
			this.txtBuscar.Location = new System.Drawing.Point(64, 106);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(216, 20);
			this.txtBuscar.TabIndex = 0;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Enabled = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(296, 105);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 1;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// uGridPagos
			// 
			this.uGridPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPagos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPagos.DataSource = this.ultraDataSource1;
			appearance5.BackColor = System.Drawing.Color.White;
			this.uGridPagos.DisplayLayout.Appearance = appearance5;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn2.CellAppearance = appearance6;
			ultraGridColumn2.Header.VisiblePosition = 6;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 100;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn3.CellAppearance = appearance7;
			ultraGridColumn3.Format = "dd/MM/yyyy";
			ultraGridColumn3.Header.VisiblePosition = 10;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 80;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.Caption = "";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 20;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance8;
			ultraGridColumn6.Header.Caption = "Tipo De Cobro";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance9;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 100;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Width = 160;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance10;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.VisiblePosition = 36;
			ultraGridColumn13.Width = 150;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance11;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 100;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance12;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 60;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance13;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 60;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn17.CellAppearance = appearance14;
			ultraGridColumn17.Format = "dd/MM/yyyy";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 90;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance15;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Valor Cuota";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 100;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance16;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 100;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance17;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 100;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn21.CellAppearance = appearance18;
			ultraGridColumn21.Format = "dd/MM/yyyy";
			ultraGridColumn21.Header.Caption = "Cobrado";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 90;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance19;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "Ab Interes";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 100;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance20;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "Ab Capital";
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 100;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance21;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "Ab Mora";
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 100;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance22;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.Caption = "Ab Dif";
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 100;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance23;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 100;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance24;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Width = 100;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn28.CellAppearance = appearance25;
			ultraGridColumn28.Header.Caption = "Estado Cuota";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Width = 120;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn29.CellAppearance = appearance26;
			ultraGridColumn29.Header.Caption = "Dias Mora";
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Width = 90;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn30.Header.Caption = "Rango Morosidad";
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Width = 160;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn31.CellAppearance = appearance27;
			ultraGridColumn31.Header.Caption = "Dias Mora Pago";
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn32.CellAppearance = appearance28;
			ultraGridColumn32.Format = "dd/MM/yyyy";
			ultraGridColumn32.Header.Caption = "Fecha Envio";
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Width = 90;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn34.CellAppearance = appearance29;
			ultraGridColumn34.Format = "dd/MM/yyyy";
			ultraGridColumn34.Header.Caption = "Fecha Sistema";
			ultraGridColumn34.Header.VisiblePosition = 35;
			ultraGridColumn34.Width = 120;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Header.VisiblePosition = 33;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn36.Header.VisiblePosition = 34;
			ultraGridColumn37.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance30.BackColorAlpha = Infragistics.Win.Alpha.Opaque;
			appearance30.Image = ((object)(resources.GetObject("appearance30.Image")));
			appearance30.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn37.CellButtonAppearance = appearance30;
			ultraGridColumn37.Header.Caption = "Ver";
			ultraGridColumn37.Header.VisiblePosition = 1;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn37.Width = 26;
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
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance31;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.uGridPagos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridPagos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPagos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPagos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance32.BackColor = System.Drawing.Color.Transparent;
			this.uGridPagos.DisplayLayout.Override.CardAreaAppearance = appearance32;
			appearance33.ForeColor = System.Drawing.Color.Black;
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPagos.DisplayLayout.Override.CellAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance34.FontData.BoldAsString = "True";
			appearance34.FontData.Name = "Arial";
			appearance34.FontData.SizeInPoints = 8.5F;
			appearance34.ForeColor = System.Drawing.Color.White;
			appearance34.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPagos.DisplayLayout.Override.HeaderAppearance = appearance34;
			this.uGridPagos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagos.DisplayLayout.Override.RowSelectorAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagos.DisplayLayout.Override.SelectedRowAppearance = appearance36;
			this.uGridPagos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPagos.Location = new System.Drawing.Point(8, 152);
			this.uGridPagos.Name = "uGridPagos";
			this.uGridPagos.Size = new System.Drawing.Size(1348, 272);
			this.uGridPagos.TabIndex = 966;
			this.uGridPagos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPagos_CellChange);
			this.uGridPagos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPagos_AfterCellUpdate);
			this.uGridPagos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridPagos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn3.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn4.DefaultValue = 0;
			ultraDataColumn5.DataType = typeof(bool);
			ultraDataColumn5.DefaultValue = false;
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn7.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn9.DefaultValue = 0;
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn15.DefaultValue = 0;
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn16.DefaultValue = 0;
			ultraDataColumn17.DataType = typeof(System.DateTime);
			ultraDataColumn17.DefaultValue = new System.DateTime(((long)(0)));
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
			ultraDataColumn21.DataType = typeof(System.DateTime);
			ultraDataColumn21.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn22.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn23.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn24.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn25.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn26.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn27.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn29.DefaultValue = 0;
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn31.DefaultValue = 0;
			ultraDataColumn32.DataType = typeof(System.DateTime);
			ultraDataColumn32.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn34.DataType = typeof(System.DateTime);
			ultraDataColumn34.DefaultValue = new System.DateTime(((long)(0)));
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
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1364, 8);
			this.groupBox1.TabIndex = 972;
			this.groupBox1.TabStop = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(784, 44);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(35, 17);
			this.label4.TabIndex = 977;
			this.label4.Text = "Desde";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(920, 44);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(33, 17);
			this.label5.TabIndex = 976;
			this.label5.Text = "Hasta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFMora
			// 
			this.txtFMora.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFMora.Enabled = false;
			this.txtFMora.Location = new System.Drawing.Point(968, 41);
			this.txtFMora.MaxValue = 105;
			this.txtFMora.MinValue = 0;
			this.txtFMora.Name = "txtFMora";
			this.txtFMora.PromptChar = ' ';
			this.txtFMora.Size = new System.Drawing.Size(72, 22);
			this.txtFMora.TabIndex = 3;
			// 
			// txtIMora
			// 
			this.txtIMora.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIMora.Enabled = false;
			this.txtIMora.Location = new System.Drawing.Point(832, 41);
			this.txtIMora.MaxValue = 105;
			this.txtIMora.MinValue = 0;
			this.txtIMora.Name = "txtIMora";
			this.txtIMora.PromptChar = ' ';
			this.txtIMora.Size = new System.Drawing.Size(72, 22);
			this.txtIMora.TabIndex = 2;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(320, 12);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 973;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// optSeleccionados
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.optSeleccionados.Appearance = appearance37;
			this.optSeleccionados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optSeleccionados.CheckedIndex = 2;
			this.optSeleccionados.Enabled = false;
			this.optSeleccionados.ItemAppearance = appearance38;
			this.optSeleccionados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem8.DataValue = 0;
			valueListItem8.DisplayText = "Todos";
			valueListItem9.DataValue = 1;
			valueListItem9.DisplayText = "No Seleccionados";
			valueListItem10.DataValue = 2;
			valueListItem10.DisplayText = "Seleccionados";
			this.optSeleccionados.Items.Add(valueListItem8);
			this.optSeleccionados.Items.Add(valueListItem9);
			this.optSeleccionados.Items.Add(valueListItem10);
			this.optSeleccionados.ItemSpacingVertical = 10;
			this.optSeleccionados.Location = new System.Drawing.Point(8, 72);
			this.optSeleccionados.Name = "optSeleccionados";
			this.optSeleccionados.Size = new System.Drawing.Size(304, 26);
			this.optSeleccionados.TabIndex = 978;
			this.optSeleccionados.Text = "Seleccionados";
			this.optSeleccionados.ValueChanged += new System.EventHandler(this.optSeleccionados_ValueChanged);
			// 
			// btnContabilizar
			// 
			this.btnContabilizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnContabilizar.CausesValidation = false;
			this.btnContabilizar.Enabled = false;
			this.btnContabilizar.Image = ((System.Drawing.Image)(resources.GetObject("btnContabilizar.Image")));
			this.btnContabilizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnContabilizar.Location = new System.Drawing.Point(1196, 104);
			this.btnContabilizar.Name = "btnContabilizar";
			this.btnContabilizar.Size = new System.Drawing.Size(160, 24);
			this.btnContabilizar.TabIndex = 748;
			this.btnContabilizar.Text = "Contabilizar Liquidación";
			this.btnContabilizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnContabilizar.Click += new System.EventHandler(this.btnContabilizar_Click);
			// 
			// btnActualizaPagos
			// 
			this.btnActualizaPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizaPagos.Enabled = false;
			this.btnActualizaPagos.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizaPagos.Image")));
			this.btnActualizaPagos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizaPagos.Location = new System.Drawing.Point(1196, 40);
			this.btnActualizaPagos.Name = "btnActualizaPagos";
			this.btnActualizaPagos.Size = new System.Drawing.Size(160, 24);
			this.btnActualizaPagos.TabIndex = 1019;
			this.btnActualizaPagos.Text = "Actualizar Lista De Pagos";
			this.btnActualizaPagos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizaPagos.Click += new System.EventHandler(this.btnActualizaPagos_Click);
			// 
			// btnImportarCartera
			// 
			this.btnImportarCartera.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImportarCartera.Image = ((System.Drawing.Image)(resources.GetObject("btnImportarCartera.Image")));
			this.btnImportarCartera.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportarCartera.Location = new System.Drawing.Point(1196, 8);
			this.btnImportarCartera.Name = "btnImportarCartera";
			this.btnImportarCartera.Size = new System.Drawing.Size(160, 24);
			this.btnImportarCartera.TabIndex = 1018;
			this.btnImportarCartera.Text = "Crear Liquidación";
			this.btnImportarCartera.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportarCartera.Click += new System.EventHandler(this.btnImportarCartera_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 12);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(34, 17);
			this.label3.TabIndex = 1020;
			this.label3.Text = "Fecha";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance39;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 10);
			this.dtFecha.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.TabIndex = 1021;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
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
			// optEstado_Cuota
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado_Cuota.Appearance = appearance40;
			this.optEstado_Cuota.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado_Cuota.CheckedIndex = 2;
			this.optEstado_Cuota.Enabled = false;
			this.optEstado_Cuota.ItemAppearance = appearance41;
			this.optEstado_Cuota.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem11.DataValue = 0;
			valueListItem11.DisplayText = "Todos";
			valueListItem12.DataValue = 1;
			valueListItem12.DisplayText = "Abonados";
			valueListItem13.DataValue = 2;
			valueListItem13.DisplayText = "Cancelados";
			this.optEstado_Cuota.Items.Add(valueListItem11);
			this.optEstado_Cuota.Items.Add(valueListItem12);
			this.optEstado_Cuota.Items.Add(valueListItem13);
			this.optEstado_Cuota.ItemSpacingVertical = 10;
			this.optEstado_Cuota.Location = new System.Drawing.Point(320, 39);
			this.optEstado_Cuota.Name = "optEstado_Cuota";
			this.optEstado_Cuota.Size = new System.Drawing.Size(304, 26);
			this.optEstado_Cuota.TabIndex = 1022;
			this.optEstado_Cuota.Text = "Cancelados";
			this.optEstado_Cuota.ValueChanged += new System.EventHandler(this.optEstado_Cuota_ValueChanged);
			// 
			// btnGenerarPagos
			// 
			this.btnGenerarPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerarPagos.Enabled = false;
			this.btnGenerarPagos.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerarPagos.Image")));
			this.btnGenerarPagos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerarPagos.Location = new System.Drawing.Point(1196, 73);
			this.btnGenerarPagos.Name = "btnGenerarPagos";
			this.btnGenerarPagos.Size = new System.Drawing.Size(160, 24);
			this.btnGenerarPagos.TabIndex = 1023;
			this.btnGenerarPagos.Text = "Generar Archivo De Pagos";
			this.btnGenerarPagos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerarPagos.Click += new System.EventHandler(this.btnGenerarPagos_Click);
			// 
			// btnRatio
			// 
			this.btnRatio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnRatio.CausesValidation = false;
			this.btnRatio.Enabled = false;
			this.btnRatio.Image = ((System.Drawing.Image)(resources.GetObject("btnRatio.Image")));
			this.btnRatio.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRatio.Location = new System.Drawing.Point(1060, 8);
			this.btnRatio.Name = "btnRatio";
			this.btnRatio.Size = new System.Drawing.Size(128, 24);
			this.btnRatio.TabIndex = 1024;
			this.btnRatio.Text = "Ver Ratio";
			this.btnRatio.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRatio.Click += new System.EventHandler(this.btnRatio_Click);
			// 
			// btnProyeccionPagos
			// 
			this.btnProyeccionPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnProyeccionPagos.CausesValidation = false;
			this.btnProyeccionPagos.Enabled = false;
			this.btnProyeccionPagos.Image = ((System.Drawing.Image)(resources.GetObject("btnProyeccionPagos.Image")));
			this.btnProyeccionPagos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnProyeccionPagos.Location = new System.Drawing.Point(1060, 40);
			this.btnProyeccionPagos.Name = "btnProyeccionPagos";
			this.btnProyeccionPagos.Size = new System.Drawing.Size(128, 24);
			this.btnProyeccionPagos.TabIndex = 1025;
			this.btnProyeccionPagos.Text = "Proyección De Pagos";
			this.btnProyeccionPagos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnProyeccionPagos.Click += new System.EventHandler(this.btnProyeccionPagos_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1060, 73);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(128, 24);
			this.btnExportar.TabIndex = 1026;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnSeleccionar
			// 
			this.btnSeleccionar.CausesValidation = false;
			this.btnSeleccionar.Image = ((System.Drawing.Image)(resources.GetObject("btnSeleccionar.Image")));
			this.btnSeleccionar.Location = new System.Drawing.Point(24, 147);
			this.btnSeleccionar.Name = "btnSeleccionar";
			this.btnSeleccionar.Size = new System.Drawing.Size(25, 25);
			this.btnSeleccionar.TabIndex = 1027;
			this.btnSeleccionar.Click += new System.EventHandler(this.btnSeleccionar_Click);
			// 
			// btnRescates
			// 
			this.btnRescates.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnRescates.CausesValidation = false;
			this.btnRescates.Enabled = false;
			this.btnRescates.Image = ((System.Drawing.Image)(resources.GetObject("btnRescates.Image")));
			this.btnRescates.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRescates.Location = new System.Drawing.Point(1060, 104);
			this.btnRescates.Name = "btnRescates";
			this.btnRescates.Size = new System.Drawing.Size(128, 24);
			this.btnRescates.TabIndex = 1028;
			this.btnRescates.Text = "Generar Rescates";
			this.btnRescates.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRescates.Click += new System.EventHandler(this.btnRescates_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(784, 77);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 17);
			this.label1.TabIndex = 1032;
			this.label1.Text = "Desde";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(920, 77);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 17);
			this.label2.TabIndex = 1031;
			this.label2.Text = "Hasta";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(664, 44);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(97, 17);
			this.label6.TabIndex = 1033;
			this.label6.Text = "Dias Mora Al Corte";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(664, 77);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(95, 17);
			this.label8.TabIndex = 1034;
			this.label8.Text = "Dias Mora Al Pago";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIMoraP
			// 
			this.txtIMoraP.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtIMoraP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtIMoraP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIMoraP.Enabled = false;
			this.txtIMoraP.Location = new System.Drawing.Point(832, 75);
			this.txtIMoraP.Name = "txtIMoraP";
			this.txtIMoraP.Size = new System.Drawing.Size(72, 20);
			this.txtIMoraP.TabIndex = 1035;
			this.txtIMoraP.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtIMoraP_KeyPress);
			// 
			// txtFMoraP
			// 
			this.txtFMoraP.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFMoraP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtFMoraP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFMoraP.Enabled = false;
			this.txtFMoraP.Location = new System.Drawing.Point(968, 75);
			this.txtFMoraP.Name = "txtFMoraP";
			this.txtFMoraP.Size = new System.Drawing.Size(72, 20);
			this.txtFMoraP.TabIndex = 1036;
			this.txtFMoraP.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFMoraP_KeyPress);
			// 
			// frmCre_PagosCarteraPichincha
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1364, 430);
			this.Controls.Add(this.txtFMoraP);
			this.Controls.Add(this.txtIMoraP);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.btnRescates);
			this.Controls.Add(this.btnSeleccionar);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnProyeccionPagos);
			this.Controls.Add(this.btnRatio);
			this.Controls.Add(this.btnGenerarPagos);
			this.Controls.Add(this.optEstado_Cuota);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btnActualizaPagos);
			this.Controls.Add(this.btnImportarCartera);
			this.Controls.Add(this.optSeleccionados);
			this.Controls.Add(this.txtFMora);
			this.Controls.Add(this.txtIMora);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.optTipos);
			this.Controls.Add(this.optEstado);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.uGridPagos);
			this.Controls.Add(this.btnContabilizar);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmCre_PagosCarteraPichincha";
			this.Text = "Cobranza Cartera Pichincha";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_PagosCarteraPichincha_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_PagosCarteraPichincha_Load);
			((System.ComponentModel.ISupportInitialize)(this.optTipos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFMora)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIMora)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optSeleccionados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado_Cuota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIMoraP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFMoraP)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCre_PagosCarteraPichincha_Load(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

			int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Cre_Estado_Envio_Pagos_Pichincha '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

			if (iEstado == 0) this.btnImportarCartera.Enabled = true;
			if (iEstado == 1) 
			{
				this.btnActualizaPagos.Enabled = true;
				this.btnGenerarPagos.Enabled = true;
				this.btnRescates.Enabled = true;
			}
			if (iEstado == 2) this.btnContabilizar.Enabled = true;
			
			if (iEstado > 0)
			{
				this.optEstado.Enabled = true;
				this.optEstado_Cuota.Enabled = true;
				this.optSeleccionados.Enabled = true;
				this.optTipos.Enabled = true;

				this.txtBuscar.Enabled = true;
				this.txtIMora.Enabled = true;
				this.txtIMoraP.Enabled = true;
				this.txtFMora.Enabled = true;
				this.txtFMoraP.Enabled = true;

				this.btnVer.Enabled = true;
				this.btnExportar.Enabled = true;
				this.btnRatio.Enabled = true;
				this.btnProyeccionPagos.Enabled = true;
				this.btnRescates.Enabled = true;
			}
			
			this.txtIMora.MinValue = -0;
			this.txtIMora.MaxValue = 105;
			this.txtFMora.MinValue = -0;
			this.txtFMora.MaxValue = 105;

			this.txtIMora.Value = 1;
			this.txtFMora.Value = 105;
			
//			this.txtIMoraP.MinValue = -540;
//			this.txtIMoraP.MaxValue = 105;
//			this.txtFMoraP.MinValue = -540;
//			this.txtFMoraP.MaxValue = 105;

			this.txtIMoraP.Value = "-540";
			this.txtFMoraP.Value = "105";
		}

		private void optEstado_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Consulta()
		{
			string sSQL = string.Format("Exec Cre_Pagos_Cartera_Pichincha '{0}', '{1}', {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtBuscar.Text.ToString(), 
				(int)this.txtIMora.Value, (int)this.txtFMora.Value, 
				(int)this.optTipos.Value, (int)this.optSeleccionados.Value, (int)this.optEstado.Value, (int)this.optEstado_Cuota.Value,
				Convert.ToInt32(this.txtIMoraP.Value.ToString()), Convert.ToInt32(this.txtFMoraP.Value.ToString()));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridPagos);

			this.lblContador.Text = this.uGridPagos.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnImportarCartera_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			int iContRegistros = 0;
		
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			
			iContRegistros = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select COUNT(*) From Cre_Cartera_Vigente_Pichincha Where CONVERT(DATE, Fecha) = '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));		

			if (iContRegistros > 0)
			{
				MessageBox.Show("Existen datos importados a la fecha actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, "Exec Cre_ValidaPagosPichinchaPorLiquidar") > 0)
			{
				MessageBox.Show("Existen pagos por liquidar de dias anteriores a Hoy ('01-11-2018')\n\nLiquide los pagos antes de generar un nuevo envio", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}			
			#endregion Validacion
			
			int iProceso = 0;
			
			try
			{				
				#region Archivo Cartera Vigente
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;
					
					#region Variables
					int iContVigente = 0;
					string	OPERACION	="";
					string	IDENTIFICACION	 ="";
					string	CLIENTE	="";
					decimal	MONTO_FINANCIADO	= 0;
					decimal	SALDO_CAPITAL_VENCIDO	= 0;
					decimal	SALDO_CAPITAL_X_VENCER =  0;	
					decimal	SALDO_CAPITAL_NODEV = 0;	
					decimal	INTERES_X_PAGAR	= 0;
					decimal	MORA_X_PAGAR = 0;
					decimal	OTROS_GASTOS = 0;
					int	DIAS_MORA	= 0;
					decimal	NUMERO_CUOTAS_VENCIDAS	= 0;
					decimal	V_CUOTAS_VENCIDAS	= 0;
					int	CUOTAS_CANCELADAS	= 0;
					#endregion Variables

					#region Sube informacion
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [DATOS$]", "DATOS").Tables[0].Rows)
					{	
						if (row["OPERACION"] != System.DBNull.Value)  OPERACION = row["OPERACION"].ToString();
						if (row["IDENTIFICACION"] != System.DBNull.Value) IDENTIFICACION = row["IDENTIFICACION"].ToString();
						if (row["CLIENTE"] != System.DBNull.Value) CLIENTE = row["CLIENTE"].ToString();
						if (row["MONTO_FINANCIADO"] != System.DBNull.Value) MONTO_FINANCIADO = decimal.Parse( row["MONTO_FINANCIADO"].ToString()); 
						if (row["SALDO_CAPITAL_VENCIDO"] != System.DBNull.Value) SALDO_CAPITAL_VENCIDO = decimal.Parse( row["SALDO_CAPITAL_VENCIDO"] .ToString());
						if (row["SALDO_CAPITAL_X_VENCER"] != System.DBNull.Value) SALDO_CAPITAL_X_VENCER = decimal.Parse( row["SALDO_CAPITAL_X_VENCER"].ToString());
						if (row["SALDO_CAPITAL_NODEV"] != System.DBNull.Value) SALDO_CAPITAL_NODEV = decimal.Parse( row["SALDO_CAPITAL_NODEV"] .ToString());
						if (row["INTERES_X_PAGAR"] != System.DBNull.Value) INTERES_X_PAGAR = decimal.Parse( row["INTERES_X_PAGAR"].ToString()); 
						if (row["MORA_X_PAGAR"] != System.DBNull.Value) MORA_X_PAGAR = decimal.Parse( row["MORA_X_PAGAR"].ToString()); 
						if (row["OTROS_GASTOS"] != System.DBNull.Value) OTROS_GASTOS = decimal.Parse( row["OTROS_GASTOS"].ToString()); 
						if (row["DIAS_MORA"] != System.DBNull.Value) DIAS_MORA = int.Parse( row["DIAS_MORA"].ToString()); 
						if (row["NUMERO_CUOTAS_VENCIDAS"] != System.DBNull.Value) NUMERO_CUOTAS_VENCIDAS = decimal.Parse( row["NUMERO_CUOTAS_VENCIDAS"].ToString());
						if (row["V_CUOTAS_VENCIDAS"] != System.DBNull.Value) V_CUOTAS_VENCIDAS = decimal.Parse( row["V_CUOTAS_VENCIDAS"].ToString());
						if (row["CUOTAS_CANCELADAS"] != System.DBNull.Value) CUOTAS_CANCELADAS = int.Parse( row["CUOTAS_CANCELADAS"].ToString());

						string sSQLAct = string.Format("Exec Importa_Cartera_Vigente_Pichincha '{0}', '{1}', '{2}', {3}, {4}, {5},{6},{7}, {8}, {9}, {10}, {11}, {12}, {13}", 
							OPERACION, IDENTIFICACION, CLIENTE, MONTO_FINANCIADO, SALDO_CAPITAL_VENCIDO, SALDO_CAPITAL_X_VENCER, SALDO_CAPITAL_NODEV, INTERES_X_PAGAR,
							MORA_X_PAGAR, OTROS_GASTOS, DIAS_MORA, NUMERO_CUOTAS_VENCIDAS, V_CUOTAS_VENCIDAS, CUOTAS_CANCELADAS);						
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
						
						iContVigente++;			
					}
					#endregion Sube informacion

					iProceso = 1;

					#region Archivo Cartera Vencida
					if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
					{
						this.Cursor = Cursors.WaitCursor;
					
						#region Variables
						int iContVencida = 0;
						string CEDULA = "";
						string FECHA_VCTO_CUOTA = "";
						decimal	CAPITAL_AMORTIZACION = 0;
						decimal	INTERES = 0;	
						decimal	MORA = 0;	
						decimal	GASTO_COVID	= 0;
						decimal	VALOR_A_PAGAR = 0;
						string NUMERO_CUOTA = "";
						#endregion Variables

						#region Sube informacion
						foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [DATOS$]", "DATOS").Tables[0].Rows)
						{	
							if (row["CEDULA"] != System.DBNull.Value) CEDULA = row["CEDULA"].ToString();
							if (row["OPERACION"] != System.DBNull.Value)  OPERACION = row["OPERACION"].ToString();						
							if (row["CLIENTE"] != System.DBNull.Value) CLIENTE = row["CLIENTE"].ToString();
							if (row["FECHA_VCTO_CUOTA"] != System.DBNull.Value) FECHA_VCTO_CUOTA = row["FECHA_VCTO_CUOTA"].ToString();
							if (row["MONTO_FINANCIADO"] != System.DBNull.Value) MONTO_FINANCIADO = decimal.Parse( row["MONTO_FINANCIADO"].ToString()); 
							if (row["CAPITAL_AMORTIZACION"] != System.DBNull.Value) CAPITAL_AMORTIZACION = decimal.Parse( row["CAPITAL_AMORTIZACION"] .ToString());
							if (row["INTERES"] != System.DBNull.Value) INTERES = decimal.Parse( row["INTERES"].ToString());
							if (row["MORA"] != System.DBNull.Value) MORA = decimal.Parse( row["MORA"] .ToString());
							if (row["GASTO_COVID"] != System.DBNull.Value) GASTO_COVID = decimal.Parse(row["GASTO_COVID"].ToString()); 
							if (row["VALOR_A_PAGAR"] != System.DBNull.Value) VALOR_A_PAGAR = decimal.Parse( row["VALOR_A_PAGAR"].ToString()); 
							if (row["NUMERO_CUOTA"] != System.DBNull.Value) NUMERO_CUOTA = row["NUMERO_CUOTA"].ToString();
							if (row["DIAS_MORA"] != System.DBNull.Value) DIAS_MORA = int.Parse( row["DIAS_MORA"].ToString());
							if (row["NUMERO_CUOTAS_VENCIDAS"] != System.DBNull.Value) NUMERO_CUOTAS_VENCIDAS = int.Parse(row["NUMERO_CUOTAS_VENCIDAS"].ToString());
						
							string sSQLAct = string.Format("Exec Importa_Cartera_Vencida_Pichincha '{0}', '{1}', '{2}', '{3}', {4}, {5},{6},{7}, {8}, {9}, '{10}', {11}, {12}", 
								CEDULA,	OPERACION, CLIENTE, FECHA_VCTO_CUOTA, MONTO_FINANCIADO, CAPITAL_AMORTIZACION, INTERES, MORA, GASTO_COVID, VALOR_A_PAGAR, NUMERO_CUOTA, DIAS_MORA, NUMERO_CUOTAS_VENCIDAS);						
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
						
							iContVencida++;			
						}
						#endregion Sube informacion

						iProceso = 2;

						string sSQL = string.Format("Exec Genera_Lista_Pagos_Cartera_Pichincha '{0}', 1", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			
						iProceso = 3;

						MessageBox.Show("Importación finalizada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

						this.btnImportarCartera.Enabled = false;
						this.btnActualizaPagos.Enabled = true;
						this.btnGenerarPagos.Enabled = true;
						this.btnProyeccionPagos.Enabled = true;
						this.btnRatio.Enabled = true;

						Consulta();

						this.Cursor = Cursors.Default;
					}
					else
					{
						string sSQL = string.Format("Exec Cre_Reinicia_Proceso_Pagos_Pichincha {0}, '{1}'", iProceso, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);						
					}
					#endregion Archivo Cartera Vencida
					this.Cursor = Cursors.Default;
				}			
				#endregion Archivo Cartera Vigente
			}
			catch (Exception ex)
			{
				string sSQL = string.Format("Exec Cre_Reinicia_Proceso_Pagos_Pichincha {0}, '{1}'", iProceso, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Demo", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				this.Cursor = Cursors.Default;													
			}
		}

		private void optTipos_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void optSeleccionados_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void btnActualizaPagos_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec Genera_Lista_Pagos_Cartera_Pichincha '{0}', 0", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			
			MessageBox.Show("Actualizacion finalizada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void optEstado_Cuota_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void btnGenerarPagos_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.OK == MessageBox.Show("¿Esta seguro de GENERAR EL ARCHIVO DE PAGOS?", "Demo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				SaveFileDialog sArchivo = new SaveFileDialog();
			
				sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivo.FileName = string.Format("TRAPNT_{0}.txt", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
				sArchivo.OverwritePrompt = true;
				sArchivo.Title = "Genera Archivo de Pagos Banco Pichincha";
			
				#region Pagos
				if (sArchivo.ShowDialog() == DialogResult.OK)
				{
					using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
					{
						#region Genera Archivo
						int i = 0;

						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_Tramas_Pagos_Pichincha '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"))).Rows)
						{			
							string iLinea = dr["Fecha"].ToString() + "\t" + dr["Peticion"].ToString() + "\t" + dr["NumeroLote"].ToString() + "\t" + dr["Producto"].ToString() + "\t"
								+ dr["Referencia"].ToString() + "\t" + dr["AgenciaIngreso"].ToString() + "\t" + dr["Usuario"].ToString() + "\t"
								+ dr["FechaIngreso"].ToString() + "\t" + dr["HoraIngreso"].ToString() + "\t" + dr["SistemaIngreso"].ToString() + "\t"
								+ dr["Campo_11"].ToString();			
							file.WriteLine(iLinea);
			
							i++;					
						}
						#endregion Genera Archivo
			
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_Actualiza_Estado_Envio_Pagos '{0}', 2, 0",
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

						MessageBox.Show(string.Format("{0} Tramas de Pagos Generadas Correctamente", i), "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);

						this.btnImportarCartera.Enabled = false;
						this.btnActualizaPagos.Enabled = false;
						this.btnGenerarPagos.Enabled = false;
					}
				}
				#endregion Pagos
			}
		}

		private void btnContabilizar_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de contabilizar los pagos seleccionados", "Demo", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				Cursor = Cursors.WaitCursor;

				#region Crea Asiento Liquidacion
				string sNumero = "LIqPAG_" + Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd");

				int idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Cre_Liquidacion_Pagos_Pichincha '{0}', '{1}'", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), sNumero));
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_Actualiza_Estado_Envio_Pagos '{0}', {1}, 3", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), idAsiento));
								
				string sSQLCO = string.Format("Exec AsientoCuentasDeOrdenEnvioPagos '{0}'", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
				Funcion.EjecutaSQL(cdsSeteoF, sSQLCO, true);

				string strFiltro = "{Asiento.idAsiento} = " + idAsiento.ToString();
				
				Reporte Reporte = new Reporte("Asientos.rpt", strFiltro);
				Reporte.MdiParent = this.MdiParent;			
				Reporte.Show();								
				#endregion Crea Asiento Liquidacion

				Cursor = Cursors.Default;

				MessageBox.Show("Liquidación Generada Correctamente", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);					

				this.btnImportarCartera.Enabled = false;
				this.btnActualizaPagos.Enabled = false;
				this.btnGenerarPagos.Enabled = false;
				this.btnContabilizar.Enabled = false;

				this.btnVer_Click(sender, e);					
			}
		}

		private void btnRatio_Click(object sender, System.EventArgs e)
		{
			frmCre_Ratio_Pichincha RP = new frmCre_Ratio_Pichincha();
			RP.ShowDialog();
		}

		private void btnProyeccionPagos_Click(object sender, System.EventArgs e)
		{
			frmCre_Proyeccion_Pagos_Pichincha PPP = new frmCre_Proyeccion_Pagos_Pichincha((DateTime)this.dtFecha.Value);
			PPP.ShowDialog();
		}

		private void uGridPagos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridPagos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridPagos.UpdateData();
				int iSel = 1;
				if ((bool)e.Cell.Row.Cells["Sel"].Value) iSel = 2;

				string sSQL = string.Format("Exec Cre_Actualiza_Pagos_Seleccionados {0}, {1}", 
					(int)e.Cell.Row.Cells["idCre_Lista_Pagos_Cartera_Pichincha"].Value, iSel);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridPagos);
		}

		private void btnSeleccionar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPagos.Rows.All)
				dr.Cells["Sel"].Value = !(bool)dr.Cells["Sel"].Value;
		}

		private void frmCre_PagosCarteraPichincha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridPagos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				int iSel = 1;
				if ((bool)e.Cell.Row.Cells["Sel"].Value) iSel = 2;

				string sSQL = string.Format("Exec Cre_Actualiza_Pagos_Seleccionados {0}, {1}", 
					(int)e.Cell.Row.Cells["idCre_Lista_Pagos_Cartera_Pichincha"].Value, iSel);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
		}

		private void btnRescates_Click(object sender, System.EventArgs e)
		{
			using (frmCre_PagosRescates Rescates = new frmCre_PagosRescates())
			{				
				if (DialogResult.OK == Rescates.ShowDialog())
				{
					string sSQL = string.Format("Exec Cre_GeneraRescates '{0}', {1}, {2}, {3}", 
						Convert.ToDateTime(Rescates.dtpFecha.Value).ToString("yyyyMMdd"), (int)Rescates.txtDiasMoraD.Value, (int)Rescates.txtDiasMoraH.Value, (int)Rescates.optTipo.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);

					MessageBox.Show("Pagos De Rescates Generados Correctamente", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtIMoraP_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos_1(sender, e);
		}

		private void txtFMoraP_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos_1(sender, e);
		}
	}
}
