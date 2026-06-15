using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteMargenUtilidadPorArticulo.
	/// </summary>
	public class frmReporteMargenUtilidadPorArticulo : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenCrediPOINT;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenTarjeta;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenEfectivo;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargen;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcUtCrediPoint;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcUtEfectivo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcUtTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcCrediPoint;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcEfectivo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocales;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		
		private Acceso miAcceso;

		public frmReporteMargenUtilidadPorArticulo()
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteMargenUtilidadPorArticulo));
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeVentasLocal");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaEfectivo");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeEfectivo");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaTarjeta");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeTarjeta");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaCrediPoint");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCrediPoint");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalUtilidad");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeUtilidadLocal");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadEfectivo");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeEfectivoUtilidad");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadTarjeta");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeTarjetaUtilidad");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadCrediPoint");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCrediPointUtilidad");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Margen");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenEfectivo");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenTarjeta");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenCrediPoint");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad_Vendida");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeVentasLocal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeUtilidadLocal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeEfectivoUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeTarjetaUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCrediPointUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Margen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad_Vendida");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubgrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subgrupo");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubgrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subgrupo");
			this.txtMargenCrediPOINT = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtMargenTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtMargenEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.txtMargen = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtPorcUtCrediPoint = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcUtEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcUtTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcCrediPoint = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblContador = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkTodos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label14 = new System.Windows.Forms.Label();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenCrediPOINT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtCrediPoint)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcCrediPoint)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			this.SuspendLayout();
			// 
			// txtMargenCrediPOINT
			// 
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenCrediPOINT.Appearance = appearance1;
			this.txtMargenCrediPOINT.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenCrediPOINT.Enabled = false;
			this.txtMargenCrediPOINT.FormatString = "#,##0.00";
			this.txtMargenCrediPOINT.Location = new System.Drawing.Point(1216, 80);
			this.txtMargenCrediPOINT.MinValue = -2147483647;
			this.txtMargenCrediPOINT.Name = "txtMargenCrediPOINT";
			this.txtMargenCrediPOINT.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenCrediPOINT.PromptChar = ' ';
			this.txtMargenCrediPOINT.Size = new System.Drawing.Size(80, 21);
			this.txtMargenCrediPOINT.TabIndex = 763;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(1216, 48);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(107, 16);
			this.label13.TabIndex = 764;
			this.label13.Text = "Margen CrediPOINT";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMargenTarjeta
			// 
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenTarjeta.Appearance = appearance2;
			this.txtMargenTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenTarjeta.Enabled = false;
			this.txtMargenTarjeta.FormatString = "#,##0.00";
			this.txtMargenTarjeta.Location = new System.Drawing.Point(1120, 80);
			this.txtMargenTarjeta.MinValue = -2147483647;
			this.txtMargenTarjeta.Name = "txtMargenTarjeta";
			this.txtMargenTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenTarjeta.PromptChar = ' ';
			this.txtMargenTarjeta.Size = new System.Drawing.Size(80, 21);
			this.txtMargenTarjeta.TabIndex = 761;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(1120, 48);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(81, 16);
			this.label12.TabIndex = 762;
			this.label12.Text = "Margen Tarjeta";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMargenEfectivo
			// 
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenEfectivo.Appearance = appearance3;
			this.txtMargenEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenEfectivo.Enabled = false;
			this.txtMargenEfectivo.FormatString = "#,##0.00";
			this.txtMargenEfectivo.Location = new System.Drawing.Point(1024, 80);
			this.txtMargenEfectivo.MinValue = -2147483647;
			this.txtMargenEfectivo.Name = "txtMargenEfectivo";
			this.txtMargenEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenEfectivo.PromptChar = ' ';
			this.txtMargenEfectivo.Size = new System.Drawing.Size(80, 21);
			this.txtMargenEfectivo.TabIndex = 759;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(1016, 48);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(86, 16);
			this.label11.TabIndex = 760;
			this.label11.Text = "Margen Efectivo";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMargen
			// 
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargen.Appearance = appearance4;
			this.txtMargen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargen.Enabled = false;
			this.txtMargen.FormatString = "#,##0.00";
			this.txtMargen.Location = new System.Drawing.Point(912, 80);
			this.txtMargen.MinValue = -2147483647;
			this.txtMargen.Name = "txtMargen";
			this.txtMargen.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargen.PromptChar = ' ';
			this.txtMargen.Size = new System.Drawing.Size(72, 21);
			this.txtMargen.TabIndex = 757;
			// 
			// label9
			// 
			this.label9.BackColor = System.Drawing.Color.DarkGray;
			this.label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.ForeColor = System.Drawing.Color.Firebrick;
			this.label9.Location = new System.Drawing.Point(704, 8);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(240, 32);
			this.label9.TabIndex = 756;
			this.label9.Text = "UTILIDAD";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label8
			// 
			this.label8.BackColor = System.Drawing.Color.DarkGray;
			this.label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.ForeColor = System.Drawing.Color.Firebrick;
			this.label8.Location = new System.Drawing.Point(400, 8);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(240, 32);
			this.label8.TabIndex = 755;
			this.label8.Text = "VENTAS";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(816, 48);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(93, 16);
			this.label6.TabIndex = 754;
			this.label6.Text = "% Ut CrediPOINT";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(744, 48);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(67, 16);
			this.label5.TabIndex = 753;
			this.label5.Text = "% Ut Tarjeta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(672, 48);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(72, 16);
			this.label4.TabIndex = 752;
			this.label4.Text = "% Ut Efectivo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(560, 48);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(79, 16);
			this.label3.TabIndex = 751;
			this.label3.Text = "% CrediPOINT";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(480, 48);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(53, 16);
			this.label2.TabIndex = 750;
			this.label2.Text = "% Tarjeta";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(400, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(58, 16);
			this.label1.TabIndex = 749;
			this.label1.Text = "% Efectivo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(216, 16);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 741;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(24, 82);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(39, 16);
			this.label7.TabIndex = 740;
			this.label7.Text = "Buscar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(24, 16);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 16);
			this.lblBodega.TabIndex = 739;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPorcUtCrediPoint
			// 
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcUtCrediPoint.Appearance = appearance5;
			this.txtPorcUtCrediPoint.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcUtCrediPoint.Enabled = false;
			this.txtPorcUtCrediPoint.FormatString = "#,##0.00";
			this.txtPorcUtCrediPoint.Location = new System.Drawing.Point(824, 80);
			this.txtPorcUtCrediPoint.MinValue = -2147483647;
			this.txtPorcUtCrediPoint.Name = "txtPorcUtCrediPoint";
			this.txtPorcUtCrediPoint.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcUtCrediPoint.PromptChar = ' ';
			this.txtPorcUtCrediPoint.Size = new System.Drawing.Size(80, 21);
			this.txtPorcUtCrediPoint.TabIndex = 748;
			// 
			// txtPorcUtEfectivo
			// 
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcUtEfectivo.Appearance = appearance6;
			this.txtPorcUtEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcUtEfectivo.Enabled = false;
			this.txtPorcUtEfectivo.FormatString = "#,##0.00";
			this.txtPorcUtEfectivo.Location = new System.Drawing.Point(664, 80);
			this.txtPorcUtEfectivo.MinValue = -2147483647;
			this.txtPorcUtEfectivo.Name = "txtPorcUtEfectivo";
			this.txtPorcUtEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcUtEfectivo.PromptChar = ' ';
			this.txtPorcUtEfectivo.Size = new System.Drawing.Size(72, 21);
			this.txtPorcUtEfectivo.TabIndex = 746;
			// 
			// txtPorcUtTarjeta
			// 
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcUtTarjeta.Appearance = appearance7;
			this.txtPorcUtTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcUtTarjeta.Enabled = false;
			this.txtPorcUtTarjeta.FormatString = "#,##0.00";
			this.txtPorcUtTarjeta.Location = new System.Drawing.Point(744, 80);
			this.txtPorcUtTarjeta.MinValue = -2147483647;
			this.txtPorcUtTarjeta.Name = "txtPorcUtTarjeta";
			this.txtPorcUtTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcUtTarjeta.PromptChar = ' ';
			this.txtPorcUtTarjeta.Size = new System.Drawing.Size(72, 21);
			this.txtPorcUtTarjeta.TabIndex = 747;
			// 
			// txtPorcCrediPoint
			// 
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcCrediPoint.Appearance = appearance8;
			this.txtPorcCrediPoint.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcCrediPoint.Enabled = false;
			this.txtPorcCrediPoint.FormatString = "#,##0.00";
			this.txtPorcCrediPoint.Location = new System.Drawing.Point(560, 80);
			this.txtPorcCrediPoint.MinValue = -2147483647;
			this.txtPorcCrediPoint.Name = "txtPorcCrediPoint";
			this.txtPorcCrediPoint.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcCrediPoint.PromptChar = ' ';
			this.txtPorcCrediPoint.Size = new System.Drawing.Size(80, 21);
			this.txtPorcCrediPoint.TabIndex = 745;
			// 
			// txtPorcEfectivo
			// 
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcEfectivo.Appearance = appearance9;
			this.txtPorcEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcEfectivo.Enabled = false;
			this.txtPorcEfectivo.FormatString = "#,##0.00";
			this.txtPorcEfectivo.Location = new System.Drawing.Point(400, 80);
			this.txtPorcEfectivo.MinValue = -2147483647;
			this.txtPorcEfectivo.Name = "txtPorcEfectivo";
			this.txtPorcEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcEfectivo.PromptChar = ' ';
			this.txtPorcEfectivo.Size = new System.Drawing.Size(72, 21);
			this.txtPorcEfectivo.TabIndex = 743;
			// 
			// txtPorcTarjeta
			// 
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcTarjeta.Appearance = appearance10;
			this.txtPorcTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcTarjeta.Enabled = false;
			this.txtPorcTarjeta.FormatString = "#,##0.00";
			this.txtPorcTarjeta.Location = new System.Drawing.Point(480, 80);
			this.txtPorcTarjeta.MinValue = -2147483647;
			this.txtPorcTarjeta.Name = "txtPorcTarjeta";
			this.txtPorcTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcTarjeta.PromptChar = ' ';
			this.txtPorcTarjeta.Size = new System.Drawing.Size(72, 21);
			this.txtPorcTarjeta.TabIndex = 744;
			// 
			// txtBusqueda
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance11;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(80, 80);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(192, 21);
			this.txtBusqueda.TabIndex = 734;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1248, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 736;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(288, 78);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 735;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtDesde
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance12;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(80, 16);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 737;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// dtHasta
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance13;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(272, 16);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 738;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(976, 16);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 742;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(928, 48);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(43, 16);
			this.label10.TabIndex = 758;
			this.label10.Text = "Margen";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource2;
			appearance14.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance14;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 151;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.CellAppearance = appearance15;
			ultraGridColumn2.Format = "#,##0";
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 60;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance16;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 5;
			ultraGridColumn3.Width = 80;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance17;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "% Ventas";
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance18;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "V Efectivo";
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 80;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance19;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "% Efectivo";
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance20;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "V Tarjeta";
			ultraGridColumn7.Header.VisiblePosition = 9;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 80;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance21;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "% Tarjeta";
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn8.Width = 80;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance22;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "V CrediPoint";
			ultraGridColumn9.Header.VisiblePosition = 11;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance23;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "% CrediPoint";
			ultraGridColumn10.Header.VisiblePosition = 12;
			ultraGridColumn10.Width = 80;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance24;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Total Utilidad";
			ultraGridColumn11.Header.VisiblePosition = 13;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance25;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "% Utilidad";
			ultraGridColumn12.Header.VisiblePosition = 15;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance26;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "U Efectivo";
			ultraGridColumn13.Header.VisiblePosition = 14;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 80;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance27;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "% U Efectivo";
			ultraGridColumn14.Header.VisiblePosition = 16;
			ultraGridColumn14.Width = 80;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance28;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "U Tarjeta";
			ultraGridColumn15.Header.VisiblePosition = 17;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 80;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance29;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "% U Tarjeta";
			ultraGridColumn16.Header.VisiblePosition = 18;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance30;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "U CrediPoint";
			ultraGridColumn17.Header.VisiblePosition = 19;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 80;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance31;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "% U CrediPoint";
			ultraGridColumn18.Header.VisiblePosition = 20;
			ultraGridColumn18.Width = 80;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance32;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 21;
			ultraGridColumn19.Width = 80;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance33;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "M Efectivo";
			ultraGridColumn20.Header.VisiblePosition = 22;
			ultraGridColumn20.Width = 80;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance34;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "M Tarjeta";
			ultraGridColumn21.Header.VisiblePosition = 23;
			ultraGridColumn21.Width = 80;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance35;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "M CrediPoint";
			ultraGridColumn22.Header.VisiblePosition = 24;
			ultraGridColumn22.Width = 80;
			ultraGridColumn23.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 195;
			ultraGridColumn24.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Width = 250;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance36;
			ultraGridColumn25.Header.Caption = "Cant. V";
			ultraGridColumn25.Header.VisiblePosition = 3;
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
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance37.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance37;
			appearance38.ForeColor = System.Drawing.Color.Black;
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance39.FontData.BoldAsString = "True";
			appearance39.FontData.Name = "Arial";
			appearance39.FontData.SizeInPoints = 8F;
			appearance39.ForeColor = System.Drawing.Color.White;
			appearance39.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance39;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance40;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance41;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(224, 128);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1136, 248);
			this.uGridInformacion.TabIndex = 765;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn25});
			// 
			// chkTodos
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkTodos.Appearance = appearance42;
			this.chkTodos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkTodos.Checked = true;
			this.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodos.Location = new System.Drawing.Point(192, 136);
			this.chkTodos.Name = "chkTodos";
			this.chkTodos.Size = new System.Drawing.Size(13, 16);
			this.chkTodos.TabIndex = 768;
			this.chkTodos.CheckedChanged += new System.EventHandler(this.chkTodos_CheckedChanged);
			// 
			// uGridLocales
			// 
			this.uGridLocales.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocales.DataSource = this.ultraDataSource1;
			appearance43.BackColor = System.Drawing.Color.White;
			this.uGridLocales.DisplayLayout.Appearance = appearance43;
			this.uGridLocales.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 62;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 151;
			ultraGridColumn28.Header.Caption = "...";
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.Width = 49;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			ultraGridBand2.Header.Caption = "";
			this.uGridLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance44.ForeColor = System.Drawing.Color.Black;
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocales.DisplayLayout.Override.ActiveRowAppearance = appearance44;
			this.uGridLocales.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridLocales.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridLocales.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance45.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocales.DisplayLayout.Override.CardAreaAppearance = appearance45;
			appearance46.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance46.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance46.FontData.BoldAsString = "True";
			appearance46.FontData.Name = "Arial";
			appearance46.FontData.SizeInPoints = 10F;
			appearance46.ForeColor = System.Drawing.Color.White;
			appearance46.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocales.DisplayLayout.Override.HeaderAppearance = appearance46;
			appearance47.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance47.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowAlternateAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowSelectorAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.SelectedRowAppearance = appearance49;
			this.uGridLocales.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocales.Location = new System.Drawing.Point(0, 128);
			this.uGridLocales.Name = "uGridLocales";
			this.uGridLocales.Size = new System.Drawing.Size(221, 248);
			this.uGridLocales.TabIndex = 767;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(bool);
			ultraDataColumn28.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1354, 8);
			this.groupBox1.TabIndex = 766;
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
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(24, 48);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(53, 16);
			this.label14.TabIndex = 769;
			this.label14.Text = "Subgrupo";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn29.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn29,
																																 ultraDataColumn30});
			// 
			// cmbSubGrupo
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance50;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataSource = this.ultraDataSource3;
			this.cmbSubGrupo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 86;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 173;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbSubGrupo.DisplayMember = "Subgrupo";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(80, 48);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(192, 21);
			this.cmbSubGrupo.TabIndex = 772;
			this.cmbSubGrupo.ValueMember = "idSubgrupo";
			// 
			// frmReporteMargenUtilidadPorArticulo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1360, 374);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.chkTodos);
			this.Controls.Add(this.uGridLocales);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtMargenCrediPOINT);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.txtMargenTarjeta);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtMargenEfectivo);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.txtMargen);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.txtPorcUtCrediPoint);
			this.Controls.Add(this.txtPorcUtEfectivo);
			this.Controls.Add(this.txtPorcUtTarjeta);
			this.Controls.Add(this.txtPorcCrediPoint);
			this.Controls.Add(this.txtPorcEfectivo);
			this.Controls.Add(this.txtPorcTarjeta);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label10);
			this.KeyPreview = true;
			this.Name = "frmReporteMargenUtilidadPorArticulo";
			this.Text = "frmReporteMargenUtilidadPorArticulo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteMargenUtilidadPorArticulo_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteMargenUtilidadPorArticulo_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtMargenCrediPOINT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtCrediPoint)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcCrediPoint)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}
			#endregion Valida Fechas
						
			string sLocales = "";			
			this.lblContador.Text = "";
						
			#region Locales
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocales.Length == 0) sLocales = dr.Cells["Bodega"].Value.ToString(); else if (sLocales.Length > 0) sLocales = sLocales + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}

			if (sLocales.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Locales

			int idSubgrupo = 0;
			if (this.cmbSubGrupo.ActiveRow != null) idSubgrupo = (int)this.cmbSubGrupo.Value;
			
			#region Consulta
			this.uGridInformacion.DataSource = ultraDataSource2;

			this.Cursor = Cursors.WaitCursor;
		
			string sSQL = string.Format("Exec ResumenMargenDeUtilidadPorArticulo '{0}', '{1}', {2}, '{3}', '{4}',{5}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 0, this.txtBusqueda.Text.ToString(), sLocales, idSubgrupo);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInformacion);
			
			this.lblContador.Text = this.uGridInformacion.Rows.Count + " Registros Encontrados";

			this.Cursor = Cursors.Default;
			#endregion Consulta

			#region Porcentajes Totales
			this.txtPorcEfectivo.Value = 0.00m;
			this.txtPorcTarjeta.Value = 0.00m;
			this.txtPorcCrediPoint.Value = 0.00m;
			this.txtPorcUtEfectivo.Value = 0.00m;
			this.txtPorcUtTarjeta.Value = 0.00m;
			this.txtPorcUtCrediPoint.Value = 0.00m;

			if (this.uGridInformacion.Rows.Count > 0)
			{
				decimal dSumaTotal = 0.00m;
				decimal dSumaEfectivo = 0.00m;
				decimal dSumaTarjeta = 0.00m;
				decimal dSumaCrediPoint = 0.00m;
				decimal dSumaTotalUtilidad = 0.00m;
				decimal dSumaUtEfectivo = 0.00m;
				decimal dSumaUtTarjeta = 0.00m;
				decimal dSumaUtCrediPoint = 0.00m;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
				{
					dSumaTotal = dSumaTotal + Convert.ToDecimal(dr.Cells["Total"].Value);
					dSumaEfectivo = dSumaEfectivo + Convert.ToDecimal(dr.Cells["VentaEfectivo"].Value);
					dSumaTarjeta = dSumaTarjeta + Convert.ToDecimal(dr.Cells["VentaTarjeta"].Value);
					dSumaCrediPoint = dSumaCrediPoint + Convert.ToDecimal(dr.Cells["VentaCrediPoint"].Value);

					dSumaTotalUtilidad = dSumaTotalUtilidad + Convert.ToDecimal(dr.Cells["TotalUtilidad"].Value);
					dSumaUtEfectivo = dSumaUtEfectivo + Convert.ToDecimal(dr.Cells["UtilidadEfectivo"].Value);
					dSumaUtTarjeta = dSumaUtTarjeta + Convert.ToDecimal(dr.Cells["UtilidadTarjeta"].Value);
					dSumaUtCrediPoint = dSumaUtCrediPoint + Convert.ToDecimal(dr.Cells["UtilidadCrediPoint"].Value);
				}			

				this.txtPorcEfectivo.Value = dSumaEfectivo/dSumaTotal * 100.00m;
				this.txtPorcTarjeta.Value = dSumaTarjeta/dSumaTotal * 100.00m;
				this.txtPorcCrediPoint.Value = dSumaCrediPoint/dSumaTotal * 100.00m;

				this.txtPorcUtEfectivo.Value = dSumaUtEfectivo/dSumaTotalUtilidad * 100.00m;
				this.txtPorcUtTarjeta.Value = dSumaUtTarjeta/dSumaTotalUtilidad * 100.00m;
				this.txtPorcUtCrediPoint.Value = dSumaUtCrediPoint/dSumaTotalUtilidad * 100.00m;

				this.txtMargen.Value = dSumaTotalUtilidad/dSumaTotal * 100.00m;

				//				this.txtMargenEfectivo.Value = dSumaEfectivo/dSumaTotal * 100.00m;
				//				this.txtMargenTarjeta.Value = dSumaTarjeta/dSumaTotal * 100.00m;
				//				this.txtMargenCrediPOINT.Value = dSumaCrediPoint/dSumaTotal * 100.00m;

				if( dSumaEfectivo > 0)
				{
					this.txtMargenEfectivo.Value = dSumaUtEfectivo/dSumaEfectivo * 100.00m;
				}
				else
				{
					this.txtMargenEfectivo.Value = 0.00m;
				}
				
				if ( dSumaTarjeta > 0)
				{
					this.txtMargenTarjeta.Value = dSumaUtTarjeta/dSumaTarjeta * 100.00m;
				}
				else
				{
					this.txtMargenTarjeta.Value = 0.00m;
				}

				if ( dSumaCrediPoint > 0)
				{
					this.txtMargenCrediPOINT.Value = dSumaUtCrediPoint/dSumaCrediPoint * 100.00m;
				}
				else
				{
				  this.txtMargenCrediPOINT.Value = 0.00m;
				}
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
				{					
					dr.Cells["PorcentajeVentasLocal"].Value = (Convert.ToDecimal(dr.Cells["Total"].Value)/dSumaTotal) * 100;

					dr.Cells["PorcentajeUtilidadLocal"].Value = (Convert.ToDecimal(dr.Cells["TotalUtilidad"].Value)/dSumaTotalUtilidad) * 100;
          
					if (Convert.ToDecimal(dr.Cells["Total"].Value)> 0.00m)
					  dr.Cells["Margen"].Value = (Convert.ToDecimal(dr.Cells["TotalUtilidad"].Value)/Convert.ToDecimal(dr.Cells["Total"].Value)) * 100;
					if (Convert.ToDecimal(dr.Cells["VentaEfectivo"].Value) > 0.00m)
						dr.Cells["MargenEfectivo"].Value = (Convert.ToDecimal(dr.Cells["UtilidadEfectivo"].Value)/Convert.ToDecimal(dr.Cells["VentaEfectivo"].Value)) * 100;//VentaEfectivo
					if (Convert.ToDecimal(dr.Cells["VentaTarjeta"].Value) > 0.00m)
						dr.Cells["MargenTarjeta"].Value = (Convert.ToDecimal(dr.Cells["UtilidadTarjeta"].Value)/Convert.ToDecimal(dr.Cells["VentaTarjeta"].Value)) * 100;//VentaTarjeta
					if (Convert.ToDecimal(dr.Cells["VentaCrediPoint"].Value) > 0.00m)
						dr.Cells["MargenCrediPoint"].Value = (Convert.ToDecimal(dr.Cells["UtilidadCrediPoint"].Value)/Convert.ToDecimal(dr.Cells["VentaCrediPoint"].Value)) * 100;//VentaCrediPoint
				}
			}
			#endregion Porcentajes Totales
		}

		private void frmReporteMargenUtilidadPorArticulo_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "071714");

			if (!Funcion.Permiso("911", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Proformas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			
				return;
			}
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;			
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;
      this.cmbSubGrupo.DataSource = FuncionesProcedimientos.dtGeneral("Select idSubgrupo, SubGrupo From ArticuloSubGrupo Order By SubGrupo");
			this.uGridLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre, Convert(Bit, 1) As Sel From Bodega Where Activo = 1 And Factura = 1 Order By Nombre");
	
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void frmReporteMargenUtilidadPorArticulo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();	
		}

		private void chkTodos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if (this.chkTodos.Checked) dr.Cells["Sel"].Value = 1;
				else dr.Cells["Sel"].Value = 0;
			}
		}
	}
}
