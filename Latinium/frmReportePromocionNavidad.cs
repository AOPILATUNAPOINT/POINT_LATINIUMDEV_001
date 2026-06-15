using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReportePromocionNavidad.
	/// </summary>
	public class frmReportePromocionNavidad : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkLocal;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocal;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optRegiones;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optConsulta;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGrirVendedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSupervisor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReportePromocionNavidad()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReportePromocionNavidad));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Personal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PORCENTAJE");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalEfectivo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalTC");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalCP");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalGARANTIA");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalTV65");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalTV70");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalTV75");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalPayJoy");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Personal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PORCENTAJE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalTC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalCP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalGARANTIA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalTV65");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalTV70");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalTV75");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalPayJoy");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Region");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Personal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaSemanal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaSemanal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BODEGA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ALMACEN");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcMen");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SumSemanal");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SumaDiaria");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorMetaCredito");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcMetaIntangible");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SumMeta");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SumaPuntos");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jefe");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BODEGA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ALMACEN");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcMen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SumSemanal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SumaDiaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorMetaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcMetaIntangible");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SumMeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SumaPuntos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jefe");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcMeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrcCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcIntangible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diario");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semanal");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PuntosMeta");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PuntosMetaCredito");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PuntosMetaIntangible");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSupervisor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcMeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrcCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcIntangible");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semanal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PuntosMeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PuntosMetaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PuntosMetaIntangible");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			this.btnExcel = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkLocal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridLocal = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.optRegiones = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optConsulta = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.uGrirVendedor = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSupervisor = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optRegiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optConsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGrirVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSupervisor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			this.SuspendLayout();
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1180, 6);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(67, 24);
			this.btnExcel.TabIndex = 1012;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(744, 10);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 1011;
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 350;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 350;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance2;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.Caption = "% Meta";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 100;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance3;
			ultraGridColumn4.Header.Caption = "Efectivo";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance4;
			ultraGridColumn5.Header.Caption = "Tarjetas ";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance5;
			ultraGridColumn6.Header.Caption = "Credi Point";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance6;
			ultraGridColumn7.Header.Caption = "Garantia";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance7;
			ultraGridColumn8.Header.Caption = "TV 65";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance8;
			ultraGridColumn9.Header.Caption = "TV 70";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance9;
			ultraGridColumn10.Header.Caption = "TV 75";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance10;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 68;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance11;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Puntaje ";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 68;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance12;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "PayJoy";
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Hidden = true;
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
																										 ultraGridColumn13});
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance15;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.uGridInformacion.Location = new System.Drawing.Point(232, 48);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1044, 328);
			this.uGridInformacion.TabIndex = 1010;
			this.uGridInformacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInformacion_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn13});
			// 
			// chkLocal
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkLocal.Appearance = appearance18;
			this.chkLocal.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkLocal.Checked = true;
			this.chkLocal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkLocal.Location = new System.Drawing.Point(192, 80);
			this.chkLocal.Name = "chkLocal";
			this.chkLocal.Size = new System.Drawing.Size(13, 16);
			this.chkLocal.TabIndex = 1009;
			this.chkLocal.CheckedChanged += new System.EventHandler(this.chkLocal_CheckedChanged);
			// 
			// uGridLocal
			// 
			this.uGridLocal.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocal.DataSource = this.ultraDataSource6;
			appearance19.BackColor = System.Drawing.Color.White;
			this.uGridLocal.DisplayLayout.Appearance = appearance19;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.Caption = "Local";
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 150;
			ultraGridColumn16.Header.Caption = "...";
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Width = 28;
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn17.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			ultraGridBand2.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand2.SummaryFooterCaption = "TOTALES";
			this.uGridLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridLocal.DisplayLayout.GroupByBox.Hidden = true;
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocal.DisplayLayout.Override.ActiveRowAppearance = appearance20;
			this.uGridLocal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLocal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLocal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance21.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocal.DisplayLayout.Override.CardAreaAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance22.FontData.BoldAsString = "True";
			appearance22.FontData.Name = "Arial";
			appearance22.FontData.SizeInPoints = 8F;
			appearance22.ForeColor = System.Drawing.Color.White;
			appearance22.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocal.DisplayLayout.Override.HeaderAppearance = appearance22;
			this.uGridLocal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance23.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance23.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocal.DisplayLayout.Override.RowAlternateAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocal.DisplayLayout.Override.RowSelectorAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocal.DisplayLayout.Override.SelectedRowAppearance = appearance25;
			this.uGridLocal.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridLocal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocal.Location = new System.Drawing.Point(16, 80);
			this.uGridLocal.Name = "uGridLocal";
			this.uGridLocal.Size = new System.Drawing.Size(208, 296);
			this.uGridLocal.TabIndex = 1008;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(bool);
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-88, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1468, 8);
			this.groupBox1.TabIndex = 1007;
			this.groupBox1.TabStop = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(184, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 1003;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(16, 10);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(37, 16);
			this.label5.TabIndex = 1002;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optRegiones
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.optRegiones.Appearance = appearance26;
			this.optRegiones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optRegiones.CheckedIndex = 0;
			this.optRegiones.ItemAppearance = appearance27;
			this.optRegiones.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Sierra";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Costa";
			this.optRegiones.Items.Add(valueListItem1);
			this.optRegiones.Items.Add(valueListItem2);
			this.optRegiones.Items.Add(valueListItem3);
			this.optRegiones.ItemSpacingVertical = 10;
			this.optRegiones.Location = new System.Drawing.Point(24, 48);
			this.optRegiones.Name = "optRegiones";
			this.optRegiones.Size = new System.Drawing.Size(192, 24);
			this.optRegiones.TabIndex = 1006;
			this.optRegiones.Text = "Todos";
			this.optRegiones.Visible = false;
			this.optRegiones.ValueChanged += new System.EventHandler(this.optRegiones_ValueChanged);
			// 
			// optConsulta
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.optConsulta.Appearance = appearance28;
			this.optConsulta.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optConsulta.CheckedIndex = 1;
			this.optConsulta.Enabled = false;
			this.optConsulta.ItemAppearance = appearance29;
			this.optConsulta.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "Vendedor";
			valueListItem5.DataValue = 2;
			valueListItem5.DisplayText = "Jefe De Local";
			valueListItem6.DataValue = 3;
			valueListItem6.DisplayText = "Supervisor";
			this.optConsulta.Items.Add(valueListItem4);
			this.optConsulta.Items.Add(valueListItem5);
			this.optConsulta.Items.Add(valueListItem6);
			this.optConsulta.ItemSpacingVertical = 10;
			this.optConsulta.Location = new System.Drawing.Point(352, 6);
			this.optConsulta.Name = "optConsulta";
			this.optConsulta.Size = new System.Drawing.Size(280, 24);
			this.optConsulta.TabIndex = 1005;
			this.optConsulta.Text = "Jefe De Local";
			this.optConsulta.ValueChanged += new System.EventHandler(this.optConsulta_ValueChanged);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(648, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 1004;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtHasta
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance30;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(232, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1001;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance31;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1000;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn20.DataType = typeof(System.DateTime);
			ultraDataColumn21.DataType = typeof(System.DateTime);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
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
			// uGrirVendedor
			// 
			this.uGrirVendedor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGrirVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGrirVendedor.DataSource = this.ultraDataSource3;
			appearance32.BackColor = System.Drawing.Color.White;
			this.uGrirVendedor.DisplayLayout.Appearance = appearance32;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.Caption = "Almacen";
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Width = 280;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance33;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "% Meta";
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Width = 280;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn21.CellAppearance = appearance34;
			ultraGridColumn21.Header.Caption = "Semanal";
			ultraGridColumn21.Header.VisiblePosition = 4;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 80;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn22.CellAppearance = appearance35;
			ultraGridColumn22.Header.Caption = "Diario";
			ultraGridColumn22.Header.VisiblePosition = 5;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 80;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn23.CellAppearance = appearance36;
			ultraGridColumn23.Header.Caption = "Credito";
			ultraGridColumn23.Header.VisiblePosition = 6;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 80;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellAppearance = appearance37;
			ultraGridColumn24.Header.Caption = "Intangible";
			ultraGridColumn24.Header.VisiblePosition = 7;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 80;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellAppearance = appearance38;
			ultraGridColumn25.Header.Caption = "Meta";
			ultraGridColumn25.Header.VisiblePosition = 8;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 80;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn26.CellAppearance = appearance39;
			ultraGridColumn26.Header.Caption = "SubTotal";
			ultraGridColumn26.Header.VisiblePosition = 9;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 80;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn27.CellAppearance = appearance40;
			ultraGridColumn27.Header.VisiblePosition = 10;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 80;
			ultraGridColumn28.Header.Caption = "Grupo";
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 280;
			ultraGridBand3.Columns.AddRange(new object[] {
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
			this.uGrirVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGrirVendedor.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGrirVendedor.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGrirVendedor.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance41.BackColor = System.Drawing.Color.Transparent;
			this.uGrirVendedor.DisplayLayout.Override.CardAreaAppearance = appearance41;
			appearance42.ForeColor = System.Drawing.Color.Black;
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGrirVendedor.DisplayLayout.Override.CellAppearance = appearance42;
			appearance43.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance43.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance43.FontData.BoldAsString = "True";
			appearance43.FontData.Name = "Arial";
			appearance43.FontData.SizeInPoints = 8F;
			appearance43.ForeColor = System.Drawing.Color.White;
			appearance43.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGrirVendedor.DisplayLayout.Override.HeaderAppearance = appearance43;
			this.uGrirVendedor.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGrirVendedor.DisplayLayout.Override.RowSelectorAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGrirVendedor.DisplayLayout.Override.SelectedRowAppearance = appearance45;
			this.uGrirVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.uGrirVendedor.Location = new System.Drawing.Point(232, 48);
			this.uGrirVendedor.Name = "uGrirVendedor";
			this.uGrirVendedor.Size = new System.Drawing.Size(1040, 328);
			this.uGrirVendedor.TabIndex = 1013;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37});
			// 
			// uGridSupervisor
			// 
			this.uGridSupervisor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridSupervisor.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSupervisor.DataSource = this.ultraDataSource4;
			appearance46.BackColor = System.Drawing.Color.White;
			this.uGridSupervisor.DisplayLayout.Appearance = appearance46;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.Caption = "Supervisor";
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 250;
			ultraGridColumn31.Header.Caption = "%Meta";
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn32.Header.Caption = "%Credito";
			ultraGridColumn32.Header.VisiblePosition = 3;
			ultraGridColumn33.Header.Caption = "%Intangible";
			ultraGridColumn33.Header.VisiblePosition = 4;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn34.CellAppearance = appearance47;
			ultraGridColumn34.Header.VisiblePosition = 5;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn35.CellAppearance = appearance48;
			ultraGridColumn35.Header.VisiblePosition = 6;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn36.CellAppearance = appearance49;
			ultraGridColumn36.Header.Caption = "Pnt. Meta";
			ultraGridColumn36.Header.VisiblePosition = 7;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn37.CellAppearance = appearance50;
			ultraGridColumn37.Header.Caption = "Pnt. Meta Credito";
			ultraGridColumn37.Header.VisiblePosition = 8;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn38.CellAppearance = appearance51;
			ultraGridColumn38.Header.Caption = "Pnt. Meta Intangible";
			ultraGridColumn38.Header.VisiblePosition = 9;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn39.CellAppearance = appearance52;
			ultraGridColumn39.Header.VisiblePosition = 10;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn40.CellAppearance = appearance53;
			ultraGridColumn40.Header.VisiblePosition = 11;
			ultraGridColumn40.Width = 80;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40});
			this.uGridSupervisor.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.uGridSupervisor.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridSupervisor.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSupervisor.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance54.BackColor = System.Drawing.Color.Transparent;
			this.uGridSupervisor.DisplayLayout.Override.CardAreaAppearance = appearance54;
			appearance55.ForeColor = System.Drawing.Color.Black;
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSupervisor.DisplayLayout.Override.CellAppearance = appearance55;
			appearance56.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance56.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance56.FontData.BoldAsString = "True";
			appearance56.FontData.Name = "Arial";
			appearance56.FontData.SizeInPoints = 8F;
			appearance56.ForeColor = System.Drawing.Color.White;
			appearance56.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSupervisor.DisplayLayout.Override.HeaderAppearance = appearance56;
			this.uGridSupervisor.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSupervisor.DisplayLayout.Override.RowSelectorAppearance = appearance57;
			appearance58.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance58.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSupervisor.DisplayLayout.Override.SelectedRowAppearance = appearance58;
			this.uGridSupervisor.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.uGridSupervisor.Location = new System.Drawing.Point(232, 48);
			this.uGridSupervisor.Name = "uGridSupervisor";
			this.uGridSupervisor.Size = new System.Drawing.Size(1008, 328);
			this.uGridSupervisor.TabIndex = 1014;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49});
			// 
			// frmReportePromocionNavidad
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1284, 390);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.chkLocal);
			this.Controls.Add(this.uGridLocal);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.optRegiones);
			this.Controls.Add(this.optConsulta);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGrirVendedor);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.uGridSupervisor);
			this.KeyPreview = true;
			this.Name = "frmReportePromocionNavidad";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Avance Presupuesto Promocion Navidad";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReportePromocionNavidad_KeyDown);
			this.Load += new System.EventHandler(this.frmReportePromocionNavidad_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optRegiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optConsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGrirVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSupervisor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}


		private void frmReportePromocionNavidad_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200529");

			if (!Funcion.Permiso("1313", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar al Reporte ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExcel.Enabled =  true;

			DateTime dtFecha = DateTime.Parse("01/04/2021");

			this.dtDesde.CalendarInfo.MinDate = dtFecha;
			this.dtHasta.CalendarInfo.MinDate = dtFecha;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.dtDesde.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
			this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);

			string sSQLb = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQLb);		

			if (iBodega == 0) 			
				this.optRegiones.Visible = true;
      
			string sSQL = string.Format("Exec ValidaLocalesRangoMorosidadActualizado {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegiones.Value);			
			this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
		}

		private void optRegiones_ValueChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				dr.Cells["Sel"].Value = false;
						
			if ((int)this.optRegiones.Value == 0)
			{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegiones.Value);		
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				{
					if ( (int) dr.Cells["Region"].Value == 1)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
			}

			if ((int)this.optRegiones.Value == 1)
			{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegiones.Value);		
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				{
					if ( (int) dr.Cells["Region"].Value == 1)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
			}

			if ((int)this.optRegiones.Value == 2)
			{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegiones.Value);		
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				{
					if ( (int) dr.Cells["Region"].Value == 2)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void chkLocal_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
			{
				dr.Cells["Sel"].Value = this.chkLocal.Checked;
			}
		}

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

			#region Valida Local

			string sLocal= "";
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocal.Length == 0) sLocal = dr.Cells["Bodega"].Value.ToString();
					else if (sLocal.Length > 0) sLocal = sLocal + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}

			if (sLocal.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Local ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			#endregion Valida Local

			this.lblContador.Text = "";
			string sSQLPa = string.Format("Exec [ReportePromocionDiaDeLaMAdre] '{0}', '{1}', '{2}', {3}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocal, (int) this.optConsulta.Value);
			if ((int) this.optConsulta.Value == 1)
			{
				FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridInformacion);
				this.uGridInformacion.Visible = true;
				this.uGrirVendedor.Visible = false;
				this.uGridSupervisor.Visible = false;
			}

			if ((int) this.optConsulta.Value == 2)
			{
				FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGrirVendedor);
				this.uGridInformacion.Visible = false;
				this.uGrirVendedor.Visible = true;
				this.uGridSupervisor.Visible = false;
			}

			if ((int) this.optConsulta.Value == 3)
			{
				if (!miAcceso.BExportar)  return;
				FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridSupervisor);
				this.uGridInformacion.Visible = false;
				this.uGrirVendedor.Visible = false;
				this.uGridSupervisor.Visible = true;
			}

			this.lblContador.Text = this.uGrirVendedor.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value != System.DBNull.Value)
			{
				this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/07/2017");
				this.dtHasta.CalendarInfo.MaxDate = DateTime.Parse("31/12/2030");

				if (Convert.ToDateTime(this.dtDesde.Value).Month == DateTime.Today.Month) 
				{
					this.dtHasta.Value = DateTime.Today;
					this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;
				}
				else
				{
					this.dtHasta.Value = DateTime.DaysInMonth(Convert.ToDateTime(this.dtDesde.Value).Year, Convert.ToDateTime(this.dtDesde.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Year.ToString();
					this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
				}

				this.dtHasta.CalendarInfo.MinDate = Convert.ToDateTime(this.dtDesde.Value);						
			}
			else this.dtHasta.Value = System.DBNull.Value;
		}

		private void frmReportePromocionNavidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void optConsulta_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click (sender, e);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
				
			if( (int) this.optConsulta.Value == 1)
			{
				FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
			}
			if ((int) this.optConsulta.Value == 2)
			{
				FuncionesProcedimientos.ExportaExcel(this.uGrirVendedor);
			}
			if ((int) this.optConsulta.Value == 3)
			{
				FuncionesProcedimientos.ExportaExcel(this.uGridSupervisor);
			}
		}

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{		
		}
	}
}
/* 
 
*/