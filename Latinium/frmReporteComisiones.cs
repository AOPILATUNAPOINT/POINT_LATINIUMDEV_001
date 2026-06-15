using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteComisiones.
	/// </summary>
	public class frmReporteComisiones : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcReportes;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl pgEntrega;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl TbBasico;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargoCarpa;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalleCarpa;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargosGX;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalleGx;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargoCredito;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalleCredito;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridContado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargoContado;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteComisiones()
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
    private Acceso miAcceso;
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_PorcentajesComisionesMensuales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCarpa");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta_BonoMulta");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GrupoPuerteo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesdeC");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HastaC");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_PorcentajesComisionesMensuales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta_BonoMulta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("GrupoPuerteo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DesdeC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HastaC");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_PorcentajesCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorComision");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDesde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaHasta");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_PorcentajesCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDesde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaHasta");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_PorcentajesCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargoCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorComision");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_PorcentajesCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorComision");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_PorcentajesComisionGE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargosGX");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorComision");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_PorcentajesComisionGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorComision");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_PorcentajesComisionesMensualesVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Inicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fin");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaDesde");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaHasta");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesdeC");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HastaC");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointRespaldo");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_PorcentajesComisionesMensualesVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Inicio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fin");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaDesde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaHasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DesdeC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HastaC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PointRespaldo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_PorcentajesComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fijo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDesde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaHasta");
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteComisiones));
			this.TbBasico = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridDetalleCredito = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCargoCredito = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbCargoCarpa = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.uGridDetalleCarpa = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.pgEntrega = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridDetalleGx = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCargosGX = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridContado = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCargoContado = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.utcReportes = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.TbBasico.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargoCredito)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargoCarpa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleCarpa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.pgEntrega.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleGx)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargosGX)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridContado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargoContado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).BeginInit();
			this.utcReportes.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// TbBasico
			// 
			this.TbBasico.Controls.Add(this.cmbCargos);
			this.TbBasico.Controls.Add(this.uGridDetalle);
			this.TbBasico.Location = new System.Drawing.Point(-10000, -10000);
			this.TbBasico.Name = "TbBasico";
			this.TbBasico.Size = new System.Drawing.Size(708, 425);
			// 
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCargos.DisplayMember = "Cargo";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargos.Location = new System.Drawing.Point(752, 448);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(24, 8);
			this.cmbCargos.TabIndex = 763;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.Visible = false;
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource5;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance1;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "SeteoF";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 126;
			ultraGridColumn4.Header.Caption = "Cargo";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn4.Width = 148;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 59;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 55;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 55;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "% Carpa";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 55;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance6;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 55;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance7;
			ultraGridColumn10.Header.Caption = "M.Bono Multa";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 64;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance8;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 62;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance9;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 74;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance10;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Width = 58;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance11;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn14.Width = 72;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn14});
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance14;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8.5F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance15;
			appearance16.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance16.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(0, 0);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(704, 424);
			this.uGridDetalle.TabIndex = 762;
			this.uGridDetalle.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalle_AfterRowInsert);
			this.uGridDetalle.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalle_BeforeRowsDeleted);
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(bool);
			ultraDataColumn9.DataType = typeof(bool);
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn12});
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.uGridDetalleCredito);
			this.ultraTabPageControl10.Controls.Add(this.cmbCargoCredito);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(708, 425);
			// 
			// uGridDetalleCredito
			// 
			this.uGridDetalleCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalleCredito.Cursor = System.Windows.Forms.Cursors.Hand;
			this.uGridDetalleCredito.DataSource = this.ultraDataSource4;
			appearance19.BackColor = System.Drawing.Color.White;
			this.uGridDetalleCredito.DisplayLayout.Appearance = appearance19;
			this.uGridDetalleCredito.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "SeteoF";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 113;
			ultraGridColumn16.Header.Caption = "Cargo";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn16.Width = 169;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance20;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Width = 93;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance21;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Width = 79;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance22;
			ultraGridColumn19.Header.VisiblePosition = 4;
			ultraGridColumn19.Width = 62;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance23;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "Valor Fijo";
			ultraGridColumn20.Header.VisiblePosition = 5;
			ultraGridColumn20.Width = 116;
			ultraGridColumn21.Header.Caption = "Fecha Desde";
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Width = 79;
			ultraGridColumn22.Header.Caption = "Fecha Hasta";
			ultraGridColumn22.Header.VisiblePosition = 7;
			ultraGridColumn22.Width = 86;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.uGridDetalleCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleCredito.DisplayLayout.Override.ActiveRowAppearance = appearance24;
			this.uGridDetalleCredito.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalleCredito.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalleCredito.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance25.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalleCredito.DisplayLayout.Override.CardAreaAppearance = appearance25;
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleCredito.DisplayLayout.Override.CellAppearance = appearance26;
			this.uGridDetalleCredito.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance27.FontData.BoldAsString = "True";
			appearance27.FontData.Name = "Arial";
			appearance27.FontData.SizeInPoints = 8.5F;
			appearance27.ForeColor = System.Drawing.Color.White;
			appearance27.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalleCredito.DisplayLayout.Override.HeaderAppearance = appearance27;
			appearance28.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance28.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleCredito.DisplayLayout.Override.RowAlternateAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleCredito.DisplayLayout.Override.RowSelectorAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleCredito.DisplayLayout.Override.SelectedRowAppearance = appearance30;
			this.uGridDetalleCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalleCredito.Location = new System.Drawing.Point(0, 0);
			this.uGridDetalleCredito.Name = "uGridDetalleCredito";
			this.uGridDetalleCredito.Size = new System.Drawing.Size(704, 416);
			this.uGridDetalleCredito.TabIndex = 771;
			this.uGridDetalleCredito.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalleCredito_AfterRowInsert);
			this.uGridDetalleCredito.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalleCredito_BeforeRowsDeleted);
			this.uGridDetalleCredito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalleCredito_KeyDown);
			this.uGridDetalleCredito.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalleCredito_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.DateTime);
			ultraDataColumn20.DataType = typeof(System.DateTime);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// cmbCargoCredito
			// 
			this.cmbCargoCredito.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 200;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbCargoCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCargoCredito.DisplayMember = "Cargo";
			this.cmbCargoCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargoCredito.Location = new System.Drawing.Point(224, 32);
			this.cmbCargoCredito.Name = "cmbCargoCredito";
			this.cmbCargoCredito.Size = new System.Drawing.Size(200, 112);
			this.cmbCargoCredito.TabIndex = 770;
			this.cmbCargoCredito.ValueMember = "idCargo";
			this.cmbCargoCredito.Visible = false;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.cmbCargoCarpa);
			this.ultraTabPageControl2.Controls.Add(this.uGridDetalleCarpa);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(708, 425);
			// 
			// cmbCargoCarpa
			// 
			this.cmbCargoCarpa.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 200;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbCargoCarpa.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCargoCarpa.DisplayMember = "Cargo";
			this.cmbCargoCarpa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargoCarpa.Location = new System.Drawing.Point(296, 120);
			this.cmbCargoCarpa.Name = "cmbCargoCarpa";
			this.cmbCargoCarpa.Size = new System.Drawing.Size(200, 112);
			this.cmbCargoCarpa.TabIndex = 769;
			this.cmbCargoCarpa.ValueMember = "idCargo";
			this.cmbCargoCarpa.Visible = false;
			// 
			// uGridDetalleCarpa
			// 
			this.uGridDetalleCarpa.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalleCarpa.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalleCarpa.DataSource = this.ultraDataSource2;
			appearance31.BackColor = System.Drawing.Color.White;
			this.uGridDetalleCarpa.DisplayLayout.Appearance = appearance31;
			this.uGridDetalleCarpa.DisplayLayout.AutoFitColumns = true;
			ultraGridBand6.AddButtonCaption = "SeteoF";
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 117;
			ultraGridColumn28.Header.Caption = "Cargo";
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn28.Width = 289;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance32;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn29.Width = 100;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance33;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 3;
			ultraGridColumn30.Width = 94;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance34;
			ultraGridColumn31.Header.VisiblePosition = 4;
			ultraGridColumn31.Width = 118;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance35;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.Caption = "Valor";
			ultraGridColumn32.Header.VisiblePosition = 5;
			ultraGridColumn32.Width = 83;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			this.uGridDetalleCarpa.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance36.ForeColor = System.Drawing.Color.Black;
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleCarpa.DisplayLayout.Override.ActiveRowAppearance = appearance36;
			this.uGridDetalleCarpa.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalleCarpa.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalleCarpa.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance37.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalleCarpa.DisplayLayout.Override.CardAreaAppearance = appearance37;
			appearance38.ForeColor = System.Drawing.Color.Black;
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleCarpa.DisplayLayout.Override.CellAppearance = appearance38;
			this.uGridDetalleCarpa.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance39.FontData.BoldAsString = "True";
			appearance39.FontData.Name = "Arial";
			appearance39.FontData.SizeInPoints = 8.5F;
			appearance39.ForeColor = System.Drawing.Color.White;
			appearance39.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalleCarpa.DisplayLayout.Override.HeaderAppearance = appearance39;
			appearance40.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance40.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleCarpa.DisplayLayout.Override.RowAlternateAppearance = appearance40;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleCarpa.DisplayLayout.Override.RowSelectorAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleCarpa.DisplayLayout.Override.SelectedRowAppearance = appearance42;
			this.uGridDetalleCarpa.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalleCarpa.Location = new System.Drawing.Point(0, 0);
			this.uGridDetalleCarpa.Name = "uGridDetalleCarpa";
			this.uGridDetalleCarpa.Size = new System.Drawing.Size(704, 416);
			this.uGridDetalleCarpa.TabIndex = 762;
			this.uGridDetalleCarpa.AfterRowsDeleted += new System.EventHandler(this.uGridDetalleCarpa_AfterRowsDeleted);
			this.uGridDetalleCarpa.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalleCarpa_AfterRowInsert);
			this.uGridDetalleCarpa.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalleCarpa_BeforeRowsDeleted);
			this.uGridDetalleCarpa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalleCarpa_KeyDown);
			this.uGridDetalleCarpa.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalleCarpa_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// pgEntrega
			// 
			this.pgEntrega.Controls.Add(this.uGridDetalleGx);
			this.pgEntrega.Controls.Add(this.cmbCargosGX);
			this.pgEntrega.Location = new System.Drawing.Point(-10000, -10000);
			this.pgEntrega.Name = "pgEntrega";
			this.pgEntrega.Size = new System.Drawing.Size(708, 425);
			this.pgEntrega.Paint += new System.Windows.Forms.PaintEventHandler(this.pgEntrega_Paint);
			// 
			// uGridDetalleGx
			// 
			this.uGridDetalleGx.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalleGx.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalleGx.DataSource = this.ultraDataSource3;
			appearance43.BackColor = System.Drawing.Color.White;
			this.uGridDetalleGx.DisplayLayout.Appearance = appearance43;
			this.uGridDetalleGx.DisplayLayout.AutoFitColumns = true;
			ultraGridBand7.AddButtonCaption = "SeteoF";
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 281;
			ultraGridColumn34.Header.Caption = "Cargo";
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn34.Width = 343;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance44;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 2;
			ultraGridColumn35.Width = 80;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance45;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.VisiblePosition = 3;
			ultraGridColumn36.Width = 70;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance46;
			ultraGridColumn37.Header.VisiblePosition = 4;
			ultraGridColumn37.Width = 110;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance47;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.Caption = "Valor";
			ultraGridColumn38.Header.VisiblePosition = 5;
			ultraGridColumn38.Width = 81;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38});
			this.uGridDetalleGx.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance48.ForeColor = System.Drawing.Color.Black;
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleGx.DisplayLayout.Override.ActiveRowAppearance = appearance48;
			this.uGridDetalleGx.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalleGx.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalleGx.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance49.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalleGx.DisplayLayout.Override.CardAreaAppearance = appearance49;
			appearance50.ForeColor = System.Drawing.Color.Black;
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleGx.DisplayLayout.Override.CellAppearance = appearance50;
			this.uGridDetalleGx.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance51.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance51.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance51.FontData.BoldAsString = "True";
			appearance51.FontData.Name = "Arial";
			appearance51.FontData.SizeInPoints = 8.5F;
			appearance51.ForeColor = System.Drawing.Color.White;
			appearance51.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalleGx.DisplayLayout.Override.HeaderAppearance = appearance51;
			appearance52.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance52.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleGx.DisplayLayout.Override.RowAlternateAppearance = appearance52;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleGx.DisplayLayout.Override.RowSelectorAppearance = appearance53;
			appearance54.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance54.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance54.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleGx.DisplayLayout.Override.SelectedRowAppearance = appearance54;
			this.uGridDetalleGx.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalleGx.Location = new System.Drawing.Point(0, 0);
			this.uGridDetalleGx.Name = "uGridDetalleGx";
			this.uGridDetalleGx.Size = new System.Drawing.Size(704, 416);
			this.uGridDetalleGx.TabIndex = 763;
			this.uGridDetalleGx.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalleGx_AfterRowInsert);
			this.uGridDetalleGx.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalleGx_BeforeRowsDeleted);
			this.uGridDetalleGx.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalleGx_KeyDown);
			this.uGridDetalleGx.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalleGx_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// cmbCargosGX
			// 
			this.cmbCargosGX.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn39.Header.VisiblePosition = 0;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 1;
			ultraGridColumn40.Width = 200;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn39,
																										 ultraGridColumn40});
			this.cmbCargosGX.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbCargosGX.DisplayMember = "Cargo";
			this.cmbCargosGX.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargosGX.Location = new System.Drawing.Point(360, 40);
			this.cmbCargosGX.Name = "cmbCargosGX";
			this.cmbCargosGX.Size = new System.Drawing.Size(200, 112);
			this.cmbCargosGX.TabIndex = 764;
			this.cmbCargosGX.ValueMember = "idCargo";
			this.cmbCargosGX.Visible = false;
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.uGridContado);
			this.ultraTabPageControl1.Controls.Add(this.cmbCargoContado);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(2, 21);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(708, 425);
			// 
			// uGridContado
			// 
			this.uGridContado.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridContado.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridContado.DataSource = this.ultraDataSource6;
			appearance55.BackColor = System.Drawing.Color.White;
			this.uGridContado.DisplayLayout.Appearance = appearance55;
			this.uGridContado.DisplayLayout.AutoFitColumns = true;
			ultraGridBand9.AddButtonCaption = "SeteoF";
			ultraGridColumn41.Header.VisiblePosition = 0;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 156;
			ultraGridColumn42.Header.Caption = "Cargo";
			ultraGridColumn42.Header.VisiblePosition = 1;
			ultraGridColumn42.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn42.Width = 56;
			ultraGridColumn43.Header.VisiblePosition = 2;
			ultraGridColumn43.Width = 57;
			ultraGridColumn44.Header.VisiblePosition = 3;
			ultraGridColumn44.Width = 59;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance56;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.VisiblePosition = 4;
			ultraGridColumn45.Width = 60;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance57;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 5;
			ultraGridColumn46.Width = 57;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance58;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.VisiblePosition = 6;
			ultraGridColumn47.Width = 61;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance59;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.VisiblePosition = 7;
			ultraGridColumn48.Width = 58;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance60;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.VisiblePosition = 8;
			ultraGridColumn49.Width = 55;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance61;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.VisiblePosition = 9;
			ultraGridColumn50.Width = 55;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance62;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.VisiblePosition = 10;
			ultraGridColumn51.Width = 55;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance63;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.VisiblePosition = 11;
			ultraGridColumn52.Width = 55;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance64;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.VisiblePosition = 12;
			ultraGridColumn53.Width = 55;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53});
			this.uGridContado.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			appearance65.ForeColor = System.Drawing.Color.Black;
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridContado.DisplayLayout.Override.ActiveRowAppearance = appearance65;
			this.uGridContado.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridContado.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridContado.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance66.BackColor = System.Drawing.Color.Transparent;
			this.uGridContado.DisplayLayout.Override.CardAreaAppearance = appearance66;
			appearance67.ForeColor = System.Drawing.Color.Black;
			appearance67.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridContado.DisplayLayout.Override.CellAppearance = appearance67;
			this.uGridContado.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance68.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance68.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance68.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance68.FontData.BoldAsString = "True";
			appearance68.FontData.Name = "Arial";
			appearance68.FontData.SizeInPoints = 8.5F;
			appearance68.ForeColor = System.Drawing.Color.White;
			appearance68.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridContado.DisplayLayout.Override.HeaderAppearance = appearance68;
			appearance69.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance69.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance69.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridContado.DisplayLayout.Override.RowAlternateAppearance = appearance69;
			appearance70.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance70.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance70.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridContado.DisplayLayout.Override.RowSelectorAppearance = appearance70;
			appearance71.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance71.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance71.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridContado.DisplayLayout.Override.SelectedRowAppearance = appearance71;
			this.uGridContado.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridContado.Location = new System.Drawing.Point(0, 0);
			this.uGridContado.Name = "uGridContado";
			this.uGridContado.Size = new System.Drawing.Size(704, 416);
			this.uGridContado.TabIndex = 772;
			this.uGridContado.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridContado_AfterRowInsert);
			this.uGridContado.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridContado_BeforeRowsDeleted);
			this.uGridContado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridContado_KeyDown);
			this.uGridContado.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridContado_InitializeLayout);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(System.DateTime);
			ultraDataColumn36.DataType = typeof(System.DateTime);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45});
			// 
			// cmbCargoContado
			// 
			this.cmbCargoContado.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn54.Header.VisiblePosition = 0;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 1;
			ultraGridColumn55.Width = 200;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn54,
																											ultraGridColumn55});
			this.cmbCargoContado.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbCargoContado.DisplayMember = "Cargo";
			this.cmbCargoContado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargoContado.Location = new System.Drawing.Point(254, 156);
			this.cmbCargoContado.Name = "cmbCargoContado";
			this.cmbCargoContado.Size = new System.Drawing.Size(200, 112);
			this.cmbCargoContado.TabIndex = 773;
			this.cmbCargoContado.ValueMember = "idCargo";
			this.cmbCargoContado.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(bool);
			ultraDataColumn49.DataType = typeof(System.Decimal);
			ultraDataColumn50.DataType = typeof(System.Decimal);
			ultraDataColumn51.DataType = typeof(System.Decimal);
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn53.DataType = typeof(System.Decimal);
			ultraDataColumn54.DataType = typeof(System.Decimal);
			ultraDataColumn55.DataType = typeof(System.Byte);
			ultraDataColumn56.DataType = typeof(System.DateTime);
			ultraDataColumn57.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50,
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54,
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57});
			// 
			// utcReportes
			// 
			this.utcReportes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance72.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcReportes.Appearance = appearance72;
			this.utcReportes.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcReportes.Controls.Add(this.pgEntrega);
			this.utcReportes.Controls.Add(this.TbBasico);
			this.utcReportes.Controls.Add(this.ultraTabPageControl10);
			this.utcReportes.Controls.Add(this.ultraTabPageControl2);
			this.utcReportes.Controls.Add(this.ultraTabPageControl1);
			this.utcReportes.Location = new System.Drawing.Point(0, 40);
			this.utcReportes.Name = "utcReportes";
			this.utcReportes.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcReportes.Size = new System.Drawing.Size(712, 448);
			this.utcReportes.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.PropertyPage2003;
			this.utcReportes.TabButtonStyle = Infragistics.Win.UIElementButtonStyle.OfficeXPToolbarButton;
			this.utcReportes.TabIndex = 229;
			ultraTab1.TabPage = this.TbBasico;
			ultraTab1.Text = "Contado";
			ultraTab1.Visible = false;
			ultraTab2.TabPage = this.ultraTabPageControl10;
			ultraTab2.Text = "Credito";
			ultraTab3.TabPage = this.ultraTabPageControl2;
			ultraTab3.Text = "Carpa";
			ultraTab4.TabPage = this.pgEntrega;
			ultraTab4.Text = "Garantia EX";
			ultraTab5.TabPage = this.ultraTabPageControl1;
			ultraTab5.Text = "Contado";
			this.utcReportes.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																													ultraTab1,
																																													ultraTab2,
																																													ultraTab3,
																																													ultraTab4,
																																													ultraTab5});
			this.utcReportes.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			this.utcReportes.MouseDown += new System.Windows.Forms.MouseEventHandler(this.utcReportes_MouseDown);
			this.utcReportes.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.utcReportes_SelectedTabChanged);
			this.utcReportes.MouseEnter += new System.EventHandler(this.utcReportes_MouseEnter);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(708, 425);
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
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(448, 8);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 769;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(616, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 770;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(536, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(72, 24);
			this.btnExcel.TabIndex = 771;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// frmReporteComisiones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(712, 486);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.utcReportes);
			this.KeyPreview = true;
			this.Name = "frmReporteComisiones";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Reporte Comisiones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteComisiones_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteComisiones_Load);
			this.TbBasico.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargoCredito)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbCargoCarpa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleCarpa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.pgEntrega.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleGx)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargosGX)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridContado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargoContado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).EndInit();
			this.utcReportes.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
		private void Consulta()
		{
			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Com_ConsultaPorcentajesComisionesMensuales");
		}
		private void ConsultaCarpa()
		{
			this.uGridDetalleCarpa.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Com_ConsultaPorcentajesCarpa");
		}
		private void ConsultaGX()
		{
			this.uGridDetalleGx.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Com_ConsultaPorcentajesComisionesGE");
		}
		private void ConsultaCredito()
		{
			this.uGridDetalleCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Com_ConsultaPorcentajesCredito");
		}
		private void ConsultaContado()
		{
			this.uGridContado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Com_ConsultaPorcentajesComisionesMensualesVentas");
		}

		private void frmReporteComisiones_Load(object sender, System.EventArgs e)
		{
//			miAcceso = new Acceso(cdsSeteoF, "052112");
//			if (miAcceso.BExportar) this.btnExcel.Enabled = true;
			#region Cargar Combo
      this.utcReportes.Tabs[1].Visible = false;
			this.cmbCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCargo, Cargo From Com_CargosDeVentas Where Comisiona = 1 Order By Orden");
			Consulta();

			this.cmbCargoCarpa.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCargo, Cargo From Com_CargosDeVentas Where Comisiona = 1 Order By Orden");
			ConsultaCarpa();

			this.cmbCargosGX.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCargo, Cargo From Com_CargosDeVentas Where Comisiona = 1 Order By Orden");
			ConsultaGX();

			this.cmbCargoCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCargo, Cargo From Com_CargosDeVentas Where Comisiona = 1 Order By Orden");
			ConsultaCredito();

			this.cmbCargoContado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCargo, Cargo From Com_CargosDeVentas Where Comisiona = 1 Order By Orden");
			ConsultaContado();
			#endregion Cargar Combo

		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
//			FuncionesProcedimientos.FormatoGrid(e, "Desde", 2);
//			FuncionesProcedimientos.FormatoGrid(e, "Hasta", 2);
//			FuncionesProcedimientos.FormatoGrid(e, "PorcentajeVentas", 2);
//			FuncionesProcedimientos.FormatoGrid(e, "PorcentajeComision", 2);
//			FuncionesProcedimientos.FormatoGrid(e, "PorcentajeCarpa", 2);
//			FuncionesProcedimientos.FormatoGrid(e, "ValorComision", 2);
		}

		private void uGridDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalle);
		}

		private void uGridDetalle_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idCom_PorcentajesComisionesMensuales"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idCom_PorcentajesComisionesMensuales"].Value > 0)
						{
							string sSQL = string.Format("Delete From Com_PorcentajesComisionesMensuales Where idCom_PorcentajesComisionesMensuales = {0}", (int)e.Rows[i].Cells["idCom_PorcentajesComisionesMensuales"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}						
					}
				}
			}

			e.DisplayPromptMsg = false;
		}

		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCom_PorcentajesComisionesMensuales"].Value = 0;
			e.Row.Cells["Desde"].Value = 0;
			e.Row.Cells["Hasta"].Value = 0.00m;
			e.Row.Cells["Desde"].Value = 0.00m;
			e.Row.Cells["Porcentaje"].Value = 0.00m;
			e.Row.Cells["PorcentajeCarpa"].Value = 0.00m;
			e.Row.Cells["Bono"].Value = 0.00m;
			e.Row.Cells["Meta_BonoMulta"].Value = 0;
			e.Row.Cells["Diferencia"].Value = 0;
			e.Row.Cells["GrupoPuerteo"].Value = 0;
			e.Row.Cells["DesdeC"].Value = 0.00m;
			e.Row.Cells["HastaC"].Value = 0.00m;
	
		}

		private void utcReportes_MouseEnter(object sender, System.EventArgs e)
		{
		
		}

		private void utcReportes_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			
		}

		private void uGridDetalleCarpa_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Desde", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Hasta", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ValorComision", 2);
		}

		private void uGridDetalleCarpa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalleCarpa);
		}

		private void uGridDetalleCarpa_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCom_PorcentajesCarpa"].Value = 0;			
			e.Row.Cells["Desde"].Value = 0.00m;			
			e.Row.Cells["Hasta"].Value = 0.00m;
			e.Row.Cells["Porcentaje"].Value = 0.00m;			
			e.Row.Cells["ValorComision"].Value = 0.00m;
			e.Row.Cells["FechaDesde"].Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			e.Row.Cells["FechaHasta"].Value = DateTime.Today;	

		}

		private void uGridDetalleCarpa_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			
		}

		private void uGridDetalleCarpa_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idCom_PorcentajesCarpa"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idCom_PorcentajesCarpa"].Value > 0)
						{
							string sSQL = string.Format("Delete From Com_PorcentajesCarpa Where idCom_PorcentajesCarpa = {0}", (int)e.Rows[i].Cells["idCom_PorcentajesCarpa"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}						
					}
				}
			}

			e.DisplayPromptMsg = false;
		}

		private void pgEntrega_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void uGridDetalleGx_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Desde", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Hasta", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ValorComision", 2);
		}

		private void uGridDetalleGx_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalleGx);
		}

		private void uGridDetalleGx_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCom_PorcentajesComisionGE"].Value = 0;			
			e.Row.Cells["Desde"].Value = 0.00m;			
			e.Row.Cells["Hasta"].Value = 0.00m;
			e.Row.Cells["Porcentaje"].Value = 0.00m;			
			e.Row.Cells["ValorComision"].Value = 0.00m;
			e.Row.Cells["FechaDesde"].Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			e.Row.Cells["FechaHasta"].Value = DateTime.Today;	
		}

		private void uGridDetalleGx_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idCom_PorcentajesComisionGE"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idCom_PorcentajesComisionGE"].Value > 0)
						{
							string sSQL = string.Format("Delete From Com_PorcentajesComisionGE Where idCom_PorcentajesComisionGE = {0}", (int)e.Rows[i].Cells["idCom_PorcentajesComisionGE"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}						
					}
				}
			}

			e.DisplayPromptMsg = false;
		}

		private void uGridDetalleCredito_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCom_PorcentajesCredito"].Value = 0;			
			e.Row.Cells["Desde"].Value = 0.00m;			
			e.Row.Cells["Hasta"].Value = 0.00m;
			e.Row.Cells["Porcentaje"].Value = 0.00m;			
			e.Row.Cells["ValorComision"].Value = 0.00m;
			e.Row.Cells["FechaDesde"].Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			e.Row.Cells["FechaHasta"].Value = DateTime.Today;	
		}

		private void uGridDetalleCredito_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idCom_PorcentajesCredito"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idCom_PorcentajesCredito"].Value > 0)
						{
							string sSQL = string.Format("Delete From Com_PorcentajesCredito Where idCom_PorcentajesCredito = {0}", (int)e.Rows[i].Cells["idCom_PorcentajesCredito"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}						
					}
				}
			}

			e.DisplayPromptMsg = false;
		}

		private void uGridDetalleCredito_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Desde", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Hasta", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ValorComision", 2);
		}

		private void uGridDetalleCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalleCredito);
		}

		private void frmReporteComisiones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();	
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Valida Comisiones
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
			{
				#region Porcentajes Desde C
				if (Convert.ToDecimal(dr.Cells["DesdeC"].Value) > Convert.ToDecimal(dr.Cells["HastaC"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[0].Selected = true;
					this.uGridDetalle.ActiveRow = dr;
					this.uGridDetalle.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final
					
				#region Cargo
				if (dr.Cells["idCargo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[0].Selected = true;
					this.uGridDetalle.ActiveRow = dr;
					this.uGridDetalle.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idCargo"].Value == 0)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDetalle.ActiveRow = dr;
					this.uGridDetalle.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo

				#region Porcentajes Inicial Final
				if (Convert.ToDecimal(dr.Cells["Desde"].Value) > Convert.ToDecimal(dr.Cells["Hasta"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[0].Selected = true;
					this.uGridDetalle.ActiveRow = dr;
					this.uGridDetalle.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final

				#region Porcentajes repetidos
				string sSQL = string.Format("Exec Com_ValidaPorcentajes {0}, {1}, {2}, {3}", 
					(int)dr.Cells["idCom_PorcentajesComisionesMensuales"].Value, (int)dr.Cells["idCargo"].Value, Convert.ToDecimal(dr.Cells["Desde"].Value), Convert.ToDecimal(dr.Cells["Hasta"].Value));

				if (Funcion.bEscalarSQL(cdsSeteoF, sSQL, true))
				{
					MessageBox.Show("Esta ingresando valores repetidos en Porcentajes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[0].Selected = true;
					this.uGridDetalle.ActiveRow = dr;
					this.uGridDetalle.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes repetidos				
			}
				
			#endregion Valida Comisiones

			#region Valida Carpa
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalleCarpa.Rows.All)
			{
				#region fecha
				//					if ((DateTime)dr.Cells["FechaDesde"].Value > (DateTime)dr.Cells["FechaHasta"].Value)
				//					{
				//						MessageBox.Show("La fecha final no puede ser menor a la fecha inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//						this.utcReportes.Tabs[2].Selected = true;
				//						this.uGridDetalleCarpa.ActiveRow = dr;
				//						this.uGridDetalleCarpa.ActiveRow.Selected = true;
				//
				//						return;
				//					}
				#endregion fecha

				#region Cargo
				if (dr.Cells["idCargo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[2].Selected = true;
					this.uGridDetalleCarpa.ActiveRow = dr;
					this.uGridDetalleCarpa.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idCargo"].Value == 0)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDetalleCarpa.ActiveRow = dr;
					this.uGridDetalleCarpa.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo

				#region Porcentajes Inicial Final
				if (Convert.ToDecimal(dr.Cells["Desde"].Value) > Convert.ToDecimal(dr.Cells["Hasta"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[2].Selected = true;
					this.uGridDetalleCarpa.ActiveRow = dr;
					this.uGridDetalleCarpa.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final			
			}
			#endregion Valida Carpa

			#region Valida Credito

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalleCredito.Rows.All)
			{
				//					#region fecha
				//					if ((DateTime)dr.Cells["FechaDesde"].Value > (DateTime)dr.Cells["FechaHasta"].Value)
				//					{
				//						MessageBox.Show("La fecha final no puede ser menor a la fecha inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//						this.utcReportes.Tabs[1].Selected = true;
				//						this.uGridDetalleCredito.ActiveRow = dr;
				//						this.uGridDetalleCredito.ActiveRow.Selected = true;
				//
				//						return;
				//					}
				//					#endregion fecha

				#region Cargo
				if (dr.Cells["idCargo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[1].Selected = true;
					this.uGridDetalleCredito.ActiveRow = dr;
					this.uGridDetalleCredito.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idCargo"].Value == 0)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDetalleCredito.ActiveRow = dr;
					this.uGridDetalleCredito.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo

				#region Porcentajes Inicial Final
				if (Convert.ToDecimal(dr.Cells["Desde"].Value) > Convert.ToDecimal(dr.Cells["Hasta"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[1].Selected = true;
					this.uGridDetalleCredito.ActiveRow = dr;
					this.uGridDetalleCredito.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final		
			}

			#endregion Valida Credito

			#region Valida GX

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalleGx.Rows.All)
			{
				#region fecha
				//					if ((DateTime)dr.Cells["FechaDesde"].Value > (DateTime)dr.Cells["FechaHasta"].Value)
				//					{
				//						MessageBox.Show("La fecha final no puede ser menor a la fecha inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//						this.utcReportes.Tabs[3].Selected = true;
				//						this.uGridDetalleGx.ActiveRow = dr;
				//						this.uGridDetalleGx.ActiveRow.Selected = true;
				//
				//						return;
				//					}
				#endregion fecha

				#region Cargo
				if (dr.Cells["idCargo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[3].Selected = true;
					this.uGridDetalleGx.ActiveRow = dr;
					this.uGridDetalleGx.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idCargo"].Value == 0)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[3].Selected = true;
					this.uGridDetalleGx.ActiveRow = dr;
					this.uGridDetalleGx.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo

				#region Porcentajes Inicial Final
				if (Convert.ToDecimal(dr.Cells["Desde"].Value) > Convert.ToDecimal(dr.Cells["Hasta"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[3].Selected = true;
					this.uGridDetalleGx.ActiveRow = dr;
					this.uGridDetalleGx.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final

				#region Porcentajes repetidos
				string sSQL = string.Format("Exec Com_ValidaPorcentajesGE {0}, {1}, {2}, {3}", 
					(int)dr.Cells["idCom_PorcentajesComisionGE"].Value, (int)dr.Cells["idCargo"].Value, Convert.ToDecimal(dr.Cells["Desde"].Value), Convert.ToDecimal(dr.Cells["Hasta"].Value));

				if (Funcion.bEscalarSQL(cdsSeteoF, sSQL, true))
				{
					MessageBox.Show("Esta ingresando valores repetidos en Porcentajes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[3].Selected = true;
					this.uGridDetalleGx.ActiveRow = dr;
					this.uGridDetalleGx.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes repetidos				
			}

			#endregion Valida GX

			#region Valida Contado
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridContado.Rows.All)
			{
				#region Porcentajes Desde C
				if (Convert.ToDecimal(dr.Cells["DesdeC"].Value) > Convert.ToDecimal(dr.Cells["HastaC"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[4].Selected = true;
					this.uGridContado.ActiveRow = dr;
					this.uGridContado.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final
					
				#region Cargo
				if (dr.Cells["idCargo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[4].Selected = true;
					this.uGridContado.ActiveRow = dr;
					this.uGridContado.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idCargo"].Value == 0)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridContado.ActiveRow = dr;
					this.uGridContado.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo

				#region Porcentajes Inicial Final
				if (Convert.ToDecimal(dr.Cells["Desde"].Value) > Convert.ToDecimal(dr.Cells["Hasta"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[4].Selected = true;
					this.uGridContado.ActiveRow = dr;
					this.uGridContado.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final

				#region Porcentajes Ventas Inicial Final
				if (Convert.ToDecimal(dr.Cells["VentaDesde"].Value) > Convert.ToDecimal(dr.Cells["VentaHasta"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcReportes.Tabs[4].Selected = true;
					this.uGridContado.ActiveRow = dr;
					this.uGridContado.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Ventas Inicial Final

				#region Porcentajes repetidos
//				string sSQL = string.Format("Exec Com_ValidaPorcentajes {0}, {1}, {2}, {3}", 
//					(int)dr.Cells["idCom_PorcentajesComisionesMensuales"].Value, (int)dr.Cells["idCargo"].Value, Convert.ToDecimal(dr.Cells["Desde"].Value), Convert.ToDecimal(dr.Cells["Hasta"].Value));
//
//				if (Funcion.bEscalarSQL(cdsSeteoF, sSQL, true))
//				{
//					MessageBox.Show("Esta ingresando valores repetidos en Porcentajes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.utcReportes.Tabs[0].Selected = true;
//					this.uGridDetalle.ActiveRow = dr;
//					this.uGridDetalle.ActiveRow.Selected = true;
//
//					return;
//				}
			#endregion Porcentajes repetidos				
			}
				
			#endregion Valida Contado

			#region Variables
			DateTime dtDesde ;
			DateTime dtHasta;

			DateTime dtDesdeC;
			DateTime dtHastaC;

			DateTime dtDesdeGX;
			DateTime dtHastaGX;

			DateTime dtDesdeGCre;
			DateTime dtHastaGCre;
			#endregion Variables

			#region guardar
			#region guarda carpa
			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Guardar los cambios", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalleCarpa.Rows.All)
				{
					//					dtDesde = (DateTime)dr.Cells["FechaDesde"].Value;
					//					dtHasta = (DateTime)dr.Cells["FechaHasta"].Value;
					string sSQLDetalle = string.Format("Exec Com_GuardaPorcentajesCarpa {0}, {1}, {2}, {3}, {4}, {5}",
						(int)dr.Cells["idCom_PorcentajesCarpa"].Value, (int)dr.Cells["idCargo"].Value, 
						(decimal)dr.Cells["Desde"].Value, (decimal)dr.Cells["Hasta"].Value, 
						(decimal)dr.Cells["Porcentaje"].Value, (decimal)dr.Cells["ValorComision"].Value) ;
					Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
				}

				ConsultaCarpa();
				#endregion guarda carpa
			}
			#region guarda comisiones 
       
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
			{
		
				string sSQLDetalle = string.Format("Exec Com_PorcentajesComisionesMensual {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8},{9},{10},{11}",
					(int)dr.Cells["idCom_PorcentajesComisionesMensuales"].Value, 
					(int)dr.Cells["idCargo"].Value, 
					(decimal)dr.Cells["Desde"].Value, 
					(decimal)dr.Cells["Hasta"].Value, 
					(decimal)dr.Cells["Porcentaje"].Value, 
					(decimal)dr.Cells["PorcentajeCarpa"].Value, 
					(decimal)dr.Cells["Bono"].Value, 
					Convert.ToDecimal((bool)dr.Cells["Meta_BonoMulta"].Value), 
					Convert.ToDecimal((bool)dr.Cells["Diferencia"].Value), 
					Convert.ToDecimal((bool)dr.Cells["GrupoPuerteo"].Value), 
					(decimal)dr.Cells["DesdeC"].Value, 
					(decimal)dr.Cells["HastaC"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
				
			}
			Consulta();
			#endregion guarda comisiones 

			#region Guarda GX
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalleGx.Rows.All)
			{
			
				string sSQLDetalle = string.Format("Exec Com_GuardaPorcentajesComisionesGE {0}, {1}, {2}, {3}, {4}, {5}",
					(int)dr.Cells["idCom_PorcentajesComisionGE"].Value, (int)dr.Cells["idCargo"].Value, 
					(decimal)dr.Cells["Desde"].Value, (decimal)dr.Cells["Hasta"].Value, 
					(decimal)dr.Cells["Porcentaje"].Value, (decimal)dr.Cells["ValorComision"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			}
			ConsultaGX();

			
			#endregion Guarda GX

			#region Guarda Credito

			//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalleCredito.Rows.All)
			//				{
			//					dtDesdeGCre = (DateTime)dr.Cells["FechaDesde"].Value;
			//					dtHastaGCre= (DateTime)dr.Cells["FechaHasta"].Value;
			//					string sSQLDetalle = string.Format("Exec Com_GuardaPorcentajesCredito {0}, {1}, {2}, {3}, {4}, {5},'{6}','{7}'",
			//						(int)dr.Cells["idCom_PorcentajesCredito"].Value, (int)dr.Cells["idCargo"].Value, 
			//						(decimal)dr.Cells["Desde"].Value, (decimal)dr.Cells["Hasta"].Value, 
			//						(decimal)dr.Cells["Porcentaje"].Value, (decimal)dr.Cells["ValorComision"].Value,
			//						Convert.ToDateTime(dtDesdeGCre).ToString("yyyyMMdd"), Convert.ToDateTime(dtHastaGCre).ToString("yyyyMMdd"));
			//					Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			//				}
			//				ConsultaCredito();

			//	}
			#endregion Guarda Credito

			#region guarda comisionesNueva
	       
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridContado.Rows.All)
				{
					int idComision = dr.Cells["idCom_PorcentajesComisionesMensualesVentas"].Value != DBNull.Value 
						? (int)dr.Cells["idCom_PorcentajesComisionesMensualesVentas"].Value 
						: 0;

					DateTime inicioFecha = dr.Cells["Inicio"].Value != DBNull.Value ? Convert.ToDateTime(dr.Cells["Inicio"].Value): new DateTime(1900, 1, 1);
					DateTime finFecha = dr.Cells["Fin"].Value != DBNull.Value ? Convert.ToDateTime(dr.Cells["Fin"].Value): new DateTime(1900, 1, 1);

					string sSQLDetalleContado = string.Format("Exec Com_PorcentajesComisionesMensualesVentasGuardar {0}, {1}, '{2}', '{3}', {4}, {5}, {6}, {7}, {8},{9},{10},{11}",
						idComision, 
						(int)dr.Cells["idCargo"].Value, 
						inicioFecha.ToString("yyyyMMdd"),
						finFecha.ToString("yyyyMMdd"),
						(decimal)dr.Cells["Porcentaje"].Value, 
						(decimal)dr.Cells["Bono"].Value, 
						(decimal)dr.Cells["Desde"].Value,
						(decimal)dr.Cells["Hasta"].Value,
						(decimal)dr.Cells["VentaDesde"].Value, 
						(decimal)dr.Cells["VentaHasta"].Value, 
						(decimal)dr.Cells["DesdeC"].Value, 
						(decimal)dr.Cells["HastaC"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalleContado);				
				}
				ConsultaContado();
				#endregion guarda comisionesNueva

			#endregion guardar

			
			this.Close();
		}


		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			#region Exportar 
			if (this.utcReportes.Tabs[0].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridDetalle);
		  if (this.utcReportes.Tabs[1].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridDetalleCredito);
			if (this.utcReportes.Tabs[2].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridDetalleCarpa);
			if (this.utcReportes.Tabs[3].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridDetalleGx);
			if (this.utcReportes.Tabs[4].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridContado);
			#endregion Exportar 
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void uGridContado_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCom_PorcentajesComisionesMensualesVentas"].Value = 0;
			e.Row.Cells["Inicio"].Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			e.Row.Cells["Fin"].Value = DateTime.Today;	
			e.Row.Cells["Desde"].Value = 0.00m;
			e.Row.Cells["Hasta"].Value = 0.00m;
			e.Row.Cells["Porcentaje"].Value = 0.00m;
			e.Row.Cells["Bono"].Value = 0.00m;
			e.Row.Cells["Desde"].Value = 0.00m;
			e.Row.Cells["Hasta"].Value = 0.00m;
			e.Row.Cells["VentaDesde"].Value = 0.00m;
			e.Row.Cells["VentaHasta"].Value = 0.00m;
			e.Row.Cells["DesdeC"].Value = 0.00m;
			e.Row.Cells["HastaC"].Value = 0.00m;
		}

		private void uGridContado_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idCom_PorcentajesComisionesMensualesVentas"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idCom_PorcentajesComisionesMensualesVentas"].Value > 0)
						{
							string sSQL = string.Format("Delete From Com_PorcentajesComisionesMensualesVentas Where idCom_PorcentajesComisionesMensualesVentas = {0}", (int)e.Rows[i].Cells["idCom_PorcentajesComisionesMensualesVentas"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}						
					}
				}
			}

			e.DisplayPromptMsg = false;
		}

		private void uGridContado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridContado);
		}

		private void uGridContado_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Desde", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Hasta", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Bono", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Desde", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Hasta", 2);
			FuncionesProcedimientos.FormatoGrid(e, "VentaDesde", 2);
			FuncionesProcedimientos.FormatoGrid(e, "VentaHasta", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DesdeC", 2);
			FuncionesProcedimientos.FormatoGrid(e, "HastaC", 2);

		}

		private void utcReportes_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}
  }	
				
}
