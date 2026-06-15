using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de OrdenMovRep.
	/// </summary>
	public class OrdenMovRep : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private C1.Data.C1DataSet cdsOrdenMov;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsAnalisis;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private Infragistics.Win.Misc.UltraButton btPantalla;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public OrdenMovRep()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FactEmitida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FactCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Guia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumOrdenMov");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaViaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Destino");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Chofer");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Placa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Equipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VCobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VPagado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gasolinera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descontar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cabezal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Combustible");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorGalon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValeTermoking");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Confirmación");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalVales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValesNoDesc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PagoGastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("GastosVarios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Peajes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mantenimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Reposicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RefReposicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalGastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HoraSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HoraLlegada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HorasViaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HLlegada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HorasTermo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("KmCargado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("KmVacio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("KmTransp");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantClase");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Clase");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Utilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Termoking");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("StCombCab");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Combustible1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorGalon1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("stCombCab1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Policia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalReposicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HorometroEN");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("KmFurgon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Otro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CarroE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CompraValor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaValor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InicioViaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FinDescargue");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FactEmitida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FactCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Guia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumOrdenMov");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaViaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Destino");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Chofer");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Placa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Equipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VCobrado");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VPagado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gasolinera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cabezal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Combustible");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorGalon");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValeTermoking");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmación");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalVales");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValesNoDesc");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoGastos");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comida");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GastosVarios");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peajes");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mantenimiento");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reposicion");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefReposicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalGastos");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraLlegada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HorasViaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HLlegada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HorasTermo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("KmCargado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("KmVacio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("KmTransp");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Utilidad");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Termoking");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("StCombCab");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Combustible1");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorGalon1");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("stCombCab1");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Policia");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalReposicion");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HorometroEN");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("KmFurgon");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Otro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CarroE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CompraValor");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaValor");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InicioViaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FinDescargue");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "VCobrado", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "VCobrado", 10, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "VPagado", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "VPagado", 11, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.cdsOrdenMov = new C1.Data.C1DataSet();
			this.btPantalla = new Infragistics.Win.Misc.UltraButton();
			this.udsAnalisis = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsOrdenMov)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAnalisis)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbDesde
			// 
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(112, 24);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 24;
			this.cmbDesde.Size = new System.Drawing.Size(112, 21);
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 0;
			this.cmbDesde.Value = new System.DateTime(2006, 1, 11, 0, 0, 0, 0);
			// 
			// cmbHasta
			// 
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(112, 56);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 24;
			this.cmbHasta.Size = new System.Drawing.Size(112, 21);
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 1;
			this.cmbHasta.Value = new System.DateTime(2006, 1, 11, 0, 0, 0, 0);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(48, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Desde";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(48, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Hasta";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(248, 24);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btExcel
			// 
			this.btExcel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btExcel.Location = new System.Drawing.Point(424, 24);
			this.btExcel.Name = "btExcel";
			this.btExcel.TabIndex = 5;
			this.btExcel.Text = "&Excel";
			this.btExcel.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cdsOrdenMov
			// 
			this.cdsOrdenMov.BindingContextCtrl = this;
			this.cdsOrdenMov.DataLibrary = "LibFacturacion";
			this.cdsOrdenMov.DataLibraryUrl = "";
			this.cdsOrdenMov.DataSetDef = "dsOrdenMov";
			this.cdsOrdenMov.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsOrdenMov.Name = "cdsOrdenMov";
			this.cdsOrdenMov.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsOrdenMov.SchemaDef = null;
			// 
			// btPantalla
			// 
			this.btPantalla.Location = new System.Drawing.Point(336, 24);
			this.btPantalla.Name = "btPantalla";
			this.btPantalla.TabIndex = 8;
			this.btPantalla.Text = "Pantalla";
			this.btPantalla.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// udsAnalisis
			// 
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(System.Double);
			ultraDataColumn12.DataType = typeof(System.Double);
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn16.DataType = typeof(System.Double);
			ultraDataColumn17.DataType = typeof(System.Double);
			ultraDataColumn18.DataType = typeof(System.Double);
			ultraDataColumn19.DataType = typeof(bool);
			ultraDataColumn20.DataType = typeof(System.Double);
			ultraDataColumn21.DataType = typeof(System.Double);
			ultraDataColumn22.DataType = typeof(System.Double);
			ultraDataColumn23.DataType = typeof(System.Double);
			ultraDataColumn24.DataType = typeof(System.Double);
			ultraDataColumn25.DataType = typeof(System.Double);
			ultraDataColumn26.DataType = typeof(System.Double);
			ultraDataColumn27.DataType = typeof(System.Double);
			ultraDataColumn28.DataType = typeof(System.Double);
			ultraDataColumn29.DataType = typeof(System.Double);
			ultraDataColumn30.DataType = typeof(System.Double);
			ultraDataColumn31.DataType = typeof(System.Double);
			ultraDataColumn32.DataType = typeof(System.Double);
			ultraDataColumn33.DataType = typeof(System.DateTime);
			ultraDataColumn34.DataType = typeof(System.DateTime);
			ultraDataColumn35.DataType = typeof(System.DateTime);
			ultraDataColumn36.DataType = typeof(System.Double);
			ultraDataColumn37.DataType = typeof(System.Double);
			ultraDataColumn38.DataType = typeof(System.Double);
			ultraDataColumn39.DataType = typeof(System.Double);
			ultraDataColumn40.DataType = typeof(System.Double);
			ultraDataColumn41.DataType = typeof(System.Double);
			ultraDataColumn42.DataType = typeof(System.Double);
			ultraDataColumn45.DataType = typeof(System.Double);
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn47.DefaultValue = 0;
			ultraDataColumn48.DataType = typeof(System.Double);
			ultraDataColumn49.DataType = typeof(System.Double);
			ultraDataColumn50.DataType = typeof(System.Double);
			ultraDataColumn51.DataType = typeof(System.Double);
			ultraDataColumn52.DataType = typeof(System.Double);
			ultraDataColumn53.DataType = typeof(System.Double);
			ultraDataColumn54.DataType = typeof(System.Double);
			ultraDataColumn55.DataType = typeof(System.Double);
			ultraDataColumn56.DataType = typeof(System.Double);
			ultraDataColumn58.DataType = typeof(System.Double);
			ultraDataColumn59.DataType = typeof(System.Double);
			ultraDataColumn60.DataType = typeof(System.DateTime);
			ultraDataColumn61.DataType = typeof(System.DateTime);
			ultraDataColumn62.DataType = typeof(System.DateTime);
			this.udsAnalisis.Band.Columns.AddRange(new object[] {
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
																														ultraDataColumn36,
																														ultraDataColumn37,
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
																														ultraDataColumn49,
																														ultraDataColumn50,
																														ultraDataColumn51,
																														ultraDataColumn52,
																														ultraDataColumn53,
																														ultraDataColumn54,
																														ultraDataColumn55,
																														ultraDataColumn56,
																														ultraDataColumn57,
																														ultraDataColumn58,
																														ultraDataColumn59,
																														ultraDataColumn60,
																														ultraDataColumn61,
																														ultraDataColumn62});
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.udsAnalisis;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 80;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 87;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn5.Format = "dd/MMM/yyyy";
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn9.Header.VisiblePosition = 10;
			ultraGridColumn10.Header.VisiblePosition = 11;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance2;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 12;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance3;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 13;
			ultraGridColumn13.Header.VisiblePosition = 14;
			ultraGridColumn14.Header.VisiblePosition = 15;
			ultraGridColumn15.Header.VisiblePosition = 16;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance4;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 17;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance5;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 18;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance6;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 20;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 25;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance7;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 26;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance8;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 27;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance9;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 28;
			ultraGridColumn22.Hidden = true;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance10;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 29;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance11;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 30;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance12;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 31;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance13;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 32;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance14;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 33;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance15;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 35;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance16;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 36;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance17;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 37;
			ultraGridColumn31.Header.VisiblePosition = 38;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance18;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.VisiblePosition = 39;
			ultraGridColumn33.Header.VisiblePosition = 42;
			ultraGridColumn34.Header.VisiblePosition = 43;
			ultraGridColumn35.Header.Caption = "Total Horas Viaje";
			ultraGridColumn35.Header.VisiblePosition = 44;
			ultraGridColumn36.Header.VisiblePosition = 46;
			ultraGridColumn37.Header.VisiblePosition = 47;
			ultraGridColumn38.Header.VisiblePosition = 48;
			ultraGridColumn39.Header.Caption = "Kilometros";
			ultraGridColumn39.Header.VisiblePosition = 49;
			ultraGridColumn40.Header.Caption = "Horometro TL";
			ultraGridColumn40.Header.VisiblePosition = 50;
			ultraGridColumn41.Header.Caption = "Kilos Transp";
			ultraGridColumn41.Header.VisiblePosition = 54;
			ultraGridColumn42.Header.VisiblePosition = 55;
			ultraGridColumn43.Header.VisiblePosition = 56;
			ultraGridColumn44.Header.VisiblePosition = 57;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance19;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.VisiblePosition = 59;
			ultraGridColumn46.Header.Caption = "Vale Termoking";
			ultraGridColumn46.Header.VisiblePosition = 21;
			ultraGridColumn47.Header.VisiblePosition = 58;
			ultraGridColumn47.Hidden = true;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance20;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.Caption = "SubT Comb Cab";
			ultraGridColumn48.Header.VisiblePosition = 19;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance21;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.VisiblePosition = 22;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance22;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.Caption = "ValorGalon Termoking";
			ultraGridColumn50.Header.VisiblePosition = 23;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance23;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.Caption = "SubT Comb Termoking";
			ultraGridColumn51.Header.VisiblePosition = 24;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance24;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.VisiblePosition = 34;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance25;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.VisiblePosition = 40;
			ultraGridColumn54.Header.Caption = "Horometro EN";
			ultraGridColumn54.Header.VisiblePosition = 51;
			ultraGridColumn55.Header.VisiblePosition = 52;
			ultraGridColumn56.Header.VisiblePosition = 53;
			ultraGridColumn57.Header.Caption = "Equipo";
			ultraGridColumn57.Header.VisiblePosition = 3;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance26;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.VisiblePosition = 60;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance27;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.VisiblePosition = 61;
			ultraGridColumn60.Header.VisiblePosition = 4;
			ultraGridColumn61.Header.VisiblePosition = 41;
			ultraGridColumn62.Header.VisiblePosition = 45;
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
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
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
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62});
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance28.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance28;
			this.ultraGrid2.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid2.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance29.FontData.BoldAsString = "True";
			appearance29.FontData.Name = "Arial";
			appearance29.FontData.SizeInPoints = 10F;
			appearance29.ForeColor = System.Drawing.Color.White;
			appearance29.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance29;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance31;
			this.ultraGrid2.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid2.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid2.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid2.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid2.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid2.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid2.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(16, 88);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(744, 280);
			this.ultraGrid2.TabIndex = 9;
			// 
			// OrdenMovRep
			// 
			this.AcceptButton = this.btPantalla;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btExcel;
			this.ClientSize = new System.Drawing.Size(776, 390);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.btPantalla);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Name = "OrdenMovRep";
			this.Text = "Reporte de Orden de Movilizacion";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.OrdenMovRep_Closing);
			this.Load += new System.EventHandler(this.OrdenMovRep_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsOrdenMov)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAnalisis)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(ultraGrid2, "OrdenMov.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "OrdenMov.xls");
		}

		private void OrdenMovRep_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Orden Mov. Rep'";
			Funcion.EjecutaSQL(cdsOrdenMov, stAudita, true);
			cmbDesde.Value = DateTime.Today;
			cmbHasta.Value = DateTime.Today;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;

			string stTitulo = "desde el " +
				dtDesde.ToString("dd/MMM/yyyy") + " hasta el " + 
				dtHasta.ToString("dd/MMM/yyyy");

			dtHasta = dtHasta.AddDays(1);
			string stReporte = "OrdenMovilizacion.Rpt";
			string stFiltro = "{OrdenMovilizacion.Fecha} >= #" +
				dtDesde.ToString("MM/dd/yyyy") + 
				"# And {OrdenMovilizacion.Fecha} < #" +
				dtHasta.ToString("MM/dd/yyyy") + "#";

			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.Show();
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFechaIni = (DateTime) cmbDesde.Value;
			DateTime dtFechaFin = (DateTime) cmbHasta.Value;
			int IdArticulo = 0;
//			string stSelect = "Select Numero From Compra Inner Join DetCompra "
//				+ " On Compra.idCompra = DetCompra.idCompra Where DetCompra.idArticulo = "
//				+ IdArticulo.ToString();
			string stSelect = "SELECT IsNull(Articulo.CodAlterno, ''), IsNull(Articulo.Codigo, ''), "
				+ "OrdenMovilizacion.Fecha, IsNull(Articulo.Articulo, ''), IsNull(ArticuloGrupo.Grupo, ''), "
				+ "IsNull(ArticuloSubGrupo.SubGrupo, ''), IsNull(Carros.Placa, ''), "
				+ "IsNull(OrdenMovilizacion.Clase, ''), Articulo.Precio1, Articulo.Precio2, "
				+ "IsNull(OrdenMovilizacion.Gasolinera, ''), IsNull(OrdenMovilizacion.DescCabezal, ''), "
				+ "IsNull(OrdenMovilizacion.Cabezal, ''), IsNull(OrdenMovilizacion.Galones, ''), "
				+ "IsNull(OrdenMovilizacion.ValorGalon, 0), IsNull(OrdenMovilizacion.DescEfectivo, 0), "
				+ "IsNull(OrdenMovilizacion.Termoking, ''), IsNull(OrdenMovilizacion.Galon1, 0), "
				+ "IsNull(OrdenMovilizacion.ValorGalon1, 0), IsNull(OrdenMovilizacion.Descontar, 0), "
				+ "IsNull(OrdenMovilizacion.Efectivo, 0), IsNull(OrdenMovilizacion.RefGasto, ''), "
				+ "IsNull(OrdenMovilizacion.Sueldo, 0), OrdenMovilizacion.Comida, "
				+ "IsNull(OrdenMovilizacion.Varios, 0), OrdenMovilizacion.Peaje, "
				+ "IsNull(OrdenMovilizacion.Policia, 0), OrdenMovilizacion.Comision, "
				+ "OrdenMovilizacion.HoraInicio, OrdenMovilizacion.HoraFinal, "
				+ "OrdenMovilizacion.Horometros, OrdenMovilizacion.HorometroL, "
				+ "OrdenMovilizacion.KmCargado, OrdenMovilizacion.KmVacio, "
				+ "OrdenMovilizacion.KmTransporte, OrdenMovilizacion.Jabas, "
				+ "IsNull(OrdenMovilizacion.Clase, ''), IsNull(OrdenMovilizacion.Observacion, ''), "
				+ "Articulo.idArticulo, "
				+ "IsNull(OrdenMovilizacion.HorometroEN, 0), IsNull(OrdenMovilizacion.KmFurgon, 0), "
				+ "IsNull(OrdenMovilizacion.Otro, 0), IsNull(Carros.Equipo, ''), "
				+ "IsNull(OrdenMovilizacion.FechaVenta, GetDate()), "
				+ "IsNull(OrdenMovilizacion.InicioViaje, GetDate()), "
				+ "IsNull(OrdenMovilizacion.FinDescargue, GetDate()) "
				+ "FROM Articulo INNER JOIN OrdenMovilizacion "
				+ "ON Articulo.idArticulo = OrdenMovilizacion.idArticulo "
				+ "INNER JOIN ArticuloGrupo ON Articulo.idGrupoArticulo = "
				+ "ArticuloGrupo.idGrupoArticulo LEFT OUTER JOIN ArticuloSubGrupo "
				+ "ON Articulo.idSubGrupo = ArticuloSubGrupo.idSubGrupo LEFT OUTER JOIN "
				+ "ArticuloMarca ON Articulo.idMarca = ArticuloMarca.idMarca "
				+ "LEFT OUTER JOIN Carros ON OrdenMovilizacion.idCarro = Carros.idCarro "
				+ "Where OrdenMovilizacion.Fecha >= '" + dtFechaIni.ToString("yyyyMMdd")
				+ "' And OrdenMovilizacion.Fecha < '" 
				+ dtFechaFin.AddDays(1).ToString("yyyyMMdd") + "'";
			SqlDataReader dr = Funcion.rEscalarSQL(cdsOrdenMov, stSelect, true);
			int iNumFilas = -1;
			udsAnalisis.Rows.Clear();
			while(dr.Read())
			{
				string stCodAlterno = "";
				stCodAlterno = dr.GetString(0);
				string stCodigo = "";
				stCodigo = dr.GetString(1);
				DateTime dtFecha;
				dtFecha = dr.GetDateTime(2);
				string stArticulo = dr.GetString(3);
				string stGrupo = "";
				stGrupo = dr.GetString(4);
				string stSubGrupo = dr.GetString(5);
				string stMarca = dr.GetString(6);
				string stUbicacion = dr.GetString(7);
				double dPrecio2 = dr.GetDouble(8);
				double dPrecio1 = dr.GetDouble(9);
				string stGasolinera = dr.GetString(10);
				bool bDescCabezal = dr.GetBoolean(11);
				string stCabezal = dr.GetString(12);
				double dGalones = dr.GetDouble(13);
				double dValorGalon = dr.GetDouble(14);
				bool bDescEfectivo = dr.GetBoolean(15);
				string stTermoking = dr.GetString(16);
				double dGalon1 = dr.GetDouble(17);
				double dValorGalon1 = dr.GetDouble(18);
				bool bDescontar = dr.GetBoolean(19);
				double dEfectivo = dr.GetDouble(20);
				string stRefGasto = dr.GetString(21);
				double dSueldo = dr.GetDouble(22);
				double dComida = dr.GetDouble(23);
				double dVarios = dr.GetDouble(24);
				double dPeaje = dr.GetDouble(25);
				double dPolicia = dr.GetDouble(26);
				double dComision = dr.GetDouble(27);
				DateTime dtHoraInicio;
				if (dr.GetValue(28) == System.DBNull.Value)
					dtHoraInicio = DateTime.Today.AddYears(-100);
				else
          dtHoraInicio = dr.GetDateTime(28);
				DateTime dtHoraFinal;
				if (dr.GetValue(29) == System.DBNull.Value)
					dtHoraFinal = DateTime.Today.AddYears(-100);
				else
          dtHoraFinal = dr.GetDateTime(29);

				double dHorometroS = dr.GetDouble(30);
				double dHorometroL = dr.GetDouble(31);
				double dKmCargado = dr.GetDouble(32);
				double dKmVacio = dr.GetDouble(33);
				double dKmTransporte = dr.GetDouble(34);
				double dJabas = dr.GetDouble(35);
				string stClase = dr.GetString(36);
				string stObservacion = dr.GetString(37);
				IdArticulo = dr.GetInt32(38);
				double dHorEN = dr.GetDouble(39);
				double dKmFurgon = dr.GetDouble(40);
				double dOtro = dr.GetDouble(41);
				string stCarroE = dr.GetString(42);
				DateTime dtFechaVenta = dr.GetDateTime(43);
				DateTime dtInicioViaje = dr.GetDateTime(44);
				DateTime dtFinDescargue = dr.GetDateTime(45);

				iNumFilas ++;
				udsAnalisis.Rows.Add();
				udsAnalisis.Rows[iNumFilas]["Guia"]=stCodAlterno;
				udsAnalisis.Rows[iNumFilas]["NumOrdenMov"]=stCodigo;
				udsAnalisis.Rows[iNumFilas]["FechaViaje"]=dtFecha;
				udsAnalisis.Rows[iNumFilas]["Destino"]=stArticulo;
				udsAnalisis.Rows[iNumFilas]["Empresa"]=stGrupo;
				udsAnalisis.Rows[iNumFilas]["Chofer"]=stSubGrupo;
				udsAnalisis.Rows[iNumFilas]["Placa"]=stMarca;
				udsAnalisis.Rows[iNumFilas]["Equipo"]=stUbicacion;
				udsAnalisis.Rows[iNumFilas]["VCobrado"]=dPrecio1;
				udsAnalisis.Rows[iNumFilas]["VPagado"]=dPrecio2;
				udsAnalisis.Rows[iNumFilas]["Gasolinera"]=stGasolinera;
				udsAnalisis.Rows[iNumFilas]["Descontar"]=bDescCabezal;
				udsAnalisis.Rows[iNumFilas]["Cabezal"]=stCabezal;
				udsAnalisis.Rows[iNumFilas]["Combustible"]=dGalones;
				udsAnalisis.Rows[iNumFilas]["ValorGalon"]=dValorGalon;
				udsAnalisis.Rows[iNumFilas]["StCombCab"]=dValorGalon*dGalones;

				udsAnalisis.Rows[iNumFilas]["Termoking"]=stTermoking;
				udsAnalisis.Rows[iNumFilas]["Confirmación"]=bDescontar;
				udsAnalisis.Rows[iNumFilas]["Efectivo"]=dEfectivo;

				udsAnalisis.Rows[iNumFilas]["Combustible1"] = dGalon1;
				udsAnalisis.Rows[iNumFilas]["ValorGalon1"] = dValorGalon1;
				double dSubtComb = dValorGalon1*dGalon1;
				udsAnalisis.Rows[iNumFilas]["StCombCab1"] = dSubtComb;

				udsAnalisis.Rows[iNumFilas]["TotalVales"] = dEfectivo 
					+ dValorGalon * dGalones + dGalon1 * dValorGalon1; // 
				//udsAnalisis.Rows[iNumFilas]["ValesNoDesc"]=stArticulo;
				udsAnalisis.Rows[iNumFilas]["PagoGastos"]=0.00;
				udsAnalisis.Rows[iNumFilas]["Sueldo"] = dSueldo;
				udsAnalisis.Rows[iNumFilas]["Comida"] = dComida;
				udsAnalisis.Rows[iNumFilas]["Peajes"] = dPeaje;
				udsAnalisis.Rows[iNumFilas]["GastosVarios"]= dVarios;
				udsAnalisis.Rows[iNumFilas]["Policia"]= dPolicia;
				udsAnalisis.Rows[iNumFilas]["Comision"]= dComision;
				//udsAnalisis.Rows[iNumFilas]["Reposicion"]=;
				//				udsAnalisis.Rows[iNumFilas]["RefReposicion"]=stArticulo;
				udsAnalisis.Rows[iNumFilas]["TotalReposicion"]=dSueldo + dComida
					+ dPeaje + dVarios + dPolicia + dComision;
				udsAnalisis.Rows[iNumFilas]["HoraSalida"]=dtHoraInicio;
				udsAnalisis.Rows[iNumFilas]["HoraLlegada"]=dtHoraFinal;
				//				udsAnalisis.Rows[iNumFilas]["HorasViaje"]=stArticulo;
				udsAnalisis.Rows[iNumFilas]["HSalida"] = dHorometroS;
				udsAnalisis.Rows[iNumFilas]["HLlegada"] = dHorometroL;
				//				udsAnalisis.Rows[iNumFilas]["HorasTermo"]=stArticulo;
				udsAnalisis.Rows[iNumFilas]["KmCargado"] = dKmCargado;
				udsAnalisis.Rows[iNumFilas]["KmVacio"] = dKmVacio;
				udsAnalisis.Rows[iNumFilas]["KmTransp"] = dKmTransporte;
				udsAnalisis.Rows[iNumFilas]["Clase"] = stClase;
				udsAnalisis.Rows[iNumFilas]["Observaciones"] = stObservacion;
				udsAnalisis.Rows[iNumFilas]["IdArticulo"] = IdArticulo;
				udsAnalisis.Rows[iNumFilas]["Utilidad"] = dPrecio1 - dPrecio2
					- double.Parse(udsAnalisis.Rows[iNumFilas]["TotalVales"].ToString())
				- double.Parse(udsAnalisis.Rows[iNumFilas]["PagoGastos"].ToString())
				- dSueldo - dComida - dPeaje - dVarios - dPolicia - dComision -dKmVacio;
				udsAnalisis.Rows[iNumFilas]["HorometroEN"] = dHorEN;
				udsAnalisis.Rows[iNumFilas]["kmFurgon"] = dKmFurgon;
				udsAnalisis.Rows[iNumFilas]["Otro"] = dOtro;
				udsAnalisis.Rows[iNumFilas]["CarroE"] = stCarroE;
				udsAnalisis.Rows[iNumFilas]["FechaVenta"] = dtFechaVenta;
				udsAnalisis.Rows[iNumFilas]["InicioViaje"] = dtInicioViaje;
				udsAnalisis.Rows[iNumFilas]["FinDescargue"] = dtFinDescargue;
				//				udsAnalisis.Rows[iNumFilas]["CantClase"]=stArticulo;
			}
			dr.Close();
			cdsOrdenMov.Schema.Connections[0].Close();
			Cursor = Cursors.Hand;

			for (int i=0; i<= iNumFilas; i++)
			{
				IdArticulo = int.Parse(udsAnalisis.Rows[i]["IdArticulo"].ToString());
				stSelect = "Select Compra.Numero "
					+ "From Compra Inner Join DetCompra "
					+ " On Compra.idCompra = DetCompra.idCompra Where "
					+ "DetCompra.idArticulo = "	+ IdArticulo.ToString()
					+ " And Compra.idTipoFactura = ";
				string stSelect1 = "Select DetCompra.Cantidad * DetCompra.Precio "
					+ "From Compra Inner Join DetCompra "
					+ " On Compra.idCompra = DetCompra.idCompra Where "
					+ "DetCompra.idArticulo = "	+ IdArticulo.ToString()
					+ " And Compra.idTipoFactura = ";
				string stNumero = Funcion.sEscalarSQL(cdsOrdenMov, stSelect + "1", false);
				udsAnalisis.Rows[i]["FactEmitida"]=stNumero;
				double dValor = Funcion.dEscalarSQL(cdsOrdenMov, stSelect1 + "1", false);
				udsAnalisis.Rows[i]["VentaValor"]=dValor;

				stNumero = Funcion.sEscalarSQL(cdsOrdenMov, stSelect + "4", false);
				udsAnalisis.Rows[i]["FactCompra"] = stNumero;
				dValor = Funcion.dEscalarSQL(cdsOrdenMov, stSelect1 + "4", false);
				udsAnalisis.Rows[i]["CompraValor"]=dValor;
			}
			Cursor = Cursors.Default;
		}

		private void OrdenMovRep_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Orden Mov Rep'";
			Funcion.EjecutaSQL(cdsOrdenMov, stAudita, true);
		}
	}
}
