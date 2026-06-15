using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_PromedioVendedoresPorLocal.
	/// </summary>
	public class frmCom_PromedioVendedoresPorLocal : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridAmbato;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridAmbato2;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridCondado;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridComiteDelPueblo;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridRecreo1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridQuinche;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridLatacunga;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridInca;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSantoDomingo;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSangolqui;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridTumbaco;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSolanda;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridRiobamba;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridRecreo3;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridPrensa;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridMachachi;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridGuamani;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridConsumidorFinal;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridCarapungo;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridCalderon;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridCayambe;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSanLuis;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridQuicentroSur;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridIbarra;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridChillogallo;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Button btnBuscar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_PromedioVendedoresPorLocal()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Promedio", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Promedio", 8, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Promedio");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance133 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance134 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance135 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance136 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance137 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance138 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance139 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance140 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance141 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance142 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance143 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance144 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance145 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance146 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance147 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance148 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance149 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand19 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance150 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance151 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance152 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance153 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance154 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance155 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance156 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance157 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand20 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance158 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance159 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance160 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance161 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance162 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance163 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance164 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance165 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand21 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance166 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance167 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance168 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance169 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance170 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance171 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance172 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance173 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand22 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn192 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn193 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn194 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn195 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn196 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn197 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance174 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn198 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance175 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance176 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance177 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance178 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance179 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance180 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance181 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand23 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn199 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn200 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn201 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn202 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn203 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn204 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn205 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn206 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance182 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn207 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance183 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance184 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance185 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance186 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance187 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance188 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance189 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand24 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn208 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn209 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn210 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn211 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn212 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn213 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn214 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn215 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance190 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn216 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance191 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance192 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance193 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance194 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance195 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance196 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance197 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand25 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn217 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn218 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn219 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn220 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn221 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn222 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn223 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn224 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance198 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn225 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vendido", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vendido", 7, true);
			Infragistics.Win.Appearance appearance199 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance200 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance201 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance202 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance203 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance204 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance205 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_PromedioVendedoresPorLocal));
			this.uGridAmbato = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridAmbato2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridCondado = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridComiteDelPueblo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridRecreo1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridQuinche = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridLatacunga = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridInca = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridSantoDomingo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridSangolqui = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridTumbaco = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridSolanda = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridRiobamba = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridRecreo3 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridPrensa = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridMachachi = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridGuamani = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridConsumidorFinal = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridCarapungo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridCalderon = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridCayambe = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridSanLuis = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridQuicentroSur = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridIbarra = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridChillogallo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.btnBuscar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridAmbato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridAmbato2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCondado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridComiteDelPueblo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRecreo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridQuinche)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLatacunga)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSantoDomingo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSangolqui)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTumbaco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSolanda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRiobamba)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRecreo3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPrensa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMachachi)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGuamani)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridConsumidorFinal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCarapungo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCalderon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCayambe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSanLuis)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridQuicentroSur)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridIbarra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridChillogallo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridAmbato
			// 
			this.uGridAmbato.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridAmbato.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridAmbato.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 178;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 39;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 50;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 55;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Width = 55;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance5;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 55;
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
			ultraGridBand1.Header.Caption = "Ambato";
			ultraGridBand1.HeaderVisible = true;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance6;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance7;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridAmbato.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridAmbato.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridAmbato.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.uGridAmbato.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 7.5F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridAmbato.DisplayLayout.Override.HeaderAppearance = appearance9;
			appearance10.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance10.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAmbato.DisplayLayout.Override.RowAlternateAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAmbato.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAmbato.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.uGridAmbato.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridAmbato.Location = new System.Drawing.Point(5, 32);
			this.uGridAmbato.Name = "uGridAmbato";
			this.uGridAmbato.Size = new System.Drawing.Size(310, 120);
			this.uGridAmbato.TabIndex = 2;
			this.uGridAmbato.DoubleClick += new System.EventHandler(this.uGridAmbato_DoubleClick);
			this.uGridAmbato.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridAmbato_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
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
																																 ultraDataColumn9});
			// 
			// uGridAmbato2
			// 
			this.uGridAmbato2.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridAmbato2.DataSource = this.ultraDataSource1;
			appearance13.BackColor = System.Drawing.Color.White;
			this.uGridAmbato2.DisplayLayout.Appearance = appearance13;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 41;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 177;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 34;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 53;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Width = 39;
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn16.Header.VisiblePosition = 7;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance14;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 5;
			ultraGridColumn17.Width = 56;
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			ultraGridBand2.Header.Caption = "Ambato 2";
			ultraGridBand2.HeaderVisible = true;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance15;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings3});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridAmbato2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridAmbato2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridAmbato2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridAmbato2.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 7.5F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridAmbato2.DisplayLayout.Override.HeaderAppearance = appearance17;
			appearance18.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance18.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAmbato2.DisplayLayout.Override.RowAlternateAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAmbato2.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAmbato2.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.uGridAmbato2.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridAmbato2.Location = new System.Drawing.Point(317, 32);
			this.uGridAmbato2.Name = "uGridAmbato2";
			this.uGridAmbato2.Size = new System.Drawing.Size(310, 120);
			this.uGridAmbato2.TabIndex = 11;
			this.uGridAmbato2.DoubleClick += new System.EventHandler(this.uGridAmbato2_DoubleClick);
			// 
			// uGridCondado
			// 
			this.uGridCondado.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCondado.DataSource = this.ultraDataSource1;
			appearance21.BackColor = System.Drawing.Color.White;
			this.uGridCondado.DisplayLayout.Appearance = appearance21;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 41;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 178;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 34;
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 53;
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn23.Width = 38;
			ultraGridColumn24.Header.VisiblePosition = 6;
			ultraGridColumn25.Header.VisiblePosition = 7;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance22;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 5;
			ultraGridColumn26.Width = 56;
			ultraGridColumn27.Header.VisiblePosition = 8;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			ultraGridBand3.Header.Caption = "Condado";
			ultraGridBand3.HeaderVisible = true;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance23;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridCondado.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridCondado.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCondado.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance24.BackColor = System.Drawing.Color.Transparent;
			this.uGridCondado.DisplayLayout.Override.CardAreaAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance25.FontData.BoldAsString = "True";
			appearance25.FontData.Name = "Arial";
			appearance25.FontData.SizeInPoints = 7.5F;
			appearance25.ForeColor = System.Drawing.Color.White;
			appearance25.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCondado.DisplayLayout.Override.HeaderAppearance = appearance25;
			appearance26.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance26.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCondado.DisplayLayout.Override.RowAlternateAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCondado.DisplayLayout.Override.RowSelectorAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCondado.DisplayLayout.Override.SelectedRowAppearance = appearance28;
			this.uGridCondado.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCondado.Location = new System.Drawing.Point(317, 152);
			this.uGridCondado.Name = "uGridCondado";
			this.uGridCondado.Size = new System.Drawing.Size(310, 120);
			this.uGridCondado.TabIndex = 28;
			this.uGridCondado.DoubleClick += new System.EventHandler(this.uGridCondado_DoubleClick);
			// 
			// uGridComiteDelPueblo
			// 
			this.uGridComiteDelPueblo.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridComiteDelPueblo.DataSource = this.ultraDataSource1;
			appearance29.BackColor = System.Drawing.Color.White;
			this.uGridComiteDelPueblo.DisplayLayout.Appearance = appearance29;
			ultraGridBand4.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Width = 178;
			ultraGridColumn30.Header.VisiblePosition = 2;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 3;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 4;
			ultraGridColumn32.Width = 39;
			ultraGridColumn33.Header.VisiblePosition = 6;
			ultraGridColumn34.Header.VisiblePosition = 7;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance30;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 5;
			ultraGridColumn35.Width = 55;
			ultraGridColumn36.Header.VisiblePosition = 8;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36});
			ultraGridBand4.Header.Caption = "Comite del Pueblo";
			ultraGridBand4.HeaderVisible = true;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance31;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings5});
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridComiteDelPueblo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.uGridComiteDelPueblo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridComiteDelPueblo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance32.BackColor = System.Drawing.Color.Transparent;
			this.uGridComiteDelPueblo.DisplayLayout.Override.CardAreaAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance33.FontData.BoldAsString = "True";
			appearance33.FontData.Name = "Arial";
			appearance33.FontData.SizeInPoints = 7.5F;
			appearance33.ForeColor = System.Drawing.Color.White;
			appearance33.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridComiteDelPueblo.DisplayLayout.Override.HeaderAppearance = appearance33;
			appearance34.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance34.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridComiteDelPueblo.DisplayLayout.Override.RowAlternateAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridComiteDelPueblo.DisplayLayout.Override.RowSelectorAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridComiteDelPueblo.DisplayLayout.Override.SelectedRowAppearance = appearance36;
			this.uGridComiteDelPueblo.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridComiteDelPueblo.Location = new System.Drawing.Point(5, 152);
			this.uGridComiteDelPueblo.Name = "uGridComiteDelPueblo";
			this.uGridComiteDelPueblo.Size = new System.Drawing.Size(310, 120);
			this.uGridComiteDelPueblo.TabIndex = 27;
			this.uGridComiteDelPueblo.DoubleClick += new System.EventHandler(this.uGridComiteDelPueblo_DoubleClick);
			// 
			// uGridRecreo1
			// 
			this.uGridRecreo1.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridRecreo1.DataSource = this.ultraDataSource1;
			appearance37.BackColor = System.Drawing.Color.White;
			this.uGridRecreo1.DisplayLayout.Appearance = appearance37;
			ultraGridColumn37.Header.VisiblePosition = 0;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 41;
			ultraGridColumn38.Header.VisiblePosition = 1;
			ultraGridColumn38.Width = 178;
			ultraGridColumn39.Header.VisiblePosition = 2;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 34;
			ultraGridColumn40.Header.VisiblePosition = 3;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 53;
			ultraGridColumn41.Header.VisiblePosition = 4;
			ultraGridColumn41.Width = 39;
			ultraGridColumn42.Header.VisiblePosition = 6;
			ultraGridColumn43.Header.VisiblePosition = 7;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance38;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.VisiblePosition = 5;
			ultraGridColumn44.Width = 55;
			ultraGridColumn45.Header.VisiblePosition = 8;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45});
			ultraGridBand5.Header.Caption = "Recreo 1";
			ultraGridBand5.HeaderVisible = true;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance39;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings6});
			ultraGridBand5.SummaryFooterCaption = "";
			this.uGridRecreo1.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.uGridRecreo1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridRecreo1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance40.BackColor = System.Drawing.Color.Transparent;
			this.uGridRecreo1.DisplayLayout.Override.CardAreaAppearance = appearance40;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance41.FontData.BoldAsString = "True";
			appearance41.FontData.Name = "Arial";
			appearance41.FontData.SizeInPoints = 7.5F;
			appearance41.ForeColor = System.Drawing.Color.White;
			appearance41.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRecreo1.DisplayLayout.Override.HeaderAppearance = appearance41;
			appearance42.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance42.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRecreo1.DisplayLayout.Override.RowAlternateAppearance = appearance42;
			appearance43.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance43.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRecreo1.DisplayLayout.Override.RowSelectorAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRecreo1.DisplayLayout.Override.SelectedRowAppearance = appearance44;
			this.uGridRecreo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRecreo1.Location = new System.Drawing.Point(944, 392);
			this.uGridRecreo1.Name = "uGridRecreo1";
			this.uGridRecreo1.Size = new System.Drawing.Size(310, 120);
			this.uGridRecreo1.TabIndex = 32;
			this.uGridRecreo1.DoubleClick += new System.EventHandler(this.uGridRecreo1_DoubleClick);
			// 
			// uGridQuinche
			// 
			this.uGridQuinche.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridQuinche.DataSource = this.ultraDataSource1;
			appearance45.BackColor = System.Drawing.Color.White;
			this.uGridQuinche.DisplayLayout.Appearance = appearance45;
			ultraGridBand6.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn46.Header.VisiblePosition = 0;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Width = 178;
			ultraGridColumn48.Header.VisiblePosition = 2;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 3;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 4;
			ultraGridColumn50.Width = 39;
			ultraGridColumn51.Header.VisiblePosition = 6;
			ultraGridColumn52.Header.VisiblePosition = 7;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance46;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.VisiblePosition = 5;
			ultraGridColumn53.Width = 55;
			ultraGridColumn54.Header.VisiblePosition = 8;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54});
			ultraGridBand6.Header.Caption = "Quinche";
			ultraGridBand6.HeaderVisible = true;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance47;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand6.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings7});
			ultraGridBand6.SummaryFooterCaption = "";
			this.uGridQuinche.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.uGridQuinche.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridQuinche.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance48.BackColor = System.Drawing.Color.Transparent;
			this.uGridQuinche.DisplayLayout.Override.CardAreaAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance49.FontData.BoldAsString = "True";
			appearance49.FontData.Name = "Arial";
			appearance49.FontData.SizeInPoints = 7.5F;
			appearance49.ForeColor = System.Drawing.Color.White;
			appearance49.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridQuinche.DisplayLayout.Override.HeaderAppearance = appearance49;
			appearance50.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance50.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridQuinche.DisplayLayout.Override.RowAlternateAppearance = appearance50;
			appearance51.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance51.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridQuinche.DisplayLayout.Override.RowSelectorAppearance = appearance51;
			appearance52.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance52.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridQuinche.DisplayLayout.Override.SelectedRowAppearance = appearance52;
			this.uGridQuinche.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridQuinche.Location = new System.Drawing.Point(632, 392);
			this.uGridQuinche.Name = "uGridQuinche";
			this.uGridQuinche.Size = new System.Drawing.Size(310, 120);
			this.uGridQuinche.TabIndex = 31;
			this.uGridQuinche.DoubleClick += new System.EventHandler(this.uGridQuinche_DoubleClick);
			// 
			// uGridLatacunga
			// 
			this.uGridLatacunga.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLatacunga.DataSource = this.ultraDataSource1;
			appearance53.BackColor = System.Drawing.Color.White;
			this.uGridLatacunga.DisplayLayout.Appearance = appearance53;
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 41;
			ultraGridColumn56.Header.VisiblePosition = 1;
			ultraGridColumn56.Width = 178;
			ultraGridColumn57.Header.VisiblePosition = 2;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 34;
			ultraGridColumn58.Header.VisiblePosition = 3;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 53;
			ultraGridColumn59.Header.VisiblePosition = 4;
			ultraGridColumn59.Width = 38;
			ultraGridColumn60.Header.VisiblePosition = 6;
			ultraGridColumn61.Header.VisiblePosition = 7;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance54;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.VisiblePosition = 5;
			ultraGridColumn62.Width = 56;
			ultraGridColumn63.Header.VisiblePosition = 8;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63});
			ultraGridBand7.Header.Caption = "Latacunga";
			ultraGridBand7.HeaderVisible = true;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance55;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand7.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings8});
			ultraGridBand7.SummaryFooterCaption = "";
			this.uGridLatacunga.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.uGridLatacunga.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLatacunga.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance56.BackColor = System.Drawing.Color.Transparent;
			this.uGridLatacunga.DisplayLayout.Override.CardAreaAppearance = appearance56;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance57.FontData.BoldAsString = "True";
			appearance57.FontData.Name = "Arial";
			appearance57.FontData.SizeInPoints = 7.5F;
			appearance57.ForeColor = System.Drawing.Color.White;
			appearance57.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLatacunga.DisplayLayout.Override.HeaderAppearance = appearance57;
			appearance58.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance58.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLatacunga.DisplayLayout.Override.RowAlternateAppearance = appearance58;
			appearance59.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance59.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance59.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLatacunga.DisplayLayout.Override.RowSelectorAppearance = appearance59;
			appearance60.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance60.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance60.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLatacunga.DisplayLayout.Override.SelectedRowAppearance = appearance60;
			this.uGridLatacunga.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLatacunga.Location = new System.Drawing.Point(632, 272);
			this.uGridLatacunga.Name = "uGridLatacunga";
			this.uGridLatacunga.Size = new System.Drawing.Size(310, 120);
			this.uGridLatacunga.TabIndex = 30;
			this.uGridLatacunga.DoubleClick += new System.EventHandler(this.uGridLatacunga_DoubleClick);
			// 
			// uGridInca
			// 
			this.uGridInca.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInca.DataSource = this.ultraDataSource1;
			appearance61.BackColor = System.Drawing.Color.White;
			this.uGridInca.DisplayLayout.Appearance = appearance61;
			ultraGridBand8.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn64.Header.VisiblePosition = 0;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 1;
			ultraGridColumn65.Width = 178;
			ultraGridColumn66.Header.VisiblePosition = 2;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 3;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 4;
			ultraGridColumn68.Width = 39;
			ultraGridColumn69.Header.VisiblePosition = 6;
			ultraGridColumn70.Header.VisiblePosition = 7;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance62;
			ultraGridColumn71.Format = "#,##0.00";
			ultraGridColumn71.Header.VisiblePosition = 5;
			ultraGridColumn71.Width = 55;
			ultraGridColumn72.Header.VisiblePosition = 8;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67,
																										 ultraGridColumn68,
																										 ultraGridColumn69,
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72});
			ultraGridBand8.Header.Caption = "Inca";
			ultraGridBand8.HeaderVisible = true;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance63;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand8.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings9});
			ultraGridBand8.SummaryFooterCaption = "";
			this.uGridInca.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.uGridInca.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInca.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance64.BackColor = System.Drawing.Color.Transparent;
			this.uGridInca.DisplayLayout.Override.CardAreaAppearance = appearance64;
			appearance65.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance65.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance65.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance65.FontData.BoldAsString = "True";
			appearance65.FontData.Name = "Arial";
			appearance65.FontData.SizeInPoints = 7.5F;
			appearance65.ForeColor = System.Drawing.Color.White;
			appearance65.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInca.DisplayLayout.Override.HeaderAppearance = appearance65;
			appearance66.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance66.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance66.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInca.DisplayLayout.Override.RowAlternateAppearance = appearance66;
			appearance67.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance67.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance67.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInca.DisplayLayout.Override.RowSelectorAppearance = appearance67;
			appearance68.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance68.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance68.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInca.DisplayLayout.Override.SelectedRowAppearance = appearance68;
			this.uGridInca.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInca.Location = new System.Drawing.Point(320, 272);
			this.uGridInca.Name = "uGridInca";
			this.uGridInca.Size = new System.Drawing.Size(310, 120);
			this.uGridInca.TabIndex = 29;
			this.uGridInca.DoubleClick += new System.EventHandler(this.uGridInca_DoubleClick);
			// 
			// uGridSantoDomingo
			// 
			this.uGridSantoDomingo.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSantoDomingo.DataSource = this.ultraDataSource1;
			appearance69.BackColor = System.Drawing.Color.White;
			this.uGridSantoDomingo.DisplayLayout.Appearance = appearance69;
			ultraGridColumn73.Header.VisiblePosition = 0;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn73.Width = 41;
			ultraGridColumn74.Header.VisiblePosition = 1;
			ultraGridColumn74.Width = 178;
			ultraGridColumn75.Header.VisiblePosition = 2;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn75.Width = 34;
			ultraGridColumn76.Header.VisiblePosition = 3;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.Width = 53;
			ultraGridColumn77.Header.VisiblePosition = 4;
			ultraGridColumn77.Width = 39;
			ultraGridColumn78.Header.VisiblePosition = 6;
			ultraGridColumn79.Header.VisiblePosition = 7;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance70;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.VisiblePosition = 5;
			ultraGridColumn80.Width = 55;
			ultraGridColumn81.Header.VisiblePosition = 8;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn73,
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79,
																										 ultraGridColumn80,
																										 ultraGridColumn81});
			ultraGridBand9.Header.Caption = "Santo Domingo";
			ultraGridBand9.HeaderVisible = true;
			appearance71.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance71;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand9.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings10});
			ultraGridBand9.SummaryFooterCaption = "";
			this.uGridSantoDomingo.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.uGridSantoDomingo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSantoDomingo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance72.BackColor = System.Drawing.Color.Transparent;
			this.uGridSantoDomingo.DisplayLayout.Override.CardAreaAppearance = appearance72;
			appearance73.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance73.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance73.FontData.BoldAsString = "True";
			appearance73.FontData.Name = "Arial";
			appearance73.FontData.SizeInPoints = 7.5F;
			appearance73.ForeColor = System.Drawing.Color.White;
			appearance73.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSantoDomingo.DisplayLayout.Override.HeaderAppearance = appearance73;
			appearance74.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance74.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance74.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSantoDomingo.DisplayLayout.Override.RowAlternateAppearance = appearance74;
			appearance75.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance75.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance75.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSantoDomingo.DisplayLayout.Override.RowSelectorAppearance = appearance75;
			appearance76.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance76.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance76.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSantoDomingo.DisplayLayout.Override.SelectedRowAppearance = appearance76;
			this.uGridSantoDomingo.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSantoDomingo.Location = new System.Drawing.Point(8, 632);
			this.uGridSantoDomingo.Name = "uGridSantoDomingo";
			this.uGridSantoDomingo.Size = new System.Drawing.Size(310, 120);
			this.uGridSantoDomingo.TabIndex = 34;
			this.uGridSantoDomingo.DoubleClick += new System.EventHandler(this.uGridSantoDomingo_DoubleClick);
			// 
			// uGridSangolqui
			// 
			this.uGridSangolqui.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSangolqui.DataSource = this.ultraDataSource1;
			appearance77.BackColor = System.Drawing.Color.White;
			this.uGridSangolqui.DisplayLayout.Appearance = appearance77;
			ultraGridBand10.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn82.Header.VisiblePosition = 0;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 1;
			ultraGridColumn83.Width = 178;
			ultraGridColumn84.Header.VisiblePosition = 2;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 3;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 4;
			ultraGridColumn86.Width = 39;
			ultraGridColumn87.Header.VisiblePosition = 6;
			ultraGridColumn88.Header.VisiblePosition = 7;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn89.CellAppearance = appearance78;
			ultraGridColumn89.Format = "#,##0.00";
			ultraGridColumn89.Header.VisiblePosition = 5;
			ultraGridColumn89.Width = 55;
			ultraGridColumn90.Header.VisiblePosition = 8;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn82,
																											ultraGridColumn83,
																											ultraGridColumn84,
																											ultraGridColumn85,
																											ultraGridColumn86,
																											ultraGridColumn87,
																											ultraGridColumn88,
																											ultraGridColumn89,
																											ultraGridColumn90});
			ultraGridBand10.Header.Caption = "Sangolqui";
			ultraGridBand10.HeaderVisible = true;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance79;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand10.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings11});
			ultraGridBand10.SummaryFooterCaption = "";
			this.uGridSangolqui.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.uGridSangolqui.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSangolqui.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance80.BackColor = System.Drawing.Color.Transparent;
			this.uGridSangolqui.DisplayLayout.Override.CardAreaAppearance = appearance80;
			appearance81.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance81.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance81.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance81.FontData.BoldAsString = "True";
			appearance81.FontData.Name = "Arial";
			appearance81.FontData.SizeInPoints = 7.5F;
			appearance81.ForeColor = System.Drawing.Color.White;
			appearance81.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSangolqui.DisplayLayout.Override.HeaderAppearance = appearance81;
			appearance82.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance82.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance82.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSangolqui.DisplayLayout.Override.RowAlternateAppearance = appearance82;
			appearance83.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance83.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance83.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSangolqui.DisplayLayout.Override.RowSelectorAppearance = appearance83;
			appearance84.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance84.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance84.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSangolqui.DisplayLayout.Override.SelectedRowAppearance = appearance84;
			this.uGridSangolqui.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSangolqui.Location = new System.Drawing.Point(944, 512);
			this.uGridSangolqui.Name = "uGridSangolqui";
			this.uGridSangolqui.Size = new System.Drawing.Size(310, 120);
			this.uGridSangolqui.TabIndex = 33;
			this.uGridSangolqui.DoubleClick += new System.EventHandler(this.uGridSangolqui_DoubleClick);
			// 
			// uGridTumbaco
			// 
			this.uGridTumbaco.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTumbaco.DataSource = this.ultraDataSource1;
			appearance85.BackColor = System.Drawing.Color.White;
			this.uGridTumbaco.DisplayLayout.Appearance = appearance85;
			ultraGridColumn91.Header.VisiblePosition = 0;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Width = 41;
			ultraGridColumn92.Header.VisiblePosition = 1;
			ultraGridColumn92.Width = 178;
			ultraGridColumn93.Header.VisiblePosition = 2;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn93.Width = 34;
			ultraGridColumn94.Header.VisiblePosition = 3;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn94.Width = 53;
			ultraGridColumn95.Header.VisiblePosition = 4;
			ultraGridColumn95.Width = 39;
			ultraGridColumn96.Header.VisiblePosition = 6;
			ultraGridColumn97.Header.VisiblePosition = 7;
			appearance86.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn98.CellAppearance = appearance86;
			ultraGridColumn98.Format = "#,##0.00";
			ultraGridColumn98.Header.VisiblePosition = 5;
			ultraGridColumn98.Width = 55;
			ultraGridColumn99.Header.VisiblePosition = 8;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn91,
																											ultraGridColumn92,
																											ultraGridColumn93,
																											ultraGridColumn94,
																											ultraGridColumn95,
																											ultraGridColumn96,
																											ultraGridColumn97,
																											ultraGridColumn98,
																											ultraGridColumn99});
			ultraGridBand11.Header.Caption = "Tumbaco";
			ultraGridBand11.HeaderVisible = true;
			appearance87.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance87;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand11.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings12});
			ultraGridBand11.SummaryFooterCaption = "";
			this.uGridTumbaco.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.uGridTumbaco.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridTumbaco.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance88.BackColor = System.Drawing.Color.Transparent;
			this.uGridTumbaco.DisplayLayout.Override.CardAreaAppearance = appearance88;
			appearance89.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance89.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance89.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance89.FontData.BoldAsString = "True";
			appearance89.FontData.Name = "Arial";
			appearance89.FontData.SizeInPoints = 7.5F;
			appearance89.ForeColor = System.Drawing.Color.White;
			appearance89.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTumbaco.DisplayLayout.Override.HeaderAppearance = appearance89;
			appearance90.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance90.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance90.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTumbaco.DisplayLayout.Override.RowAlternateAppearance = appearance90;
			appearance91.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance91.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance91.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTumbaco.DisplayLayout.Override.RowSelectorAppearance = appearance91;
			appearance92.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance92.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance92.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTumbaco.DisplayLayout.Override.SelectedRowAppearance = appearance92;
			this.uGridTumbaco.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTumbaco.Location = new System.Drawing.Point(632, 632);
			this.uGridTumbaco.Name = "uGridTumbaco";
			this.uGridTumbaco.Size = new System.Drawing.Size(310, 120);
			this.uGridTumbaco.TabIndex = 44;
			this.uGridTumbaco.DoubleClick += new System.EventHandler(this.uGridTumbaco_DoubleClick);
			// 
			// uGridSolanda
			// 
			this.uGridSolanda.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSolanda.DataSource = this.ultraDataSource1;
			appearance93.BackColor = System.Drawing.Color.White;
			this.uGridSolanda.DisplayLayout.Appearance = appearance93;
			ultraGridBand12.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn100.Header.VisiblePosition = 0;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 1;
			ultraGridColumn101.Width = 178;
			ultraGridColumn102.Header.VisiblePosition = 2;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 3;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 4;
			ultraGridColumn104.Width = 38;
			ultraGridColumn105.Header.VisiblePosition = 6;
			ultraGridColumn106.Header.VisiblePosition = 7;
			appearance94.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn107.CellAppearance = appearance94;
			ultraGridColumn107.Format = "#,##0.00";
			ultraGridColumn107.Header.VisiblePosition = 5;
			ultraGridColumn107.Width = 56;
			ultraGridColumn108.Header.VisiblePosition = 8;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn100,
																											ultraGridColumn101,
																											ultraGridColumn102,
																											ultraGridColumn103,
																											ultraGridColumn104,
																											ultraGridColumn105,
																											ultraGridColumn106,
																											ultraGridColumn107,
																											ultraGridColumn108});
			ultraGridBand12.Header.Caption = "Solanda";
			ultraGridBand12.HeaderVisible = true;
			appearance95.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance95;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand12.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings13});
			ultraGridBand12.SummaryFooterCaption = "";
			this.uGridSolanda.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.uGridSolanda.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSolanda.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance96.BackColor = System.Drawing.Color.Transparent;
			this.uGridSolanda.DisplayLayout.Override.CardAreaAppearance = appearance96;
			appearance97.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance97.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance97.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance97.FontData.BoldAsString = "True";
			appearance97.FontData.Name = "Arial";
			appearance97.FontData.SizeInPoints = 7.5F;
			appearance97.ForeColor = System.Drawing.Color.White;
			appearance97.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSolanda.DisplayLayout.Override.HeaderAppearance = appearance97;
			appearance98.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance98.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance98.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSolanda.DisplayLayout.Override.RowAlternateAppearance = appearance98;
			appearance99.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance99.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance99.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSolanda.DisplayLayout.Override.RowSelectorAppearance = appearance99;
			appearance100.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance100.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance100.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSolanda.DisplayLayout.Override.SelectedRowAppearance = appearance100;
			this.uGridSolanda.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSolanda.Location = new System.Drawing.Point(320, 632);
			this.uGridSolanda.Name = "uGridSolanda";
			this.uGridSolanda.Size = new System.Drawing.Size(310, 120);
			this.uGridSolanda.TabIndex = 43;
			this.uGridSolanda.DoubleClick += new System.EventHandler(this.uGridSolanda_DoubleClick);
			// 
			// uGridRiobamba
			// 
			this.uGridRiobamba.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridRiobamba.DataSource = this.ultraDataSource1;
			appearance101.BackColor = System.Drawing.Color.White;
			this.uGridRiobamba.DisplayLayout.Appearance = appearance101;
			ultraGridColumn109.Header.VisiblePosition = 0;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn109.Width = 41;
			ultraGridColumn110.Header.VisiblePosition = 1;
			ultraGridColumn110.Width = 178;
			ultraGridColumn111.Header.VisiblePosition = 2;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn111.Width = 34;
			ultraGridColumn112.Header.VisiblePosition = 3;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn112.Width = 53;
			ultraGridColumn113.Header.VisiblePosition = 4;
			ultraGridColumn113.Width = 38;
			ultraGridColumn114.Header.VisiblePosition = 6;
			ultraGridColumn115.Header.VisiblePosition = 7;
			appearance102.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn116.CellAppearance = appearance102;
			ultraGridColumn116.Format = "#,##0.00";
			ultraGridColumn116.Header.VisiblePosition = 5;
			ultraGridColumn116.Width = 56;
			ultraGridColumn117.Header.VisiblePosition = 8;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn109,
																											ultraGridColumn110,
																											ultraGridColumn111,
																											ultraGridColumn112,
																											ultraGridColumn113,
																											ultraGridColumn114,
																											ultraGridColumn115,
																											ultraGridColumn116,
																											ultraGridColumn117});
			ultraGridBand13.Header.Caption = "Riobamba";
			ultraGridBand13.HeaderVisible = true;
			appearance103.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance103;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand13.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings14});
			ultraGridBand13.SummaryFooterCaption = "";
			this.uGridRiobamba.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.uGridRiobamba.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridRiobamba.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance104.BackColor = System.Drawing.Color.Transparent;
			this.uGridRiobamba.DisplayLayout.Override.CardAreaAppearance = appearance104;
			appearance105.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance105.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance105.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance105.FontData.BoldAsString = "True";
			appearance105.FontData.Name = "Arial";
			appearance105.FontData.SizeInPoints = 7.5F;
			appearance105.ForeColor = System.Drawing.Color.White;
			appearance105.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRiobamba.DisplayLayout.Override.HeaderAppearance = appearance105;
			appearance106.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance106.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance106.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRiobamba.DisplayLayout.Override.RowAlternateAppearance = appearance106;
			appearance107.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance107.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance107.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRiobamba.DisplayLayout.Override.RowSelectorAppearance = appearance107;
			appearance108.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance108.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance108.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRiobamba.DisplayLayout.Override.SelectedRowAppearance = appearance108;
			this.uGridRiobamba.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRiobamba.Location = new System.Drawing.Point(320, 512);
			this.uGridRiobamba.Name = "uGridRiobamba";
			this.uGridRiobamba.Size = new System.Drawing.Size(310, 120);
			this.uGridRiobamba.TabIndex = 42;
			this.uGridRiobamba.DoubleClick += new System.EventHandler(this.uGridRiobamba_DoubleClick);
			// 
			// uGridRecreo3
			// 
			this.uGridRecreo3.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridRecreo3.DataSource = this.ultraDataSource1;
			appearance109.BackColor = System.Drawing.Color.White;
			this.uGridRecreo3.DisplayLayout.Appearance = appearance109;
			ultraGridBand14.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn118.Header.VisiblePosition = 0;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 1;
			ultraGridColumn119.Width = 178;
			ultraGridColumn120.Header.VisiblePosition = 2;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 3;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 4;
			ultraGridColumn122.Width = 39;
			ultraGridColumn123.Header.VisiblePosition = 6;
			ultraGridColumn124.Header.VisiblePosition = 7;
			appearance110.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn125.CellAppearance = appearance110;
			ultraGridColumn125.Format = "#,##0.00";
			ultraGridColumn125.Header.VisiblePosition = 5;
			ultraGridColumn125.Width = 55;
			ultraGridColumn126.Header.VisiblePosition = 8;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn118,
																											ultraGridColumn119,
																											ultraGridColumn120,
																											ultraGridColumn121,
																											ultraGridColumn122,
																											ultraGridColumn123,
																											ultraGridColumn124,
																											ultraGridColumn125,
																											ultraGridColumn126});
			ultraGridBand14.Header.Caption = "Recreo 3";
			ultraGridBand14.HeaderVisible = true;
			appearance111.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance111;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand14.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings15});
			ultraGridBand14.SummaryFooterCaption = "";
			this.uGridRecreo3.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.uGridRecreo3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridRecreo3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance112.BackColor = System.Drawing.Color.Transparent;
			this.uGridRecreo3.DisplayLayout.Override.CardAreaAppearance = appearance112;
			appearance113.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance113.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance113.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance113.FontData.BoldAsString = "True";
			appearance113.FontData.Name = "Arial";
			appearance113.FontData.SizeInPoints = 7.5F;
			appearance113.ForeColor = System.Drawing.Color.White;
			appearance113.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRecreo3.DisplayLayout.Override.HeaderAppearance = appearance113;
			appearance114.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance114.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance114.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRecreo3.DisplayLayout.Override.RowAlternateAppearance = appearance114;
			appearance115.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance115.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance115.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRecreo3.DisplayLayout.Override.RowSelectorAppearance = appearance115;
			appearance116.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance116.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance116.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRecreo3.DisplayLayout.Override.SelectedRowAppearance = appearance116;
			this.uGridRecreo3.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRecreo3.Location = new System.Drawing.Point(8, 512);
			this.uGridRecreo3.Name = "uGridRecreo3";
			this.uGridRecreo3.Size = new System.Drawing.Size(310, 120);
			this.uGridRecreo3.TabIndex = 41;
			this.uGridRecreo3.DoubleClick += new System.EventHandler(this.uGridRecreo3_DoubleClick);
			// 
			// uGridPrensa
			// 
			this.uGridPrensa.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPrensa.DataSource = this.ultraDataSource1;
			appearance117.BackColor = System.Drawing.Color.White;
			this.uGridPrensa.DisplayLayout.Appearance = appearance117;
			ultraGridColumn127.Header.VisiblePosition = 0;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn127.Width = 41;
			ultraGridColumn128.Header.VisiblePosition = 1;
			ultraGridColumn128.Width = 178;
			ultraGridColumn129.Header.VisiblePosition = 2;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn129.Width = 34;
			ultraGridColumn130.Header.VisiblePosition = 3;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn130.Width = 53;
			ultraGridColumn131.Header.VisiblePosition = 4;
			ultraGridColumn131.Width = 39;
			ultraGridColumn132.Header.VisiblePosition = 6;
			ultraGridColumn133.Header.VisiblePosition = 7;
			appearance118.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn134.CellAppearance = appearance118;
			ultraGridColumn134.Format = "#,##0.00";
			ultraGridColumn134.Header.VisiblePosition = 5;
			ultraGridColumn134.Width = 55;
			ultraGridColumn135.Header.VisiblePosition = 8;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn127,
																											ultraGridColumn128,
																											ultraGridColumn129,
																											ultraGridColumn130,
																											ultraGridColumn131,
																											ultraGridColumn132,
																											ultraGridColumn133,
																											ultraGridColumn134,
																											ultraGridColumn135});
			ultraGridBand15.Header.Caption = "Prensa";
			ultraGridBand15.HeaderVisible = true;
			appearance119.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance119;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand15.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings16});
			ultraGridBand15.SummaryFooterCaption = "";
			this.uGridPrensa.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.uGridPrensa.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPrensa.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance120.BackColor = System.Drawing.Color.Transparent;
			this.uGridPrensa.DisplayLayout.Override.CardAreaAppearance = appearance120;
			appearance121.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance121.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance121.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance121.FontData.BoldAsString = "True";
			appearance121.FontData.Name = "Arial";
			appearance121.FontData.SizeInPoints = 7.5F;
			appearance121.ForeColor = System.Drawing.Color.White;
			appearance121.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPrensa.DisplayLayout.Override.HeaderAppearance = appearance121;
			appearance122.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance122.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance122.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPrensa.DisplayLayout.Override.RowAlternateAppearance = appearance122;
			appearance123.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance123.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance123.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPrensa.DisplayLayout.Override.RowSelectorAppearance = appearance123;
			appearance124.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance124.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance124.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPrensa.DisplayLayout.Override.SelectedRowAppearance = appearance124;
			this.uGridPrensa.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPrensa.Location = new System.Drawing.Point(8, 392);
			this.uGridPrensa.Name = "uGridPrensa";
			this.uGridPrensa.Size = new System.Drawing.Size(310, 120);
			this.uGridPrensa.TabIndex = 40;
			this.uGridPrensa.DoubleClick += new System.EventHandler(this.uGridPrensa_DoubleClick);
			// 
			// uGridMachachi
			// 
			this.uGridMachachi.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridMachachi.DataSource = this.ultraDataSource1;
			appearance125.BackColor = System.Drawing.Color.White;
			this.uGridMachachi.DisplayLayout.Appearance = appearance125;
			ultraGridBand16.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn136.Header.VisiblePosition = 0;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 1;
			ultraGridColumn137.Width = 178;
			ultraGridColumn138.Header.VisiblePosition = 2;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn139.Header.VisiblePosition = 3;
			ultraGridColumn139.Hidden = true;
			ultraGridColumn140.Header.VisiblePosition = 4;
			ultraGridColumn140.Width = 38;
			ultraGridColumn141.Header.VisiblePosition = 6;
			ultraGridColumn142.Header.VisiblePosition = 7;
			appearance126.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn143.CellAppearance = appearance126;
			ultraGridColumn143.Format = "#,##0.00";
			ultraGridColumn143.Header.VisiblePosition = 5;
			ultraGridColumn143.Width = 56;
			ultraGridColumn144.Header.VisiblePosition = 8;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn136,
																											ultraGridColumn137,
																											ultraGridColumn138,
																											ultraGridColumn139,
																											ultraGridColumn140,
																											ultraGridColumn141,
																											ultraGridColumn142,
																											ultraGridColumn143,
																											ultraGridColumn144});
			ultraGridBand16.Header.Caption = "Machachi";
			ultraGridBand16.HeaderVisible = true;
			appearance127.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance127;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand16.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings17});
			ultraGridBand16.SummaryFooterCaption = "";
			this.uGridMachachi.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.uGridMachachi.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridMachachi.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance128.BackColor = System.Drawing.Color.Transparent;
			this.uGridMachachi.DisplayLayout.Override.CardAreaAppearance = appearance128;
			appearance129.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance129.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance129.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance129.FontData.BoldAsString = "True";
			appearance129.FontData.Name = "Arial";
			appearance129.FontData.SizeInPoints = 7.5F;
			appearance129.ForeColor = System.Drawing.Color.White;
			appearance129.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridMachachi.DisplayLayout.Override.HeaderAppearance = appearance129;
			appearance130.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance130.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance130.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMachachi.DisplayLayout.Override.RowAlternateAppearance = appearance130;
			appearance131.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance131.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance131.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMachachi.DisplayLayout.Override.RowSelectorAppearance = appearance131;
			appearance132.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance132.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance132.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMachachi.DisplayLayout.Override.SelectedRowAppearance = appearance132;
			this.uGridMachachi.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridMachachi.Location = new System.Drawing.Point(944, 272);
			this.uGridMachachi.Name = "uGridMachachi";
			this.uGridMachachi.Size = new System.Drawing.Size(310, 120);
			this.uGridMachachi.TabIndex = 39;
			this.uGridMachachi.DoubleClick += new System.EventHandler(this.uGridMachachi_DoubleClick);
			// 
			// uGridGuamani
			// 
			this.uGridGuamani.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGuamani.DataSource = this.ultraDataSource1;
			appearance133.BackColor = System.Drawing.Color.White;
			this.uGridGuamani.DisplayLayout.Appearance = appearance133;
			ultraGridColumn145.Header.VisiblePosition = 0;
			ultraGridColumn145.Hidden = true;
			ultraGridColumn145.Width = 41;
			ultraGridColumn146.Header.VisiblePosition = 1;
			ultraGridColumn146.Width = 178;
			ultraGridColumn147.Header.VisiblePosition = 2;
			ultraGridColumn147.Hidden = true;
			ultraGridColumn147.Width = 34;
			ultraGridColumn148.Header.VisiblePosition = 3;
			ultraGridColumn148.Hidden = true;
			ultraGridColumn148.Width = 53;
			ultraGridColumn149.Header.VisiblePosition = 4;
			ultraGridColumn149.Width = 39;
			ultraGridColumn150.Header.VisiblePosition = 6;
			ultraGridColumn151.Header.VisiblePosition = 7;
			appearance134.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn152.CellAppearance = appearance134;
			ultraGridColumn152.Format = "#,##0.00";
			ultraGridColumn152.Header.VisiblePosition = 5;
			ultraGridColumn152.Width = 55;
			ultraGridColumn153.Header.VisiblePosition = 8;
			ultraGridBand17.Columns.AddRange(new object[] {
																											ultraGridColumn145,
																											ultraGridColumn146,
																											ultraGridColumn147,
																											ultraGridColumn148,
																											ultraGridColumn149,
																											ultraGridColumn150,
																											ultraGridColumn151,
																											ultraGridColumn152,
																											ultraGridColumn153});
			ultraGridBand17.Header.Caption = "Guamani";
			ultraGridBand17.HeaderVisible = true;
			appearance135.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance135;
			summarySettings18.DisplayFormat = "{0: #,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand17.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings18});
			ultraGridBand17.SummaryFooterCaption = "";
			this.uGridGuamani.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			this.uGridGuamani.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridGuamani.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance136.BackColor = System.Drawing.Color.Transparent;
			this.uGridGuamani.DisplayLayout.Override.CardAreaAppearance = appearance136;
			appearance137.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance137.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance137.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance137.FontData.BoldAsString = "True";
			appearance137.FontData.Name = "Arial";
			appearance137.FontData.SizeInPoints = 7.5F;
			appearance137.ForeColor = System.Drawing.Color.White;
			appearance137.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGuamani.DisplayLayout.Override.HeaderAppearance = appearance137;
			appearance138.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance138.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance138.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGuamani.DisplayLayout.Override.RowAlternateAppearance = appearance138;
			appearance139.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance139.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance139.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGuamani.DisplayLayout.Override.RowSelectorAppearance = appearance139;
			appearance140.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance140.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance140.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGuamani.DisplayLayout.Override.SelectedRowAppearance = appearance140;
			this.uGridGuamani.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGuamani.Location = new System.Drawing.Point(941, 152);
			this.uGridGuamani.Name = "uGridGuamani";
			this.uGridGuamani.Size = new System.Drawing.Size(310, 120);
			this.uGridGuamani.TabIndex = 38;
			this.uGridGuamani.DoubleClick += new System.EventHandler(this.uGridGuamani_DoubleClick);
			// 
			// uGridConsumidorFinal
			// 
			this.uGridConsumidorFinal.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridConsumidorFinal.DataSource = this.ultraDataSource1;
			appearance141.BackColor = System.Drawing.Color.White;
			this.uGridConsumidorFinal.DisplayLayout.Appearance = appearance141;
			ultraGridBand18.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn154.Header.VisiblePosition = 0;
			ultraGridColumn154.Hidden = true;
			ultraGridColumn155.Header.VisiblePosition = 1;
			ultraGridColumn155.Width = 178;
			ultraGridColumn156.Header.VisiblePosition = 2;
			ultraGridColumn156.Hidden = true;
			ultraGridColumn157.Header.VisiblePosition = 3;
			ultraGridColumn157.Hidden = true;
			ultraGridColumn158.Header.VisiblePosition = 4;
			ultraGridColumn158.Width = 38;
			ultraGridColumn159.Header.VisiblePosition = 6;
			ultraGridColumn160.Header.VisiblePosition = 7;
			appearance142.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn161.CellAppearance = appearance142;
			ultraGridColumn161.Format = "#,##0.00";
			ultraGridColumn161.Header.VisiblePosition = 5;
			ultraGridColumn161.Width = 56;
			ultraGridColumn162.Header.VisiblePosition = 8;
			ultraGridBand18.Columns.AddRange(new object[] {
																											ultraGridColumn154,
																											ultraGridColumn155,
																											ultraGridColumn156,
																											ultraGridColumn157,
																											ultraGridColumn158,
																											ultraGridColumn159,
																											ultraGridColumn160,
																											ultraGridColumn161,
																											ultraGridColumn162});
			ultraGridBand18.Header.Caption = "Consumidor Final";
			ultraGridBand18.HeaderVisible = true;
			appearance143.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance143;
			summarySettings19.DisplayFormat = "{0: #,##0.00}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand18.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings19});
			ultraGridBand18.SummaryFooterCaption = "";
			this.uGridConsumidorFinal.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			this.uGridConsumidorFinal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridConsumidorFinal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance144.BackColor = System.Drawing.Color.Transparent;
			this.uGridConsumidorFinal.DisplayLayout.Override.CardAreaAppearance = appearance144;
			appearance145.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance145.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance145.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance145.FontData.BoldAsString = "True";
			appearance145.FontData.Name = "Arial";
			appearance145.FontData.SizeInPoints = 7.5F;
			appearance145.ForeColor = System.Drawing.Color.White;
			appearance145.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridConsumidorFinal.DisplayLayout.Override.HeaderAppearance = appearance145;
			appearance146.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance146.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance146.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConsumidorFinal.DisplayLayout.Override.RowAlternateAppearance = appearance146;
			appearance147.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance147.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance147.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConsumidorFinal.DisplayLayout.Override.RowSelectorAppearance = appearance147;
			appearance148.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance148.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance148.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConsumidorFinal.DisplayLayout.Override.SelectedRowAppearance = appearance148;
			this.uGridConsumidorFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridConsumidorFinal.Location = new System.Drawing.Point(629, 152);
			this.uGridConsumidorFinal.Name = "uGridConsumidorFinal";
			this.uGridConsumidorFinal.Size = new System.Drawing.Size(310, 120);
			this.uGridConsumidorFinal.TabIndex = 37;
			this.uGridConsumidorFinal.DoubleClick += new System.EventHandler(this.uGridConsumidorFinal_DoubleClick);
			// 
			// uGridCarapungo
			// 
			this.uGridCarapungo.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCarapungo.DataSource = this.ultraDataSource1;
			appearance149.BackColor = System.Drawing.Color.White;
			this.uGridCarapungo.DisplayLayout.Appearance = appearance149;
			ultraGridColumn163.Header.VisiblePosition = 0;
			ultraGridColumn163.Hidden = true;
			ultraGridColumn163.Width = 41;
			ultraGridColumn164.Header.VisiblePosition = 1;
			ultraGridColumn164.Width = 178;
			ultraGridColumn165.Header.VisiblePosition = 2;
			ultraGridColumn165.Hidden = true;
			ultraGridColumn165.Width = 34;
			ultraGridColumn166.Header.VisiblePosition = 3;
			ultraGridColumn166.Hidden = true;
			ultraGridColumn166.Width = 53;
			ultraGridColumn167.Header.VisiblePosition = 4;
			ultraGridColumn167.Width = 39;
			ultraGridColumn168.Header.VisiblePosition = 6;
			ultraGridColumn169.Header.VisiblePosition = 7;
			appearance150.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn170.CellAppearance = appearance150;
			ultraGridColumn170.Format = "#,##0.00";
			ultraGridColumn170.Header.VisiblePosition = 5;
			ultraGridColumn170.Width = 55;
			ultraGridColumn171.Header.VisiblePosition = 8;
			ultraGridBand19.Columns.AddRange(new object[] {
																											ultraGridColumn163,
																											ultraGridColumn164,
																											ultraGridColumn165,
																											ultraGridColumn166,
																											ultraGridColumn167,
																											ultraGridColumn168,
																											ultraGridColumn169,
																											ultraGridColumn170,
																											ultraGridColumn171});
			ultraGridBand19.Header.Caption = "Carapungo";
			ultraGridBand19.HeaderVisible = true;
			appearance151.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance151;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand19.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings20});
			ultraGridBand19.SummaryFooterCaption = "";
			this.uGridCarapungo.DisplayLayout.BandsSerializer.Add(ultraGridBand19);
			this.uGridCarapungo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCarapungo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance152.BackColor = System.Drawing.Color.Transparent;
			this.uGridCarapungo.DisplayLayout.Override.CardAreaAppearance = appearance152;
			appearance153.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance153.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance153.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance153.FontData.BoldAsString = "True";
			appearance153.FontData.Name = "Arial";
			appearance153.FontData.SizeInPoints = 7.5F;
			appearance153.ForeColor = System.Drawing.Color.White;
			appearance153.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCarapungo.DisplayLayout.Override.HeaderAppearance = appearance153;
			appearance154.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance154.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance154.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCarapungo.DisplayLayout.Override.RowAlternateAppearance = appearance154;
			appearance155.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance155.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance155.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCarapungo.DisplayLayout.Override.RowSelectorAppearance = appearance155;
			appearance156.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance156.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance156.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCarapungo.DisplayLayout.Override.SelectedRowAppearance = appearance156;
			this.uGridCarapungo.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCarapungo.Location = new System.Drawing.Point(941, 32);
			this.uGridCarapungo.Name = "uGridCarapungo";
			this.uGridCarapungo.Size = new System.Drawing.Size(310, 120);
			this.uGridCarapungo.TabIndex = 36;
			this.uGridCarapungo.DoubleClick += new System.EventHandler(this.uGridCarapungo_DoubleClick);
			// 
			// uGridCalderon
			// 
			this.uGridCalderon.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCalderon.DataSource = this.ultraDataSource1;
			appearance157.BackColor = System.Drawing.Color.White;
			this.uGridCalderon.DisplayLayout.Appearance = appearance157;
			ultraGridBand20.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn172.Header.VisiblePosition = 0;
			ultraGridColumn172.Hidden = true;
			ultraGridColumn173.Header.VisiblePosition = 1;
			ultraGridColumn173.Width = 178;
			ultraGridColumn174.Header.VisiblePosition = 2;
			ultraGridColumn174.Hidden = true;
			ultraGridColumn175.Header.VisiblePosition = 3;
			ultraGridColumn175.Hidden = true;
			ultraGridColumn176.Header.VisiblePosition = 4;
			ultraGridColumn176.Width = 39;
			ultraGridColumn177.Header.VisiblePosition = 6;
			ultraGridColumn178.Header.VisiblePosition = 7;
			appearance158.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn179.CellAppearance = appearance158;
			ultraGridColumn179.Format = "#,##0.00";
			ultraGridColumn179.Header.VisiblePosition = 5;
			ultraGridColumn179.Width = 55;
			ultraGridColumn180.Header.VisiblePosition = 8;
			ultraGridBand20.Columns.AddRange(new object[] {
																											ultraGridColumn172,
																											ultraGridColumn173,
																											ultraGridColumn174,
																											ultraGridColumn175,
																											ultraGridColumn176,
																											ultraGridColumn177,
																											ultraGridColumn178,
																											ultraGridColumn179,
																											ultraGridColumn180});
			ultraGridBand20.Header.Caption = "Calderón";
			ultraGridBand20.HeaderVisible = true;
			appearance159.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance159;
			summarySettings21.DisplayFormat = "{0: #,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand20.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings21});
			ultraGridBand20.SummaryFooterCaption = "";
			this.uGridCalderon.DisplayLayout.BandsSerializer.Add(ultraGridBand20);
			this.uGridCalderon.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCalderon.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance160.BackColor = System.Drawing.Color.Transparent;
			this.uGridCalderon.DisplayLayout.Override.CardAreaAppearance = appearance160;
			appearance161.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance161.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance161.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance161.FontData.BoldAsString = "True";
			appearance161.FontData.Name = "Arial";
			appearance161.FontData.SizeInPoints = 7.5F;
			appearance161.ForeColor = System.Drawing.Color.White;
			appearance161.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCalderon.DisplayLayout.Override.HeaderAppearance = appearance161;
			appearance162.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance162.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance162.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCalderon.DisplayLayout.Override.RowAlternateAppearance = appearance162;
			appearance163.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance163.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance163.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCalderon.DisplayLayout.Override.RowSelectorAppearance = appearance163;
			appearance164.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance164.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance164.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCalderon.DisplayLayout.Override.SelectedRowAppearance = appearance164;
			this.uGridCalderon.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCalderon.Location = new System.Drawing.Point(629, 32);
			this.uGridCalderon.Name = "uGridCalderon";
			this.uGridCalderon.Size = new System.Drawing.Size(310, 120);
			this.uGridCalderon.TabIndex = 35;
			this.uGridCalderon.DoubleClick += new System.EventHandler(this.uGridCalderon_DoubleClick);
			// 
			// uGridCayambe
			// 
			this.uGridCayambe.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCayambe.DataSource = this.ultraDataSource1;
			appearance165.BackColor = System.Drawing.Color.White;
			this.uGridCayambe.DisplayLayout.Appearance = appearance165;
			ultraGridColumn181.Header.VisiblePosition = 0;
			ultraGridColumn181.Hidden = true;
			ultraGridColumn181.Width = 41;
			ultraGridColumn182.Header.VisiblePosition = 1;
			ultraGridColumn182.Width = 178;
			ultraGridColumn183.Header.VisiblePosition = 2;
			ultraGridColumn183.Hidden = true;
			ultraGridColumn183.Width = 34;
			ultraGridColumn184.Header.VisiblePosition = 3;
			ultraGridColumn184.Hidden = true;
			ultraGridColumn184.Width = 53;
			ultraGridColumn185.Header.VisiblePosition = 4;
			ultraGridColumn185.Width = 38;
			ultraGridColumn186.Header.VisiblePosition = 6;
			ultraGridColumn187.Header.VisiblePosition = 7;
			appearance166.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn188.CellAppearance = appearance166;
			ultraGridColumn188.Format = "#,##0.00";
			ultraGridColumn188.Header.VisiblePosition = 5;
			ultraGridColumn188.Width = 56;
			ultraGridColumn189.Header.VisiblePosition = 8;
			ultraGridBand21.Columns.AddRange(new object[] {
																											ultraGridColumn181,
																											ultraGridColumn182,
																											ultraGridColumn183,
																											ultraGridColumn184,
																											ultraGridColumn185,
																											ultraGridColumn186,
																											ultraGridColumn187,
																											ultraGridColumn188,
																											ultraGridColumn189});
			ultraGridBand21.Header.Caption = "Cayambe";
			ultraGridBand21.HeaderVisible = true;
			appearance167.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance167;
			summarySettings22.DisplayFormat = "{0: #,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand21.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings22});
			ultraGridBand21.SummaryFooterCaption = "";
			this.uGridCayambe.DisplayLayout.BandsSerializer.Add(ultraGridBand21);
			this.uGridCayambe.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCayambe.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance168.BackColor = System.Drawing.Color.Transparent;
			this.uGridCayambe.DisplayLayout.Override.CardAreaAppearance = appearance168;
			appearance169.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance169.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance169.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance169.FontData.BoldAsString = "True";
			appearance169.FontData.Name = "Arial";
			appearance169.FontData.SizeInPoints = 7.5F;
			appearance169.ForeColor = System.Drawing.Color.White;
			appearance169.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCayambe.DisplayLayout.Override.HeaderAppearance = appearance169;
			appearance170.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance170.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance170.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCayambe.DisplayLayout.Override.RowAlternateAppearance = appearance170;
			appearance171.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance171.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance171.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCayambe.DisplayLayout.Override.RowSelectorAppearance = appearance171;
			appearance172.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance172.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance172.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCayambe.DisplayLayout.Override.SelectedRowAppearance = appearance172;
			this.uGridCayambe.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCayambe.Location = new System.Drawing.Point(944, 632);
			this.uGridCayambe.Name = "uGridCayambe";
			this.uGridCayambe.Size = new System.Drawing.Size(310, 120);
			this.uGridCayambe.TabIndex = 49;
			this.uGridCayambe.DoubleClick += new System.EventHandler(this.uGridCayambe_DoubleClick);
			// 
			// uGridSanLuis
			// 
			this.uGridSanLuis.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSanLuis.DataSource = this.ultraDataSource1;
			appearance173.BackColor = System.Drawing.Color.White;
			this.uGridSanLuis.DisplayLayout.Appearance = appearance173;
			ultraGridColumn190.Header.VisiblePosition = 0;
			ultraGridColumn190.Hidden = true;
			ultraGridColumn190.Width = 41;
			ultraGridColumn191.Header.VisiblePosition = 1;
			ultraGridColumn191.Width = 178;
			ultraGridColumn192.Header.VisiblePosition = 2;
			ultraGridColumn192.Hidden = true;
			ultraGridColumn192.Width = 34;
			ultraGridColumn193.Header.VisiblePosition = 3;
			ultraGridColumn193.Hidden = true;
			ultraGridColumn193.Width = 53;
			ultraGridColumn194.Header.VisiblePosition = 4;
			ultraGridColumn194.Width = 39;
			ultraGridColumn195.Header.VisiblePosition = 6;
			ultraGridColumn196.Header.VisiblePosition = 7;
			appearance174.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn197.CellAppearance = appearance174;
			ultraGridColumn197.Format = "#,##0.00";
			ultraGridColumn197.Header.VisiblePosition = 5;
			ultraGridColumn197.Width = 55;
			ultraGridColumn198.Header.VisiblePosition = 8;
			ultraGridBand22.Columns.AddRange(new object[] {
																											ultraGridColumn190,
																											ultraGridColumn191,
																											ultraGridColumn192,
																											ultraGridColumn193,
																											ultraGridColumn194,
																											ultraGridColumn195,
																											ultraGridColumn196,
																											ultraGridColumn197,
																											ultraGridColumn198});
			ultraGridBand22.Header.Caption = "San Luis";
			ultraGridBand22.HeaderVisible = true;
			appearance175.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance175;
			summarySettings23.DisplayFormat = "{0: #,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand22.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings23});
			ultraGridBand22.SummaryFooterCaption = "";
			this.uGridSanLuis.DisplayLayout.BandsSerializer.Add(ultraGridBand22);
			this.uGridSanLuis.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSanLuis.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance176.BackColor = System.Drawing.Color.Transparent;
			this.uGridSanLuis.DisplayLayout.Override.CardAreaAppearance = appearance176;
			appearance177.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance177.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance177.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance177.FontData.BoldAsString = "True";
			appearance177.FontData.Name = "Arial";
			appearance177.FontData.SizeInPoints = 7.5F;
			appearance177.ForeColor = System.Drawing.Color.White;
			appearance177.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSanLuis.DisplayLayout.Override.HeaderAppearance = appearance177;
			appearance178.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance178.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance178.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSanLuis.DisplayLayout.Override.RowAlternateAppearance = appearance178;
			appearance179.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance179.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance179.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSanLuis.DisplayLayout.Override.RowSelectorAppearance = appearance179;
			appearance180.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance180.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance180.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSanLuis.DisplayLayout.Override.SelectedRowAppearance = appearance180;
			this.uGridSanLuis.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSanLuis.Location = new System.Drawing.Point(632, 512);
			this.uGridSanLuis.Name = "uGridSanLuis";
			this.uGridSanLuis.Size = new System.Drawing.Size(310, 120);
			this.uGridSanLuis.TabIndex = 48;
			this.uGridSanLuis.DoubleClick += new System.EventHandler(this.uGridSanLuis_DoubleClick);
			// 
			// uGridQuicentroSur
			// 
			this.uGridQuicentroSur.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridQuicentroSur.DataSource = this.ultraDataSource1;
			appearance181.BackColor = System.Drawing.Color.White;
			this.uGridQuicentroSur.DisplayLayout.Appearance = appearance181;
			ultraGridColumn199.Header.VisiblePosition = 0;
			ultraGridColumn199.Hidden = true;
			ultraGridColumn199.Width = 41;
			ultraGridColumn200.Header.VisiblePosition = 1;
			ultraGridColumn200.Width = 178;
			ultraGridColumn201.Header.VisiblePosition = 2;
			ultraGridColumn201.Hidden = true;
			ultraGridColumn201.Width = 34;
			ultraGridColumn202.Header.VisiblePosition = 3;
			ultraGridColumn202.Hidden = true;
			ultraGridColumn202.Width = 53;
			ultraGridColumn203.Header.VisiblePosition = 4;
			ultraGridColumn203.Width = 38;
			ultraGridColumn204.Header.VisiblePosition = 6;
			ultraGridColumn205.Header.VisiblePosition = 7;
			appearance182.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn206.CellAppearance = appearance182;
			ultraGridColumn206.Format = "#,##0.00";
			ultraGridColumn206.Header.VisiblePosition = 5;
			ultraGridColumn206.Width = 56;
			ultraGridColumn207.Header.VisiblePosition = 8;
			ultraGridBand23.Columns.AddRange(new object[] {
																											ultraGridColumn199,
																											ultraGridColumn200,
																											ultraGridColumn201,
																											ultraGridColumn202,
																											ultraGridColumn203,
																											ultraGridColumn204,
																											ultraGridColumn205,
																											ultraGridColumn206,
																											ultraGridColumn207});
			ultraGridBand23.Header.Caption = "Quicentro Sur";
			ultraGridBand23.HeaderVisible = true;
			appearance183.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance183;
			summarySettings24.DisplayFormat = "{0: #,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand23.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings24});
			ultraGridBand23.SummaryFooterCaption = "";
			this.uGridQuicentroSur.DisplayLayout.BandsSerializer.Add(ultraGridBand23);
			this.uGridQuicentroSur.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridQuicentroSur.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance184.BackColor = System.Drawing.Color.Transparent;
			this.uGridQuicentroSur.DisplayLayout.Override.CardAreaAppearance = appearance184;
			appearance185.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance185.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance185.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance185.FontData.BoldAsString = "True";
			appearance185.FontData.Name = "Arial";
			appearance185.FontData.SizeInPoints = 7.5F;
			appearance185.ForeColor = System.Drawing.Color.White;
			appearance185.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridQuicentroSur.DisplayLayout.Override.HeaderAppearance = appearance185;
			appearance186.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance186.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance186.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridQuicentroSur.DisplayLayout.Override.RowAlternateAppearance = appearance186;
			appearance187.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance187.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance187.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridQuicentroSur.DisplayLayout.Override.RowSelectorAppearance = appearance187;
			appearance188.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance188.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance188.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridQuicentroSur.DisplayLayout.Override.SelectedRowAppearance = appearance188;
			this.uGridQuicentroSur.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridQuicentroSur.Location = new System.Drawing.Point(320, 392);
			this.uGridQuicentroSur.Name = "uGridQuicentroSur";
			this.uGridQuicentroSur.Size = new System.Drawing.Size(310, 120);
			this.uGridQuicentroSur.TabIndex = 47;
			this.uGridQuicentroSur.DoubleClick += new System.EventHandler(this.uGridQuicentroSur_DoubleClick);
			// 
			// uGridIbarra
			// 
			this.uGridIbarra.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridIbarra.DataSource = this.ultraDataSource1;
			appearance189.BackColor = System.Drawing.Color.White;
			this.uGridIbarra.DisplayLayout.Appearance = appearance189;
			ultraGridColumn208.Header.VisiblePosition = 0;
			ultraGridColumn208.Hidden = true;
			ultraGridColumn208.Width = 41;
			ultraGridColumn209.Header.VisiblePosition = 1;
			ultraGridColumn209.Width = 178;
			ultraGridColumn210.Header.VisiblePosition = 2;
			ultraGridColumn210.Hidden = true;
			ultraGridColumn210.Width = 34;
			ultraGridColumn211.Header.VisiblePosition = 3;
			ultraGridColumn211.Hidden = true;
			ultraGridColumn211.Width = 53;
			ultraGridColumn212.Header.VisiblePosition = 4;
			ultraGridColumn212.Width = 39;
			ultraGridColumn213.Header.VisiblePosition = 6;
			ultraGridColumn214.Header.VisiblePosition = 7;
			appearance190.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn215.CellAppearance = appearance190;
			ultraGridColumn215.Format = "#,##0.00";
			ultraGridColumn215.Header.VisiblePosition = 5;
			ultraGridColumn215.Width = 55;
			ultraGridColumn216.Header.VisiblePosition = 8;
			ultraGridBand24.Columns.AddRange(new object[] {
																											ultraGridColumn208,
																											ultraGridColumn209,
																											ultraGridColumn210,
																											ultraGridColumn211,
																											ultraGridColumn212,
																											ultraGridColumn213,
																											ultraGridColumn214,
																											ultraGridColumn215,
																											ultraGridColumn216});
			ultraGridBand24.Header.Caption = "Ibarra";
			ultraGridBand24.HeaderVisible = true;
			appearance191.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance191;
			summarySettings25.DisplayFormat = "{0: #,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand24.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings25});
			ultraGridBand24.SummaryFooterCaption = "";
			this.uGridIbarra.DisplayLayout.BandsSerializer.Add(ultraGridBand24);
			this.uGridIbarra.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridIbarra.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance192.BackColor = System.Drawing.Color.Transparent;
			this.uGridIbarra.DisplayLayout.Override.CardAreaAppearance = appearance192;
			appearance193.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance193.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance193.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance193.FontData.BoldAsString = "True";
			appearance193.FontData.Name = "Arial";
			appearance193.FontData.SizeInPoints = 7.5F;
			appearance193.ForeColor = System.Drawing.Color.White;
			appearance193.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridIbarra.DisplayLayout.Override.HeaderAppearance = appearance193;
			appearance194.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance194.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance194.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIbarra.DisplayLayout.Override.RowAlternateAppearance = appearance194;
			appearance195.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance195.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance195.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIbarra.DisplayLayout.Override.RowSelectorAppearance = appearance195;
			appearance196.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance196.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance196.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIbarra.DisplayLayout.Override.SelectedRowAppearance = appearance196;
			this.uGridIbarra.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridIbarra.Location = new System.Drawing.Point(8, 272);
			this.uGridIbarra.Name = "uGridIbarra";
			this.uGridIbarra.Size = new System.Drawing.Size(310, 120);
			this.uGridIbarra.TabIndex = 46;
			this.uGridIbarra.DoubleClick += new System.EventHandler(this.uGridIbarra_DoubleClick);
			// 
			// uGridChillogallo
			// 
			this.uGridChillogallo.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridChillogallo.DataSource = this.ultraDataSource1;
			appearance197.BackColor = System.Drawing.Color.White;
			this.uGridChillogallo.DisplayLayout.Appearance = appearance197;
			ultraGridColumn217.Header.VisiblePosition = 0;
			ultraGridColumn217.Hidden = true;
			ultraGridColumn217.Width = 41;
			ultraGridColumn218.Header.VisiblePosition = 1;
			ultraGridColumn218.Width = 178;
			ultraGridColumn219.Header.VisiblePosition = 2;
			ultraGridColumn219.Hidden = true;
			ultraGridColumn219.Width = 34;
			ultraGridColumn220.Header.VisiblePosition = 3;
			ultraGridColumn220.Hidden = true;
			ultraGridColumn220.Width = 53;
			ultraGridColumn221.Header.VisiblePosition = 4;
			ultraGridColumn221.Width = 39;
			ultraGridColumn222.Header.VisiblePosition = 6;
			ultraGridColumn223.Header.VisiblePosition = 7;
			appearance198.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn224.CellAppearance = appearance198;
			ultraGridColumn224.Format = "#,##0.00";
			ultraGridColumn224.Header.VisiblePosition = 5;
			ultraGridColumn224.Width = 55;
			ultraGridColumn225.Header.VisiblePosition = 8;
			ultraGridBand25.Columns.AddRange(new object[] {
																											ultraGridColumn217,
																											ultraGridColumn218,
																											ultraGridColumn219,
																											ultraGridColumn220,
																											ultraGridColumn221,
																											ultraGridColumn222,
																											ultraGridColumn223,
																											ultraGridColumn224,
																											ultraGridColumn225});
			ultraGridBand25.Header.Caption = "Chillogallo";
			ultraGridBand25.HeaderVisible = true;
			appearance199.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance199;
			summarySettings26.DisplayFormat = "{0: #,##0.00}";
			summarySettings26.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand25.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings26});
			ultraGridBand25.SummaryFooterCaption = "";
			this.uGridChillogallo.DisplayLayout.BandsSerializer.Add(ultraGridBand25);
			this.uGridChillogallo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridChillogallo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance200.BackColor = System.Drawing.Color.Transparent;
			this.uGridChillogallo.DisplayLayout.Override.CardAreaAppearance = appearance200;
			appearance201.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance201.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance201.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance201.FontData.BoldAsString = "True";
			appearance201.FontData.Name = "Arial";
			appearance201.FontData.SizeInPoints = 7.5F;
			appearance201.ForeColor = System.Drawing.Color.White;
			appearance201.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridChillogallo.DisplayLayout.Override.HeaderAppearance = appearance201;
			appearance202.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance202.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance202.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridChillogallo.DisplayLayout.Override.RowAlternateAppearance = appearance202;
			appearance203.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance203.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance203.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridChillogallo.DisplayLayout.Override.RowSelectorAppearance = appearance203;
			appearance204.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance204.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance204.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridChillogallo.DisplayLayout.Override.SelectedRowAppearance = appearance204;
			this.uGridChillogallo.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridChillogallo.Location = new System.Drawing.Point(1253, 32);
			this.uGridChillogallo.Name = "uGridChillogallo";
			this.uGridChillogallo.Size = new System.Drawing.Size(310, 120);
			this.uGridChillogallo.TabIndex = 45;
			this.uGridChillogallo.DoubleClick += new System.EventHandler(this.uGridChillogallo_DoubleClick);
			// 
			// dtFecha
			// 
			appearance205.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance205;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 6);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 168;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 17);
			this.label6.TabIndex = 169;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(176, 5);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(77, 23);
			this.btnBuscar.TabIndex = 298;
			this.btnBuscar.Text = "&Ver";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// frmCom_PromedioVendedoresPorLocal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1370, 746);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.uGridCayambe);
			this.Controls.Add(this.uGridSanLuis);
			this.Controls.Add(this.uGridQuicentroSur);
			this.Controls.Add(this.uGridIbarra);
			this.Controls.Add(this.uGridChillogallo);
			this.Controls.Add(this.uGridTumbaco);
			this.Controls.Add(this.uGridSolanda);
			this.Controls.Add(this.uGridRiobamba);
			this.Controls.Add(this.uGridRecreo3);
			this.Controls.Add(this.uGridPrensa);
			this.Controls.Add(this.uGridMachachi);
			this.Controls.Add(this.uGridGuamani);
			this.Controls.Add(this.uGridConsumidorFinal);
			this.Controls.Add(this.uGridCarapungo);
			this.Controls.Add(this.uGridCalderon);
			this.Controls.Add(this.uGridSantoDomingo);
			this.Controls.Add(this.uGridSangolqui);
			this.Controls.Add(this.uGridRecreo1);
			this.Controls.Add(this.uGridQuinche);
			this.Controls.Add(this.uGridLatacunga);
			this.Controls.Add(this.uGridInca);
			this.Controls.Add(this.uGridCondado);
			this.Controls.Add(this.uGridComiteDelPueblo);
			this.Controls.Add(this.uGridAmbato2);
			this.Controls.Add(this.uGridAmbato);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCom_PromedioVendedoresPorLocal";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Promedio Vendedores PorLocal";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCom_PromedioVendedoresPorLocal_KeyDown);
			this.Load += new System.EventHandler(this.frmCom_PromedioVendedoresPorLocal_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridAmbato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridAmbato2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCondado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridComiteDelPueblo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRecreo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridQuinche)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLatacunga)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSantoDomingo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSangolqui)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTumbaco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSolanda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRiobamba)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRecreo3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPrensa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMachachi)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGuamani)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridConsumidorFinal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCarapungo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCalderon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCayambe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSanLuis)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridQuicentroSur)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridIbarra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridChillogallo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCom_PromedioVendedoresPorLocal_Load(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Parse(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
			this.AutoScroll = true;
		}

		private void AbreInfoVentas(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			frmReporteVendedoresPorFecha RVF = new frmReporteVendedoresPorFecha(uGrid.ActiveRow.Cells["Vendedor"].Value.ToString(), 0);
      RVF.ShowDialog();																																						 
		}

		private void frmCom_PromedioVendedoresPorLocal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void SumatoriaPromedios(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			uGrid.DisplayLayout.Bands[0].Summaries.Add("Promedio", Infragistics.Win.UltraWinGrid.SummaryType.Sum, uGrid.DisplayLayout.Bands[0].Columns["Promedio"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
			uGrid.DisplayLayout.Bands[0].Summaries["Promedio"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
			uGrid.DisplayLayout.Bands[0].Summaries["Promedio"].DisplayFormat = "{0: #,##0.00}";
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = Convert.ToDateTime(this.dtFecha.Value);
						
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 43), this.uGridAmbato);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 74), this.uGridAmbato2);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 11), this.uGridCalderon);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 63), this.uGridCarapungo);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 70), this.uGridChillogallo);

			this.SumatoriaPromedios(this.uGridAmbato2);
			this.SumatoriaPromedios(this.uGridCalderon);
			this.SumatoriaPromedios(this.uGridCarapungo);
			this.SumatoriaPromedios(this.uGridChillogallo);

			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 66), this.uGridComiteDelPueblo);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 16), this.uGridCondado);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 12), this.uGridConsumidorFinal);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 6), this.uGridGuamani);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 25), this.uGridIbarra);

			this.SumatoriaPromedios(this.uGridComiteDelPueblo);
			this.SumatoriaPromedios(this.uGridCondado);
			this.SumatoriaPromedios(this.uGridConsumidorFinal);
			this.SumatoriaPromedios(this.uGridGuamani);
			this.SumatoriaPromedios(this.uGridIbarra);

			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 71), this.uGridInca);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 60), this.uGridLatacunga);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 76), this.uGridMachachi);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 73), this.uGridPrensa);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 29), this.uGridQuicentroSur);

			this.SumatoriaPromedios(this.uGridInca);
			this.SumatoriaPromedios(this.uGridLatacunga);
			this.SumatoriaPromedios(this.uGridMachachi);
			this.SumatoriaPromedios(this.uGridPrensa);
			this.SumatoriaPromedios(this.uGridQuicentroSur);

			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 75), this.uGridQuinche);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 31), this.uGridRecreo1);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 67), this.uGridRecreo3);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 64), this.uGridRiobamba);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 38), this.uGridSanLuis);

			this.SumatoriaPromedios(this.uGridQuinche);
			this.SumatoriaPromedios(this.uGridRecreo1);
			this.SumatoriaPromedios(this.uGridRecreo3);
			this.SumatoriaPromedios(this.uGridRiobamba);
			this.SumatoriaPromedios(this.uGridSanLuis);

			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 37), this.uGridSangolqui);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 72), this.uGridSantoDomingo);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 65), this.uGridSolanda);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 39), this.uGridTumbaco);
			FuncionesProcedimientos.dsGeneral(string.Format("Exec Com_ListaVendedoresLocales '{0}', {1}", dtFecha.ToString("yyyyMMdd"), 77), this.uGridCayambe);

			this.SumatoriaPromedios(this.uGridSangolqui);
			this.SumatoriaPromedios(this.uGridSantoDomingo);
			this.SumatoriaPromedios(this.uGridSolanda);
			this.SumatoriaPromedios(this.uGridTumbaco);
			this.SumatoriaPromedios(this.uGridCayambe);
		}

		private void uGridAmbato_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridAmbato);
		}

		private void uGridAmbato2_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridAmbato2);
		}

		private void uGridCalderon_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridCalderon);
		}

		private void uGridCarapungo_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridCarapungo);
		}

		private void uGridChillogallo_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridChillogallo);
		}

		private void uGridComiteDelPueblo_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridComiteDelPueblo);
		}

		private void uGridCondado_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridCondado);
		}

		private void uGridConsumidorFinal_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridConsumidorFinal);
		}

		private void uGridGuamani_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridGuamani);
		}

		private void uGridIbarra_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridIbarra);
		}

		private void uGridInca_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridInca);
		}

		private void uGridLatacunga_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridLatacunga);
		}

		private void uGridMachachi_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridMachachi);
		}

		private void uGridPrensa_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridPrensa);
		}

		private void uGridQuicentroSur_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridQuicentroSur);
		}

		private void uGridQuinche_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridQuinche);
		}

		private void uGridRecreo1_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridRecreo1);
		}

		private void uGridRecreo3_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridRecreo3);
		}

		private void uGridRiobamba_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridRiobamba);
		}

		private void uGridSanLuis_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridSanLuis);
		}

		private void uGridSangolqui_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridSangolqui);
		}

		private void uGridSantoDomingo_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridSantoDomingo);
		}

		private void uGridSolanda_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridSolanda);
		}

		private void uGridTumbaco_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridTumbaco);
		}

		private void uGridCayambe_DoubleClick(object sender, System.EventArgs e)
		{
			AbreInfoVentas(this.uGridCayambe);
		}

		private void uGridAmbato_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
