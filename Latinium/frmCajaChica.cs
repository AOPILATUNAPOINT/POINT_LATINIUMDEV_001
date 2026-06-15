using System;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using LibContabilidad.DataObjects;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCajaChica.
	/// </summary>
	public class frmCajaChica : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private System.Windows.Forms.Label lblComprobante;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblSubTotal;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCajaChica;
		private C1.Data.C1DataSet cdsCajaChica;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataSet cdsSeteoC;
		private C1.Data.C1DataSet cdsSeteoF;

		#region Variables
		C1.Data.C1DataRow drCajaChica;
		private C1DataRow drSeteoF;
		private C1DataRow drSeteoC;
		bool bEliminaAlValidar = false;
		bool bEmpresa = false;
		private Acceso miAcceso;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblEstado;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCajaChica;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbArticulo2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbArticulo1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtResponsable;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCentroDeCosto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRetencionesRenta1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRetencionesIVA1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRetencionesIVA2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRetencionesRenta2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private System.Windows.Forms.Button btnActualizar;
		private System.Windows.Forms.Button btnCalcular;
		private System.Windows.Forms.Button btnEgreso;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Button btnProcesar;
		private System.Windows.Forms.Button btnAutorizar;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidProyecto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodega;
		CultureInfo us = new CultureInfo("en-US");
		#endregion Variables

		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPersonal;

		int iBodegaPredef = 0;
		public frmCajaChica()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idCajaChica = 0;
		int idBloqueaTransacciones = 0;

		public frmCajaChica(int IdCajaChica)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idCajaChica = IdCajaChica;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCajaChica));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_DetalleCajaChica", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleCajaChica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCajaChica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RUC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RazonSocial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorización");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo1", -1, "cmbArticulo1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA1");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal1");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo2", -1, "cmbArticulo2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA2");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal2");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbCentroDeCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SF");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizacionRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Imprime");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Suma");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI1", -1, "cmbRetencionesRenta1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValRetRenta1");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIVA1", -1, "cmbRetencionesIVA1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValRetIVA1");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI2", -1, "cmbRetencionesRenta2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValRetRenta2");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIVA2", -1, "cmbRetencionesIVA2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValRetIVA2");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetenidoRenta");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetenidoIVA");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetencionAsumida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbPersonal");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 18, true, "_DetalleCajaChica", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 18, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retenido", 27, true, "_DetalleCajaChica", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retenido", 27, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Suma", 32, true, "_DetalleCajaChica", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Suma", 32, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetenidoRenta", 41, true, "_DetalleCajaChica", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetenidoRenta", 41, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetenidoIVA", 42, true, "_DetalleCajaChica", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetenidoIVA", 42, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloSRI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIVA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIVA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloSRI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCajaChica = new C1.Data.C1DataSet();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtIdCajaChica = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cmbArticulo2 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.txtCajaChica = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbArticulo1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.txtResponsable = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCentroDeCosto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.cmbRetencionesRenta1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbRetencionesIVA1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbRetencionesIVA2 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbRetencionesRenta2 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.btnCalcular = new System.Windows.Forms.Button();
			this.btnEgreso = new System.Windows.Forms.Button();
			this.label5 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.btnProcesar = new System.Windows.Forms.Button();
			this.btnAutorizar = new System.Windows.Forms.Button();
			this.label8 = new System.Windows.Forms.Label();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtidProyecto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCajaChica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCajaChica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCajaChica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetencionesRenta1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetencionesIVA1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetencionesIVA2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetencionesRenta2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.AccessibleDescription = resources.GetString("ultraGrid1.AccessibleDescription");
			this.ultraGrid1.AccessibleName = resources.GetString("ultraGrid1.AccessibleName");
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("ultraGrid1.Anchor")));
			this.ultraGrid1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("ultraGrid1.BackgroundImage")));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_CajaChica.CajaChica - DetalleCajaChica";
			this.ultraGrid1.DataSource = this.cdsCajaChica;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = resources.GetString("ultraGridBand1.AddButtonCaption");
			ultraGridBand1.AddButtonToolTipText = resources.GetString("ultraGridBand1.AddButtonToolTipText");
			ultraGridColumn1.Format = resources.GetString("ultraGridColumn1.Format");
			ultraGridColumn1.Header.Caption = resources.GetString("resource.Caption");
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.MaskInput = resources.GetString("ultraGridColumn1.MaskInput");
			ultraGridColumn1.NullText = resources.GetString("ultraGridColumn1.NullText");
			ultraGridColumn2.Format = resources.GetString("ultraGridColumn2.Format");
			ultraGridColumn2.Header.Caption = resources.GetString("resource.Caption1");
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.MaskInput = resources.GetString("ultraGridColumn2.MaskInput");
			ultraGridColumn2.NullText = resources.GetString("ultraGridColumn2.NullText");
			ultraGridColumn3.Format = resources.GetString("ultraGridColumn3.Format");
			ultraGridColumn3.Header.Caption = resources.GetString("resource.Caption2");
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.MaskInput = resources.GetString("ultraGridColumn3.MaskInput");
			ultraGridColumn3.NullText = resources.GetString("ultraGridColumn3.NullText");
			ultraGridColumn4.Format = resources.GetString("ultraGridColumn4.Format");
			ultraGridColumn4.Header.Caption = resources.GetString("resource.Caption3");
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.MaskInput = resources.GetString("ultraGridColumn4.MaskInput");
			ultraGridColumn4.NullText = resources.GetString("ultraGridColumn4.NullText");
			ultraGridColumn5.FieldLen = 13;
			ultraGridColumn5.Format = resources.GetString("ultraGridColumn5.Format");
			ultraGridColumn5.Header.Caption = resources.GetString("resource.Caption4");
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.MaskInput = resources.GetString("ultraGridColumn5.MaskInput");
			ultraGridColumn5.NullText = resources.GetString("ultraGridColumn5.NullText");
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Format = resources.GetString("ultraGridColumn6.Format");
			ultraGridColumn6.Header.Caption = resources.GetString("resource.Caption5");
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.MaskInput = resources.GetString("ultraGridColumn6.MaskInput");
			ultraGridColumn6.NullText = resources.GetString("ultraGridColumn6.NullText");
			ultraGridColumn6.Width = 180;
			ultraGridColumn7.Format = resources.GetString("ultraGridColumn7.Format");
			ultraGridColumn7.Header.Caption = resources.GetString("resource.Caption6");
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.MaskInput = resources.GetString("ultraGridColumn7.MaskInput");
			ultraGridColumn7.NullText = resources.GetString("ultraGridColumn7.NullText");
			ultraGridColumn7.Width = 62;
			ultraGridColumn8.FieldLen = 6;
			ultraGridColumn8.Format = resources.GetString("ultraGridColumn8.Format");
			ultraGridColumn8.Header.Caption = resources.GetString("resource.Caption7");
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn8.MaskInput = resources.GetString("ultraGridColumn8.MaskInput");
			ultraGridColumn8.NullText = resources.GetString("ultraGridColumn8.NullText");
			ultraGridColumn8.Width = 60;
			ultraGridColumn9.FieldLen = 9;
			ultraGridColumn9.Format = resources.GetString("ultraGridColumn9.Format");
			ultraGridColumn9.Header.Caption = resources.GetString("resource.Caption8");
			ultraGridColumn9.Header.VisiblePosition = 11;
			ultraGridColumn9.MaskInput = resources.GetString("ultraGridColumn9.MaskInput");
			ultraGridColumn9.NullText = resources.GetString("ultraGridColumn9.NullText");
			ultraGridColumn9.Width = 100;
			ultraGridColumn10.FieldLen = 37;
			ultraGridColumn10.Format = resources.GetString("ultraGridColumn10.Format");
			ultraGridColumn10.Header.Caption = resources.GetString("resource.Caption9");
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.MaskInput = resources.GetString("ultraGridColumn10.MaskInput");
			ultraGridColumn10.NullText = resources.GetString("ultraGridColumn10.NullText");
			ultraGridColumn10.Width = 73;
			ultraGridColumn11.Format = resources.GetString("ultraGridColumn11.Format");
			ultraGridColumn11.Header.Caption = resources.GetString("resource.Caption10");
			ultraGridColumn11.Header.VisiblePosition = 15;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.MaskInput = resources.GetString("ultraGridColumn11.MaskInput");
			ultraGridColumn11.NullText = resources.GetString("ultraGridColumn11.NullText");
			ultraGridColumn12.Format = resources.GetString("ultraGridColumn12.Format");
			ultraGridColumn12.Header.Caption = resources.GetString("resource.Caption11");
			ultraGridColumn12.Header.VisiblePosition = 16;
			ultraGridColumn12.MaskInput = resources.GetString("ultraGridColumn12.MaskInput");
			ultraGridColumn12.NullText = resources.GetString("ultraGridColumn12.NullText");
			ultraGridColumn12.Width = 120;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance2;
			ultraGridColumn13.Format = resources.GetString("ultraGridColumn13.Format");
			ultraGridColumn13.Header.Caption = resources.GetString("resource.Caption12");
			ultraGridColumn13.Header.VisiblePosition = 17;
			ultraGridColumn13.MaskInput = resources.GetString("ultraGridColumn13.MaskInput");
			ultraGridColumn13.NullText = resources.GetString("ultraGridColumn13.NullText");
			ultraGridColumn13.PromptChar = ' ';
			ultraGridColumn13.Width = 40;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance3;
			ultraGridColumn14.Format = resources.GetString("ultraGridColumn14.Format");
			ultraGridColumn14.Header.Caption = resources.GetString("resource.Caption13");
			ultraGridColumn14.Header.VisiblePosition = 18;
			ultraGridColumn14.MaskInput = resources.GetString("ultraGridColumn14.MaskInput");
			ultraGridColumn14.NullText = resources.GetString("ultraGridColumn14.NullText");
			ultraGridColumn14.PromptChar = ' ';
			ultraGridColumn14.Width = 70;
			ultraGridColumn15.Format = resources.GetString("ultraGridColumn15.Format");
			ultraGridColumn15.Header.Caption = resources.GetString("resource.Caption14");
			ultraGridColumn15.Header.VisiblePosition = 19;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.MaskInput = resources.GetString("ultraGridColumn15.MaskInput");
			ultraGridColumn15.NullText = resources.GetString("ultraGridColumn15.NullText");
			ultraGridColumn16.Format = resources.GetString("ultraGridColumn16.Format");
			ultraGridColumn16.Header.Caption = resources.GetString("resource.Caption15");
			ultraGridColumn16.Header.VisiblePosition = 20;
			ultraGridColumn16.MaskInput = resources.GetString("ultraGridColumn16.MaskInput");
			ultraGridColumn16.NullText = resources.GetString("ultraGridColumn16.NullText");
			ultraGridColumn16.Width = 120;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance4;
			ultraGridColumn17.Format = resources.GetString("ultraGridColumn17.Format");
			ultraGridColumn17.Header.Caption = resources.GetString("resource.Caption16");
			ultraGridColumn17.Header.VisiblePosition = 21;
			ultraGridColumn17.MaskInput = resources.GetString("ultraGridColumn17.MaskInput");
			ultraGridColumn17.NullText = resources.GetString("ultraGridColumn17.NullText");
			ultraGridColumn17.PromptChar = ' ';
			ultraGridColumn17.Width = 40;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance5;
			ultraGridColumn18.Format = resources.GetString("ultraGridColumn18.Format");
			ultraGridColumn18.Header.Caption = resources.GetString("resource.Caption17");
			ultraGridColumn18.Header.VisiblePosition = 22;
			ultraGridColumn18.MaskInput = resources.GetString("ultraGridColumn18.MaskInput");
			ultraGridColumn18.NullText = resources.GetString("ultraGridColumn18.NullText");
			ultraGridColumn18.PromptChar = ' ';
			ultraGridColumn18.Width = 70;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance6;
			ultraGridColumn19.Format = resources.GetString("ultraGridColumn19.Format");
			ultraGridColumn19.Header.Caption = resources.GetString("resource.Caption18");
			ultraGridColumn19.Header.VisiblePosition = 23;
			ultraGridColumn19.MaskInput = resources.GetString("ultraGridColumn19.MaskInput");
			ultraGridColumn19.NullText = resources.GetString("ultraGridColumn19.NullText");
			ultraGridColumn19.PromptChar = ' ';
			ultraGridColumn19.Width = 70;
			ultraGridColumn20.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn20.Format = resources.GetString("ultraGridColumn20.Format");
			ultraGridColumn20.Header.Caption = resources.GetString("resource.Caption19");
			ultraGridColumn20.Header.VisiblePosition = 12;
			ultraGridColumn20.MaskInput = resources.GetString("ultraGridColumn20.MaskInput");
			ultraGridColumn20.NullText = resources.GetString("ultraGridColumn20.NullText");
			ultraGridColumn20.Width = 219;
			ultraGridColumn21.Format = resources.GetString("ultraGridColumn21.Format");
			ultraGridColumn21.Header.Caption = resources.GetString("resource.Caption20");
			ultraGridColumn21.Header.VisiblePosition = 13;
			ultraGridColumn21.MaskInput = resources.GetString("ultraGridColumn21.MaskInput");
			ultraGridColumn21.NullText = resources.GetString("ultraGridColumn21.NullText");
			ultraGridColumn21.Width = 140;
			ultraGridColumn22.Format = resources.GetString("ultraGridColumn22.Format");
			ultraGridColumn22.Header.Caption = resources.GetString("resource.Caption21");
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.MaskInput = resources.GetString("ultraGridColumn22.MaskInput");
			ultraGridColumn22.NullText = resources.GetString("ultraGridColumn22.NullText");
			ultraGridColumn22.Width = 30;
			ultraGridColumn23.Format = resources.GetString("ultraGridColumn23.Format");
			ultraGridColumn23.Header.Caption = resources.GetString("resource.Caption22");
			ultraGridColumn23.Header.VisiblePosition = 24;
			ultraGridColumn23.MaskInput = resources.GetString("ultraGridColumn23.MaskInput");
			ultraGridColumn23.NullText = resources.GetString("ultraGridColumn23.NullText");
			ultraGridColumn24.Format = resources.GetString("ultraGridColumn24.Format");
			ultraGridColumn24.Header.Caption = resources.GetString("resource.Caption23");
			ultraGridColumn24.Header.VisiblePosition = 25;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.MaskInput = resources.GetString("ultraGridColumn24.MaskInput");
			ultraGridColumn24.NullText = resources.GetString("ultraGridColumn24.NullText");
			ultraGridColumn25.Format = resources.GetString("ultraGridColumn25.Format");
			ultraGridColumn25.Header.Caption = resources.GetString("resource.Caption24");
			ultraGridColumn25.Header.VisiblePosition = 28;
			ultraGridColumn25.MaskInput = resources.GetString("ultraGridColumn25.MaskInput");
			ultraGridColumn25.NullText = resources.GetString("ultraGridColumn25.NullText");
			ultraGridColumn25.Width = 120;
			ultraGridColumn26.Format = resources.GetString("ultraGridColumn26.Format");
			ultraGridColumn26.Header.Caption = resources.GetString("resource.Caption25");
			ultraGridColumn26.Header.VisiblePosition = 27;
			ultraGridColumn26.MaskInput = resources.GetString("ultraGridColumn26.MaskInput");
			ultraGridColumn26.NullText = resources.GetString("ultraGridColumn26.NullText");
			ultraGridColumn26.Width = 60;
			ultraGridColumn27.Format = resources.GetString("ultraGridColumn27.Format");
			ultraGridColumn27.Header.Caption = resources.GetString("resource.Caption26");
			ultraGridColumn27.Header.VisiblePosition = 29;
			ultraGridColumn27.MaskInput = resources.GetString("ultraGridColumn27.MaskInput");
			ultraGridColumn27.NullText = resources.GetString("ultraGridColumn27.NullText");
			ultraGridColumn27.Width = 100;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance7;
			ultraGridColumn28.Format = resources.GetString("ultraGridColumn28.Format");
			ultraGridColumn28.Header.Caption = resources.GetString("resource.Caption27");
			ultraGridColumn28.Header.VisiblePosition = 40;
			ultraGridColumn28.MaskInput = resources.GetString("ultraGridColumn28.MaskInput");
			ultraGridColumn28.NullText = resources.GetString("ultraGridColumn28.NullText");
			ultraGridColumn28.PromptChar = ' ';
			ultraGridColumn29.Format = resources.GetString("ultraGridColumn29.Format");
			ultraGridColumn29.Header.Caption = resources.GetString("resource.Caption28");
			ultraGridColumn29.Header.VisiblePosition = 41;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.MaskInput = resources.GetString("ultraGridColumn29.MaskInput");
			ultraGridColumn29.NullText = resources.GetString("ultraGridColumn29.NullText");
			ultraGridColumn30.Format = resources.GetString("ultraGridColumn30.Format");
			ultraGridColumn30.Header.Caption = resources.GetString("resource.Caption29");
			ultraGridColumn30.Header.VisiblePosition = 42;
			ultraGridColumn30.MaskInput = resources.GetString("ultraGridColumn30.MaskInput");
			ultraGridColumn30.NullText = resources.GetString("ultraGridColumn30.NullText");
			ultraGridColumn31.Format = resources.GetString("ultraGridColumn31.Format");
			ultraGridColumn31.Header.Caption = resources.GetString("resource.Caption30");
			ultraGridColumn31.Header.VisiblePosition = 44;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.MaskInput = resources.GetString("ultraGridColumn31.MaskInput");
			ultraGridColumn31.NullText = resources.GetString("ultraGridColumn31.NullText");
			ultraGridColumn32.Format = resources.GetString("ultraGridColumn32.Format");
			ultraGridColumn32.Header.Caption = resources.GetString("resource.Caption31");
			ultraGridColumn32.Header.VisiblePosition = 43;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.MaskInput = resources.GetString("ultraGridColumn32.MaskInput");
			ultraGridColumn32.NullText = resources.GetString("ultraGridColumn32.NullText");
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance8;
			ultraGridColumn33.Format = resources.GetString("ultraGridColumn33.Format");
			ultraGridColumn33.Header.Caption = resources.GetString("resource.Caption32");
			ultraGridColumn33.Header.VisiblePosition = 45;
			ultraGridColumn33.MaskInput = resources.GetString("ultraGridColumn33.MaskInput");
			ultraGridColumn33.NullText = resources.GetString("ultraGridColumn33.NullText");
			ultraGridColumn33.Width = 110;
			ultraGridColumn34.Format = resources.GetString("ultraGridColumn34.Format");
			ultraGridColumn34.Header.Caption = resources.GetString("resource.Caption33");
			ultraGridColumn34.Header.VisiblePosition = 30;
			ultraGridColumn34.MaskInput = resources.GetString("ultraGridColumn34.MaskInput");
			ultraGridColumn34.NullText = resources.GetString("ultraGridColumn34.NullText");
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance9;
			ultraGridColumn35.Format = resources.GetString("ultraGridColumn35.Format");
			ultraGridColumn35.Header.Caption = resources.GetString("resource.Caption34");
			ultraGridColumn35.Header.VisiblePosition = 31;
			ultraGridColumn35.MaskInput = resources.GetString("ultraGridColumn35.MaskInput");
			ultraGridColumn35.NullText = resources.GetString("ultraGridColumn35.NullText");
			ultraGridColumn36.Format = resources.GetString("ultraGridColumn36.Format");
			ultraGridColumn36.Header.Caption = resources.GetString("resource.Caption35");
			ultraGridColumn36.Header.VisiblePosition = 32;
			ultraGridColumn36.MaskInput = resources.GetString("ultraGridColumn36.MaskInput");
			ultraGridColumn36.NullText = resources.GetString("ultraGridColumn36.NullText");
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance10;
			ultraGridColumn37.Format = resources.GetString("ultraGridColumn37.Format");
			ultraGridColumn37.Header.Caption = resources.GetString("resource.Caption36");
			ultraGridColumn37.Header.VisiblePosition = 33;
			ultraGridColumn37.MaskInput = resources.GetString("ultraGridColumn37.MaskInput");
			ultraGridColumn37.NullText = resources.GetString("ultraGridColumn37.NullText");
			ultraGridColumn38.Format = resources.GetString("ultraGridColumn38.Format");
			ultraGridColumn38.Header.Caption = resources.GetString("resource.Caption37");
			ultraGridColumn38.Header.VisiblePosition = 34;
			ultraGridColumn38.MaskInput = resources.GetString("ultraGridColumn38.MaskInput");
			ultraGridColumn38.NullText = resources.GetString("ultraGridColumn38.NullText");
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance11;
			ultraGridColumn39.Format = resources.GetString("ultraGridColumn39.Format");
			ultraGridColumn39.Header.Caption = resources.GetString("resource.Caption38");
			ultraGridColumn39.Header.VisiblePosition = 35;
			ultraGridColumn39.MaskInput = resources.GetString("ultraGridColumn39.MaskInput");
			ultraGridColumn39.NullText = resources.GetString("ultraGridColumn39.NullText");
			ultraGridColumn40.Format = resources.GetString("ultraGridColumn40.Format");
			ultraGridColumn40.Header.Caption = resources.GetString("resource.Caption39");
			ultraGridColumn40.Header.VisiblePosition = 36;
			ultraGridColumn40.MaskInput = resources.GetString("ultraGridColumn40.MaskInput");
			ultraGridColumn40.NullText = resources.GetString("ultraGridColumn40.NullText");
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance12;
			ultraGridColumn41.Format = resources.GetString("ultraGridColumn41.Format");
			ultraGridColumn41.Header.Caption = resources.GetString("resource.Caption40");
			ultraGridColumn41.Header.VisiblePosition = 37;
			ultraGridColumn41.MaskInput = resources.GetString("ultraGridColumn41.MaskInput");
			ultraGridColumn41.NullText = resources.GetString("ultraGridColumn41.NullText");
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance13;
			ultraGridColumn42.Format = resources.GetString("ultraGridColumn42.Format");
			ultraGridColumn42.Header.Caption = resources.GetString("resource.Caption41");
			ultraGridColumn42.Header.VisiblePosition = 38;
			ultraGridColumn42.MaskInput = resources.GetString("ultraGridColumn42.MaskInput");
			ultraGridColumn42.NullText = resources.GetString("ultraGridColumn42.NullText");
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance14;
			ultraGridColumn43.Format = resources.GetString("ultraGridColumn43.Format");
			ultraGridColumn43.Header.Caption = resources.GetString("resource.Caption42");
			ultraGridColumn43.Header.VisiblePosition = 39;
			ultraGridColumn43.MaskInput = resources.GetString("ultraGridColumn43.MaskInput");
			ultraGridColumn43.NullText = resources.GetString("ultraGridColumn43.NullText");
			ultraGridColumn44.Format = resources.GetString("ultraGridColumn44.Format");
			ultraGridColumn44.Header.Caption = resources.GetString("resource.Caption43");
			ultraGridColumn44.Header.VisiblePosition = 4;
			ultraGridColumn44.MaskInput = resources.GetString("ultraGridColumn44.MaskInput");
			ultraGridColumn44.NullText = resources.GetString("ultraGridColumn44.NullText");
			ultraGridColumn45.Format = resources.GetString("ultraGridColumn45.Format");
			ultraGridColumn45.Header.Caption = resources.GetString("resource.Caption44");
			ultraGridColumn45.Header.VisiblePosition = 46;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.MaskInput = resources.GetString("ultraGridColumn45.MaskInput");
			ultraGridColumn45.NullText = resources.GetString("ultraGridColumn45.NullText");
			ultraGridColumn46.Format = resources.GetString("ultraGridColumn46.Format");
			ultraGridColumn46.Header.Caption = resources.GetString("resource.Caption45");
			ultraGridColumn46.Header.VisiblePosition = 26;
			ultraGridColumn46.MaskInput = resources.GetString("ultraGridColumn46.MaskInput");
			ultraGridColumn46.NullText = resources.GetString("ultraGridColumn46.NullText");
			ultraGridColumn47.Format = resources.GetString("ultraGridColumn47.Format");
			ultraGridColumn47.Header.Caption = resources.GetString("resource.Caption46");
			ultraGridColumn47.Header.VisiblePosition = 14;
			ultraGridColumn47.MaskInput = resources.GetString("ultraGridColumn47.MaskInput");
			ultraGridColumn47.NullText = resources.GetString("ultraGridColumn47.NullText");
			ultraGridColumn47.Width = 150;
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
																										 ultraGridColumn47});
			appearance15.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance15;
			ultraGridBand1.Header.Caption = resources.GetString("resource.Caption47");
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand1.Override.GroupByRowDescriptionMask = resources.GetString("resource.GroupByRowDescriptionMask");
			ultraGridBand1.Override.NullText = resources.GetString("resource.NullText");
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance16;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance17;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance18;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance19;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance20;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5});
			ultraGridBand1.SummaryFooterCaption = resources.GetString("ultraGridBand1.SummaryFooterCaption");
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance21.ForeColor = System.Drawing.Color.Black;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance21;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance22;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid1.DisplayLayout.Override.GroupByRowDescriptionMask = resources.GetString("resource.GroupByRowDescriptionMask1");
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance23.FontData.BoldAsString = resources.GetString("resource.BoldAsString");
			appearance23.FontData.ItalicAsString = resources.GetString("resource.ItalicAsString");
			appearance23.FontData.Name = resources.GetString("resource.Name");
			appearance23.FontData.SizeInPoints = ((System.Single)(resources.GetObject("resource.SizeInPoints")));
			appearance23.FontData.StrikeoutAsString = resources.GetString("resource.StrikeoutAsString");
			appearance23.FontData.UnderlineAsString = resources.GetString("resource.UnderlineAsString");
			appearance23.ForeColor = System.Drawing.Color.White;
			appearance23.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance23;
			this.ultraGrid1.DisplayLayout.Override.NullText = resources.GetString("resource.NullText1");
			appearance24.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance24.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance26;
			this.ultraGrid1.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("ultraGrid1.Dock")));
			this.ultraGrid1.Enabled = ((bool)(resources.GetObject("ultraGrid1.Enabled")));
			this.ultraGrid1.Font = ((System.Drawing.Font)(resources.GetObject("ultraGrid1.Font")));
			this.ultraGrid1.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("ultraGrid1.ImeMode")));
			this.ultraGrid1.Location = ((System.Drawing.Point)(resources.GetObject("ultraGrid1.Location")));
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("ultraGrid1.RightToLeft")));
			this.ultraGrid1.Size = ((System.Drawing.Size)(resources.GetObject("ultraGrid1.Size")));
			this.ultraGrid1.TabIndex = ((int)(resources.GetObject("ultraGrid1.TabIndex")));
			this.ultraGrid1.Text = resources.GetString("ultraGrid1.Text");
			this.ultraGrid1.Visible = ((bool)(resources.GetObject("ultraGrid1.Visible")));
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// cdsCajaChica
			// 
			this.cdsCajaChica.BindingContextCtrl = this;
			this.cdsCajaChica.DataLibrary = "LibContabilidad";
			this.cdsCajaChica.DataLibraryUrl = "";
			this.cdsCajaChica.DataSetDef = "dsCajaChica";
			this.cdsCajaChica.EnforceConstraints = false;
			this.cdsCajaChica.FillOnRequest = false;
			this.cdsCajaChica.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCajaChica.Name = "cdsCajaChica";
			this.cdsCajaChica.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCajaChica.SchemaDef = null;
			this.cdsCajaChica.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsCajaChica_PositionChanged);
			this.cdsCajaChica.BeforeFill += new C1.Data.FillEventHandler(this.cdsCajaChica_BeforeFill);
			// 
			// label6
			// 
			this.label6.AccessibleDescription = resources.GetString("label6.AccessibleDescription");
			this.label6.AccessibleName = resources.GetString("label6.AccessibleName");
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("label6.Anchor")));
			this.label6.AutoSize = ((bool)(resources.GetObject("label6.AutoSize")));
			this.label6.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("label6.Dock")));
			this.label6.Enabled = ((bool)(resources.GetObject("label6.Enabled")));
			this.label6.Font = ((System.Drawing.Font)(resources.GetObject("label6.Font")));
			this.label6.Image = ((System.Drawing.Image)(resources.GetObject("label6.Image")));
			this.label6.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label6.ImageAlign")));
			this.label6.ImageIndex = ((int)(resources.GetObject("label6.ImageIndex")));
			this.label6.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("label6.ImeMode")));
			this.label6.Location = ((System.Drawing.Point)(resources.GetObject("label6.Location")));
			this.label6.Name = "label6";
			this.label6.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("label6.RightToLeft")));
			this.label6.Size = ((System.Drawing.Size)(resources.GetObject("label6.Size")));
			this.label6.TabIndex = ((int)(resources.GetObject("label6.TabIndex")));
			this.label6.Text = resources.GetString("label6.Text");
			this.label6.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label6.TextAlign")));
			this.label6.Visible = ((bool)(resources.GetObject("label6.Visible")));
			// 
			// dtFecha
			// 
			this.dtFecha.AccessibleDescription = resources.GetString("dtFecha.AccessibleDescription");
			this.dtFecha.AccessibleName = resources.GetString("dtFecha.AccessibleName");
			this.dtFecha.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("dtFecha.Anchor")));
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance27;
			this.dtFecha.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("dtFecha.BackgroundImage")));
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCajaChica, "_CajaChica.Fecha"));
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("dtFecha.Dock")));
			this.dtFecha.Enabled = ((bool)(resources.GetObject("dtFecha.Enabled")));
			this.dtFecha.Font = ((System.Drawing.Font)(resources.GetObject("dtFecha.Font")));
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("dtFecha.ImeMode")));
			this.dtFecha.Location = ((System.Drawing.Point)(resources.GetObject("dtFecha.Location")));
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("dtFecha.RightToLeft")));
			this.dtFecha.Size = ((System.Drawing.Size)(resources.GetObject("dtFecha.Size")));
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = ((int)(resources.GetObject("dtFecha.TabIndex")));
			this.dtFecha.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			this.dtFecha.Visible = ((bool)(resources.GetObject("dtFecha.Visible")));
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.AccessibleDescription = resources.GetString("cmbProyecto.AccessibleDescription");
			this.cmbProyecto.AccessibleName = resources.GetString("cmbProyecto.AccessibleName");
			this.cmbProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("cmbProyecto.Anchor")));
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance28;
			this.cmbProyecto.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("cmbProyecto.BackgroundImage")));
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCajaChica, "_CajaChica.idProyecto"));
			ultraGridBand2.AddButtonCaption = resources.GetString("ultraGridBand2.AddButtonCaption");
			ultraGridBand2.AddButtonToolTipText = resources.GetString("ultraGridBand2.AddButtonToolTipText");
			ultraGridColumn48.Format = resources.GetString("ultraGridColumn48.Format");
			ultraGridColumn48.Header.Caption = resources.GetString("resource.Caption48");
			ultraGridColumn48.Header.VisiblePosition = 0;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.MaskInput = resources.GetString("ultraGridColumn48.MaskInput");
			ultraGridColumn48.NullText = resources.GetString("ultraGridColumn48.NullText");
			ultraGridColumn49.Format = resources.GetString("ultraGridColumn49.Format");
			ultraGridColumn49.Header.Caption = resources.GetString("resource.Caption49");
			ultraGridColumn49.Header.VisiblePosition = 1;
			ultraGridColumn49.MaskInput = resources.GetString("ultraGridColumn49.MaskInput");
			ultraGridColumn49.NullText = resources.GetString("ultraGridColumn49.NullText");
			ultraGridColumn49.Width = 408;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn48,
																										 ultraGridColumn49});
			ultraGridBand2.SummaryFooterCaption = resources.GetString("ultraGridBand2.SummaryFooterCaption");
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("cmbProyecto.Dock")));
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Enabled = ((bool)(resources.GetObject("cmbProyecto.Enabled")));
			this.cmbProyecto.Font = ((System.Drawing.Font)(resources.GetObject("cmbProyecto.Font")));
			this.cmbProyecto.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("cmbProyecto.ImeMode")));
			this.cmbProyecto.Location = ((System.Drawing.Point)(resources.GetObject("cmbProyecto.Location")));
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("cmbProyecto.RightToLeft")));
			this.cmbProyecto.Size = ((System.Drawing.Size)(resources.GetObject("cmbProyecto.Size")));
			this.cmbProyecto.TabIndex = ((int)(resources.GetObject("cmbProyecto.TabIndex")));
			this.cmbProyecto.Text = resources.GetString("cmbProyecto.Text");
			this.cmbProyecto.ValueMember = "idProyecto";
			this.cmbProyecto.Visible = ((bool)(resources.GetObject("cmbProyecto.Visible")));
			this.cmbProyecto.Leave += new System.EventHandler(this.cmbProyecto_Leave);
			this.cmbProyecto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbProyecto_KeyPress);
			this.cmbProyecto.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbProyecto_InitializeLayout);
			// 
			// lblComprobante
			// 
			this.lblComprobante.AccessibleDescription = resources.GetString("lblComprobante.AccessibleDescription");
			this.lblComprobante.AccessibleName = resources.GetString("lblComprobante.AccessibleName");
			this.lblComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("lblComprobante.Anchor")));
			this.lblComprobante.AutoSize = ((bool)(resources.GetObject("lblComprobante.AutoSize")));
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("lblComprobante.Dock")));
			this.lblComprobante.Enabled = ((bool)(resources.GetObject("lblComprobante.Enabled")));
			this.lblComprobante.Font = ((System.Drawing.Font)(resources.GetObject("lblComprobante.Font")));
			this.lblComprobante.Image = ((System.Drawing.Image)(resources.GetObject("lblComprobante.Image")));
			this.lblComprobante.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("lblComprobante.ImageAlign")));
			this.lblComprobante.ImageIndex = ((int)(resources.GetObject("lblComprobante.ImageIndex")));
			this.lblComprobante.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("lblComprobante.ImeMode")));
			this.lblComprobante.Location = ((System.Drawing.Point)(resources.GetObject("lblComprobante.Location")));
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("lblComprobante.RightToLeft")));
			this.lblComprobante.Size = ((System.Drawing.Size)(resources.GetObject("lblComprobante.Size")));
			this.lblComprobante.TabIndex = ((int)(resources.GetObject("lblComprobante.TabIndex")));
			this.lblComprobante.Text = resources.GetString("lblComprobante.Text");
			this.lblComprobante.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("lblComprobante.TextAlign")));
			this.lblComprobante.Visible = ((bool)(resources.GetObject("lblComprobante.Visible")));
			// 
			// label1
			// 
			this.label1.AccessibleDescription = resources.GetString("label1.AccessibleDescription");
			this.label1.AccessibleName = resources.GetString("label1.AccessibleName");
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("label1.Anchor")));
			this.label1.AutoSize = ((bool)(resources.GetObject("label1.AutoSize")));
			this.label1.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("label1.Dock")));
			this.label1.Enabled = ((bool)(resources.GetObject("label1.Enabled")));
			this.label1.Font = ((System.Drawing.Font)(resources.GetObject("label1.Font")));
			this.label1.Image = ((System.Drawing.Image)(resources.GetObject("label1.Image")));
			this.label1.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label1.ImageAlign")));
			this.label1.ImageIndex = ((int)(resources.GetObject("label1.ImageIndex")));
			this.label1.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("label1.ImeMode")));
			this.label1.Location = ((System.Drawing.Point)(resources.GetObject("label1.Location")));
			this.label1.Name = "label1";
			this.label1.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("label1.RightToLeft")));
			this.label1.Size = ((System.Drawing.Size)(resources.GetObject("label1.Size")));
			this.label1.TabIndex = ((int)(resources.GetObject("label1.TabIndex")));
			this.label1.Text = resources.GetString("label1.Text");
			this.label1.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label1.TextAlign")));
			this.label1.Visible = ((bool)(resources.GetObject("label1.Visible")));
			// 
			// txtNotas
			// 
			this.txtNotas.AccessibleDescription = resources.GetString("txtNotas.AccessibleDescription");
			this.txtNotas.AccessibleName = resources.GetString("txtNotas.AccessibleName");
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("txtNotas.Anchor")));
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance29;
			this.txtNotas.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("txtNotas.BackgroundImage")));
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCajaChica, "_CajaChica.Nota"));
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("txtNotas.Dock")));
			this.txtNotas.Enabled = ((bool)(resources.GetObject("txtNotas.Enabled")));
			this.txtNotas.Font = ((System.Drawing.Font)(resources.GetObject("txtNotas.Font")));
			this.txtNotas.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("txtNotas.ImeMode")));
			this.txtNotas.Location = ((System.Drawing.Point)(resources.GetObject("txtNotas.Location")));
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("txtNotas.RightToLeft")));
			this.txtNotas.Size = ((System.Drawing.Size)(resources.GetObject("txtNotas.Size")));
			this.txtNotas.TabIndex = ((int)(resources.GetObject("txtNotas.TabIndex")));
			this.txtNotas.Text = resources.GetString("txtNotas.Text");
			this.txtNotas.Visible = ((bool)(resources.GetObject("txtNotas.Visible")));
			this.txtNotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotas_KeyPress);
			this.txtNotas.ValueChanged += new System.EventHandler(this.txtNotas_ValueChanged);
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.AccessibleDescription = resources.GetString("lblSubTotal.AccessibleDescription");
			this.lblSubTotal.AccessibleName = resources.GetString("lblSubTotal.AccessibleName");
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("lblSubTotal.Anchor")));
			this.lblSubTotal.AutoSize = ((bool)(resources.GetObject("lblSubTotal.AutoSize")));
			this.lblSubTotal.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("lblSubTotal.Dock")));
			this.lblSubTotal.Enabled = ((bool)(resources.GetObject("lblSubTotal.Enabled")));
			this.lblSubTotal.Font = ((System.Drawing.Font)(resources.GetObject("lblSubTotal.Font")));
			this.lblSubTotal.Image = ((System.Drawing.Image)(resources.GetObject("lblSubTotal.Image")));
			this.lblSubTotal.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("lblSubTotal.ImageAlign")));
			this.lblSubTotal.ImageIndex = ((int)(resources.GetObject("lblSubTotal.ImageIndex")));
			this.lblSubTotal.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("lblSubTotal.ImeMode")));
			this.lblSubTotal.Location = ((System.Drawing.Point)(resources.GetObject("lblSubTotal.Location")));
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("lblSubTotal.RightToLeft")));
			this.lblSubTotal.Size = ((System.Drawing.Size)(resources.GetObject("lblSubTotal.Size")));
			this.lblSubTotal.TabIndex = ((int)(resources.GetObject("lblSubTotal.TabIndex")));
			this.lblSubTotal.Text = resources.GetString("lblSubTotal.Text");
			this.lblSubTotal.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("lblSubTotal.TextAlign")));
			this.lblSubTotal.Visible = ((bool)(resources.GetObject("lblSubTotal.Visible")));
			// 
			// txtTotal
			// 
			this.txtTotal.AccessibleDescription = resources.GetString("txtTotal.AccessibleDescription");
			this.txtTotal.AccessibleName = resources.GetString("txtTotal.AccessibleName");
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("txtTotal.Anchor")));
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance30;
			this.txtTotal.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("txtTotal.BackgroundImage")));
			this.txtTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCajaChica, "_CajaChica.Total"));
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("txtTotal.Dock")));
			this.txtTotal.Enabled = ((bool)(resources.GetObject("txtTotal.Enabled")));
			this.txtTotal.Font = ((System.Drawing.Font)(resources.GetObject("txtTotal.Font")));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("txtTotal.ImeMode")));
			this.txtTotal.Location = ((System.Drawing.Point)(resources.GetObject("txtTotal.Location")));
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("txtTotal.RightToLeft")));
			this.txtTotal.Size = ((System.Drawing.Size)(resources.GetObject("txtTotal.Size")));
			this.txtTotal.TabIndex = ((int)(resources.GetObject("txtTotal.TabIndex")));
			this.txtTotal.Visible = ((bool)(resources.GetObject("txtTotal.Visible")));
			// 
			// barraDato1
			// 
			this.barraDato1.AccessibleDescription = resources.GetString("barraDato1.AccessibleDescription");
			this.barraDato1.AccessibleName = resources.GetString("barraDato1.AccessibleName");
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("barraDato1.Anchor")));
			this.barraDato1.AutoScroll = ((bool)(resources.GetObject("barraDato1.AutoScroll")));
			this.barraDato1.AutoScrollMargin = ((System.Drawing.Size)(resources.GetObject("barraDato1.AutoScrollMargin")));
			this.barraDato1.AutoScrollMinSize = ((System.Drawing.Size)(resources.GetObject("barraDato1.AutoScrollMinSize")));
			this.barraDato1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("barraDato1.BackgroundImage")));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_CajaChica";
			this.barraDato1.DataNombreId = "idCajaChica";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsCajaChica;
			this.barraDato1.DataTabla = "CajaChica";
			this.barraDato1.DataTablaHija = "DetalleCajaChica";
			this.barraDato1.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("barraDato1.Dock")));
			this.barraDato1.Enabled = ((bool)(resources.GetObject("barraDato1.Enabled")));
			this.barraDato1.Font = ((System.Drawing.Font)(resources.GetObject("barraDato1.Font")));
			this.barraDato1.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("barraDato1.ImeMode")));
			this.barraDato1.Location = ((System.Drawing.Point)(resources.GetObject("barraDato1.Location")));
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("barraDato1.RightToLeft")));
			this.barraDato1.Size = ((System.Drawing.Size)(resources.GetObject("barraDato1.Size")));
			this.barraDato1.TabIndex = ((int)(resources.GetObject("barraDato1.TabIndex")));
			this.barraDato1.Visible = ((bool)(resources.GetObject("barraDato1.Visible")));
			this.barraDato1.VisibleBorrar = false;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtIdCajaChica
			// 
			this.txtIdCajaChica.AccessibleDescription = resources.GetString("txtIdCajaChica.AccessibleDescription");
			this.txtIdCajaChica.AccessibleName = resources.GetString("txtIdCajaChica.AccessibleName");
			this.txtIdCajaChica.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("txtIdCajaChica.Anchor")));
			this.txtIdCajaChica.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("txtIdCajaChica.BackgroundImage")));
			this.txtIdCajaChica.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCajaChica, "_CajaChica.idCajaChica"));
			this.txtIdCajaChica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCajaChica.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("txtIdCajaChica.Dock")));
			this.txtIdCajaChica.Enabled = ((bool)(resources.GetObject("txtIdCajaChica.Enabled")));
			this.txtIdCajaChica.Font = ((System.Drawing.Font)(resources.GetObject("txtIdCajaChica.Font")));
			this.txtIdCajaChica.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("txtIdCajaChica.ImeMode")));
			this.txtIdCajaChica.Location = ((System.Drawing.Point)(resources.GetObject("txtIdCajaChica.Location")));
			this.txtIdCajaChica.Name = "txtIdCajaChica";
			this.txtIdCajaChica.PromptChar = ' ';
			this.txtIdCajaChica.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("txtIdCajaChica.RightToLeft")));
			this.txtIdCajaChica.Size = ((System.Drawing.Size)(resources.GetObject("txtIdCajaChica.Size")));
			this.txtIdCajaChica.TabIndex = ((int)(resources.GetObject("txtIdCajaChica.TabIndex")));
			this.txtIdCajaChica.Visible = ((bool)(resources.GetObject("txtIdCajaChica.Visible")));
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// cmbArticulo2
			// 
			this.cmbArticulo2.AccessibleDescription = resources.GetString("cmbArticulo2.AccessibleDescription");
			this.cmbArticulo2.AccessibleName = resources.GetString("cmbArticulo2.AccessibleName");
			this.cmbArticulo2.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("cmbArticulo2.Anchor")));
			this.cmbArticulo2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("cmbArticulo2.BackgroundImage")));
			this.cmbArticulo2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo2.DataSource = this.ultraDataSource2;
			this.cmbArticulo2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = resources.GetString("ultraGridBand3.AddButtonCaption");
			ultraGridBand3.AddButtonToolTipText = resources.GetString("ultraGridBand3.AddButtonToolTipText");
			ultraGridColumn50.Format = resources.GetString("ultraGridColumn50.Format");
			ultraGridColumn50.Header.Caption = resources.GetString("resource.Caption50");
			ultraGridColumn50.Header.VisiblePosition = 0;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.MaskInput = resources.GetString("ultraGridColumn50.MaskInput");
			ultraGridColumn50.NullText = resources.GetString("ultraGridColumn50.NullText");
			ultraGridColumn51.Format = resources.GetString("ultraGridColumn51.Format");
			ultraGridColumn51.Header.Caption = resources.GetString("resource.Caption51");
			ultraGridColumn51.Header.VisiblePosition = 1;
			ultraGridColumn51.MaskInput = resources.GetString("ultraGridColumn51.MaskInput");
			ultraGridColumn51.NullText = resources.GetString("ultraGridColumn51.NullText");
			ultraGridColumn51.Width = 185;
			ultraGridColumn52.Format = resources.GetString("ultraGridColumn52.Format");
			ultraGridColumn52.Header.Caption = resources.GetString("resource.Caption52");
			ultraGridColumn52.Header.VisiblePosition = 2;
			ultraGridColumn52.MaskInput = resources.GetString("ultraGridColumn52.MaskInput");
			ultraGridColumn52.NullText = resources.GetString("ultraGridColumn52.NullText");
			ultraGridColumn52.Width = 245;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52});
			ultraGridBand3.SummaryFooterCaption = resources.GetString("ultraGridBand3.SummaryFooterCaption");
			this.cmbArticulo2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbArticulo2.DisplayMember = "Articulo";
			this.cmbArticulo2.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("cmbArticulo2.Dock")));
			this.cmbArticulo2.Enabled = ((bool)(resources.GetObject("cmbArticulo2.Enabled")));
			this.cmbArticulo2.Font = ((System.Drawing.Font)(resources.GetObject("cmbArticulo2.Font")));
			this.cmbArticulo2.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("cmbArticulo2.ImeMode")));
			this.cmbArticulo2.Location = ((System.Drawing.Point)(resources.GetObject("cmbArticulo2.Location")));
			this.cmbArticulo2.Name = "cmbArticulo2";
			this.cmbArticulo2.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("cmbArticulo2.RightToLeft")));
			this.cmbArticulo2.Size = ((System.Drawing.Size)(resources.GetObject("cmbArticulo2.Size")));
			this.cmbArticulo2.TabIndex = ((int)(resources.GetObject("cmbArticulo2.TabIndex")));
			this.cmbArticulo2.Text = resources.GetString("cmbArticulo2.Text");
			this.cmbArticulo2.ValueMember = "idArticulo";
			this.cmbArticulo2.Visible = ((bool)(resources.GetObject("cmbArticulo2.Visible")));
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			this.cdsSeteoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoC_BeforeFill);
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
			// txtNumero
			// 
			this.txtNumero.AccessibleDescription = resources.GetString("txtNumero.AccessibleDescription");
			this.txtNumero.AccessibleName = resources.GetString("txtNumero.AccessibleName");
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("txtNumero.Anchor")));
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance31;
			this.txtNumero.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("txtNumero.BackgroundImage")));
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCajaChica, "_CajaChica.Numero"));
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("txtNumero.Dock")));
			this.txtNumero.Enabled = ((bool)(resources.GetObject("txtNumero.Enabled")));
			this.txtNumero.Font = ((System.Drawing.Font)(resources.GetObject("txtNumero.Font")));
			this.txtNumero.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("txtNumero.ImeMode")));
			this.txtNumero.Location = ((System.Drawing.Point)(resources.GetObject("txtNumero.Location")));
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("txtNumero.RightToLeft")));
			this.txtNumero.Size = ((System.Drawing.Size)(resources.GetObject("txtNumero.Size")));
			this.txtNumero.TabIndex = ((int)(resources.GetObject("txtNumero.TabIndex")));
			this.txtNumero.Text = resources.GetString("txtNumero.Text");
			this.txtNumero.Visible = ((bool)(resources.GetObject("txtNumero.Visible")));
			// 
			// label3
			// 
			this.label3.AccessibleDescription = resources.GetString("label3.AccessibleDescription");
			this.label3.AccessibleName = resources.GetString("label3.AccessibleName");
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("label3.Anchor")));
			this.label3.AutoSize = ((bool)(resources.GetObject("label3.AutoSize")));
			this.label3.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("label3.Dock")));
			this.label3.Enabled = ((bool)(resources.GetObject("label3.Enabled")));
			this.label3.Font = ((System.Drawing.Font)(resources.GetObject("label3.Font")));
			this.label3.Image = ((System.Drawing.Image)(resources.GetObject("label3.Image")));
			this.label3.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label3.ImageAlign")));
			this.label3.ImageIndex = ((int)(resources.GetObject("label3.ImageIndex")));
			this.label3.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("label3.ImeMode")));
			this.label3.Location = ((System.Drawing.Point)(resources.GetObject("label3.Location")));
			this.label3.Name = "label3";
			this.label3.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("label3.RightToLeft")));
			this.label3.Size = ((System.Drawing.Size)(resources.GetObject("label3.Size")));
			this.label3.TabIndex = ((int)(resources.GetObject("label3.TabIndex")));
			this.label3.Text = resources.GetString("label3.Text");
			this.label3.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label3.TextAlign")));
			this.label3.Visible = ((bool)(resources.GetObject("label3.Visible")));
			// 
			// lblEstado
			// 
			this.lblEstado.AccessibleDescription = resources.GetString("lblEstado.AccessibleDescription");
			this.lblEstado.AccessibleName = resources.GetString("lblEstado.AccessibleName");
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("lblEstado.Anchor")));
			this.lblEstado.AutoSize = ((bool)(resources.GetObject("lblEstado.AutoSize")));
			this.lblEstado.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("lblEstado.Dock")));
			this.lblEstado.Enabled = ((bool)(resources.GetObject("lblEstado.Enabled")));
			this.lblEstado.Font = ((System.Drawing.Font)(resources.GetObject("lblEstado.Font")));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Image = ((System.Drawing.Image)(resources.GetObject("lblEstado.Image")));
			this.lblEstado.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("lblEstado.ImageAlign")));
			this.lblEstado.ImageIndex = ((int)(resources.GetObject("lblEstado.ImageIndex")));
			this.lblEstado.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("lblEstado.ImeMode")));
			this.lblEstado.Location = ((System.Drawing.Point)(resources.GetObject("lblEstado.Location")));
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("lblEstado.RightToLeft")));
			this.lblEstado.Size = ((System.Drawing.Size)(resources.GetObject("lblEstado.Size")));
			this.lblEstado.TabIndex = ((int)(resources.GetObject("lblEstado.TabIndex")));
			this.lblEstado.Text = resources.GetString("lblEstado.Text");
			this.lblEstado.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("lblEstado.TextAlign")));
			this.lblEstado.Visible = ((bool)(resources.GetObject("lblEstado.Visible")));
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// label2
			// 
			this.label2.AccessibleDescription = resources.GetString("label2.AccessibleDescription");
			this.label2.AccessibleName = resources.GetString("label2.AccessibleName");
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("label2.Anchor")));
			this.label2.AutoSize = ((bool)(resources.GetObject("label2.AutoSize")));
			this.label2.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("label2.Dock")));
			this.label2.Enabled = ((bool)(resources.GetObject("label2.Enabled")));
			this.label2.Font = ((System.Drawing.Font)(resources.GetObject("label2.Font")));
			this.label2.Image = ((System.Drawing.Image)(resources.GetObject("label2.Image")));
			this.label2.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label2.ImageAlign")));
			this.label2.ImageIndex = ((int)(resources.GetObject("label2.ImageIndex")));
			this.label2.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("label2.ImeMode")));
			this.label2.Location = ((System.Drawing.Point)(resources.GetObject("label2.Location")));
			this.label2.Name = "label2";
			this.label2.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("label2.RightToLeft")));
			this.label2.Size = ((System.Drawing.Size)(resources.GetObject("label2.Size")));
			this.label2.TabIndex = ((int)(resources.GetObject("label2.TabIndex")));
			this.label2.Text = resources.GetString("label2.Text");
			this.label2.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label2.TextAlign")));
			this.label2.Visible = ((bool)(resources.GetObject("label2.Visible")));
			// 
			// txtCajaChica
			// 
			this.txtCajaChica.AccessibleDescription = resources.GetString("txtCajaChica.AccessibleDescription");
			this.txtCajaChica.AccessibleName = resources.GetString("txtCajaChica.AccessibleName");
			this.txtCajaChica.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("txtCajaChica.Anchor")));
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCajaChica.Appearance = appearance32;
			this.txtCajaChica.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("txtCajaChica.BackgroundImage")));
			this.txtCajaChica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCajaChica.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("txtCajaChica.Dock")));
			this.txtCajaChica.Enabled = ((bool)(resources.GetObject("txtCajaChica.Enabled")));
			this.txtCajaChica.Font = ((System.Drawing.Font)(resources.GetObject("txtCajaChica.Font")));
			this.txtCajaChica.FormatString = "#,##0.00";
			this.txtCajaChica.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("txtCajaChica.ImeMode")));
			this.txtCajaChica.Location = ((System.Drawing.Point)(resources.GetObject("txtCajaChica.Location")));
			this.txtCajaChica.Name = "txtCajaChica";
			this.txtCajaChica.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCajaChica.PromptChar = ' ';
			this.txtCajaChica.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("txtCajaChica.RightToLeft")));
			this.txtCajaChica.Size = ((System.Drawing.Size)(resources.GetObject("txtCajaChica.Size")));
			this.txtCajaChica.TabIndex = ((int)(resources.GetObject("txtCajaChica.TabIndex")));
			this.txtCajaChica.Visible = ((bool)(resources.GetObject("txtCajaChica.Visible")));
			// 
			// cmbArticulo1
			// 
			this.cmbArticulo1.AccessibleDescription = resources.GetString("cmbArticulo1.AccessibleDescription");
			this.cmbArticulo1.AccessibleName = resources.GetString("cmbArticulo1.AccessibleName");
			this.cmbArticulo1.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("cmbArticulo1.Anchor")));
			this.cmbArticulo1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("cmbArticulo1.BackgroundImage")));
			this.cmbArticulo1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo1.DataSource = this.ultraDataSource1;
			this.cmbArticulo1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = resources.GetString("ultraGridBand4.AddButtonCaption");
			ultraGridBand4.AddButtonToolTipText = resources.GetString("ultraGridBand4.AddButtonToolTipText");
			ultraGridColumn53.Format = resources.GetString("ultraGridColumn53.Format");
			ultraGridColumn53.Header.Caption = resources.GetString("resource.Caption53");
			ultraGridColumn53.Header.VisiblePosition = 0;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.MaskInput = resources.GetString("ultraGridColumn53.MaskInput");
			ultraGridColumn53.NullText = resources.GetString("ultraGridColumn53.NullText");
			ultraGridColumn54.Format = resources.GetString("ultraGridColumn54.Format");
			ultraGridColumn54.Header.Caption = resources.GetString("resource.Caption54");
			ultraGridColumn54.Header.VisiblePosition = 1;
			ultraGridColumn54.MaskInput = resources.GetString("ultraGridColumn54.MaskInput");
			ultraGridColumn54.NullText = resources.GetString("ultraGridColumn54.NullText");
			ultraGridColumn54.Width = 185;
			ultraGridColumn55.Format = resources.GetString("ultraGridColumn55.Format");
			ultraGridColumn55.Header.Caption = resources.GetString("resource.Caption55");
			ultraGridColumn55.Header.VisiblePosition = 2;
			ultraGridColumn55.MaskInput = resources.GetString("ultraGridColumn55.MaskInput");
			ultraGridColumn55.NullText = resources.GetString("ultraGridColumn55.NullText");
			ultraGridColumn55.Width = 245;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55});
			ultraGridBand4.SummaryFooterCaption = resources.GetString("ultraGridBand4.SummaryFooterCaption");
			this.cmbArticulo1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbArticulo1.DisplayMember = "Articulo";
			this.cmbArticulo1.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("cmbArticulo1.Dock")));
			this.cmbArticulo1.Enabled = ((bool)(resources.GetObject("cmbArticulo1.Enabled")));
			this.cmbArticulo1.Font = ((System.Drawing.Font)(resources.GetObject("cmbArticulo1.Font")));
			this.cmbArticulo1.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("cmbArticulo1.ImeMode")));
			this.cmbArticulo1.Location = ((System.Drawing.Point)(resources.GetObject("cmbArticulo1.Location")));
			this.cmbArticulo1.Name = "cmbArticulo1";
			this.cmbArticulo1.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("cmbArticulo1.RightToLeft")));
			this.cmbArticulo1.Size = ((System.Drawing.Size)(resources.GetObject("cmbArticulo1.Size")));
			this.cmbArticulo1.TabIndex = ((int)(resources.GetObject("cmbArticulo1.TabIndex")));
			this.cmbArticulo1.Text = resources.GetString("cmbArticulo1.Text");
			this.cmbArticulo1.ValueMember = "idArticulo";
			this.cmbArticulo1.Visible = ((bool)(resources.GetObject("cmbArticulo1.Visible")));
			// 
			// ultraDataSource1
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// label4
			// 
			this.label4.AccessibleDescription = resources.GetString("label4.AccessibleDescription");
			this.label4.AccessibleName = resources.GetString("label4.AccessibleName");
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("label4.Anchor")));
			this.label4.AutoSize = ((bool)(resources.GetObject("label4.AutoSize")));
			this.label4.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("label4.Dock")));
			this.label4.Enabled = ((bool)(resources.GetObject("label4.Enabled")));
			this.label4.Font = ((System.Drawing.Font)(resources.GetObject("label4.Font")));
			this.label4.Image = ((System.Drawing.Image)(resources.GetObject("label4.Image")));
			this.label4.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label4.ImageAlign")));
			this.label4.ImageIndex = ((int)(resources.GetObject("label4.ImageIndex")));
			this.label4.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("label4.ImeMode")));
			this.label4.Location = ((System.Drawing.Point)(resources.GetObject("label4.Location")));
			this.label4.Name = "label4";
			this.label4.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("label4.RightToLeft")));
			this.label4.Size = ((System.Drawing.Size)(resources.GetObject("label4.Size")));
			this.label4.TabIndex = ((int)(resources.GetObject("label4.TabIndex")));
			this.label4.Text = resources.GetString("label4.Text");
			this.label4.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label4.TextAlign")));
			this.label4.Visible = ((bool)(resources.GetObject("label4.Visible")));
			// 
			// txtResponsable
			// 
			this.txtResponsable.AccessibleDescription = resources.GetString("txtResponsable.AccessibleDescription");
			this.txtResponsable.AccessibleName = resources.GetString("txtResponsable.AccessibleName");
			this.txtResponsable.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("txtResponsable.Anchor")));
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtResponsable.Appearance = appearance33;
			this.txtResponsable.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("txtResponsable.BackgroundImage")));
			this.txtResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtResponsable.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCajaChica, "_CajaChica.Nombre"));
			this.txtResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResponsable.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("txtResponsable.Dock")));
			this.txtResponsable.Enabled = ((bool)(resources.GetObject("txtResponsable.Enabled")));
			this.txtResponsable.Font = ((System.Drawing.Font)(resources.GetObject("txtResponsable.Font")));
			this.txtResponsable.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("txtResponsable.ImeMode")));
			this.txtResponsable.Location = ((System.Drawing.Point)(resources.GetObject("txtResponsable.Location")));
			this.txtResponsable.MaxLength = 50;
			this.txtResponsable.Name = "txtResponsable";
			this.txtResponsable.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("txtResponsable.RightToLeft")));
			this.txtResponsable.Size = ((System.Drawing.Size)(resources.GetObject("txtResponsable.Size")));
			this.txtResponsable.TabIndex = ((int)(resources.GetObject("txtResponsable.TabIndex")));
			this.txtResponsable.Text = resources.GetString("txtResponsable.Text");
			this.txtResponsable.Visible = ((bool)(resources.GetObject("txtResponsable.Visible")));
			this.txtResponsable.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtResponsable_KeyPress);
			this.txtResponsable.ValueChanged += new System.EventHandler(this.txtResponsable_ValueChanged);
			// 
			// cmbCentroDeCosto
			// 
			this.cmbCentroDeCosto.AccessibleDescription = resources.GetString("cmbCentroDeCosto.AccessibleDescription");
			this.cmbCentroDeCosto.AccessibleName = resources.GetString("cmbCentroDeCosto.AccessibleName");
			this.cmbCentroDeCosto.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("cmbCentroDeCosto.Anchor")));
			this.cmbCentroDeCosto.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("cmbCentroDeCosto.BackgroundImage")));
			this.cmbCentroDeCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroDeCosto.DataSource = this.ultraDataSource3;
			this.cmbCentroDeCosto.DisplayLayout.AutoFitColumns = true;
			ultraGridBand5.AddButtonCaption = resources.GetString("ultraGridBand5.AddButtonCaption");
			ultraGridBand5.AddButtonToolTipText = resources.GetString("ultraGridBand5.AddButtonToolTipText");
			ultraGridColumn56.Format = resources.GetString("ultraGridColumn56.Format");
			ultraGridColumn56.Header.Caption = resources.GetString("resource.Caption56");
			ultraGridColumn56.Header.VisiblePosition = 0;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.MaskInput = resources.GetString("ultraGridColumn56.MaskInput");
			ultraGridColumn56.NullText = resources.GetString("ultraGridColumn56.NullText");
			ultraGridColumn56.Width = 178;
			ultraGridColumn57.Format = resources.GetString("ultraGridColumn57.Format");
			ultraGridColumn57.Header.Caption = resources.GetString("resource.Caption57");
			ultraGridColumn57.Header.VisiblePosition = 1;
			ultraGridColumn57.MaskInput = resources.GetString("ultraGridColumn57.MaskInput");
			ultraGridColumn57.NullText = resources.GetString("ultraGridColumn57.NullText");
			ultraGridColumn57.Width = 430;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn56,
																										 ultraGridColumn57});
			ultraGridBand5.SummaryFooterCaption = resources.GetString("ultraGridBand5.SummaryFooterCaption");
			this.cmbCentroDeCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCentroDeCosto.DisplayMember = "Nombre";
			this.cmbCentroDeCosto.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("cmbCentroDeCosto.Dock")));
			this.cmbCentroDeCosto.Enabled = ((bool)(resources.GetObject("cmbCentroDeCosto.Enabled")));
			this.cmbCentroDeCosto.Font = ((System.Drawing.Font)(resources.GetObject("cmbCentroDeCosto.Font")));
			this.cmbCentroDeCosto.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("cmbCentroDeCosto.ImeMode")));
			this.cmbCentroDeCosto.Location = ((System.Drawing.Point)(resources.GetObject("cmbCentroDeCosto.Location")));
			this.cmbCentroDeCosto.Name = "cmbCentroDeCosto";
			this.cmbCentroDeCosto.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("cmbCentroDeCosto.RightToLeft")));
			this.cmbCentroDeCosto.Size = ((System.Drawing.Size)(resources.GetObject("cmbCentroDeCosto.Size")));
			this.cmbCentroDeCosto.TabIndex = ((int)(resources.GetObject("cmbCentroDeCosto.TabIndex")));
			this.cmbCentroDeCosto.Text = resources.GetString("cmbCentroDeCosto.Text");
			this.cmbCentroDeCosto.ValueMember = "idProyecto";
			this.cmbCentroDeCosto.Visible = ((bool)(resources.GetObject("cmbCentroDeCosto.Visible")));
			this.cmbCentroDeCosto.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCentroDeCosto_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// btnImprimir
			// 
			this.btnImprimir.AccessibleDescription = resources.GetString("btnImprimir.AccessibleDescription");
			this.btnImprimir.AccessibleName = resources.GetString("btnImprimir.AccessibleName");
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnImprimir.Anchor")));
			this.btnImprimir.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnImprimir.BackgroundImage")));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnImprimir.Dock")));
			this.btnImprimir.Enabled = ((bool)(resources.GetObject("btnImprimir.Enabled")));
			this.btnImprimir.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnImprimir.FlatStyle")));
			this.btnImprimir.Font = ((System.Drawing.Font)(resources.GetObject("btnImprimir.Font")));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnImprimir.ImageAlign")));
			this.btnImprimir.ImageIndex = ((int)(resources.GetObject("btnImprimir.ImageIndex")));
			this.btnImprimir.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnImprimir.ImeMode")));
			this.btnImprimir.Location = ((System.Drawing.Point)(resources.GetObject("btnImprimir.Location")));
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnImprimir.RightToLeft")));
			this.btnImprimir.Size = ((System.Drawing.Size)(resources.GetObject("btnImprimir.Size")));
			this.btnImprimir.TabIndex = ((int)(resources.GetObject("btnImprimir.TabIndex")));
			this.btnImprimir.Text = resources.GetString("btnImprimir.Text");
			this.btnImprimir.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnImprimir.TextAlign")));
			this.btnImprimir.Visible = ((bool)(resources.GetObject("btnImprimir.Visible")));
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.AccessibleDescription = resources.GetString("btnExportar.AccessibleDescription");
			this.btnExportar.AccessibleName = resources.GetString("btnExportar.AccessibleName");
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnExportar.Anchor")));
			this.btnExportar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnExportar.BackgroundImage")));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnExportar.Dock")));
			this.btnExportar.Enabled = ((bool)(resources.GetObject("btnExportar.Enabled")));
			this.btnExportar.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnExportar.FlatStyle")));
			this.btnExportar.Font = ((System.Drawing.Font)(resources.GetObject("btnExportar.Font")));
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnExportar.ImageAlign")));
			this.btnExportar.ImageIndex = ((int)(resources.GetObject("btnExportar.ImageIndex")));
			this.btnExportar.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnExportar.ImeMode")));
			this.btnExportar.Location = ((System.Drawing.Point)(resources.GetObject("btnExportar.Location")));
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnExportar.RightToLeft")));
			this.btnExportar.Size = ((System.Drawing.Size)(resources.GetObject("btnExportar.Size")));
			this.btnExportar.TabIndex = ((int)(resources.GetObject("btnExportar.TabIndex")));
			this.btnExportar.Text = resources.GetString("btnExportar.Text");
			this.btnExportar.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnExportar.TextAlign")));
			this.btnExportar.Visible = ((bool)(resources.GetObject("btnExportar.Visible")));
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// cmbRetencionesRenta1
			// 
			this.cmbRetencionesRenta1.AccessibleDescription = resources.GetString("cmbRetencionesRenta1.AccessibleDescription");
			this.cmbRetencionesRenta1.AccessibleName = resources.GetString("cmbRetencionesRenta1.AccessibleName");
			this.cmbRetencionesRenta1.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("cmbRetencionesRenta1.Anchor")));
			this.cmbRetencionesRenta1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("cmbRetencionesRenta1.BackgroundImage")));
			this.cmbRetencionesRenta1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRetencionesRenta1.DataSource = this.ultraDataSource4;
			this.cmbRetencionesRenta1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand6.AddButtonCaption = resources.GetString("ultraGridBand6.AddButtonCaption");
			ultraGridBand6.AddButtonToolTipText = resources.GetString("ultraGridBand6.AddButtonToolTipText");
			ultraGridColumn58.Format = resources.GetString("ultraGridColumn58.Format");
			ultraGridColumn58.Header.Caption = resources.GetString("resource.Caption58");
			ultraGridColumn58.Header.VisiblePosition = 0;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.MaskInput = resources.GetString("ultraGridColumn58.MaskInput");
			ultraGridColumn58.NullText = resources.GetString("ultraGridColumn58.NullText");
			ultraGridColumn58.Width = 99;
			ultraGridColumn59.Format = resources.GetString("ultraGridColumn59.Format");
			ultraGridColumn59.Header.Caption = resources.GetString("resource.Caption59");
			ultraGridColumn59.Header.VisiblePosition = 1;
			ultraGridColumn59.MaskInput = resources.GetString("ultraGridColumn59.MaskInput");
			ultraGridColumn59.NullText = resources.GetString("ultraGridColumn59.NullText");
			ultraGridColumn59.Width = 52;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn60.CellAppearance = appearance34;
			ultraGridColumn60.Format = resources.GetString("ultraGridColumn60.Format");
			ultraGridColumn60.Header.Caption = resources.GetString("resource.Caption60");
			ultraGridColumn60.Header.VisiblePosition = 2;
			ultraGridColumn60.MaskInput = resources.GetString("ultraGridColumn60.MaskInput");
			ultraGridColumn60.NullText = resources.GetString("ultraGridColumn60.NullText");
			ultraGridColumn60.Width = 50;
			ultraGridColumn61.Format = resources.GetString("ultraGridColumn61.Format");
			ultraGridColumn61.Header.Caption = resources.GetString("resource.Caption61");
			ultraGridColumn61.Header.VisiblePosition = 3;
			ultraGridColumn61.MaskInput = resources.GetString("ultraGridColumn61.MaskInput");
			ultraGridColumn61.NullText = resources.GetString("ultraGridColumn61.NullText");
			ultraGridColumn61.Width = 304;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61});
			ultraGridBand6.SummaryFooterCaption = resources.GetString("ultraGridBand6.SummaryFooterCaption");
			this.cmbRetencionesRenta1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbRetencionesRenta1.DisplayMember = "Codigo";
			this.cmbRetencionesRenta1.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("cmbRetencionesRenta1.Dock")));
			this.cmbRetencionesRenta1.Enabled = ((bool)(resources.GetObject("cmbRetencionesRenta1.Enabled")));
			this.cmbRetencionesRenta1.Font = ((System.Drawing.Font)(resources.GetObject("cmbRetencionesRenta1.Font")));
			this.cmbRetencionesRenta1.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("cmbRetencionesRenta1.ImeMode")));
			this.cmbRetencionesRenta1.Location = ((System.Drawing.Point)(resources.GetObject("cmbRetencionesRenta1.Location")));
			this.cmbRetencionesRenta1.Name = "cmbRetencionesRenta1";
			this.cmbRetencionesRenta1.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("cmbRetencionesRenta1.RightToLeft")));
			this.cmbRetencionesRenta1.Size = ((System.Drawing.Size)(resources.GetObject("cmbRetencionesRenta1.Size")));
			this.cmbRetencionesRenta1.TabIndex = ((int)(resources.GetObject("cmbRetencionesRenta1.TabIndex")));
			this.cmbRetencionesRenta1.Text = resources.GetString("cmbRetencionesRenta1.Text");
			this.cmbRetencionesRenta1.ValueMember = "idArticuloSRI";
			this.cmbRetencionesRenta1.Visible = ((bool)(resources.GetObject("cmbRetencionesRenta1.Visible")));
			// 
			// ultraDataSource4
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(System.Double);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// cmbRetencionesIVA1
			// 
			this.cmbRetencionesIVA1.AccessibleDescription = resources.GetString("cmbRetencionesIVA1.AccessibleDescription");
			this.cmbRetencionesIVA1.AccessibleName = resources.GetString("cmbRetencionesIVA1.AccessibleName");
			this.cmbRetencionesIVA1.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("cmbRetencionesIVA1.Anchor")));
			this.cmbRetencionesIVA1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("cmbRetencionesIVA1.BackgroundImage")));
			this.cmbRetencionesIVA1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRetencionesIVA1.DataSource = this.ultraDataSource6;
			this.cmbRetencionesIVA1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand7.AddButtonCaption = resources.GetString("ultraGridBand7.AddButtonCaption");
			ultraGridBand7.AddButtonToolTipText = resources.GetString("ultraGridBand7.AddButtonToolTipText");
			ultraGridColumn62.Format = resources.GetString("ultraGridColumn62.Format");
			ultraGridColumn62.Header.Caption = resources.GetString("resource.Caption62");
			ultraGridColumn62.Header.VisiblePosition = 0;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.MaskInput = resources.GetString("ultraGridColumn62.MaskInput");
			ultraGridColumn62.NullText = resources.GetString("ultraGridColumn62.NullText");
			ultraGridColumn62.Width = 79;
			ultraGridColumn63.Format = resources.GetString("ultraGridColumn63.Format");
			ultraGridColumn63.Header.Caption = resources.GetString("resource.Caption63");
			ultraGridColumn63.Header.VisiblePosition = 1;
			ultraGridColumn63.MaskInput = resources.GetString("ultraGridColumn63.MaskInput");
			ultraGridColumn63.NullText = resources.GetString("ultraGridColumn63.NullText");
			ultraGridColumn63.Width = 50;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance35;
			ultraGridColumn64.Format = resources.GetString("ultraGridColumn64.Format");
			ultraGridColumn64.Header.Caption = resources.GetString("resource.Caption64");
			ultraGridColumn64.Header.VisiblePosition = 2;
			ultraGridColumn64.MaskInput = resources.GetString("ultraGridColumn64.MaskInput");
			ultraGridColumn64.NullText = resources.GetString("ultraGridColumn64.NullText");
			ultraGridColumn64.Width = 50;
			ultraGridColumn65.Format = resources.GetString("ultraGridColumn65.Format");
			ultraGridColumn65.Header.Caption = resources.GetString("resource.Caption65");
			ultraGridColumn65.Header.VisiblePosition = 3;
			ultraGridColumn65.MaskInput = resources.GetString("ultraGridColumn65.MaskInput");
			ultraGridColumn65.NullText = resources.GetString("ultraGridColumn65.NullText");
			ultraGridColumn65.Width = 306;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65});
			ultraGridBand7.SummaryFooterCaption = resources.GetString("ultraGridBand7.SummaryFooterCaption");
			this.cmbRetencionesIVA1.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbRetencionesIVA1.DisplayMember = "Codigo";
			this.cmbRetencionesIVA1.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("cmbRetencionesIVA1.Dock")));
			this.cmbRetencionesIVA1.Enabled = ((bool)(resources.GetObject("cmbRetencionesIVA1.Enabled")));
			this.cmbRetencionesIVA1.Font = ((System.Drawing.Font)(resources.GetObject("cmbRetencionesIVA1.Font")));
			this.cmbRetencionesIVA1.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("cmbRetencionesIVA1.ImeMode")));
			this.cmbRetencionesIVA1.Location = ((System.Drawing.Point)(resources.GetObject("cmbRetencionesIVA1.Location")));
			this.cmbRetencionesIVA1.Name = "cmbRetencionesIVA1";
			this.cmbRetencionesIVA1.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("cmbRetencionesIVA1.RightToLeft")));
			this.cmbRetencionesIVA1.Size = ((System.Drawing.Size)(resources.GetObject("cmbRetencionesIVA1.Size")));
			this.cmbRetencionesIVA1.TabIndex = ((int)(resources.GetObject("cmbRetencionesIVA1.TabIndex")));
			this.cmbRetencionesIVA1.Text = resources.GetString("cmbRetencionesIVA1.Text");
			this.cmbRetencionesIVA1.ValueMember = "idRetIVA";
			this.cmbRetencionesIVA1.Visible = ((bool)(resources.GetObject("cmbRetencionesIVA1.Visible")));
			// 
			// ultraDataSource6
			// 
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(System.Double);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// cmbRetencionesIVA2
			// 
			this.cmbRetencionesIVA2.AccessibleDescription = resources.GetString("cmbRetencionesIVA2.AccessibleDescription");
			this.cmbRetencionesIVA2.AccessibleName = resources.GetString("cmbRetencionesIVA2.AccessibleName");
			this.cmbRetencionesIVA2.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("cmbRetencionesIVA2.Anchor")));
			this.cmbRetencionesIVA2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("cmbRetencionesIVA2.BackgroundImage")));
			this.cmbRetencionesIVA2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRetencionesIVA2.DataSource = this.ultraDataSource7;
			this.cmbRetencionesIVA2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand8.AddButtonCaption = resources.GetString("ultraGridBand8.AddButtonCaption");
			ultraGridBand8.AddButtonToolTipText = resources.GetString("ultraGridBand8.AddButtonToolTipText");
			ultraGridColumn66.Format = resources.GetString("ultraGridColumn66.Format");
			ultraGridColumn66.Header.Caption = resources.GetString("resource.Caption66");
			ultraGridColumn66.Header.VisiblePosition = 0;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn66.MaskInput = resources.GetString("ultraGridColumn66.MaskInput");
			ultraGridColumn66.NullText = resources.GetString("ultraGridColumn66.NullText");
			ultraGridColumn67.Format = resources.GetString("ultraGridColumn67.Format");
			ultraGridColumn67.Header.Caption = resources.GetString("resource.Caption67");
			ultraGridColumn67.Header.VisiblePosition = 1;
			ultraGridColumn67.MaskInput = resources.GetString("ultraGridColumn67.MaskInput");
			ultraGridColumn67.NullText = resources.GetString("ultraGridColumn67.NullText");
			ultraGridColumn67.Width = 52;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance36;
			ultraGridColumn68.Format = resources.GetString("ultraGridColumn68.Format");
			ultraGridColumn68.Header.Caption = resources.GetString("resource.Caption68");
			ultraGridColumn68.Header.VisiblePosition = 2;
			ultraGridColumn68.MaskInput = resources.GetString("ultraGridColumn68.MaskInput");
			ultraGridColumn68.NullText = resources.GetString("ultraGridColumn68.NullText");
			ultraGridColumn68.Width = 50;
			ultraGridColumn69.Format = resources.GetString("ultraGridColumn69.Format");
			ultraGridColumn69.Header.Caption = resources.GetString("resource.Caption69");
			ultraGridColumn69.Header.VisiblePosition = 3;
			ultraGridColumn69.MaskInput = resources.GetString("ultraGridColumn69.MaskInput");
			ultraGridColumn69.NullText = resources.GetString("ultraGridColumn69.NullText");
			ultraGridColumn69.Width = 304;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn66,
																										 ultraGridColumn67,
																										 ultraGridColumn68,
																										 ultraGridColumn69});
			ultraGridBand8.SummaryFooterCaption = resources.GetString("ultraGridBand8.SummaryFooterCaption");
			this.cmbRetencionesIVA2.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbRetencionesIVA2.DisplayMember = "Codigo";
			this.cmbRetencionesIVA2.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("cmbRetencionesIVA2.Dock")));
			this.cmbRetencionesIVA2.Enabled = ((bool)(resources.GetObject("cmbRetencionesIVA2.Enabled")));
			this.cmbRetencionesIVA2.Font = ((System.Drawing.Font)(resources.GetObject("cmbRetencionesIVA2.Font")));
			this.cmbRetencionesIVA2.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("cmbRetencionesIVA2.ImeMode")));
			this.cmbRetencionesIVA2.Location = ((System.Drawing.Point)(resources.GetObject("cmbRetencionesIVA2.Location")));
			this.cmbRetencionesIVA2.Name = "cmbRetencionesIVA2";
			this.cmbRetencionesIVA2.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("cmbRetencionesIVA2.RightToLeft")));
			this.cmbRetencionesIVA2.Size = ((System.Drawing.Size)(resources.GetObject("cmbRetencionesIVA2.Size")));
			this.cmbRetencionesIVA2.TabIndex = ((int)(resources.GetObject("cmbRetencionesIVA2.TabIndex")));
			this.cmbRetencionesIVA2.Text = resources.GetString("cmbRetencionesIVA2.Text");
			this.cmbRetencionesIVA2.ValueMember = "idRetIVA";
			this.cmbRetencionesIVA2.Visible = ((bool)(resources.GetObject("cmbRetencionesIVA2.Visible")));
			this.cmbRetencionesIVA2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbRetencionesIVA2_InitializeLayout);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.Double);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// cmbRetencionesRenta2
			// 
			this.cmbRetencionesRenta2.AccessibleDescription = resources.GetString("cmbRetencionesRenta2.AccessibleDescription");
			this.cmbRetencionesRenta2.AccessibleName = resources.GetString("cmbRetencionesRenta2.AccessibleName");
			this.cmbRetencionesRenta2.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("cmbRetencionesRenta2.Anchor")));
			this.cmbRetencionesRenta2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("cmbRetencionesRenta2.BackgroundImage")));
			this.cmbRetencionesRenta2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRetencionesRenta2.DataSource = this.ultraDataSource5;
			this.cmbRetencionesRenta2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand9.AddButtonCaption = resources.GetString("ultraGridBand9.AddButtonCaption");
			ultraGridBand9.AddButtonToolTipText = resources.GetString("ultraGridBand9.AddButtonToolTipText");
			ultraGridColumn70.Format = resources.GetString("ultraGridColumn70.Format");
			ultraGridColumn70.Header.Caption = resources.GetString("resource.Caption70");
			ultraGridColumn70.Header.VisiblePosition = 0;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.MaskInput = resources.GetString("ultraGridColumn70.MaskInput");
			ultraGridColumn70.NullText = resources.GetString("ultraGridColumn70.NullText");
			ultraGridColumn70.Width = 99;
			ultraGridColumn71.Format = resources.GetString("ultraGridColumn71.Format");
			ultraGridColumn71.Header.Caption = resources.GetString("resource.Caption71");
			ultraGridColumn71.Header.VisiblePosition = 1;
			ultraGridColumn71.MaskInput = resources.GetString("ultraGridColumn71.MaskInput");
			ultraGridColumn71.NullText = resources.GetString("ultraGridColumn71.NullText");
			ultraGridColumn71.Width = 52;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance37;
			ultraGridColumn72.Format = resources.GetString("ultraGridColumn72.Format");
			ultraGridColumn72.Header.Caption = resources.GetString("resource.Caption72");
			ultraGridColumn72.Header.VisiblePosition = 2;
			ultraGridColumn72.MaskInput = resources.GetString("ultraGridColumn72.MaskInput");
			ultraGridColumn72.NullText = resources.GetString("ultraGridColumn72.NullText");
			ultraGridColumn72.Width = 50;
			ultraGridColumn73.Format = resources.GetString("ultraGridColumn73.Format");
			ultraGridColumn73.Header.Caption = resources.GetString("resource.Caption73");
			ultraGridColumn73.Header.VisiblePosition = 3;
			ultraGridColumn73.MaskInput = resources.GetString("ultraGridColumn73.MaskInput");
			ultraGridColumn73.NullText = resources.GetString("ultraGridColumn73.NullText");
			ultraGridColumn73.Width = 304;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73});
			ultraGridBand9.SummaryFooterCaption = resources.GetString("ultraGridBand9.SummaryFooterCaption");
			this.cmbRetencionesRenta2.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbRetencionesRenta2.DisplayMember = "Codigo";
			this.cmbRetencionesRenta2.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("cmbRetencionesRenta2.Dock")));
			this.cmbRetencionesRenta2.Enabled = ((bool)(resources.GetObject("cmbRetencionesRenta2.Enabled")));
			this.cmbRetencionesRenta2.Font = ((System.Drawing.Font)(resources.GetObject("cmbRetencionesRenta2.Font")));
			this.cmbRetencionesRenta2.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("cmbRetencionesRenta2.ImeMode")));
			this.cmbRetencionesRenta2.Location = ((System.Drawing.Point)(resources.GetObject("cmbRetencionesRenta2.Location")));
			this.cmbRetencionesRenta2.Name = "cmbRetencionesRenta2";
			this.cmbRetencionesRenta2.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("cmbRetencionesRenta2.RightToLeft")));
			this.cmbRetencionesRenta2.Size = ((System.Drawing.Size)(resources.GetObject("cmbRetencionesRenta2.Size")));
			this.cmbRetencionesRenta2.TabIndex = ((int)(resources.GetObject("cmbRetencionesRenta2.TabIndex")));
			this.cmbRetencionesRenta2.Text = resources.GetString("cmbRetencionesRenta2.Text");
			this.cmbRetencionesRenta2.ValueMember = "idArticuloSRI";
			this.cmbRetencionesRenta2.Visible = ((bool)(resources.GetObject("cmbRetencionesRenta2.Visible")));
			// 
			// ultraDataSource5
			// 
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(System.Double);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// btnActualizar
			// 
			this.btnActualizar.AccessibleDescription = resources.GetString("btnActualizar.AccessibleDescription");
			this.btnActualizar.AccessibleName = resources.GetString("btnActualizar.AccessibleName");
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnActualizar.Anchor")));
			this.btnActualizar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnActualizar.BackgroundImage")));
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnActualizar.Dock")));
			this.btnActualizar.Enabled = ((bool)(resources.GetObject("btnActualizar.Enabled")));
			this.btnActualizar.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnActualizar.FlatStyle")));
			this.btnActualizar.Font = ((System.Drawing.Font)(resources.GetObject("btnActualizar.Font")));
			this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnActualizar.ImageAlign")));
			this.btnActualizar.ImageIndex = ((int)(resources.GetObject("btnActualizar.ImageIndex")));
			this.btnActualizar.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnActualizar.ImeMode")));
			this.btnActualizar.Location = ((System.Drawing.Point)(resources.GetObject("btnActualizar.Location")));
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnActualizar.RightToLeft")));
			this.btnActualizar.Size = ((System.Drawing.Size)(resources.GetObject("btnActualizar.Size")));
			this.btnActualizar.TabIndex = ((int)(resources.GetObject("btnActualizar.TabIndex")));
			this.btnActualizar.Text = resources.GetString("btnActualizar.Text");
			this.btnActualizar.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnActualizar.TextAlign")));
			this.btnActualizar.Visible = ((bool)(resources.GetObject("btnActualizar.Visible")));
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// btnCalcular
			// 
			this.btnCalcular.AccessibleDescription = resources.GetString("btnCalcular.AccessibleDescription");
			this.btnCalcular.AccessibleName = resources.GetString("btnCalcular.AccessibleName");
			this.btnCalcular.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnCalcular.Anchor")));
			this.btnCalcular.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnCalcular.BackgroundImage")));
			this.btnCalcular.CausesValidation = false;
			this.btnCalcular.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnCalcular.Dock")));
			this.btnCalcular.Enabled = ((bool)(resources.GetObject("btnCalcular.Enabled")));
			this.btnCalcular.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnCalcular.FlatStyle")));
			this.btnCalcular.Font = ((System.Drawing.Font)(resources.GetObject("btnCalcular.Font")));
			this.btnCalcular.Image = ((System.Drawing.Image)(resources.GetObject("btnCalcular.Image")));
			this.btnCalcular.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnCalcular.ImageAlign")));
			this.btnCalcular.ImageIndex = ((int)(resources.GetObject("btnCalcular.ImageIndex")));
			this.btnCalcular.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnCalcular.ImeMode")));
			this.btnCalcular.Location = ((System.Drawing.Point)(resources.GetObject("btnCalcular.Location")));
			this.btnCalcular.Name = "btnCalcular";
			this.btnCalcular.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnCalcular.RightToLeft")));
			this.btnCalcular.Size = ((System.Drawing.Size)(resources.GetObject("btnCalcular.Size")));
			this.btnCalcular.TabIndex = ((int)(resources.GetObject("btnCalcular.TabIndex")));
			this.btnCalcular.Text = resources.GetString("btnCalcular.Text");
			this.btnCalcular.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnCalcular.TextAlign")));
			this.btnCalcular.Visible = ((bool)(resources.GetObject("btnCalcular.Visible")));
			this.btnCalcular.Click += new System.EventHandler(this.btnCalcular_Click);
			// 
			// btnEgreso
			// 
			this.btnEgreso.AccessibleDescription = resources.GetString("btnEgreso.AccessibleDescription");
			this.btnEgreso.AccessibleName = resources.GetString("btnEgreso.AccessibleName");
			this.btnEgreso.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnEgreso.Anchor")));
			this.btnEgreso.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnEgreso.BackgroundImage")));
			this.btnEgreso.CausesValidation = false;
			this.btnEgreso.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnEgreso.Dock")));
			this.btnEgreso.Enabled = ((bool)(resources.GetObject("btnEgreso.Enabled")));
			this.btnEgreso.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnEgreso.FlatStyle")));
			this.btnEgreso.Font = ((System.Drawing.Font)(resources.GetObject("btnEgreso.Font")));
			this.btnEgreso.Image = ((System.Drawing.Image)(resources.GetObject("btnEgreso.Image")));
			this.btnEgreso.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnEgreso.ImageAlign")));
			this.btnEgreso.ImageIndex = ((int)(resources.GetObject("btnEgreso.ImageIndex")));
			this.btnEgreso.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnEgreso.ImeMode")));
			this.btnEgreso.Location = ((System.Drawing.Point)(resources.GetObject("btnEgreso.Location")));
			this.btnEgreso.Name = "btnEgreso";
			this.btnEgreso.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnEgreso.RightToLeft")));
			this.btnEgreso.Size = ((System.Drawing.Size)(resources.GetObject("btnEgreso.Size")));
			this.btnEgreso.TabIndex = ((int)(resources.GetObject("btnEgreso.TabIndex")));
			this.btnEgreso.Text = resources.GetString("btnEgreso.Text");
			this.btnEgreso.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnEgreso.TextAlign")));
			this.btnEgreso.Visible = ((bool)(resources.GetObject("btnEgreso.Visible")));
			this.btnEgreso.Click += new System.EventHandler(this.btnEgreso_Click);
			// 
			// label5
			// 
			this.label5.AccessibleDescription = resources.GetString("label5.AccessibleDescription");
			this.label5.AccessibleName = resources.GetString("label5.AccessibleName");
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("label5.Anchor")));
			this.label5.AutoSize = ((bool)(resources.GetObject("label5.AutoSize")));
			this.label5.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("label5.Dock")));
			this.label5.Enabled = ((bool)(resources.GetObject("label5.Enabled")));
			this.label5.Font = ((System.Drawing.Font)(resources.GetObject("label5.Font")));
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Image = ((System.Drawing.Image)(resources.GetObject("label5.Image")));
			this.label5.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label5.ImageAlign")));
			this.label5.ImageIndex = ((int)(resources.GetObject("label5.ImageIndex")));
			this.label5.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("label5.ImeMode")));
			this.label5.Location = ((System.Drawing.Point)(resources.GetObject("label5.Location")));
			this.label5.Name = "label5";
			this.label5.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("label5.RightToLeft")));
			this.label5.Size = ((System.Drawing.Size)(resources.GetObject("label5.Size")));
			this.label5.TabIndex = ((int)(resources.GetObject("label5.TabIndex")));
			this.label5.Text = resources.GetString("label5.Text");
			this.label5.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label5.TextAlign")));
			this.label5.Visible = ((bool)(resources.GetObject("label5.Visible")));
			// 
			// label7
			// 
			this.label7.AccessibleDescription = resources.GetString("label7.AccessibleDescription");
			this.label7.AccessibleName = resources.GetString("label7.AccessibleName");
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("label7.Anchor")));
			this.label7.AutoSize = ((bool)(resources.GetObject("label7.AutoSize")));
			this.label7.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("label7.Dock")));
			this.label7.Enabled = ((bool)(resources.GetObject("label7.Enabled")));
			this.label7.Font = ((System.Drawing.Font)(resources.GetObject("label7.Font")));
			this.label7.ForeColor = System.Drawing.Color.Firebrick;
			this.label7.Image = ((System.Drawing.Image)(resources.GetObject("label7.Image")));
			this.label7.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label7.ImageAlign")));
			this.label7.ImageIndex = ((int)(resources.GetObject("label7.ImageIndex")));
			this.label7.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("label7.ImeMode")));
			this.label7.Location = ((System.Drawing.Point)(resources.GetObject("label7.Location")));
			this.label7.Name = "label7";
			this.label7.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("label7.RightToLeft")));
			this.label7.Size = ((System.Drawing.Size)(resources.GetObject("label7.Size")));
			this.label7.TabIndex = ((int)(resources.GetObject("label7.TabIndex")));
			this.label7.Text = resources.GetString("label7.Text");
			this.label7.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label7.TextAlign")));
			this.label7.Visible = ((bool)(resources.GetObject("label7.Visible")));
			// 
			// btnProcesar
			// 
			this.btnProcesar.AccessibleDescription = resources.GetString("btnProcesar.AccessibleDescription");
			this.btnProcesar.AccessibleName = resources.GetString("btnProcesar.AccessibleName");
			this.btnProcesar.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnProcesar.Anchor")));
			this.btnProcesar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnProcesar.BackgroundImage")));
			this.btnProcesar.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnProcesar.Dock")));
			this.btnProcesar.Enabled = ((bool)(resources.GetObject("btnProcesar.Enabled")));
			this.btnProcesar.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnProcesar.FlatStyle")));
			this.btnProcesar.Font = ((System.Drawing.Font)(resources.GetObject("btnProcesar.Font")));
			this.btnProcesar.Image = ((System.Drawing.Image)(resources.GetObject("btnProcesar.Image")));
			this.btnProcesar.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnProcesar.ImageAlign")));
			this.btnProcesar.ImageIndex = ((int)(resources.GetObject("btnProcesar.ImageIndex")));
			this.btnProcesar.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnProcesar.ImeMode")));
			this.btnProcesar.Location = ((System.Drawing.Point)(resources.GetObject("btnProcesar.Location")));
			this.btnProcesar.Name = "btnProcesar";
			this.btnProcesar.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnProcesar.RightToLeft")));
			this.btnProcesar.Size = ((System.Drawing.Size)(resources.GetObject("btnProcesar.Size")));
			this.btnProcesar.TabIndex = ((int)(resources.GetObject("btnProcesar.TabIndex")));
			this.btnProcesar.Text = resources.GetString("btnProcesar.Text");
			this.btnProcesar.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnProcesar.TextAlign")));
			this.btnProcesar.Visible = ((bool)(resources.GetObject("btnProcesar.Visible")));
			this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
			// 
			// btnAutorizar
			// 
			this.btnAutorizar.AccessibleDescription = resources.GetString("btnAutorizar.AccessibleDescription");
			this.btnAutorizar.AccessibleName = resources.GetString("btnAutorizar.AccessibleName");
			this.btnAutorizar.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnAutorizar.Anchor")));
			this.btnAutorizar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnAutorizar.BackgroundImage")));
			this.btnAutorizar.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnAutorizar.Dock")));
			this.btnAutorizar.Enabled = ((bool)(resources.GetObject("btnAutorizar.Enabled")));
			this.btnAutorizar.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnAutorizar.FlatStyle")));
			this.btnAutorizar.Font = ((System.Drawing.Font)(resources.GetObject("btnAutorizar.Font")));
			this.btnAutorizar.Image = ((System.Drawing.Image)(resources.GetObject("btnAutorizar.Image")));
			this.btnAutorizar.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnAutorizar.ImageAlign")));
			this.btnAutorizar.ImageIndex = ((int)(resources.GetObject("btnAutorizar.ImageIndex")));
			this.btnAutorizar.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnAutorizar.ImeMode")));
			this.btnAutorizar.Location = ((System.Drawing.Point)(resources.GetObject("btnAutorizar.Location")));
			this.btnAutorizar.Name = "btnAutorizar";
			this.btnAutorizar.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnAutorizar.RightToLeft")));
			this.btnAutorizar.Size = ((System.Drawing.Size)(resources.GetObject("btnAutorizar.Size")));
			this.btnAutorizar.TabIndex = ((int)(resources.GetObject("btnAutorizar.TabIndex")));
			this.btnAutorizar.Text = resources.GetString("btnAutorizar.Text");
			this.btnAutorizar.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnAutorizar.TextAlign")));
			this.btnAutorizar.Visible = ((bool)(resources.GetObject("btnAutorizar.Visible")));
			this.btnAutorizar.Click += new System.EventHandler(this.btnAutorizar_Click);
			// 
			// label8
			// 
			this.label8.AccessibleDescription = resources.GetString("label8.AccessibleDescription");
			this.label8.AccessibleName = resources.GetString("label8.AccessibleName");
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("label8.Anchor")));
			this.label8.AutoSize = ((bool)(resources.GetObject("label8.AutoSize")));
			this.label8.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("label8.Dock")));
			this.label8.Enabled = ((bool)(resources.GetObject("label8.Enabled")));
			this.label8.Font = ((System.Drawing.Font)(resources.GetObject("label8.Font")));
			this.label8.Image = ((System.Drawing.Image)(resources.GetObject("label8.Image")));
			this.label8.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label8.ImageAlign")));
			this.label8.ImageIndex = ((int)(resources.GetObject("label8.ImageIndex")));
			this.label8.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("label8.ImeMode")));
			this.label8.Location = ((System.Drawing.Point)(resources.GetObject("label8.Location")));
			this.label8.Name = "label8";
			this.label8.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("label8.RightToLeft")));
			this.label8.Size = ((System.Drawing.Size)(resources.GetObject("label8.Size")));
			this.label8.TabIndex = ((int)(resources.GetObject("label8.TabIndex")));
			this.label8.Text = resources.GetString("label8.Text");
			this.label8.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("label8.TextAlign")));
			this.label8.Visible = ((bool)(resources.GetObject("label8.Visible")));
			// 
			// cmbNombre
			// 
			this.cmbNombre.AccessibleDescription = resources.GetString("cmbNombre.AccessibleDescription");
			this.cmbNombre.AccessibleName = resources.GetString("cmbNombre.AccessibleName");
			this.cmbNombre.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("cmbNombre.Anchor")));
			this.cmbNombre.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("cmbNombre.BackgroundImage")));
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCajaChica, "_CajaChica.idpersonal"));
			this.cmbNombre.DisplayLayout.AutoFitColumns = true;
			ultraGridBand10.AddButtonCaption = resources.GetString("ultraGridBand10.AddButtonCaption");
			ultraGridBand10.AddButtonToolTipText = resources.GetString("ultraGridBand10.AddButtonToolTipText");
			ultraGridColumn74.Format = resources.GetString("ultraGridColumn74.Format");
			ultraGridColumn74.Header.Caption = resources.GetString("resource.Caption74");
			ultraGridColumn74.Header.VisiblePosition = 1;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn74.MaskInput = resources.GetString("ultraGridColumn74.MaskInput");
			ultraGridColumn74.NullText = resources.GetString("ultraGridColumn74.NullText");
			ultraGridColumn75.Format = resources.GetString("ultraGridColumn75.Format");
			ultraGridColumn75.Header.Caption = resources.GetString("resource.Caption75");
			ultraGridColumn75.Header.VisiblePosition = 0;
			ultraGridColumn75.MaskInput = resources.GetString("ultraGridColumn75.MaskInput");
			ultraGridColumn75.NullText = resources.GetString("ultraGridColumn75.NullText");
			ultraGridColumn75.Width = 301;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn74,
																											ultraGridColumn75});
			ultraGridBand10.SummaryFooterCaption = resources.GetString("ultraGridBand10.SummaryFooterCaption");
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombre.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("cmbNombre.Dock")));
			this.cmbNombre.Enabled = ((bool)(resources.GetObject("cmbNombre.Enabled")));
			this.cmbNombre.Font = ((System.Drawing.Font)(resources.GetObject("cmbNombre.Font")));
			this.cmbNombre.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("cmbNombre.ImeMode")));
			this.cmbNombre.Location = ((System.Drawing.Point)(resources.GetObject("cmbNombre.Location")));
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("cmbNombre.RightToLeft")));
			this.cmbNombre.Size = ((System.Drawing.Size)(resources.GetObject("cmbNombre.Size")));
			this.cmbNombre.TabIndex = ((int)(resources.GetObject("cmbNombre.TabIndex")));
			this.cmbNombre.Text = resources.GetString("cmbNombre.Text");
			this.cmbNombre.ValueMember = "idPersonal";
			this.cmbNombre.Visible = ((bool)(resources.GetObject("cmbNombre.Visible")));
			// 
			// txtidProyecto
			// 
			this.txtidProyecto.AccessibleDescription = resources.GetString("txtidProyecto.AccessibleDescription");
			this.txtidProyecto.AccessibleName = resources.GetString("txtidProyecto.AccessibleName");
			this.txtidProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("txtidProyecto.Anchor")));
			this.txtidProyecto.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("txtidProyecto.BackgroundImage")));
			this.txtidProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidProyecto.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("txtidProyecto.Dock")));
			this.txtidProyecto.Enabled = ((bool)(resources.GetObject("txtidProyecto.Enabled")));
			this.txtidProyecto.Font = ((System.Drawing.Font)(resources.GetObject("txtidProyecto.Font")));
			this.txtidProyecto.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("txtidProyecto.ImeMode")));
			this.txtidProyecto.Location = ((System.Drawing.Point)(resources.GetObject("txtidProyecto.Location")));
			this.txtidProyecto.Name = "txtidProyecto";
			this.txtidProyecto.PromptChar = ' ';
			this.txtidProyecto.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("txtidProyecto.RightToLeft")));
			this.txtidProyecto.Size = ((System.Drawing.Size)(resources.GetObject("txtidProyecto.Size")));
			this.txtidProyecto.TabIndex = ((int)(resources.GetObject("txtidProyecto.TabIndex")));
			this.txtidProyecto.Visible = ((bool)(resources.GetObject("txtidProyecto.Visible")));
			this.txtidProyecto.ValueChanged += new System.EventHandler(this.txtidProyecto_ValueChanged);
			// 
			// txtBodega
			// 
			this.txtBodega.AccessibleDescription = resources.GetString("txtBodega.AccessibleDescription");
			this.txtBodega.AccessibleName = resources.GetString("txtBodega.AccessibleName");
			this.txtBodega.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("txtBodega.Anchor")));
			this.txtBodega.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("txtBodega.BackgroundImage")));
			this.txtBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodega.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("txtBodega.Dock")));
			this.txtBodega.Enabled = ((bool)(resources.GetObject("txtBodega.Enabled")));
			this.txtBodega.Font = ((System.Drawing.Font)(resources.GetObject("txtBodega.Font")));
			this.txtBodega.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("txtBodega.ImeMode")));
			this.txtBodega.Location = ((System.Drawing.Point)(resources.GetObject("txtBodega.Location")));
			this.txtBodega.Name = "txtBodega";
			this.txtBodega.PromptChar = ' ';
			this.txtBodega.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("txtBodega.RightToLeft")));
			this.txtBodega.Size = ((System.Drawing.Size)(resources.GetObject("txtBodega.Size")));
			this.txtBodega.TabIndex = ((int)(resources.GetObject("txtBodega.TabIndex")));
			this.txtBodega.Visible = ((bool)(resources.GetObject("txtBodega.Visible")));
			this.txtBodega.ValueChanged += new System.EventHandler(this.txtBodega_ValueChanged);
			// 
			// cmbPersonal
			// 
			this.cmbPersonal.AccessibleDescription = resources.GetString("cmbPersonal.AccessibleDescription");
			this.cmbPersonal.AccessibleName = resources.GetString("cmbPersonal.AccessibleName");
			this.cmbPersonal.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("cmbPersonal.Anchor")));
			this.cmbPersonal.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("cmbPersonal.BackgroundImage")));
			this.cmbPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonal.DataSource = this.ultraDataSource8;
			this.cmbPersonal.DisplayLayout.AutoFitColumns = true;
			ultraGridBand11.AddButtonCaption = resources.GetString("ultraGridBand11.AddButtonCaption");
			ultraGridBand11.AddButtonToolTipText = resources.GetString("ultraGridBand11.AddButtonToolTipText");
			ultraGridColumn76.Format = resources.GetString("ultraGridColumn76.Format");
			ultraGridColumn76.Header.Caption = resources.GetString("resource.Caption76");
			ultraGridColumn76.Header.VisiblePosition = 0;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.MaskInput = resources.GetString("ultraGridColumn76.MaskInput");
			ultraGridColumn76.NullText = resources.GetString("ultraGridColumn76.NullText");
			ultraGridColumn76.Width = 168;
			ultraGridColumn77.Format = resources.GetString("ultraGridColumn77.Format");
			ultraGridColumn77.Header.Caption = resources.GetString("resource.Caption77");
			ultraGridColumn77.Header.VisiblePosition = 1;
			ultraGridColumn77.MaskInput = resources.GetString("ultraGridColumn77.MaskInput");
			ultraGridColumn77.NullText = resources.GetString("ultraGridColumn77.NullText");
			ultraGridColumn77.Width = 406;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn76,
																											ultraGridColumn77});
			ultraGridBand11.SummaryFooterCaption = resources.GetString("ultraGridBand11.SummaryFooterCaption");
			this.cmbPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbPersonal.DisplayMember = "Nombre";
			this.cmbPersonal.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("cmbPersonal.Dock")));
			this.cmbPersonal.Enabled = ((bool)(resources.GetObject("cmbPersonal.Enabled")));
			this.cmbPersonal.Font = ((System.Drawing.Font)(resources.GetObject("cmbPersonal.Font")));
			this.cmbPersonal.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("cmbPersonal.ImeMode")));
			this.cmbPersonal.Location = ((System.Drawing.Point)(resources.GetObject("cmbPersonal.Location")));
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("cmbPersonal.RightToLeft")));
			this.cmbPersonal.Size = ((System.Drawing.Size)(resources.GetObject("cmbPersonal.Size")));
			this.cmbPersonal.TabIndex = ((int)(resources.GetObject("cmbPersonal.TabIndex")));
			this.cmbPersonal.Text = resources.GetString("cmbPersonal.Text");
			this.cmbPersonal.ValueMember = "idPersonal";
			this.cmbPersonal.Visible = ((bool)(resources.GetObject("cmbPersonal.Visible")));
			// 
			// ultraDataSource8
			// 
			ultraDataColumn25.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// frmCajaChica
			// 
			this.AccessibleDescription = resources.GetString("$this.AccessibleDescription");
			this.AccessibleName = resources.GetString("$this.AccessibleName");
			this.AutoScaleBaseSize = ((System.Drawing.Size)(resources.GetObject("$this.AutoScaleBaseSize")));
			this.AutoScroll = ((bool)(resources.GetObject("$this.AutoScroll")));
			this.AutoScrollMargin = ((System.Drawing.Size)(resources.GetObject("$this.AutoScrollMargin")));
			this.AutoScrollMinSize = ((System.Drawing.Size)(resources.GetObject("$this.AutoScrollMinSize")));
			this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
			this.ClientSize = ((System.Drawing.Size)(resources.GetObject("$this.ClientSize")));
			this.Controls.Add(this.cmbPersonal);
			this.Controls.Add(this.txtidProyecto);
			this.Controls.Add(this.txtBodega);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.btnAutorizar);
			this.Controls.Add(this.btnProcesar);
			this.Controls.Add(this.btnEgreso);
			this.Controls.Add(this.btnCalcular);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.cmbRetencionesIVA2);
			this.Controls.Add(this.cmbRetencionesRenta2);
			this.Controls.Add(this.cmbRetencionesIVA1);
			this.Controls.Add(this.cmbRetencionesRenta1);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.cmbCentroDeCosto);
			this.Controls.Add(this.txtResponsable);
			this.Controls.Add(this.cmbArticulo1);
			this.Controls.Add(this.txtCajaChica);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.cmbArticulo2);
			this.Controls.Add(this.txtIdCajaChica);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.ultraGrid1);
			this.Enabled = ((bool)(resources.GetObject("$this.Enabled")));
			this.Font = ((System.Drawing.Font)(resources.GetObject("$this.Font")));
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("$this.ImeMode")));
			this.KeyPreview = true;
			this.Location = ((System.Drawing.Point)(resources.GetObject("$this.Location")));
			this.MaximumSize = ((System.Drawing.Size)(resources.GetObject("$this.MaximumSize")));
			this.MinimumSize = ((System.Drawing.Size)(resources.GetObject("$this.MinimumSize")));
			this.Name = "frmCajaChica";
			this.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("$this.RightToLeft")));
			this.StartPosition = ((System.Windows.Forms.FormStartPosition)(resources.GetObject("$this.StartPosition")));
			this.Text = resources.GetString("$this.Text");
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCajaChica_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmCajaChica_Closing);
			this.Load += new System.EventHandler(this.frmCajaChica_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCajaChica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCajaChica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCajaChica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetencionesRenta1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetencionesIVA1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetencionesIVA2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetencionesRenta2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				this.btnAutorizar.Enabled = false;
				this.btnProcesar.Enabled = false;
				this.btnEgreso.Enabled = false;
				this.btnImprimir.Enabled = false;

				drCajaChica = cdsCajaChica.TableViews["CajaChica"].Rows[0];

				if ((int)drCajaChica["Estado"] == 0) this.lblEstado.Text = "EN ESPERA";
				if ((int)drCajaChica["Estado"] == 1) this.lblEstado.Text = "PROCESADO";
				if ((int)drCajaChica["Estado"] == 2) this.lblEstado.Text = "AUTORIZADO EL PAGO";
				if ((int)drCajaChica["Estado"] == 3) this.lblEstado.Text = "CANCELADO EL PAGO";
				if ((int)drCajaChica["Estado"] == 4) this.lblEstado.Text = "AUTORIZADO";
				if ((int)drCajaChica["Estado"] == 5) this.lblEstado.Text = "ANULADO";

				this.txtCajaChica.Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select CajaChica From Proyecto Where idProyecto = {0}", (int)this.cmbProyecto.Value));
				
				if ((int)drCajaChica["Estado"] == 0)
				{
					this.btnAutorizar.Text = "&Autorizar";
					if (miAcceso.BAPAutorizar) this.btnAutorizar.Enabled = true; 
				}

				if ((int)drCajaChica["Estado"] == 4)
				{
					if (miAcceso.Autorizacion) this.btnProcesar.Enabled = true; 
					if (miAcceso.BAPAutorizar) this.btnAutorizar.Enabled = true;

					this.btnAutorizar.Text = "&Quitar Autorización";
				}
				
				if ((int)drCajaChica["Estado"] == 3)
					if (miAcceso.Autorizacion)
						this.btnEgreso.Enabled = true;
				
				if ((int)drCajaChica["Estado"] != 1 && (int)drCajaChica["Estado"] != 4) 
					this.btnImprimir.Enabled = true;  

				this.Text = "Flujo de Caja Chica - " + this.ultraGrid1.Rows.Count + " Registros Ingresados";	
			}
			catch{};
		}

		private void cdsCajaChica_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			if (e.DataMember.ToString() == "_CajaChica")
			{
				drCajaChica = e.Row;
			}
		}

		private void cdsCajaChica_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCajaChica.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCuenta.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoC.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirConta);	
		}

		private void Totales()
		{
			decimal dIVA1 = 0.00m;
			decimal dSubtotal1 = 0.00m;
			decimal dRetencionIVA1 = 0.00m;
			decimal dRetencionIVA2 = 0.00m;
			decimal dRetenidoIVA = 0.00m;
			decimal dRetencionRenta1 = 0.00m;
			decimal dRetencionRenta2 = 0.00m;
			decimal dRetenidoRenta = 0.00m;
			decimal dIVA2 = 0.00m;
			decimal dSubtotal2 = 0.00m;
			decimal dCalcSubtotal1 = 0.00m;
			decimal dCalcSubtotal2 = 0.00m;
			decimal dRetenido = 0.00m;
			
			drCajaChica["Total"] = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{				
				dr.Cells["RetenidoRenta"].Value = 0.00m;
				dr.Cells["RetenidoIVA"].Value = 0.00m;
				dr.Cells["Retenido"].Value = 0.00m;

				dIVA1 = Convert.ToDecimal(dr.Cells["IVA1"].Value);
				dIVA2 = Convert.ToDecimal(dr.Cells["IVA2"].Value);
				dSubtotal1 = Convert.ToDecimal(dr.Cells["Subtotal1"].Value);
				dSubtotal2 = Convert.ToDecimal(dr.Cells["Subtotal2"].Value);
				
				dCalcSubtotal1 = dSubtotal1 + Math.Round((dSubtotal1 * dIVA1) / 100, 2);
				dCalcSubtotal2 = dSubtotal2 + Math.Round((dSubtotal2 * dIVA2) / 100, 2);

				dRetencionRenta1 = Convert.ToDecimal(dr.Cells["ValRetRenta1"].Value);
				dRetencionRenta2 = Convert.ToDecimal(dr.Cells["ValRetRenta2"].Value);
				dRetencionIVA1 = Convert.ToDecimal(dr.Cells["ValRetIVA1"].Value);
				dRetencionIVA2 = Convert.ToDecimal(dr.Cells["ValRetIVA2"].Value);

				dRetenidoRenta = dRetencionRenta1 + dRetencionRenta2;
				dRetenidoIVA = dRetencionIVA1 + dRetencionIVA2;        
				dRetenido = dRetenidoRenta + dRetenidoIVA;

				dr.Cells["RetenidoRenta"].Value = dRetenidoRenta;
				dr.Cells["RetenidoIVA"].Value = dRetenidoIVA;
				dr.Cells["Retenido"].Value = dRetenido;

				dr.Cells["Total"].Value = dCalcSubtotal1 + dCalcSubtotal2;
				//if (Math.Round(Convert.ToDecimal(dr.Cells["Total"].Value, 2) >= 40) dr.Cells["Retencion"].Value = true;  
				if (!(bool)dr.Cells["RetencionAsumida"].Value) dr.Cells["Suma"].Value = dCalcSubtotal1 + dCalcSubtotal2 - dRetenido;
				drCajaChica["Total"] = Math.Round(Convert.ToDecimal(drCajaChica["Total"]), 2) + Math.Round(Convert.ToDecimal(dr.Cells["Suma"].Value), 2);
			}
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void Seguridades()
		{
			string stTipo = "0425";
			string stPermiso = "67";
			
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			
			if (!Funcion.Permiso(stPermiso,cdsSeteo))
			{
				MessageBox.Show("No puede ingresar a Flujo de Caja Chica.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if(miAcceso.BAccesoBoton) bEmpresa = true;
		} 

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}
		private void CargaLocal()
		{
			if(bEmpresa)
			{
				string sSQLBP = string.Format("Exec BodegaPorUsuarioCajaChica {0}", MenuLatinium.IdUsuario);		
				iBodegaPredef = Funcion.iEscalarSQL(cdsSeteoF, sSQLBP);

				if (iBodegaPredef > 0) 
				{
					string Periodo  = DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
					this.txtidProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select isNull(idProyecto, 0) From Bodega Where Bodega = {0}", iBodegaPredef));
					this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaProyectos {0}", (int)this.txtidProyecto.Value));
					this.cmbCentroDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaProyectos {0}", (int)this.txtidProyecto.Value));
					this.cmbNombre.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaPersonalCajaChica {0},'{1}'",iBodegaPredef,Periodo));
					this.cmbPersonal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaPersonalCajaChica {0},'{1}'",iBodegaPredef,Periodo));
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["SF"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["AP"].CellActivation = Activation.NoEdit;

					#region Agencias
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idDetalleCajaChica"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idCajaChica"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idCompra"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idProveedor"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["RUC"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["RazonSocial"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Fecha"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Serie"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Numero"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Autorización"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idDetCompra1"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idArticulo1"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["IVA1"].CellActivation = Activation.NoEdit;
					//this.ultraGrid1.DisplayLayout.Bands[0].Columns["Subtotal1"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idDetCompra2"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idArticulo2"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["IVA2"].CellActivation = Activation.NoEdit;
					//this.ultraGrid1.DisplayLayout.Bands[0].Columns["Subtotal2"].CellActivation = Activation.ActivateOnly;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.NoEdit;
					//this.ultraGrid1.DisplayLayout.Bands[0].Columns["Detalle"].CellActivation = Activation.ActivateOnly;
					//this.ultraGrid1.DisplayLayout.Bands[0].Columns["idProyecto"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["SF"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Retencion"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idRetencion"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["AutorizacionRet"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["SerieRet"].CellActivation = Activation.AllowEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["NumeroRet"].CellActivation = Activation.AllowEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Retenido"].CellActivation = Activation.AllowEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Facturado"].CellActivation = Activation.AllowEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Imprime"].CellActivation = Activation.AllowEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["FechaSistema"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Usuario"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Suma"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idArticuloSRI1"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["ValRetRenta1"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idRetIVA1"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["ValRetIVA1"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idArticuloSRI2"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["ValRetRenta2"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idRetIVA2"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["ValRetIVA2"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["RetenidoRenta"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["RetenidoIVA"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["AP"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["idAsiento"].CellActivation = Activation.NoEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["RetencionAsumida"].CellActivation = Activation.NoEdit;					
					#endregion Agencias 
				}	
			}
		}		

		int idArtAP = 0;
		private void frmCajaChica_Load(object sender, System.EventArgs e)
		{
//			if (MenuLatinium.iNivel == 23 || MenuLatinium.iNivel == 25 || MenuLatinium.iNivel == 13 )
//			{
//				if (MenuLatinium.stIdDB==1 || MenuLatinium.stIdDB==3 )
//				{
//				}
//			}

			Seguridades();

			#region Lectura de Configuracion
			drSeteoC = cdsSeteoC.TableViews["SeteoC"].Rows[0];
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			#endregion Lectura de Configuracion
			
			barraDato1.PosUltima();
			barraDato1.HabilitaControles(false);
			ModificaGrilla(false);

			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			
			if (!bEmpresa)
			{
				this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order by Nombre");
				this.cmbCentroDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order by Nombre");
				this.cmbNombre.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From Personal Order By Nombre");
			}

			CargaLocal();
			
			this.cmbArticulo1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArticulo, Codigo, Articulo From Articulo Where idTipoGrupo In (2, 3, 13, 16) order by Articulo");
			this.cmbArticulo2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArticulo, Codigo, Articulo From Articulo Where idTipoGrupo In (2, 3, 13, 16)order by Articulo");
			this.cmbRetencionesRenta1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArticuloSRI, Codigo, Porcentaje, Detalle From ArticuloSRI Where Activo = 1 Order By Codigo");
			this.cmbRetencionesRenta2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArticuloSRI, Codigo, Porcentaje, Detalle From ArticuloSRI Where Activo = 1 Order By Codigo");
			this.cmbRetencionesIVA1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idRetIva, Codigo, Porcentaje, Detalle From RetencionIva Where idRetIva > 1 Order By Codigo");
			this.cmbRetencionesIVA2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idRetIva, Codigo, Porcentaje, Detalle From RetencionIva Where idRetIva > 1 Order By Codigo");	
			this.cmbPersonal.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"Select idPersonal, Nombre From Personal Order by Nombre");
			idArtAP = Funcion.iEscalarSQL(cdsSeteoF, "Select idArticulo From Articulo Where DetpContabilidad = 1");

			if (idCajaChica > 0)
			{
				barraDato1.IdRegistro = idCajaChica;
				barraDato1.ProximoId(5);
			}
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene Acceso a Crear.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!miAcceso.CambiarFecha)
			{ 
				this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;   
				this.dtFecha.CalendarInfo.MinDate = DateTime.Today; 
			}

			this.barraDato1.CrearRegistro();

			CargaLocal();

			int idCC = 0;
			idCC = Funcion.iEscalarSQL(cdsSeteoF, "Exec ProyectoPredefinido 49, 1", true);
			if (idCC > 0)
			{
				this.cmbProyecto.Value = idCC;	
				this.cmbProyecto.Enabled = false;
			}

			this.txtTotal.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtCajaChica.Enabled = false;
			this.txtCajaChica.Value = 0;

			ModificaGrilla(true);

			drCajaChica["Fecha"] = DateTime.Today;
			drCajaChica["Estado"] = 0;

			this.btnCalcular.Enabled = true;		
			this.btnAutorizar.Enabled = false;
			this.btnExportar.Enabled = false;

			this.cmbProyecto.Focus();

		}
		
		bool bEliminoFila = false;

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			if (bEliminoFila)
			{
				MessageBox.Show("Elimino Filas Debe Grabar la Caja Chica", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else
			{
				this.barraDato1.DeshacerRegistro();
				ModificaGrilla(false);
				this.btnCalcular.Enabled = false;
				this.btnExportar.Enabled = true;
				
				if (!this.barraDato1.bNuevo)
				{
					string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 2", idBloqueaTransacciones, (int)drCajaChica["idCajaChica"]);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo);
				}

				idBloqueaTransacciones = 0;
			}
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.IdRegistro = (int)drCajaChica["idCajaChica"];
			this.barraDato1.ProximoId(5);

			int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From CajaChica Where idCajaChica = {0}", (int)drCajaChica["idCajaChica"]));

			#region EC Valida que pertenesca al CDC
			if (bEmpresa)
			{
				int IBodega = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec BodegaAsignadaPorUsuarioBuscaCajaChica {0}", MenuLatinium.IdUsuario));

				if ( IBodega > 0 )
				{
					int Existencia = Funcion.iEscalarSQL(cdsSeteoF,string.Format("select idProyecto from dbo.bodega where Bodega = {0}", IBodega));
					if ((int) this.cmbProyecto.Value != Existencia)
					{
						MessageBox.Show("No puede editar una Caja Chica diferente al Asignado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
						return;
					}
				}
			}
			#endregion EC Valida que pertenesca al CDC

			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene Acceso a Editar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}			
			else if (iEstado == 1)
			{
				MessageBox.Show("No puede editar una Caja Chica Procesada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if (iEstado == 2)
			{
				MessageBox.Show("No puede editar una Caja Chica Autorizado el Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if (iEstado == 3)
			{
				MessageBox.Show("No puede editar una Caja Chica que ya fue pagada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if (iEstado == 4)
			{
				MessageBox.Show("No puede editar una Caja Chica que ya fue autorizada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede Editar esta Caja Chica, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			else if (!Funcion.ValidaBloqueo((int)drCajaChica["idCajaChica"], 2, cdsSeteoF))
			{			
				this.barraDato1.EditarRegistro();				

				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 2", (int)drCajaChica["idCajaChica"]);
				idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

				this.txtTotal.Enabled = false;
				this.txtNumero.Enabled = false;
				this.txtCajaChica.Enabled = false;
				this.btnCalcular.Enabled = true;

				if (!miAcceso.CambiarFecha)
					this.dtFecha.Enabled = false;
	
				//this.cmbProyecto.Enabled = false;
				this.btnProcesar.Enabled = false;
				this.btnAutorizar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnExportar.Enabled = false;

				//if (Funcion.iEscalarSQL(cdsSeteoF, "Exec ProyectoPredefinido 49, 1", true) > 0) this.cmbProyecto.Enabled = false;

				if (iEstado == 0) 
					this.ModificaGrilla(true);
			}
		}

		private void Validacion()
		{
			#region Validación Cabecera
			#region Centro de costo
			if (this.cmbProyecto.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione el Centro de Costo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.cmbProyecto.Focus();
				return;
			}
			#endregion Centro de costo
			#region Fecha
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.dtFecha.Focus();
				return;
			}
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			#endregion Centro de costo
			#endregion Validación Cabecera

			#region Validación Detalle
			#region Elimina filas vacias
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drEl in ultraGrid1.Rows.All)
			{			
				if (Convert.ToDecimal(drEl.Cells["IVA1"].Value) == 0 && Convert.ToDecimal(drEl.Cells["IVA2"].Value) == 0 && Convert.ToDecimal(drEl.Cells["Subtotal1"].Value) == 0 && Convert.ToDecimal(drEl.Cells["Subtotal2"].Value) == 0)
				{
					bEliminaAlValidar = true;
					drEl.Delete();					
				}
			}
			#endregion Elimina filas vacias

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{				
				#region Ingreso Con Factura
				if (!(bool)dr.Cells["SF"].Value && !(bool)dr.Cells["AP"].Value)
				{
					#region Valida Ingreso de RUC
					if (dr.Cells["RUC"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el RUC del proveedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Valida Ingreso de RUC

					#region Valida Personal
					if (dr.Cells["idPersonal"].Value == System.DBNull.Value || (int)dr.Cells["idPersonal"].Value == 0)

					{
						MessageBox.Show("Ingrese el Nombre del Vendedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Valida Personal

					#region Valida identificación
//					int idTipoRUC = 1;
//					if (dr.Cells["RUC"].Value.ToString().Length == 10) idTipoRUC = 2;
//					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", dr.Cells["RUC"].Value.ToString(), idTipoRUC);
//					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
//				
//					if (VlCed.StartsWith("Error"))
//					{
//						MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//						this.ultraGrid1.ActiveRow = dr;
//						this.ultraGrid1.ActiveRow.Selected = true;
//						return;
//					}
					#endregion Valida identificación

					#region Busca Proveedor
					if (dr.Cells["RUC"].Value.ToString().Length > 0)
					{
						string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", dr.Cells["RUC"].Value.ToString(), 4, 0, false);
						SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
						drC.Read();
						if (!drC.HasRows)
						{
							drC.Close();
							MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						drC.Close();
					}
					#endregion Busca Proveedor

					#region Autorización
					if (dr.Cells["Autorización"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese la Autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Autorización"].Value.ToString().Length != 10)
					{
						if (dr.Cells["Autorización"].Value.ToString().Length != 37)
						{
							MessageBox.Show("La Autorización es de 10 o 37 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					}
					#endregion Autorización

					#region Serie
					if (dr.Cells["Serie"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese la Serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Serie"].Value.ToString().Length < 6)
					{
						MessageBox.Show("La Serie debe ser de 6 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Serie

					#region Numero
					if (dr.Cells["Numero"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el número de la Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}

					char Pad = '0';
					dr.Cells["Numero"].Value = dr.Cells["Numero"].Value.ToString().PadLeft(9, Pad);
					
					int idDCCValN = 0;
					int idCompraValN = 0;

					if (dr.Cells["idDetalleCajaChica"].Value != System.DBNull.Value) idDCCValN = (int)dr.Cells["idDetalleCajaChica"].Value; 
					if (dr.Cells["idCompra"].Value != System.DBNull.Value) idCompraValN = (int)dr.Cells["idCompra"].Value; 

					string sSQLValNum = string.Format("Exec ValidaNumeracionCompraCajaChica {0}, {1}, {2}, '{3}', '{4}', '{5}'", 
						idDCCValN, idCompraValN, (int)dr.Cells["idProveedor"].Value, dr.Cells["Serie"].Value.ToString(), dr.Cells["Numero"].Value.ToString(), dr.Cells["Autorización"].Value.ToString());
					string sMensValNum = Funcion.sEscalarSQL(cdsSeteoF, sSQLValNum, true);
					if (sMensValNum.Length > 0)
					{
						MessageBox.Show(sMensValNum, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Numero
				}
				#endregion Ingreso Con Factura

				if (!(bool)dr.Cells["SF"].Value && (bool)dr.Cells["AP"].Value)
				{
					#region Valida Ingreso de RUC
					if (dr.Cells["RUC"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el RUC del proveedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Valida Ingreso de RUC

					#region Valida identificación
//					int idTipoRUC = 1;
//					if (dr.Cells["RUC"].Value.ToString().Length == 10) idTipoRUC = 2;
//					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", dr.Cells["RUC"].Value.ToString(), idTipoRUC);
//					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
//				
//					if (VlCed.StartsWith("Error"))
//					{
//						MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//						this.ultraGrid1.ActiveRow = dr;
//						this.ultraGrid1.ActiveRow.Selected = true;
//						return;
//					}
					#endregion Valida identificación

					#region Busca Proveedor
					if (dr.Cells["RUC"].Value.ToString().Length > 0)
					{
						string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", dr.Cells["RUC"].Value.ToString(), 4, 0, false);
						SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
						drC.Read();
						if (!drC.HasRows)
						{
							drC.Close();
							MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						drC.Close();
					}
					#endregion Busca Proveedor
				}

				#region Fecha
				if (dr.Cells["Fecha"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
				if (Funcion.ValidaPeriodo((DateTime)dr.Cells["Fecha"].Value, cdsSeteoF, "Compra"))
				{
					DateTime dtFechaVP = (DateTime)dr.Cells["Fecha"].Value;
					MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
				#endregion Fecha

				#region Detalle
				if (dr.Cells["Detalle"].Value.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Detalle de la Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}					
				#endregion Detalle

				#region Articulo
				if (dr.Cells["idArticulo1"].Value == System.DBNull.Value && dr.Cells["idArticulo2"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Articulo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}					
				#endregion Articulo

				#region Base Imponible
				if (Convert.ToDecimal(dr.Cells["Subtotal1"].Value) == 0 && Convert.ToDecimal(dr.Cells["Subtotal2"].Value) == 0)
				{
					MessageBox.Show("Ingrese el Valor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}					
				#endregion Base Imponible

				#region Valida Valores de Gastos
				if (dr.Cells["idArticulo1"].Value == System.DBNull.Value)
				{
					if (Convert.ToDecimal(dr.Cells["IVA1"].Value) > 0)
					{
						MessageBox.Show("Seleccione un Gasto o Ponga el valor a Cero en el IVA del Gasto 1", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}

					if (Convert.ToDecimal(dr.Cells["Subtotal1"].Value) > 0)
					{
						MessageBox.Show("Seleccione un Gasto o Ponga el valor a Cero en el Valor del Gasto 1", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}

				if (dr.Cells["idArticulo2"].Value == System.DBNull.Value)
				{
					if (Convert.ToDecimal(dr.Cells["IVA2"].Value) > 0)
					{
						MessageBox.Show("Seleccione un Gasto o Ponga el valor a Cero en el IVA del Gasto 2", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}

					if (Convert.ToDecimal(dr.Cells["Subtotal2"].Value) > 0)
					{
						MessageBox.Show("Seleccione un Gasto o Ponga el valor a Cero en el Valor del Gasto 2", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Valida Valores de Gastos
				
				#region Porcentaje IVA
				if ((bool)dr.Cells["SF"].Value)
				{
					if (Convert.ToDecimal(dr.Cells["IVA1"].Value) > 0)
					{
						MessageBox.Show("Si es un Gasto sin Factura el Porcentaje de IVA es 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						dr.Cells["IVA2"].Value = 0;
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
						//Totales();
					}
					if (Convert.ToDecimal(dr.Cells["IVA2"].Value) > 0)
					{
						MessageBox.Show("Si es un Gasto sin Factura el Porcentaje de IVA es 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						dr.Cells["IVA2"].Value = 0;
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
						//Totales();
					}
				}
				#endregion Porcentaje IVA

				#region Retenciones
				#region Retencion Asumida
				//				if (!(bool)dr.Cells["SF"].Value && !(bool)dr.Cells["AP"].Value)
				//				{
				//					if (!(bool)dr.Cells["Retencion"].Value || !(bool)dr.Cells["RetencionAsumida"].Value)
				//					{
				//						if (Convert.ToDecimal(dr.Cells["Total"].Value >= 40.00)
				//						{
				//							MessageBox.Show("Para Facturas con Valor Mayor a 40.00 debe Crear Retención", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//							this.ultraGrid1.ActiveRow = dr;
				//							this.ultraGrid1.ActiveRow.Selected = true;
				//							return;
				//						}
				//					}
				//				}
				#endregion Retencion Asumida

				#region Valida Retenciones
				if (!(bool)dr.Cells["SF"].Value && !(bool)dr.Cells["AP"].Value)
				{
					if ((bool)dr.Cells["Retencion"].Value)
					{
						#region Autorización
						if (dr.Cells["AutorizacionRet"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Ingrese la Autorización de la retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						if (dr.Cells["AutorizacionRet"].Value.ToString().Length < 10)
						{
							MessageBox.Show("La Autorización de la retención es de 10 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						#endregion Autorización

						#region Serie
						if (dr.Cells["SerieRet"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Ingrese la Serie de la retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						if (dr.Cells["SerieRet"].Value.ToString().Length < 6)
						{
							MessageBox.Show("La Serie de la retención debe ser de 6 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						#endregion Serie

						#region Numero
						if (dr.Cells["NumeroRet"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Ingrese el número de la retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					
						char Padd = '0';
						dr.Cells["NumeroRet"].Value = dr.Cells["NumeroRet"].Value.ToString().PadLeft(9, Padd);

						#region Valida Numero de retencion repetida
						int idDCCValR = 0;
						int idCompraValR = 0;

						if (dr.Cells["idDetalleCajaChica"].Value != System.DBNull.Value) idDCCValR = (int)dr.Cells["idDetalleCajaChica"].Value; 
						if (dr.Cells["idCompra"].Value != System.DBNull.Value) idCompraValR = (int)dr.Cells["idCompra"].Value; 
						
						string sSQLValida = string.Format("Exec ValidaNumeroRetencionCompraCajaChica {0}, {1}, '{2}', '{3}', '{4}'", 
							idDCCValR, idCompraValR, dr.Cells["SerieRet"].Value.ToString(), dr.Cells["NumeroRet"].Value.ToString(), dr.Cells["AutorizacionRet"].Value.ToString());
						string sMensValNumRet = Funcion.sEscalarSQL(cdsSeteoF, sSQLValida, true);
						if (sMensValNumRet.Length > 0)
						{
							MessageBox.Show(sMensValNumRet, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						#endregion Valida Numero de retencion repetida
						#endregion Numero
					}
				}
				#endregion Valida Retenciones
				#endregion Retenciones
			}
			#endregion Validación Detalle
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			this.Totales();

			#region Validación Cabecera
			#region Centro de costo
			if (this.cmbProyecto.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Centro de Costo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.cmbProyecto.Focus();
				return;
			}
			#endregion Centro de costo
			#region Fecha
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.dtFecha.Focus();
				return;
			}
			//			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			//			{
			//				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
			//				MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				this.dtFecha.Focus();
			//				return;
			//			}
			#endregion Centro de costo
			#region Nombre
			if (this.cmbNombre.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el nombre", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.cmbNombre.Focus();
				return;
			}
			#endregion Nombre
			#endregion Validación Cabecera
			
			#region Validación Detalle
			#region Elimina filas vacias
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drEl in ultraGrid1.Rows.All)
			{			
				if (Convert.ToDecimal(drEl.Cells["IVA1"].Value) == 0 && Convert.ToDecimal(drEl.Cells["IVA2"].Value) == 0 && Convert.ToDecimal(drEl.Cells["Subtotal1"].Value) == 0 && Convert.ToDecimal(drEl.Cells["Subtotal2"].Value) == 0)
				{
					bEliminaAlValidar = true;
					drEl.Delete();					
				}
			}
			#endregion Elimina filas vacias

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["SF"].Value && !(bool)dr.Cells["AP"].Value)
				{
					if (dr.Cells["idProveedor"].Value != System.DBNull.Value)
					{
						if ((int)dr.Cells["idProveedor"].Value > 0)
						{
							MessageBox.Show("Ha ingresado un proveedor pero el registro va sin factura y no es un anticipo proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					}
				}
			
				#region Ingreso Con Factura
				if (!(bool)dr.Cells["SF"].Value && !(bool)dr.Cells["AP"].Value)
				{
					#region Valida Ingreso de RUC
					if (dr.Cells["RUC"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el RUC del proveedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Valida Ingreso de RUC

					#region Valida identificación
					//					int idTipoRUC = 1;
					//					if (dr.Cells["RUC"].Value.ToString().Length == 10) idTipoRUC = 2;
					//					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", dr.Cells["RUC"].Value.ToString(), idTipoRUC);
					//					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
					//				
					//					if (VlCed.StartsWith("Error"))
					//					{
					//						MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					//						this.ultraGrid1.ActiveRow = dr;
					//						this.ultraGrid1.ActiveRow.Selected = true;
					//						return;
					//					}
					#endregion Valida identificación

					#region Busca Proveedor
					if (dr.Cells["RUC"].Value.ToString().Length > 0)
					{
						string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", dr.Cells["RUC"].Value.ToString(), 4, 0, false);
						SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
						drC.Read();
						if (!drC.HasRows)
						{
							drC.Close();
							MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						drC.Close();
					}
					#endregion Busca Proveedor

					#region Autorización
					if (dr.Cells["Autorización"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese la Autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Autorización"].Value.ToString().Length != 10)
					{
						if (dr.Cells["Autorización"].Value.ToString().Length != 37)
						{
							MessageBox.Show("La Autorización es de 10 o 37 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					}
					#endregion Autorización

					#region Serie
					if (dr.Cells["Serie"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese la Serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Serie"].Value.ToString().Length < 6)
					{
						MessageBox.Show("La Serie debe ser de 6 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Serie

					#region Numero
					if (dr.Cells["Numero"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el número de la Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}

					char Pad = '0';
					dr.Cells["Numero"].Value = dr.Cells["Numero"].Value.ToString().PadLeft(9, Pad);
					
					int idDCCValN = 0;
					int idCompraValN = 0;

					if (dr.Cells["idDetalleCajaChica"].Value != System.DBNull.Value) idDCCValN = (int)dr.Cells["idDetalleCajaChica"].Value; 
					if (dr.Cells["idCompra"].Value != System.DBNull.Value) idCompraValN = (int)dr.Cells["idCompra"].Value; 

					string sSQLValNum = string.Format("Exec ValidaNumeracionCompraCajaChica {0}, {1}, {2}, '{3}', '{4}', '{5}'", 
						idDCCValN, idCompraValN, (int)dr.Cells["idProveedor"].Value, dr.Cells["Serie"].Value.ToString(), dr.Cells["Numero"].Value.ToString(), dr.Cells["Autorización"].Value.ToString());
					string sMensValNum = Funcion.sEscalarSQL(cdsSeteoF, sSQLValNum, true);
					if (sMensValNum.Length > 0)
					{
						MessageBox.Show(sMensValNum, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Numero
				}
				#endregion Ingreso Con Factura

				if (!bEmpresa)
				{
					#region Valida Personal
					if (!(bool)dr.Cells["SF"].Value && !(bool)dr.Cells["AP"].Value)
					{
						if (dr.Cells["idPersonal"].Value == System.DBNull.Value || (int)dr.Cells["idPersonal"].Value == 0)
						{
							MessageBox.Show("Ingrese el Nombre del Vendedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					}
					#endregion Valida Personal

					if (!(bool)dr.Cells["SF"].Value && (bool)dr.Cells["AP"].Value)
					{
						#region Valida Ingreso de RUC
						if (dr.Cells["RUC"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Ingrese el RUC del proveedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}

						#endregion Valida Ingreso de RUC
						
						#region Valida identificación
						//					int idTipoRUC = 1;
						//					if (dr.Cells["RUC"].Value.ToString().Length == 10) idTipoRUC = 2;
						//					
						//
						//					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", dr.Cells["RUC"].Value.ToString(), idTipoRUC);
						//					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
						//				
						//					if (VlCed.StartsWith("Error"))
						//					{
						//						MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						//						this.ultraGrid1.ActiveRow = dr;
						//						this.ultraGrid1.ActiveRow.Selected = true;
						//						return;
						//					}
						#endregion Valida identificación

						#region Busca Proveedor
						if (dr.Cells["RUC"].Value.ToString().Length > 0)
						{
							string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", dr.Cells["RUC"].Value.ToString(), 4, 0, false);
							SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
							drC.Read();
							if (!drC.HasRows)
							{
								drC.Close();
								MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.ultraGrid1.ActiveRow = dr;
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}
							drC.Close();
						}
						#endregion Busca Proveedor
					}

					#region Fecha
					if (dr.Cells["Fecha"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese la fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					if (Funcion.ValidaPeriodo((DateTime)dr.Cells["Fecha"].Value, cdsSeteoF, "Compra"))
					{
						DateTime dtFechaVP = (DateTime)dr.Cells["Fecha"].Value;
						MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					if (!miAcceso.CambiarFecha)
					{
						if (MenuLatinium.stIdDB == 1)
						{
							if (Convert.ToDateTime( dr.Cells["Fecha"].Value) != Convert.ToDateTime(this.dtFecha.Value))
							{
								MessageBox.Show("La Fecha del detalle no puede ser diferente de la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
								this.ultraGrid1.ActiveRow = dr;
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}
						}
					}
					#endregion Fecha

					#region Detalle
					if (dr.Cells["Detalle"].Value.ToString().Length < 5)
					{
						MessageBox.Show("Ingrese el Detalle de la Factura, caracteres minimos 5.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}					
					#endregion Detalle

					#region Articulo
					if (dr.Cells["idArticulo1"].Value == System.DBNull.Value && dr.Cells["idArticulo2"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione un Articulo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}			
					#endregion Articulo

					#region Base Imponible
					if (Convert.ToDecimal(dr.Cells["Subtotal1"].Value) == 0 && Convert.ToDecimal(dr.Cells["Subtotal2"].Value) == 0)
					{
						MessageBox.Show("Ingrese el Valor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}					
					#endregion Base Imponible

					#region Valida Valores de Gastos
					if (dr.Cells["idArticulo1"].Value == System.DBNull.Value)
					{
						if (Convert.ToDecimal(dr.Cells["IVA1"].Value) > 0)
						{
							MessageBox.Show("Seleccione un Gasto o Ponga el valor a Cero en el IVA del Gasto 1", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}

						if (Convert.ToDecimal(dr.Cells["Subtotal1"].Value) > 0)
						{
							MessageBox.Show("Seleccione un Gasto o Ponga el valor a Cero en el Valor del Gasto 1", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					}

					if (dr.Cells["idArticulo2"].Value == System.DBNull.Value)
					{
						if (Convert.ToDecimal(dr.Cells["IVA2"].Value) > 0)
						{
							MessageBox.Show("Seleccione un Gasto o Ponga el valor a Cero en el IVA del Gasto 2", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}

						if (Convert.ToDecimal(dr.Cells["Subtotal2"].Value) > 0)
						{
							MessageBox.Show("Seleccione un Gasto o Ponga el valor a Cero en el Valor del Gasto 2", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					}
					#endregion Valida Valores de Gastos
				
					#region Porcentaje IVA
					if ((bool)dr.Cells["SF"].Value)
					{
						if (Convert.ToDecimal(dr.Cells["IVA1"].Value) > 0)
						{
							MessageBox.Show("Si es un Gasto sin Factura el Porcentaje de IVA es 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							dr.Cells["IVA2"].Value = 0;
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
							//Totales();
						}
						if (Convert.ToDecimal(dr.Cells["IVA2"].Value) > 0)
						{
							MessageBox.Show("Si es un Gasto sin Factura el Porcentaje de IVA es 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							dr.Cells["IVA2"].Value = 0;
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
							//Totales();
						}
					}
					#endregion Porcentaje IVA

					#region Retenciones
					#region Retencion Asumida
					//				if (!(bool)dr.Cells["SF"].Value && !(bool)dr.Cells["AP"].Value)
					//				{
					//					if (!(bool)dr.Cells["Retencion"].Value || !(bool)dr.Cells["RetencionAsumida"].Value)
					//					{
					//						if (Convert.ToDecimal(dr.Cells["Total"].Value >= 40.00)
					//						{
					//							MessageBox.Show("Para Facturas con Valor Mayor a 40.00 debe Crear Retención", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//							this.ultraGrid1.ActiveRow = dr;
					//							this.ultraGrid1.ActiveRow.Selected = true;
					//							return;
					//						}
					//					}
					//				}
					#endregion Retencion Asumida

					#region Valida Retenciones
					if (!(bool)dr.Cells["SF"].Value && !(bool)dr.Cells["AP"].Value)
					{
						if ((bool)dr.Cells["Retencion"].Value)
						{
							#region Autorización
							if (dr.Cells["AutorizacionRet"].Value.ToString().Length == 0)
							{
								MessageBox.Show("Ingrese la Autorización de la retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.ultraGrid1.ActiveRow = dr;
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}
							if (dr.Cells["AutorizacionRet"].Value.ToString().Length < 10)
							{
								MessageBox.Show("La Autorización de la retención es de 10 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.ultraGrid1.ActiveRow = dr;
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}
							#endregion Autorización

							#region Serie
							if (dr.Cells["SerieRet"].Value.ToString().Length == 0)
							{
								MessageBox.Show("Ingrese la Serie de la retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.ultraGrid1.ActiveRow = dr;
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}
							if (dr.Cells["SerieRet"].Value.ToString().Length < 6)
							{
								MessageBox.Show("La Serie de la retención debe ser de 6 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.ultraGrid1.ActiveRow = dr;
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}
							#endregion Serie

							#region Numero
							if (dr.Cells["NumeroRet"].Value.ToString().Length == 0)
							{
								MessageBox.Show("Ingrese el número de la retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.ultraGrid1.ActiveRow = dr;
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}
					
							char Padd = '0';
							dr.Cells["NumeroRet"].Value = dr.Cells["NumeroRet"].Value.ToString().PadLeft(9, Padd);

							#region Valida Numero de retencion repetida
							int idDCCValR = 0;
							int idCompraValR = 0;

							if (dr.Cells["idDetalleCajaChica"].Value != System.DBNull.Value) idDCCValR = (int)dr.Cells["idDetalleCajaChica"].Value; 
							if (dr.Cells["idCompra"].Value != System.DBNull.Value) idCompraValR = (int)dr.Cells["idCompra"].Value; 
						
							string sSQLValida = string.Format("Exec ValidaNumeroRetencionCompraCajaChica {0}, {1}, '{2}', '{3}', '{4}'", 
								idDCCValR, idCompraValR, dr.Cells["SerieRet"].Value.ToString(), dr.Cells["NumeroRet"].Value.ToString(), dr.Cells["AutorizacionRet"].Value.ToString());
							string sMensValNumRet = Funcion.sEscalarSQL(cdsSeteoF, sSQLValida, true);
							if (sMensValNumRet.Length > 0)
							{
								MessageBox.Show(sMensValNumRet, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.ultraGrid1.ActiveRow = dr;
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}
							#endregion Valida Numero de retencion repetida
							#endregion Numero
						}
					}
					#endregion Valida Retenciones
					#endregion Retenciones
				}
				else
				{
					#region Detalle
					if (dr.Cells["Detalle"].Value.ToString().Length < 5)
					{
						MessageBox.Show("Ingrese el Detalle de la Factura, caracteres minimos 5.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}					
					#endregion Detalle
				}
			}
			#endregion Validación Detalle
		
			#region valida Presupuesto Articulo 1
				
			//				if (MenuLatinium.stIdDB == 1)// solo para point 
			//				{
			//					int iCentroDeCosto = 0 ;
			//					int idArticulo1  = 0 ;
			//					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			//					{
			//						if (dr.Cells["idProyecto"].Value != System.DBNull.Value) iCentroDeCosto = (int)dr.Cells["idProyecto"].Value;
			//						if (dr.Cells["idArticulo1"].Value != System.DBNull.Value) idArticulo1 = (int)dr.Cells["idArticulo1"].Value;
			//						if ( iCentroDeCosto > 0 )
			//						{
			//							if ( idArticulo1 > 0 )
			//							{
			//								int idCuenta = 0;
			//								idCuenta = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [RetornaCuentaCajaChica] {0}",idArticulo1));	
			//								if( idCuenta > 0 )
			//								{
			//									int idCuentaExi = 0;
			//									idCuentaExi = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCuentaPresupuestoAsiento] {0}, {1}, '{2}'",
			//										idCuenta, iCentroDeCosto, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));	
			//
			//									if (idCuentaExi > 0 )
			//									{
			//										int iProyectoS = 0;
			//										iProyectoS = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCCPresupuestoAsiento] {0}, '{1}'",
			//											iCentroDeCosto,Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd") ));
			//
			//										if ( iProyectoS > 0 )
			//										{
			//											#region valida valores
			//											decimal Presupuesto = 0;
			//											decimal Utilizado = 0;
			//											decimal Disponible = 0;
			//											string Cuenta = "";
			//											if ( Convert.ToDecimal (dr.Cells["SubTotal1"].Value ) > 0 )
			//											{
			//												string sSQL = string.Format("Exec [ValidaCCPresupuestoAsientoValores] '{0}', {1}, {2}", 
			//													Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), idCuenta, (int) iCentroDeCosto );
			//												SqlDataReader dra = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			//												dra.Read();
			//												if (dra.HasRows)
			//												{
			//													if (dra.GetValue(0) != System.DBNull.Value)Presupuesto = dra.GetDecimal (0);
			//													if (dra.GetValue(1) != System.DBNull.Value)Utilizado = dra.GetDecimal (1);
			//													if (dra.GetValue(2) != System.DBNull.Value)Disponible = dra.GetDecimal (2);
			//													if (dra.GetValue(3) != System.DBNull.Value)Cuenta = dra.GetString (3);
			//												}
			//												dra.Close();
			//											}
			//											if ( Convert.ToDecimal (dr.Cells["SubTotal1"].Value ) > Disponible )
			//											{
			//												this.ultraGrid1.ActiveRow = dr;
			//												this.ultraGrid1.ActiveRow.Selected = true;
			//												MessageBox.Show(string.Format("Ha superado el presupuesto de '{0}' \n\n Presupuesto : {1}  \n\n Utilizado : {2} \n\n Solicitado : {3}", Cuenta, Presupuesto, Utilizado, Convert.ToDecimal (dr.Cells["SubTotal1"].Value )), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);											
			//												return;
			//											}
			//											#endregion valida valores
			//										}           
			//									}
			//								}
			//							}
			//						}
			//					}
			//				}
			#endregion valida Presupuesto Articulo 1

			#region valida Presupuesto Articulo 2
				
			//				if (MenuLatinium.stIdDB == 1)// solo para point 
			//				{
			//					int iCentroDeCosto = 0 ;
			//					int idArticulo1  = 0 ;
			//					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			//					{
			//						if (dr.Cells["idProyecto"].Value != System.DBNull.Value) iCentroDeCosto = (int)dr.Cells["idProyecto"].Value;
			//						if (dr.Cells["idArticulo2"].Value != System.DBNull.Value) idArticulo1 = (int)dr.Cells["idArticulo2"].Value;
			//						if ( iCentroDeCosto > 0 )
			//						{
			//							if ( idArticulo1 > 0 )
			//							{
			//								int idCuenta = 0;
			//								idCuenta = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [RetornaCuentaCajaChica] {0}",idArticulo1));	
			//								if( idCuenta > 0 )
			//								{
			//									int idCuentaExi = 0;
			//									idCuentaExi = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCuentaPresupuestoAsiento] {0}, {1}, '{2}'",
			//										idCuenta, iCentroDeCosto, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));	
			//
			//									if (idCuentaExi > 0 )
			//									{
			//										int iProyectoS = 0;
			//										iProyectoS = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCCPresupuestoAsiento] {0}, '{1}'",
			//											iCentroDeCosto,Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd") ));
			//
			//										if ( iProyectoS > 0 )
			//										{
			//											#region valida valores
			//											decimal Presupuesto = 0;
			//											decimal Utilizado = 0;
			//											decimal Disponible = 0;
			//											string Cuenta = "";
			//											if ( Convert.ToDecimal (dr.Cells["SubTotal2"].Value ) > 0 )
			//											{
			//												string sSQL = string.Format("Exec [ValidaCCPresupuestoAsientoValores] '{0}', {1}, {2}", 
			//													Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), idCuenta, (int) iCentroDeCosto );
			//												SqlDataReader dra = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			//												dra.Read();
			//												if (dra.HasRows)
			//												{
			//													if (dra.GetValue(0) != System.DBNull.Value)Presupuesto = dra.GetDecimal (0);
			//													if (dra.GetValue(1) != System.DBNull.Value)Utilizado = dra.GetDecimal (1);
			//													if (dra.GetValue(2) != System.DBNull.Value)Disponible = dra.GetDecimal (2);
			//													if (dra.GetValue(3) != System.DBNull.Value)Cuenta = dra.GetString (3);
			//												}
			//												dra.Close();
			//											}
			//											if ( Convert.ToDecimal (dr.Cells["SubTotal2"].Value ) > Disponible )
			//											{
			//												this.ultraGrid1.ActiveRow = dr;
			//												this.ultraGrid1.ActiveRow.Selected = true;
			//												MessageBox.Show(string.Format("Ha superado el presupuesto de '{0}' \n\n Presupuesto : {1}  \n\n Utilizado : {2} \n\n Solicitado : {3}", Cuenta, Presupuesto, Utilizado, Convert.ToDecimal (dr.Cells["SubTotal1"].Value )), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);											
			//												return;
			//											}
			//											#endregion valida valores
			//										}           
			//									}
			//								}
			//							}
			//						}
			//					}
			//				}
			#endregion valida Presupuesto Articulo 2		
			
			#region Grabar
			if (this.barraDato1.bNuevo) 
			{				
				string sSQLNumero = string.Format("Exec NumeracionLocales 49, {0}, 0", (int)this.cmbProyecto.Value);
				drCajaChica["Numero"] = Funcion.sEscalarSQL(cdsSeteoF, sSQLNumero);
			}

			this.barraDato1.GrabaRegistro();

			if (!this.barraDato1.bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 2", idBloqueaTransacciones, (int)drCajaChica["idCajaChica"]);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo);
			}

			MessageBox.Show("Registro guardado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.barraDato1.ProximoId(5);
			#endregion Grabar

			this.btnCalcular.Enabled = false;	
			this.btnExportar.Enabled = true;

			ModificaGrilla(false);

			bEliminoFila = false;
			
			idBloqueaTransacciones = 0;
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Buscar)
			{
				MessageBox.Show("No tiene Acceso a Buscar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);			
			}			
			else
			{
				using (frmBuscaCajaChica miBusqueda = new frmBuscaCajaChica())
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						try
						{
							barraDato1.IdRegistro = miBusqueda.IdCajaChica;
							barraDato1.ProximoId(5);
							CargaLocal();
							ModificaGrilla(false);
						}
						catch{}
					}
				}
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{				
				switch(e.KeyValue)
				{	
					#region Proveedores
					case (int) Keys.F3:
						if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "RUC" && !(bool)this.ultraGrid1.ActiveRow.Cells["SF"].Value && barraDato1.bEditar)
						{
							using (BuscaClientes miBusqueda = new BuscaClientes(this.ultraGrid1.ActiveRow.Cells["RUC"].Value.ToString(), 4, 0, 1, 0, false))
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{	
									//this.ultraGrid1.ActiveRow.Cells["idProveedor"].Value = miBusqueda.idCliente;
									this.ultraGrid1.ActiveRow.Cells["RUC"].Value = miBusqueda.Ruc;
									//this.ultraGrid1.ActiveRow.Cells["RazonSocial"].Value = miBusqueda.Nombre;
								}
							}
						}
						if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "IDCUENTA" && barraDato1.bEditar)
						{
							using (frmBuscaCuentas miBusqueda = new frmBuscaCuentas())
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{
									this.ultraGrid1.ActiveRow.Cells["idCuenta"].Value = miBusqueda.idCuenta;
								}
							}
						}
						break;
					#endregion Proveedores
					#region Page Down Abre lista desplegable
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
						break;
						#endregion Page Down Abre lista desplegable
					#region Flechas
					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
						#endregion Flechas
					#region Enter
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Point Technology");
			}
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "RUC" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "SERIE" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "AUTORIZACIÓN"
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "SERIERET" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERORET" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "AUTORIZACIONRET") Funcion.CamposNumericos(sender, e);
			
			if ((bool)this.ultraGrid1.ActiveRow.Cells["SF"].Value && (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "RUC" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "RAZONSOCIAL" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "AUTORIZACIÓN" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "SERIE" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "AUTORIZACIONRET" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "SERIERET" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERORET")) e.Handled = true;

			if ((bool)this.ultraGrid1.ActiveRow.Cells["AP"].Value && (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "AUTORIZACIÓN" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "SERIE" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "AUTORIZACIONRET" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "SERIERET" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERORET")) e.Handled = true;

			if (!(bool)this.ultraGrid1.ActiveRow.Cells["Retencion"].Value && (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "AUTORIZACIONRET" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "SERIERET" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERORET")) e.Handled = true;
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
						
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico

			#region Decimales para el total
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decTotal"]);
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decTotal"]);
			e.Layout.Bands[0].Columns["IVA1"].Format = stFormato; 
			e.Layout.Bands[0].Columns["IVA1"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["IVA2"].Format = stFormato; 
			e.Layout.Bands[0].Columns["IVA2"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Subtotal1"].Format = stFormato; 
			e.Layout.Bands[0].Columns["Subtotal1"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Subtotal2"].Format = stFormato; 
			e.Layout.Bands[0].Columns["Subtotal2"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Total"].Format = stFormato; 
			e.Layout.Bands[0].Columns["Total"].MaskInput = stInput;
			#endregion Decimales para el total
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			bActualiza = true;

			if (e.Cell.Column.ToString() == "SF")
			{
				if ((bool)e.Cell.Row.Cells["SF"].Value)
				{
					e.Cell.Row.Cells["AP"].Value = false;
				}
			}

			if (e.Cell.Column.ToString() == "AP")
			{
				if ((bool)e.Cell.Row.Cells["AP"].Value)
				{
					e.Cell.Row.Cells["SF"].Value = false;
				}
			}

			if (e.Cell.Column.ToString() == "Fecha")
			{
				if ((DateTime)e.Cell.Row.Cells["Fecha"].Value > (DateTime)this.dtFecha.Value)
				{
					MessageBox.Show("La Fecha del Documento no puede ser Mayor a la Fecha de la Caja Chica", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cell.Row.Cells["Fecha"].Value = System.DBNull.Value;
				}
			}

			if (e.Cell.Column.ToString() == "RUC")
			{
				#region Valida identificación
//				if (e.Cell.Row.Cells["RUC"].Value.ToString().Length > 0)
//				{
//					int idTipoRUC = 1;
//					if (e.Cell.Row.Cells["RUC"].Value.ToString().Length == 10) idTipoRUC = 2;
//					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", e.Cell.Row.Cells["RUC"].Value.ToString(), idTipoRUC);
//					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
//				
//					if (VlCed.StartsWith("Error"))
//					{
//						MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
//						e.Cell.Row.Cells["RUC"].Value = "";
//						return;
//					}
//				}
				#endregion Valida identificación

				#region Busca Proveedor
				if (e.Cell.Row.Cells["RUC"].Value.ToString().Length > 0)
				{
					string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", e.Cell.Row.Cells["RUC"].Value.ToString(), 4, 0, false);
					SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
					drC.Read();
					if (drC.HasRows)
					{
						e.Cell.Row.Cells["idProveedor"].Value = (int) drC.GetValue(0);
						e.Cell.Row.Cells["RazonSocial"].Value = drC.GetString(2);	
			
						drC.Close();

						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaGrupoProveedor '{0}'", e.Cell.Row.Cells["RUC"].Value.ToString())) == 0)
						{
							MessageBox.Show("En la Ficha Proveedores Seleccione el Grupo al cual pertenece el Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cell.Row.Cells["idProveedor"].Value = System.DBNull.Value;
							e.Cell.Row.Cells["RUC"].Value = "";
							e.Cell.Row.Cells["RazonSocial"].Value = "";
							return;
						}
					}
					else
					{
						drC.Close();
						MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						e.Cell.Row.Cells["RUC"].Value = "";
						return;
					}
					drC.Close();
				}
				#endregion Busca Proveedor
			}

			#region Autorización
			if (e.Cell.Column.ToString() == "Autorización")
			{
				if (e.Cell.Row.Cells["Autorización"].Value.ToString().Length > 0)
				{
					if (e.Cell.Row.Cells["Autorización"].Value.ToString().Length != 10)
					{
						if (e.Cell.Row.Cells["Autorización"].Value.ToString().Length != 37)
						{
							MessageBox.Show("La Autorización debe ser de 10 o 37 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							e.Cell.Row.Cells["Autorización"].Value = "";
							return;
						}
					}
				}
			}
			#endregion Autorización

			#region Serie
			if (e.Cell.Column.ToString() == "Serie")
			{
				if (e.Cell.Row.Cells["Serie"].Value.ToString().Length > 0)
				{
					if (e.Cell.Row.Cells["Serie"].Value.ToString().Length < 6)
					{
						MessageBox.Show("La Serie debe ser de 6 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						e.Cell.Row.Cells["Serie"].Value = "";
						return;
					}
				}
			}
			#endregion Serie

			if (e.Cell.Column.ToString() == "IVA1")
			{
				if (!bEmpresa)
				{
					if (e.Cell.Row.Cells["idArticulo1"].Value == System.DBNull.Value)
					{
						e.Cell.Row.Cells["IVA1"].Value = 0;
						MessageBox.Show("Seleccione un Gasto antes de ingresar el Porcentaje de IVA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);										
					}
				
					if (e.Cell.Row.Cells["Fecha"].Value == System.DBNull.Value)
					{
						e.Cell.Row.Cells["IVA1"].Value = 0;
						MessageBox.Show("Seleccione la Fecha del documento antes de ingresar el Porcentaje de IVA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);										
					}
				}

				decimal dIVA1 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(e.Cell.Row.Cells["Fecha"].Value).ToString("yyyyMMdd")));
	
				if (Convert.ToDecimal(e.Cell.Row.Cells["IVA1"].Value) != 0 && Convert.ToDecimal(e.Cell.Row.Cells["IVA1"].Value) != dIVA1)
				{
					e.Cell.Row.Cells["IVA1"].Value = dIVA1;
					MessageBox.Show(string.Format("El valor de IVA es 0 o {0}", dIVA1), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
				}
			}

			if (e.Cell.Column.ToString() == "Subtotal1")
			{
				if (e.Cell.Row.Cells["idArticulo1"].Value == System.DBNull.Value)
				{
					if (Convert.ToDecimal(e.Cell.Row.Cells["Subtotal1"].Value) > 0)
					{
						if (bEmpresa==false)
						{
							e.Cell.Row.Cells["Subtotal1"].Value = 0;
							MessageBox.Show("Seleccione un Gasto antes de ingresar el Valor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
						}
					}
				}
			}

			if (e.Cell.Column.ToString() == "IVA2")
			{				
				if (e.Cell.Row.Cells["idArticulo2"].Value == System.DBNull.Value)
				{	
					if (bEmpresa==false)
					{
						if (Convert.ToDecimal(e.Cell.Row.Cells["IVA2"].Value) > 0)
						{
							e.Cell.Row.Cells["IVA2"].Value = 0;
							MessageBox.Show("Seleccione un Gasto antes de ingresar el Porcentaje de IVA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
						}

						if (e.Cell.Row.Cells["Fecha"].Value == System.DBNull.Value)
						{
							e.Cell.Row.Cells["IVA1"].Value = 0;
							MessageBox.Show("Seleccione la Fecha del documento antes de ingresar el Porcentaje de IVA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);										
						}
					}

					decimal dIVA2 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(e.Cell.Row.Cells["Fecha"].Value).ToString("yyyyMMdd")));
	
					if (Convert.ToDecimal(e.Cell.Row.Cells["IVA2"].Value) != 0 && Convert.ToDecimal(e.Cell.Row.Cells["IVA2"].Value) != dIVA2)
					{
						e.Cell.Row.Cells["IVA2"].Value = dIVA2;
						MessageBox.Show(string.Format("El valor de IVA es 0 o {1}", dIVA2), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
					}
				}
			}

			if (e.Cell.Column.ToString() == "Subtotal2")
			{
				if (e.Cell.Row.Cells["idArticulo2"].Value == System.DBNull.Value)
				{
					if (Convert.ToDecimal(e.Cell.Row.Cells["Subtotal2"].Value) > 0)
					{
						if (bEmpresa==false)
						{
							e.Cell.Row.Cells["Subtotal2"].Value = 0;
							MessageBox.Show("Seleccione un Gasto antes de ingresar el Valor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
					}
				}
			}

			#region Base Imponible - Iva
			if (e.Cell.Column.ToString() == "IVA1")
			{
				if ((bool)e.Cell.Row.Cells["SF"].Value)
				{
					if (Convert.ToDecimal(e.Cell.Row.Cells["IVA1"].Value) > 0) 
					{
						MessageBox.Show("Gasto sin Factura va sin IVA.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						e.Cell.Row.Cells["IVA1"].Value = 0;
						return;
					}
				}

				Totales();
			}

			if (e.Cell.Column.ToString() == "Subtotal1")
			{
				Totales();
			}

			if (e.Cell.Column.ToString() == "IVA2")
			{
				if ((bool)e.Cell.Row.Cells["SF"].Value)
				{
					if (Convert.ToDecimal(e.Cell.Row.Cells["IVA2"].Value) > 0) 
					{
						MessageBox.Show("Gasto sin Factura va sin IVA.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						e.Cell.Row.Cells["IVA2"].Value = 0;
						return;
					}
				}
				Totales();
			}

			if (e.Cell.Column.ToString() == "Subtotal2")
			{
				Totales();
			}
			#endregion Base Imponible - Iva

			#region Autorización de Retencion
			if (e.Cell.Column.ToString() == "AutorizacionRet")
			{
				if (e.Cell.Row.Cells["AutorizacionRet"].Value.ToString().Length > 0)
				{
					if (e.Cell.Row.Cells["AutorizacionRet"].Value.ToString().Length < 10)
					{
						MessageBox.Show("La Autorización de retención debe ser de 10 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						e.Cell.Row.Cells["AutorizacionRet"].Value = "";
						return;
					}
				}
			}
			#endregion Autorización de Retencion

			#region Serie de Retencion
			if (e.Cell.Column.ToString() == "SerieRet")
			{
				if (e.Cell.Row.Cells["SerieRet"].Value.ToString().Length > 0)
				{
					if (e.Cell.Row.Cells["SerieRet"].Value.ToString().Length < 6)
					{
						MessageBox.Show("La Serie de retención debe ser de 6 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						e.Cell.Row.Cells["SerieRet"].Value = "";
						return;
					}
				}
			}
			#endregion Serie de Retencion

			decimal dPorcRenta1 = 0;
			decimal dPorcRenta2 = 0;
			decimal dPorcIVA1 = 0;
			decimal dPorcIVA2 = 0;

			if (e.Cell.Column.ToString() == "idArticuloSRI1")
			{
				if (e.Cell.Row.Cells["idArticuloSRI1"].Value != System.DBNull.Value)
				{
					dPorcRenta1 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Porcentaje From ArticuloSRI Where idArticuloSRI = {0}", (int)e.Cell.Row.Cells["idArticuloSRI1"].Value));
					e.Cell.Row.Cells["ValRetRenta1"].Value = Math.Round((Convert.ToDecimal(e.Cell.Row.Cells["Subtotal1"].Value) * dPorcRenta1)/100, 2);

					Totales();
				}
				else 
				{
					e.Cell.Row.Cells["ValRetRenta1"].Value = 0;
					Totales();
				}
			}

			if (e.Cell.Column.ToString() == "idArticuloSRI2")
			{
				if (e.Cell.Row.Cells["idArticuloSRI2"].Value != System.DBNull.Value)
				{
					dPorcRenta2 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Porcentaje From ArticuloSRI Where idArticuloSRI = {0}", (int)e.Cell.Row.Cells["idArticuloSRI2"].Value));
					e.Cell.Row.Cells["ValRetRenta2"].Value = Math.Round((Convert.ToDecimal(e.Cell.Row.Cells["Subtotal2"].Value) * dPorcRenta2)/100, 2);

					Totales();
				}
				else 
				{
					e.Cell.Row.Cells["ValRetRenta2"].Value = 0;
					Totales();
				}
			}

			if (e.Cell.Column.ToString() == "idRetIVA1")
			{
				if (Convert.ToDecimal(e.Cell.Row.Cells["IVA1"].Value) > 0)
				{
					if (e.Cell.Row.Cells["idRetIVA1"].Value != System.DBNull.Value)
					{
						dPorcIVA1 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Porcentaje From RetencionIVA Where idRetIva = {0}", (int)e.Cell.Row.Cells["idRetIVA1"].Value));
						e.Cell.Row.Cells["ValRetIVA1"].Value = Math.Round((Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["Subtotal1"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["IVA1"].Value))/100), 2) * dPorcIVA1)/100, 2);

						Totales();
					}
					else 
					{
						e.Cell.Row.Cells["ValRetIVA1"].Value = 0;
						Totales();
					}
				}
			}

			if (e.Cell.Column.ToString() == "idRetIVA2")
			{
				if (Convert.ToDecimal(e.Cell.Row.Cells["IVA2"].Value) > 0)
				{
					if (e.Cell.Row.Cells["idRetIVA2"].Value != System.DBNull.Value)
					{
						dPorcIVA2 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Porcentaje From RetencionIVA Where idRetIva = {0}", (int)e.Cell.Row.Cells["idRetIVA2"].Value));
						e.Cell.Row.Cells["ValRetIVA2"].Value = Math.Round((Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["Subtotal2"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["IVA2"].Value))/100), 2) * dPorcIVA2)/100, 2);

						Totales();
					}
					else 
					{
						e.Cell.Row.Cells["ValRetIVA2"].Value = 0;
						Totales();
					}
				}
			}

			//Totales();

			/*if (e.Cell.Column.ToString() == "NumeroRet")
			{
				if (e.Cell.Row.Cells["NumeroRet"].Value.ToString().Length > 0)
				{
					char Pad = '0';
					e.Cell.Row.Cells["NumeroRet"].Value = e.Cell.Row.Cells["NumeroRet"].Value.ToString().PadLeft(9, Pad);
				}

				#region Valida Numero de retencion repetida
				string sSQLValida = string.Format("Select COUNT(*) From Retencion Where Numero = '{0}' And Autorizacion = '{1}' And Serie = '{2}'", e.Cell.Row.Cells["NumeroRet"].Value.ToString(), e.Cell.Row.Cells["AutorizacionRet"].Value.ToString(), e.Cell.Row.Cells["SerieRet"].Value.ToString());
				if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValida) > 0)
				{
					MessageBox.Show("El número de retención ya esta ingresado en el sistema.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					e.Cell.Row.Cells["NumeroRet"].Value = "";
					return;
				}
				#endregion Valida Numero de retencion repetida
			}*/
			bActualiza = false;
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			this.ultraGrid1.UpdateData();
						
			if (e.Cell.Column.ToString() == "SF")
			{
//				if (this.ultraGrid1.ActiveRow.Cells["RUC"].Value.ToString().Length > 0 || this.ultraGrid1.ActiveRow.Cells["RazonSocial"].Value.ToString().Length > 0 
//					|| this.ultraGrid1.ActiveRow.Cells["Autorización"].Value.ToString().Length > 0 || this.ultraGrid1.ActiveRow.Cells["Serie"].Value.ToString().Length > 0 
//					|| this.ultraGrid1.ActiveRow.Cells["Numero"].Value.ToString().Length > 0)
//				{
//					MessageBox.Show("No puede crear un gasto sin Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//					e.Cancel= true;
//					return;
//				}
			}
			
			if (e.Cell.Column.ToString() == "Retencion")
			{
				if ((bool)e.Cell.Row.Cells["SF"].Value)
				{
					MessageBox.Show("No puede crear retención sin Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					e.Cancel= true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "idArticuloSRI1" || e.Cell.Column.ToString() == "idArticuloSRI2" || e.Cell.Column.ToString() == "idRetIVA1" || e.Cell.Column.ToString() == "idRetIVA2")
			{
				if ((bool)e.Cell.Row.Cells["SF"].Value)
				{					
					MessageBox.Show("No puede crear retención sin Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					e.Cancel= true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "idRetIVA1")
			{
				if (Convert.ToDecimal(e.Cell.Row.Cells["IVA1"].Value) == 0.00m)
				{					
					MessageBox.Show("No puede retener IVA si el Porcentaje de IVA es 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					e.Cancel= true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "idRetIVA2")
			{
				if (Convert.ToDecimal(e.Cell.Row.Cells["IVA2"].Value) == 0.00m)
				{					
					MessageBox.Show("No puede retener IVA si el Porcentaje de IVA es 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					e.Cancel= true;
					return;
				}
			}

//			if (e.Cell.Column.ToString() == "Retencion")
//			{
//				if ((bool) e.Cell.Row.Cells["Retencion"].Value)
//				{
//					if (Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value, 2) >= 40)
//					{
//						MessageBox.Show("No puede dejar sin Retención Factura con valor mayor o igual a 40 Dolares", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//						e.Cancel= true;
//						return;
//					}
//				}
//			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{				
			if (this.barraDato1.bEditar)
			{
				if (!bEliminaAlValidar)
				{
					if ((int)drCajaChica["Estado"] == 0) 
					{
						if (miAcceso.BEliminaFilaCCEnEspera)
						{
							if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar la línea seleccionada.?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
							{
								for (int i=0; i<e.Rows.Length; i++)
								{						
									string stExec = string.Format("Exec EliminaFacturasCajaChica {0}, {1}, {2}, {3}", 
										(int)e.Rows[i].Cells["idDetalleCajaChica"].Value, (bool)e.Rows[i].Cells["SF"].Value, (bool)e.Rows[i].Cells["AP"].Value, (int)drCajaChica["Estado"]);							
									Funcion.EjecutaSQL(cdsSeteoF, stExec, true);							
								}
								e.DisplayPromptMsg = false;
							}
							else e.Cancel = true;
						}
						else 						
						{
							MessageBox.Show("No tiene Acceso a Eliminar Filas \n\n Consulte al Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cancel = true;
						}
					}
					else if ((int)drCajaChica["Estado"] == 1) 
					{
						if (miAcceso.BEliminaFilaCCProcesada)
						{
							if (DialogResult.Yes ==	MessageBox.Show("Esta Caja Chica ya esta PROCESADA.\n\n¿Esta seguro de borrar la línea seleccionada?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
							{
								for (int i=0; i<e.Rows.Length; i++)
								{
									string stExec = string.Format("Exec EliminaFacturasCajaChica {0}, {1}, {2}, {3}", 
										(int)e.Rows[i].Cells["idDetalleCajaChica"].Value, (bool)e.Rows[i].Cells["SF"].Value, (bool)e.Rows[i].Cells["AP"].Value, (int)drCajaChica["Estado"]);
									Funcion.EjecutaSQL(cdsSeteoF, stExec, true);							
								}
								bEliminoFila = true;
								e.DisplayPromptMsg = false;
							}
							else e.Cancel = true;
						}
						else 
						{
							MessageBox.Show("No puede eliminar Filas De Cajas Chicas Con El Pago Autorizado o Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cancel = true;
						}
					}
					else if ((int)drCajaChica["Estado"] > 1) 
					{
						MessageBox.Show("No puede eliminar Filas De Cajas Chicas Con El Pago Autorizado o Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
					}
				}
				else e.DisplayPromptMsg = false;
			}
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void Imprimir(int iTipo)
		{
			#region Seguridad para imprimir y reimprimir y control de vacios
			if (drCajaChica["idCajaChica"] == System.DBNull.Value)
			{
				MessageBox.Show("No hay informacion para imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir " + Text.Trim(), "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!miAcceso.ReImprimir && (bool) drCajaChica["Impreso"] == true)
			{
				MessageBox.Show("No tiene acceso a ReImprimir " + this.Text.Trim(), "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}			
			#endregion Seguridad para imprimir y reimprimir			ADMINISTRADOR

			Cursor = Cursors.WaitCursor;
			#region Nombre de Reporte y Filtro
			string stFiltro = "{CajaChica.idCajaChica} = " + drCajaChica["idCajaChica"].ToString();
			#endregion Nombre de Reporte y Filtro			

			#region Imprime el reporte
			string stReporte ="FlujoCajaChicaOE.rpt";
			if (MenuLatinium.stIdDB==1  )
			{
				stReporte ="FlujoCajaChica.rpt";
			}

			using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			{			
				// Muestra en pantalla
				if (iTipo == 0) miReporte.ShowDialog();
				if (iTipo >= 1) // Imprime una o mas copias
				{
					string ImpresoraPredefinida = MenuLatinium.ImpresoraNombre;
					string stImpresora = Funcion.sEscalarSQL(cdsSeteoF, "Exec CompraReporteImpresora 'FlujoCajaChica.rpt'");
					if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = stImpresora;
					miReporte.ImprimeRep(1);
					if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = ImpresoraPredefinida;
				}
			}
			#endregion Imprime el reporte

			#region Marca como impreso
			string stGrabaImprime = "Update CajaChica Set Impreso = 1 Where idCajaChica = " + drCajaChica["idCajaChica"].ToString();
			Funcion.EjecutaSQL(cdsCajaChica, stGrabaImprime);
			drCajaChica["Impreso"] = true;
			#endregion Marca como impreso
			Cursor = Cursors.Default;
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			Imprimir(0);
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
			#region Numero de Copias de CompraNumero
			string stCopia = "Select IsNull(Copias, 0) From CompraNumero Where idTipoFactura = 49";
			int iCopias = Funcion.iEscalarSQL(cdsSeteoF, stCopia);
			if (iCopias < 1) iCopias = 1;
			#endregion Numero de Copias de CompraNumero
			
			for (int j=0;j<iCopias;j++)
			{
				Imprimir(iCopias);
			}
		}

		private void cmbProyecto_Leave(object sender, System.EventArgs e)
		{
			int idProyecto = 0;
			if (this.cmbProyecto.Text.ToString().Length > 0) 
			{
				idProyecto = (int)this.cmbProyecto.Value;
				this.txtCajaChica.Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select CajaChica From Proyecto Where idProyecto = {0}", idProyecto));
			}
		}

		private void cmbProyecto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmCajaChica_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) 
			{
				if (this.barraDato1.bEditar)
				{
					if (DialogResult.No == MessageBox.Show("¿Esta seguro de cancelar?, no se guardarán los cambios.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
					this.barraDato1_Deshacer(sender, e);
				}
				else
					Close();
			}
			if (e.KeyCode == Keys.N)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Crear(sender, e);
				}
			}
			if (e.KeyCode == Keys.E)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Editar(sender, e);
				}
			}
			if (e.KeyCode == Keys.B)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Buscar(sender, e);
				}
			}
			if (e.KeyCode == Keys.I)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Imprimir(sender, e);
				}
			}
			if (e.KeyCode == Keys.F12)
			{
				if (this.barraDato1.bEditar)
				{
					this.barraDato1_Grabar(sender, e);
				}
			}
			if (e.KeyCode == Keys.F5)
			{
				if (!this.barraDato1.bEditar)
				{
					barraDato1.IdRegistro = (int)drCajaChica["idCajaChica"];
					barraDato1.ProximoId(5);					
				}
			}
		}

		private void cmbProyecto_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFecha.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtResponsable.Focus();
		}

		private void txtNotas_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["SF"];
			}
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtResponsable_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtResponsable_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNotas.Focus();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{			
			if (this.barraDato1.bEditar)
			{
				MessageBox.Show("No puede imprimir en edición.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)drCajaChica["Estado"] == 0)
			{
				MessageBox.Show("No ha procesado las facturas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (!(bool)dr.Cells["SF"].Value)
				{
					iCont ++;
				}
			}

			if (iCont == 0)
			{
				MessageBox.Show("No existen facturas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			int iContImp = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Imprime"].Value)
				{
					iContImp ++;
				}
			}

			if (iContImp == 0)
			{
				MessageBox.Show("No existen facturas con retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			Cursor = Cursors.WaitCursor;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{	
				int IdCompra = 0;
				if (!(bool)dr.Cells["SF"].Value && (bool)dr.Cells["Imprime"].Value) IdCompra = (int)dr.Cells["idCompra"].Value;

				#region Impresion
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfRet = new ParameterField ();
				ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
				pfRet.ParameterFieldName = "@idCompra";
				idCompraRet.Value = IdCompra;
				pfRet.CurrentValues.Add (idCompraRet);
				paramFields.Add (pfRet);
				
				Reporte miRepor = new Reporte("Retenciones.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Retenciones");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				#endregion Impresion
				
				miRepor.Show();						
				Cursor = Cursors.Default;
			}
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order by Nombre");
			this.cmbCentroDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order by Nombre");
			this.cmbArticulo1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArticulo, Codigo, Articulo From Articulo Where idTipoGrupo In (2, 3, 13, 16)");
			this.cmbArticulo2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArticulo, Codigo, Articulo From Articulo Where idTipoGrupo In (2, 3, 13, 16)");
			this.cmbRetencionesRenta1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArticuloSRI, Codigo, Porcentaje, Detalle From ArticuloSRI Order By Codigo");
			this.cmbRetencionesRenta2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArticuloSRI, Codigo, Porcentaje, Detalle From ArticuloSRI Order By Codigo");
			this.cmbRetencionesIVA1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idRetIva, Codigo, Porcentaje, Detalle From RetencionIva Where idRetIva > 1 Order By Codigo");
			this.cmbRetencionesIVA2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idRetIva, Codigo, Porcentaje, Detalle From RetencionIva Where idRetIva > 1 Order By Codigo");
			this.cmbPersonal.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"Select idPersonal, Nombre From Personal Order by Nombre");}
		private void btnCalcular_Click(object sender, System.EventArgs e)
		{			
			Totales();
		}

		private void btnEgreso_Click(object sender, System.EventArgs e)
		{
			if (drCajaChica["idAsiento"] == System.DBNull.Value)
			{
				MessageBox.Show("No existe un egreso creado de esta transacción");	
				return;
			}
			
			Cursor = Cursors.WaitCursor;
			frmAsientos miAsiento = new frmAsientos((int)drCajaChica["idAsiento"]);
			miAsiento.MdiParent = MdiParent;
			miAsiento.Show();
			Cursor = Cursors.Default;			
		}

		private void frmCajaChica_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar)
			{
				MessageBox.Show("No puede Salir La Caja Chica esta en Edicion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
			}
		}

		private void btnProcesar_Click(object sender, System.EventArgs e)
		{	
			this.barraDato1.IdRegistro = (int)drCajaChica["idCajaChica"];
			this.barraDato1.ProximoId(5);

			if ((int)drCajaChica["Estado"] != 4)
			{
				MessageBox.Show("Para procesar una caja chica debe estar autorizada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}			

			if (Funcion.ValidaBloqueo((int)drCajaChica["idCajaChica"], 2, cdsSeteoF)) return;

			#region Validación Cabecera
			#region Centro de costo
			if (this.cmbProyecto.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione el Centro de Costo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.cmbProyecto.Focus();
				return;
			}
			#endregion Centro de costo
			#region Fecha
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.dtFecha.Focus();
				return;
			}
//			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
//			{
//				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
//				MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.dtFecha.Focus();
//				return;
//			}
			#endregion Centro de costo
			#endregion Validación Cabecera

			#region Validación Detalle
			#region Elimina filas vacias
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drEl in ultraGrid1.Rows.All)
			{			
				if (Convert.ToDecimal(drEl.Cells["IVA1"].Value) == 0 && Convert.ToDecimal(drEl.Cells["IVA2"].Value) == 0 && Convert.ToDecimal(drEl.Cells["Subtotal1"].Value) == 0 && Convert.ToDecimal(drEl.Cells["Subtotal2"].Value) == 0)
				{
					bEliminaAlValidar = true;
					drEl.Delete();					
				}
			}
			#endregion Elimina filas vacias

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{				
				#region Ingreso Con Factura
				if (!(bool)dr.Cells["SF"].Value && !(bool)dr.Cells["AP"].Value)
				{
					#region Valida Ingreso de RUC
					if (dr.Cells["RUC"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el RUC del proveedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Valida Ingreso de RUC

					#region Valida Personal
					if (dr.Cells["idPersonal"].Value == System.DBNull.Value || (int)dr.Cells["idPersonal"].Value == 0)

					{
						MessageBox.Show("Ingrese el Nombre del Vendedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Valida Personal
					#region Valida identificación
//					int idTipoRUC = 1;
//					if (dr.Cells["RUC"].Value.ToString().Length == 10) idTipoRUC = 2;
//					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", dr.Cells["RUC"].Value.ToString(), idTipoRUC);
//					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
//				
//					if (VlCed.StartsWith("Error"))
//					{
//						MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//						this.ultraGrid1.ActiveRow = dr;
//						this.ultraGrid1.ActiveRow.Selected = true;
//						return;
//					}
					#endregion Valida identificación

					#region Busca Proveedor
					if (dr.Cells["RUC"].Value.ToString().Length > 0)
					{
						string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", dr.Cells["RUC"].Value.ToString(), 4, 0, false);
						SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
						drC.Read();
						if (!drC.HasRows)
						{
							drC.Close();
							MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						drC.Close();
					}
					#endregion Busca Proveedor

					#region Autorización
					if (dr.Cells["Autorización"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese la Autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Autorización"].Value.ToString().Length != 10)
					{
						if (dr.Cells["Autorización"].Value.ToString().Length != 37)
						{
							MessageBox.Show("La Autorización es de 10 o 37 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					}
					#endregion Autorización

					#region Serie
					if (dr.Cells["Serie"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese la Serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Serie"].Value.ToString().Length < 6)
					{
						MessageBox.Show("La Serie debe ser de 6 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Serie

					#region Numero
					if (dr.Cells["Numero"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el número de la Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}

					char Pad = '0';
					dr.Cells["Numero"].Value = dr.Cells["Numero"].Value.ToString().PadLeft(9, Pad);
					
					int idDCCValN = 0;
					int idCompraValN = 0;

					if (dr.Cells["idDetalleCajaChica"].Value != System.DBNull.Value) idDCCValN = (int)dr.Cells["idDetalleCajaChica"].Value; 
					if (dr.Cells["idCompra"].Value != System.DBNull.Value) idCompraValN = (int)dr.Cells["idCompra"].Value; 

					string sSQLValNum = string.Format("Exec ValidaNumeracionCompraCajaChica {0}, {1}, {2}, '{3}', '{4}', '{5}'", 
						idDCCValN, idCompraValN, (int)dr.Cells["idProveedor"].Value, dr.Cells["Serie"].Value.ToString(), dr.Cells["Numero"].Value.ToString(), dr.Cells["Autorización"].Value.ToString());
					string sMensValNum = Funcion.sEscalarSQL(cdsSeteoF, sSQLValNum, true);
					if (sMensValNum.Length > 0)
					{
						MessageBox.Show(sMensValNum, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Numero
				}
				#endregion Ingreso Con Factura

				if (!(bool)dr.Cells["SF"].Value && (bool)dr.Cells["AP"].Value)
				{
					#region Valida Ingreso de RUC
					if (dr.Cells["RUC"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el RUC del proveedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Valida Ingreso de RUC

					#region Valida identificación
//					int idTipoRUC = 1;
//					if (dr.Cells["RUC"].Value.ToString().Length == 10) idTipoRUC = 2;
//					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", dr.Cells["RUC"].Value.ToString(), idTipoRUC);
//					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
//				
//					if (VlCed.StartsWith("Error"))
//					{
//						MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//						this.ultraGrid1.ActiveRow = dr;
//						this.ultraGrid1.ActiveRow.Selected = true;
//						return;
//					}
					#endregion Valida identificación

					#region Busca Proveedor
					if (dr.Cells["RUC"].Value.ToString().Length > 0)
					{
						string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", dr.Cells["RUC"].Value.ToString(), 4, 0, false);
						SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
						drC.Read();
						if (!drC.HasRows)
						{
							drC.Close();
							MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						drC.Close();
					}
					#endregion Busca Proveedor
				}

				#region Fecha
				if (dr.Cells["Fecha"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
				if (Funcion.ValidaPeriodo((DateTime)dr.Cells["Fecha"].Value, cdsSeteoF, "Compra"))
				{
					DateTime dtFechaVP = (DateTime)dr.Cells["Fecha"].Value;
					MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
				#endregion Fecha

				#region Detalle
				if (dr.Cells["Detalle"].Value.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Detalle de la Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}					
				#endregion Detalle

				#region Articulo
				if (dr.Cells["idArticulo1"].Value == System.DBNull.Value && dr.Cells["idArticulo2"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Articulo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}					
				#endregion Articulo

				#region Base Imponible
				if (Convert.ToDecimal(dr.Cells["Subtotal1"].Value) == 0 && Convert.ToDecimal(dr.Cells["Subtotal2"].Value) == 0)
				{
					MessageBox.Show("Ingrese el Valor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}					
				#endregion Base Imponible

				#region Valida Valores de Gastos
				if (dr.Cells["idArticulo1"].Value == System.DBNull.Value)
				{
					if (Convert.ToDecimal(dr.Cells["IVA1"].Value) > 0)
					{
						MessageBox.Show("Seleccione un Gasto o Ponga el valor a Cero en el IVA del Gasto 1", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}

					if (Convert.ToDecimal(dr.Cells["Subtotal1"].Value) > 0)
					{
						MessageBox.Show("Seleccione un Gasto o Ponga el valor a Cero en el Valor del Gasto 1", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}

				if (dr.Cells["idArticulo2"].Value == System.DBNull.Value)
				{
					if (Convert.ToDecimal(dr.Cells["IVA2"].Value) > 0)
					{
						MessageBox.Show("Seleccione un Gasto o Ponga el valor a Cero en el IVA del Gasto 2", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}

					if (Convert.ToDecimal(dr.Cells["Subtotal2"].Value) > 0)
					{
						MessageBox.Show("Seleccione un Gasto o Ponga el valor a Cero en el Valor del Gasto 2", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Valida Valores de Gastos
				
				#region Porcentaje IVA
				if ((bool)dr.Cells["SF"].Value)
				{
					if (Convert.ToDecimal(dr.Cells["IVA1"].Value) > 0)
					{
						MessageBox.Show("Si es un Gasto sin Factura el Porcentaje de IVA es 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						dr.Cells["IVA2"].Value = 0;
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
						//Totales();
					}
					if (Convert.ToDecimal(dr.Cells["IVA2"].Value) > 0)
					{
						MessageBox.Show("Si es un Gasto sin Factura el Porcentaje de IVA es 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						dr.Cells["IVA2"].Value = 0;
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
						//Totales();
					}
				}
				#endregion Porcentaje IVA

				#region Retenciones
				#region Retencion Asumida
				//				if (!(bool)dr.Cells["SF"].Value && !(bool)dr.Cells["AP"].Value)
				//				{
				//					if (!(bool)dr.Cells["Retencion"].Value || !(bool)dr.Cells["RetencionAsumida"].Value)
				//					{
				//						if (Convert.ToDecimal(dr.Cells["Total"].Value >= 40.00)
				//						{
				//							MessageBox.Show("Para Facturas con Valor Mayor a 40.00 debe Crear Retención", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//							this.ultraGrid1.ActiveRow = dr;
				//							this.ultraGrid1.ActiveRow.Selected = true;
				//							return;
				//						}
				//					}
				//				}
				#endregion Retencion Asumida

				#region Valida Retenciones
				if (!(bool)dr.Cells["SF"].Value && !(bool)dr.Cells["AP"].Value)
				{
					if ((bool)dr.Cells["Retencion"].Value)
					{
						#region Autorización
						if (dr.Cells["AutorizacionRet"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Ingrese la Autorización de la retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						if (dr.Cells["AutorizacionRet"].Value.ToString().Length < 10)
						{
							MessageBox.Show("La Autorización de la retención es de 10 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						#endregion Autorización

						#region Serie
						if (dr.Cells["SerieRet"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Ingrese la Serie de la retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						if (dr.Cells["SerieRet"].Value.ToString().Length < 6)
						{
							MessageBox.Show("La Serie de la retención debe ser de 6 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						#endregion Serie

						#region Numero
						if (dr.Cells["NumeroRet"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Ingrese el número de la retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					
						char Padd = '0';
						dr.Cells["NumeroRet"].Value = dr.Cells["NumeroRet"].Value.ToString().PadLeft(9, Padd);

						#region Valida Numero de retencion repetida
						int idDCCValR = 0;
						int idCompraValR = 0;

						if (dr.Cells["idDetalleCajaChica"].Value != System.DBNull.Value) idDCCValR = (int)dr.Cells["idDetalleCajaChica"].Value; 
						if (dr.Cells["idCompra"].Value != System.DBNull.Value) idCompraValR = (int)dr.Cells["idCompra"].Value; 
						
						string sSQLValida = string.Format("Exec ValidaNumeroRetencionCompraCajaChica {0}, {1}, '{2}', '{3}', '{4}'", 
							idDCCValR, idCompraValR, dr.Cells["SerieRet"].Value.ToString(), dr.Cells["NumeroRet"].Value.ToString(), dr.Cells["AutorizacionRet"].Value.ToString());
						string sMensValNumRet = Funcion.sEscalarSQL(cdsSeteoF, sSQLValida, true);
						if (sMensValNumRet.Length > 0)
						{
							MessageBox.Show(sMensValNumRet, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						#endregion Valida Numero de retencion repetida
						#endregion Numero
					}
				}
				#endregion Valida Retenciones
				#endregion Retenciones
			}
			#endregion Validación Detalle

			if (DialogResult.Yes == MessageBox.Show("¿Desea Procesar el Registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3))
			{	
				//				using (frmIngresoBeneficiario miPrestamo = new frmIngresoBeneficiario (7, Convert.ToDecimal(this.txtTotal.Value), (int)drCajaChica["idCajaChica"]))
				//				{
				//					if (DialogResult.OK == miPrestamo.ShowDialog())
				//					{

				DateTime dtFecha = DateTime.Today;
				string sSQLFac = "";
				decimal dBase0 = 0;
				decimal dBaseImp = 0;
				decimal dIVATotal = 0;

				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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
						
						#region Procesado
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drFac in ultraGrid1.Rows.All)
						{	
							#region Compras
							dBase0 = 0;
							dBaseImp = 0;
							dIVATotal = 0;
							
							if (!(bool)drFac.Cells["SF"].Value && !(bool)drFac.Cells["AP"].Value) // && !(bool)drFac.Cells["Facturado"].Value)
							{						
								dtFecha = (DateTime)drFac.Cells["Fecha"].Value;

								int idArticulo1 = 0;
								if (drFac.Cells["idArticulo1"].Value != System.DBNull.Value) idArticulo1 = (int)drFac.Cells["idArticulo1"].Value;
								int idArticulo2 = 0;
								if (drFac.Cells["idArticulo2"].Value != System.DBNull.Value) idArticulo2 = (int)drFac.Cells["idArticulo2"].Value;

								if (Convert.ToDecimal(drFac.Cells["IVA1"].Value) > 0.00m)
								{
									dIVATotal = dIVATotal + ((Convert.ToDecimal(drFac.Cells["Subtotal1"].Value) * Convert.ToDecimal(drFac.Cells["IVA1"].Value))/100);
									dBaseImp = dBaseImp + Convert.ToDecimal(drFac.Cells["Subtotal1"].Value);
								}
								else if (Convert.ToDecimal(drFac.Cells["IVA1"].Value) == 0) dBase0 = dBase0 + Convert.ToDecimal(drFac.Cells["Subtotal1"].Value);

								if (Convert.ToDecimal(drFac.Cells["IVA2"].Value) > 0.00m)
								{
									dIVATotal = dIVATotal + ((Convert.ToDecimal(drFac.Cells["Subtotal2"].Value) * Convert.ToDecimal(drFac.Cells["IVA2"].Value))/100);
									dBaseImp = dBaseImp + Convert.ToDecimal(drFac.Cells["Subtotal2"].Value);
								}
								else if (Convert.ToDecimal(drFac.Cells["IVA2"].Value) == 0.00m) dBase0 = dBase0 + Convert.ToDecimal(drFac.Cells["Subtotal2"].Value);

								int IdCompra = 0;
								int IdDetCompra1 = 0;
								int IdDetCompra2 = 0;
								int IdRetencion = 0;

								if (drFac.Cells["idCompra"].Value != System.DBNull.Value) IdCompra = (int)drFac.Cells["idCompra"].Value; 
								if (drFac.Cells["idDetCompra1"].Value != System.DBNull.Value) IdDetCompra1 = (int)drFac.Cells["idDetCompra1"].Value; 
								if (drFac.Cells["idDetCompra2"].Value != System.DBNull.Value) IdDetCompra2 = (int)drFac.Cells["idDetCompra2"].Value; 
								if (drFac.Cells["idRetencion"].Value != System.DBNull.Value) IdRetencion = (int)drFac.Cells["idRetencion"].Value; 

								sSQLFac = string.Format("Exec CreaCompraDeCajaChica {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}, '{24}', '{25}', '{26}', {27}, '{28}'", 
									(int)drFac.Cells["idDetalleCajaChica"].Value, // 0
									IdCompra, // 1
									(int)drFac.Cells["idProyecto"].Value, // 2
									(int)drFac.Cells["idProveedor"].Value, // 3
									drFac.Cells["RUC"].Value.ToString(), // 4
									drFac.Cells["RazonSocial"].Value.ToString(), // 5
									dtFecha.ToString("yyyyMMdd"), // 6
									drFac.Cells["Autorización"].Value.ToString(), // 7
									drFac.Cells["Serie"].Value.ToString(), // 8
									drFac.Cells["Numero"].Value.ToString(), // 9 
									drFac.Cells["Detalle"].Value.ToString(), // 10
									dBase0, dBaseImp, dIVATotal, Convert.ToDecimal(drFac.Cells["Total"].Value), //14
									IdDetCompra1, idArticulo1, Convert.ToDecimal(drFac.Cells["IVA1"].Value), Convert.ToDecimal(drFac.Cells["Subtotal1"].Value), //  18
									IdDetCompra2, idArticulo2, Convert.ToDecimal(drFac.Cells["IVA2"].Value), Convert.ToDecimal(drFac.Cells["Subtotal2"].Value), // 22
									IdRetencion, drFac.Cells["SerieRet"].Value.ToString(), drFac.Cells["AutorizacionRet"].Value.ToString(), drFac.Cells["NumeroRet"].Value.ToString(), (bool)drFac.Cells["Retencion"].Value, this.txtNumero.Text.ToString());
								oCmdActualiza.CommandText = sSQLFac;
								oCmdActualiza.ExecuteNonQuery();
								//Funcion.EjecutaSQL(cdsSeteoF, sSQLFac);
								//IdCompra = Funcion.iEscalarSQL(cdsSeteoF, sSQLFac);
								
								oCmdActualiza.CommandText = string.Format("Select ISNULL(idCompra, 0) From DetalleCajaChica Where idDetalleCajaChica = {0}", (int)drFac.Cells["idDetalleCajaChica"].Value);
								IdCompra = (int)oCmdActualiza.ExecuteScalar();

								if ((bool)drFac.Cells["Retencion"].Value)
								{
									int idArticuloSRI1 = 0;
									int idArticuloSRI2 = 0;
									int idRetIVA1 = 0;
									int idRetIVA2 = 0;

									if (drFac.Cells["idArticuloSRI1"].Value != System.DBNull.Value) idArticuloSRI1 = (int)drFac.Cells["idArticuloSRI1"].Value; 
									if (drFac.Cells["idArticuloSRI2"].Value != System.DBNull.Value) idArticuloSRI2 = (int)drFac.Cells["idArticuloSRI2"].Value; 
									if (drFac.Cells["idRetIVA1"].Value != System.DBNull.Value) idRetIVA1 = (int)drFac.Cells["idRetIVA1"].Value; 
									if (drFac.Cells["idRetIVA2"].Value != System.DBNull.Value) idRetIVA2 = (int)drFac.Cells["idRetIVA2"].Value; 

									decimal dBaseRetIVA1 = 0;
									decimal dBaseRetIVA2 = 0;

									if (idRetIVA1 > 0) dBaseRetIVA1 = Math.Round((Convert.ToDecimal(drFac.Cells["Subtotal1"].Value) * 12)/100, 2);
									if (idRetIVA2 > 0) dBaseRetIVA2 = Math.Round((Convert.ToDecimal(drFac.Cells["Subtotal2"].Value) * 12)/100, 2);

									string sSQLRetenciones = string.Format("Exec CreaRetencionesComprasCajaChica {0}, {1}, '{2}', '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}",
										IdRetencion, IdCompra, drFac.Cells["SerieRet"].Value.ToString(), drFac.Cells["AutorizacionRet"].Value.ToString(), drFac.Cells["NumeroRet"].Value.ToString(), 
										idArticuloSRI1, Convert.ToDecimal(drFac.Cells["Subtotal1"].Value), Convert.ToDecimal(drFac.Cells["ValRetRenta1"].Value), 
										idRetIVA1, dBaseRetIVA1, Convert.ToDecimal(drFac.Cells["ValRetIVA1"].Value), 
										idArticuloSRI2, Convert.ToDecimal(drFac.Cells["Subtotal2"].Value), Convert.ToDecimal(drFac.Cells["ValRetRenta2"].Value),
										idRetIVA2, dBaseRetIVA2, Convert.ToDecimal(drFac.Cells["ValRetIVA2"].Value), 
										Convert.ToDecimal(drFac.Cells["RetenidoRenta"].Value), Convert.ToDecimal(drFac.Cells["RetenidoIVA"].Value), Convert.ToDecimal(drFac.Cells["Retenido"].Value), (bool)drFac.Cells["RetencionAsumida"].Value,
										(int)drFac.Cells["idProyecto"].Value);
									oCmdActualiza.CommandText = sSQLRetenciones;
									IdRetencion = (int)oCmdActualiza.ExecuteScalar();
									//IdRetencion = Funcion.iEscalarSQL(cdsSeteoF, sSQLRetenciones, true);
								}
								else
								{
									IdRetencion = 0;
									if (drFac.Cells["idRetencion"].Value != System.DBNull.Value) IdRetencion = (int)drFac.Cells["idRetencion"].Value;
									
									oCmdActualiza.CommandText = string.Format("Exec AnulaRetencionComprasCajaChica {0}, {1}", (int)drFac.Cells["idDetalleCajaChica"].Value, IdRetencion);
									oCmdActualiza.ExecuteNonQuery();
								}
								
								oCmdActualiza.CommandText = string.Format("Update DetalleCajaChica Set Facturado = 1, idRetencion = {1} Where idDetalleCajaChica = {0}", (int)drFac.Cells["idDetalleCajaChica"].Value, IdRetencion);
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Compras

							#region Anticipos
							if ((bool)drFac.Cells["AP"].Value && !(bool)drFac.Cells["SF"].Value)
							{						
								DateTime dtFechaAnticipo = DateTime.Today;// (DateTime)drFac.Cells["Fecha"].Value;
								
								int IdAnticipo = 0;
								int idArticulo1 = 0;
								int IdProyecto = 0;
								if (drFac.Cells["idProyecto"].Value != System.DBNull.Value) IdProyecto = (int)drFac.Cells["idProyecto"].Value;
								if (drFac.Cells["idArticulo1"].Value != System.DBNull.Value) idArticulo1 = (int)drFac.Cells["idArticulo1"].Value;
								if (drFac.Cells["idCompra"].Value != System.DBNull.Value) IdAnticipo = (int)drFac.Cells["idCompra"].Value; 

								oCmdActualiza.CommandText = string.Format("Declare @Bodega Int If Exists(Select Bodega From Bodega Where idProyecto = {0}) Select @Bodega = Bodega From Bodega Where idProyecto = {0} Else Set @Bodega = 1 Select @Bodega", IdProyecto);
								int IBodega = (int)oCmdActualiza.ExecuteScalar();
								
								string sNumero = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 25, 1, 0"));

								string sSQLAnticipos = string.Format("Exec GuardaAnticipos {0}, 1, '{1}', {2}, '{3}', {4}, 1, 0, '{5}', 1, {6}", 
									IdAnticipo, dtFechaAnticipo.ToString("yyyyMMdd"), IBodega, sNumero, 
									(int)drFac.Cells["idProveedor"].Value, drFac.Cells["Detalle"].Value.ToString(), Convert.ToDecimal(drFac.Cells["Total"].Value));		
								oCmdActualiza.CommandText = sSQLAnticipos;
								IdAnticipo = (int)oCmdActualiza.ExecuteScalar();
								
								oCmdActualiza.CommandText = string.Format("Update DetalleCajaChica Set idCompra = {0} Where idDetalleCajaChica = {1}", IdAnticipo, (int)drFac.Cells["idDetalleCajaChica"].Value);
								oCmdActualiza.ExecuteNonQuery();								
							}
							#endregion Anticipos

							#region Gastos
							if (!(bool)drFac.Cells["AP"].Value && (bool)drFac.Cells["SF"].Value)
							{						
								DateTime dtFechaGasto = (DateTime)drFac.Cells["Fecha"].Value;
								int IdAsientoGasto = 0;
								if (drFac.Cells["idAsiento"].Value != System.DBNull.Value) IdAsientoGasto = (int)drFac.Cells["idAsiento"].Value; 			
							}
							#endregion Gastos							
						}
						#endregion Procesado
					
						#region Autorización de Pagos
						dtFecha = (DateTime)this.dtFecha.Value;
						string sSQL = string.Format("Exec CreaAutorizacionesDePago 7, {0}, '{1}', '{2}', {3}, 2, '1 DE 1', 'REPOSICIÓN CAJA CHICA N. DOCUMENTO {5} - {4}', 0", 
							(int)drCajaChica["idCajaChica"], dtFecha.ToString("yyyyMMdd"), dtFecha.ToString("yyyyMMdd"), Convert.ToDecimal(drCajaChica["Total"]), this.cmbProyecto.Text.ToString(), drCajaChica["Numero"].ToString());						
						oCmdActualiza.CommandText = sSQL;
						oCmdActualiza.ExecuteNonQuery();	
						#endregion Autorización de Pagos

						oCmdActualiza.CommandText = string.Format("Update CajaChica Set Estado = 1 Where idCajaChica = {0}", (int)drCajaChica["idCajaChica"]);
						oCmdActualiza.ExecuteNonQuery();								

						oTransaction.Commit();
						
						oConexion.Close();

						MessageBox.Show("Proceso Finalizado Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					catch (Exception ex)
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);						
					}					
				}
				
				bEliminoFila = false;

				this.barraDato1.IdRegistro = (int)drCajaChica["idCajaChica"];
				this.barraDato1.ProximoId(5);
				//					}
				//				}
			}
		}

		private void btnAutorizar_Click(object sender, System.EventArgs e)
		{			
			decimal dSumaTotal = 0.00m;
			if (this.ultraGrid1.Rows.Count > 0)
			{
				if ((int)drCajaChica["Estado"] == 4)
				{
					this.barraDato1.IdRegistro = (int)drCajaChica["idCajaChica"];
					this.barraDato1.ProximoId(5);
				
					if ((int)drCajaChica["Estado"] != 4)
					{
						MessageBox.Show("Para autorizar una caja chica el estado debe ser EN ESPERA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de quitar la Autorizacion de esta caja chica?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3))
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update CajaChica Set Estado = 0 Where idCajaChica = {0}", (int)drCajaChica["idCajaChica"]));

						this.barraDato1.IdRegistro = (int)drCajaChica["idCajaChica"];
						this.barraDato1.ProximoId(5);
					}

					this.btnAutorizar.Text = "&Autorizar";
					return;
				}

				this.barraDato1.IdRegistro = (int)drCajaChica["idCajaChica"];
				this.barraDato1.ProximoId(5);
				
				if ((int)drCajaChica["Estado"] > 0)
				{
					MessageBox.Show("Para autorizar una caja chica el estado debe ser EN ESPERA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (Funcion.ValidaBloqueo((int)drCajaChica["idCajaChica"], 2, cdsSeteoF)) return;					
				
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					dSumaTotal = dSumaTotal + Convert.ToDecimal(dr.Cells["Suma"].Value);
			
				if(Convert.ToDecimal(this.txtTotal.Value) != Convert.ToDecimal(dSumaTotal))
				{
					MessageBox.Show("El Total debe Ser mismo que el Total-Retencion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
								
				if (DialogResult.Yes == MessageBox.Show("Antes de Autorizar La Caja Chica Asegurese que los datos registrados son los Correctos\n\n¿Esta seguro de Autorizar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3))
				{
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update CajaChica Set Estado = 4 Where idCajaChica = {0}", (int)drCajaChica["idCajaChica"]));

					this.btnAutorizar.Text = "&Quitar Autorización";

					this.barraDato1.IdRegistro = (int)drCajaChica["idCajaChica"];
					this.barraDato1.ProximoId(5);
				}
			}				
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			this.Totales();
		}

		private void cmbCentroDeCosto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbRetencionesIVA2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtidProyecto_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtBodega_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
      bool bActualiza = false;

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
//			if (bActualiza) return;
//
//			if (e.Cell.Column.ToString() == "SF")
//			{
//				bActualiza = true;
//				this.ultraGrid1.UpdateData();
//
//				bool bDisponible = false;		
//				if ((bool)e.Cell.Row.Cells["SF"].Value) bDisponible = true;
//				else bDisponible = false;
//				if (!bDisponible)
//				{
//					e.Cell.Row.Cells["SF"].Value = (bool)e.Cell.OriginalValue;
//				}
//				bActualiza = false;
//			}

			if (e.Cell.Column.ToString() == "AP")
			{
				this.ultraGrid1.UpdateData();

				if ((bool)e.Cell.Row.Cells["AP"].Value)
				{					
					e.Cell.Row.Cells["idArticulo1"].Value = idArtAP;
				}
				else 
				{				
					e.Cell.Row.Cells["idArticulo1"].Value = System.DBNull.Value;
				}
			}
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (this.barraDato1.bEditar)
			{
				e.Row.Cells["Retencion"].Value = false;
				//e.Row.Cells["Fecha"].Value = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF).Date;
				e.Row.Cells["Fecha"].Value = (DateTime)this.dtFecha.Value;

				if (bEmpresa)
				{				
					//e.Row.Cells["Detalle"].Value = " ";
					e.Row.Cells["idProyecto"].Value = (int) this.cmbProyecto.Value;								
					e.Row.Cells["SF"].Value = true;
				}
			}
		}
	}
}
