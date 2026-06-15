using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de AnexoTransAvanzado.
	/// </summary>
	public class AnexoTransAvanzado :  DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsCompraTabla;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbComprobante;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCompra;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbComprob1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCredTribut;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCredTrib1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNumero;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.Misc.UltraButton btReemplazar;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCliente;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkComprob;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCredTrib;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo CmbComprobR;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCredTribR;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaHasta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl Marco;
		private C1.Data.C1DataView cdvCompra;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCliente1;
		private Infragistics.Win.Misc.UltraButton btSerie;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerie;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutoriza;
		private Infragistics.Win.Misc.UltraButton btAutorizacion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbCaducidad;
		private Infragistics.Win.Misc.UltraButton btCaducidad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVacios;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.Misc.UltraButton btCedula;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsCliente;
		private Infragistics.Win.Misc.UltraButton btClienteProv;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroCosto;
		private C1.Data.C1DataSet cdsCentroCosto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCentroC;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkBodega;
		private DevExpress.XtraEditors.SpinEdit spnBodega;
		private DevExpress.XtraGrid.Views.Grid.GridView grvBalances;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView4;
		private DevExpress.XtraGrid.GridControl grCliente;
		private DevExpress.XtraGrid.Columns.GridColumn idCliente;
		private DevExpress.XtraGrid.Columns.GridColumn Nombre;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn3;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public AnexoTransAvanzado()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraComprobante", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraNumero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraCredTribut", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CredTributario");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraCredTribut", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CredTributario");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraComprobante", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraNumero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Compra", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoLot");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoPer");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente", -1, "cmbCliente1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPorcIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenProd");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpleado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoMoneda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante", -1, "cmbComprob1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario", -1, "cmbCredTrib1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRecibe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContadoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorDolar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorInteres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRevision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Revisado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Otro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalIva");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalExcento");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DerechoDevolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Refrendo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flete");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagosCada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasInicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutImprenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcIva0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Prioridad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutPrincipal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutPresupuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FletePorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BienServicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConvenioInternac");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vuelto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCaducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaServicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaRetDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprob1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depart");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pedido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bulto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoAutorizado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoAut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloquear");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cif");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CifIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Propina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Servicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutoImpSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNiffIngProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NiifValorIng");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IntImplicito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TasaCambio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoDif");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("subTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComisionConfiteca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasAux");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Garantia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipoObra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nivePrioridadRuta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sistemaMonitoreado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sistemaSeguridad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoNIIF");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FactAfectaCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_AnticipoV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_MultasV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FondoGV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_AnticipoP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_MultasP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FondoGP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_AnticipoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_MultasT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FondoGT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngAfectaCostoNiif");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Iva", 33, true, "Compra", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Iva", 33, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalExcento", 54, true, "Compra", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalExcento", 54, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalIva", 53, true, "Compra", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalIva", 53, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Count, null, "idCliente", 4, true, "Compra", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "idCliente", 4, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Descuento", 31, true, "Compra", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Descuento", 31, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Descuento0", 32, true, "Compra", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Descuento0", 32, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraComprobante", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn192 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn193 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn194 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraNumero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn195 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraCredTribut", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn196 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn197 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn198 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CredTributario");
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn199 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn200 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn201 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn202 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn203 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn204 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn205 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn206 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.spnBodega = new DevExpress.XtraEditors.SpinEdit();
			this.chkBodega = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbCentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCentroCosto = new C1.Data.C1DataSet();
			this.chkCentroC = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.chkCredTrib = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkComprob = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbComprobante = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCompraTabla = new C1.Data.C1DataSet();
			this.cmbFechaHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbCredTribut = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.cmbFechaDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbNumero = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.chkVacios = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btCaducidad = new Infragistics.Win.Misc.UltraButton();
			this.cmbCaducidad = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtAutoriza = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btAutorizacion = new Infragistics.Win.Misc.UltraButton();
			this.txtSerie = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btSerie = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.cmbCredTribR = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btReemplazar = new Infragistics.Win.Misc.UltraButton();
			this.CmbComprobR = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btClienteProv = new Infragistics.Win.Misc.UltraButton();
			this.udsCliente = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btCedula = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cmbComprob1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCredTrib1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.Marco = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdvCompra = new C1.Data.C1DataView();
			this.cmbCliente1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.grCliente = new DevExpress.XtraGrid.GridControl();
			this.grvBalances = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView4 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.idCliente = new DevExpress.XtraGrid.Columns.GridColumn();
			this.Nombre = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn3 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNumero)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbCaducidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutoriza)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribR)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbComprobR)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprob1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTrib1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Marco)).BeginInit();
			this.Marco.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBalances)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.spnBodega);
			this.ultraTabPageControl1.Controls.Add(this.chkBodega);
			this.ultraTabPageControl1.Controls.Add(this.cmbCentroCosto);
			this.ultraTabPageControl1.Controls.Add(this.chkCentroC);
			this.ultraTabPageControl1.Controls.Add(this.btExcel);
			this.ultraTabPageControl1.Controls.Add(this.cmbCliente);
			this.ultraTabPageControl1.Controls.Add(this.chkCredTrib);
			this.ultraTabPageControl1.Controls.Add(this.chkComprob);
			this.ultraTabPageControl1.Controls.Add(this.chkCliente);
			this.ultraTabPageControl1.Controls.Add(this.cmbComprobante);
			this.ultraTabPageControl1.Controls.Add(this.cmbFechaHasta);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel3);
			this.ultraTabPageControl1.Controls.Add(this.btGrabar);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel4);
			this.ultraTabPageControl1.Controls.Add(this.cmbCredTribut);
			this.ultraTabPageControl1.Controls.Add(this.btFiltro);
			this.ultraTabPageControl1.Controls.Add(this.cmbFechaDesde);
			this.ultraTabPageControl1.Controls.Add(this.cmbNumero);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(644, 120);
			// 
			// spnBodega
			// 
			this.spnBodega.EditValue = new System.Decimal(new int[] {
																																1,
																																0,
																																0,
																																0});
			this.spnBodega.Location = new System.Drawing.Point(552, 9);
			this.spnBodega.Name = "spnBodega";
			// 
			// spnBodega.Properties
			// 
			this.spnBodega.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnBodega.Properties.MaxValue = new System.Decimal(new int[] {
																																					999,
																																					0,
																																					0,
																																					0});
			this.spnBodega.Properties.MinValue = new System.Decimal(new int[] {
																																					1,
																																					0,
																																					0,
																																					0});
			this.spnBodega.Size = new System.Drawing.Size(72, 20);
			this.spnBodega.TabIndex = 74;
			// 
			// chkBodega
			// 
			this.chkBodega.BackColor = System.Drawing.Color.Transparent;
			this.chkBodega.Checked = true;
			this.chkBodega.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkBodega.Location = new System.Drawing.Point(376, 9);
			this.chkBodega.Name = "chkBodega";
			this.chkBodega.Size = new System.Drawing.Size(168, 21);
			this.chkBodega.TabIndex = 21;
			this.chkBodega.Text = "Toda Bodega";
			this.chkBodega.CheckedChanged += new System.EventHandler(this.chkBodega_CheckedChanged);
			// 
			// cmbCentroCosto
			// 
			this.cmbCentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroCosto.DataMember = "Proyecto";
			this.cmbCentroCosto.DataSource = this.cdsCentroCosto;
			ultraGridBand1.AddButtonCaption = "CompraComprobante";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 296;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
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
																										 ultraGridColumn10});
			this.cmbCentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCentroCosto.DisplayMember = "Nombre";
			this.cmbCentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentroCosto.Enabled = false;
			this.cmbCentroCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroCosto.Location = new System.Drawing.Point(144, 86);
			this.cmbCentroCosto.Name = "cmbCentroCosto";
			this.cmbCentroCosto.Size = new System.Drawing.Size(216, 21);
			this.cmbCentroCosto.TabIndex = 20;
			this.cmbCentroCosto.ValueMember = "idProyecto";
			// 
			// cdsCentroCosto
			// 
			this.cdsCentroCosto.BindingContextCtrl = this;
			this.cdsCentroCosto.DataLibrary = "LibContabilidad";
			this.cdsCentroCosto.DataLibraryUrl = "";
			this.cdsCentroCosto.DataSetDef = "dsProyecto";
			this.cdsCentroCosto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCentroCosto.Name = "cdsCentroCosto";
			this.cdsCentroCosto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCentroCosto.SchemaDef = null;
			// 
			// chkCentroC
			// 
			this.chkCentroC.BackColor = System.Drawing.Color.Transparent;
			this.chkCentroC.Checked = true;
			this.chkCentroC.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCentroC.Location = new System.Drawing.Point(16, 86);
			this.chkCentroC.Name = "chkCentroC";
			this.chkCentroC.Size = new System.Drawing.Size(128, 22);
			this.chkCentroC.TabIndex = 19;
			this.chkCentroC.Text = "Todo Centro Costo";
			this.chkCentroC.CheckedChanged += new System.EventHandler(this.chkCentroC_CheckedChanged);
			// 
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(552, 86);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(75, 25);
			this.btExcel.TabIndex = 18;
			this.btExcel.Text = "Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataMember = "Cliente";
			this.cmbCliente.DataSource = this.cdsCliente;
			ultraGridBand2.AddButtonCaption = "CompraComprobante";
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Width = 296;
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 6;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Enabled = false;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(144, 60);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(216, 21);
			this.cmbCliente.TabIndex = 17;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			// 
			// chkCredTrib
			// 
			this.chkCredTrib.BackColor = System.Drawing.Color.Transparent;
			this.chkCredTrib.Checked = true;
			this.chkCredTrib.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCredTrib.Location = new System.Drawing.Point(16, 34);
			this.chkCredTrib.Name = "chkCredTrib";
			this.chkCredTrib.Size = new System.Drawing.Size(120, 22);
			this.chkCredTrib.TabIndex = 16;
			this.chkCredTrib.Text = "Todo Cred. Tribut.";
			this.chkCredTrib.CheckedChanged += new System.EventHandler(this.chkCredTrib_CheckedChanged);
			// 
			// chkComprob
			// 
			this.chkComprob.BackColor = System.Drawing.Color.Transparent;
			this.chkComprob.Checked = true;
			this.chkComprob.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkComprob.Location = new System.Drawing.Point(16, 9);
			this.chkComprob.Name = "chkComprob";
			this.chkComprob.Size = new System.Drawing.Size(120, 21);
			this.chkComprob.TabIndex = 15;
			this.chkComprob.Text = "Todo Comprobante";
			this.chkComprob.CheckedChanged += new System.EventHandler(this.chkComprob_CheckedChanged);
			// 
			// chkCliente
			// 
			this.chkCliente.BackColor = System.Drawing.Color.Transparent;
			this.chkCliente.Checked = true;
			this.chkCliente.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCliente.Location = new System.Drawing.Point(16, 60);
			this.chkCliente.Name = "chkCliente";
			this.chkCliente.Size = new System.Drawing.Size(88, 22);
			this.chkCliente.TabIndex = 14;
			this.chkCliente.Text = "Todo Cliente";
			this.chkCliente.CheckedChanged += new System.EventHandler(this.chkCliente_CheckedChanged);
			// 
			// cmbComprobante
			// 
			this.cmbComprobante.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbComprobante.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbComprobante.DataMember = "CompraComprobante";
			this.cmbComprobante.DataSource = this.cdsCompraTabla;
			appearance1.BackColor = System.Drawing.Color.White;
			this.cmbComprobante.DisplayLayout.Appearance = appearance1;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 52;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Width = 345;
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn23.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbComprobante.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.cmbComprobante.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbComprobante.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbComprobante.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbComprobante.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.cmbComprobante.DisplayMember = "Comprobante";
			this.cmbComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbComprobante.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbComprobante.Enabled = false;
			this.cmbComprobante.Location = new System.Drawing.Point(144, 9);
			this.cmbComprobante.Name = "cmbComprobante";
			this.cmbComprobante.Size = new System.Drawing.Size(216, 22);
			this.cmbComprobante.TabIndex = 0;
			this.cmbComprobante.ValueMember = "idComprobante";
			this.cmbComprobante.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbComprobante_InitializeLayout);
			// 
			// cdsCompraTabla
			// 
			this.cdsCompraTabla.BindingContextCtrl = this;
			this.cdsCompraTabla.DataLibrary = "LibFacturacion";
			this.cdsCompraTabla.DataLibraryUrl = "";
			this.cdsCompraTabla.DataSetDef = "dsCompraTabla";
			this.cdsCompraTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompraTabla.Name = "cdsCompraTabla";
			this.cdsCompraTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraTabla.SchemaDef = null;
			// 
			// cmbFechaHasta
			// 
			dateButton1.Caption = "Today";
			this.cmbFechaHasta.DateButtons.Add(dateButton1);
			this.cmbFechaHasta.Format = "dd/MMM/yyyy";
			this.cmbFechaHasta.Location = new System.Drawing.Point(424, 60);
			this.cmbFechaHasta.Name = "cmbFechaHasta";
			this.cmbFechaHasta.NonAutoSizeHeight = 23;
			this.cmbFechaHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbFechaHasta.SpinButtonsVisible = true;
			this.cmbFechaHasta.TabIndex = 8;
			this.cmbFechaHasta.Value = new System.DateTime(2006, 6, 17, 0, 0, 0, 0);
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel3.Location = new System.Drawing.Point(376, 34);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(35, 15);
			this.ultraLabel3.TabIndex = 9;
			this.ultraLabel3.Text = "Desde";
			// 
			// btGrabar
			// 
			this.btGrabar.Location = new System.Drawing.Point(552, 60);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(75, 25);
			this.btGrabar.TabIndex = 13;
			this.btGrabar.Text = "Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Location = new System.Drawing.Point(376, 60);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(33, 15);
			this.ultraLabel4.TabIndex = 10;
			this.ultraLabel4.Text = "Hasta";
			// 
			// cmbCredTribut
			// 
			this.cmbCredTribut.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCredTribut.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCredTribut.DataMember = "CompraCredTribut";
			this.cmbCredTribut.DataSource = this.cdsCompraTabla;
			appearance6.BackColor = System.Drawing.Color.White;
			this.cmbCredTribut.DisplayLayout.Appearance = appearance6;
			ultraGridBand4.AddButtonCaption = "CompraComprobante";
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 59;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Width = 255;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbCredTribut.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.cmbCredTribut.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbCredTribut.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCredTribut.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCredTribut.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.cmbCredTribut.DisplayMember = "CredTributario";
			this.cmbCredTribut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCredTribut.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCredTribut.Enabled = false;
			this.cmbCredTribut.Location = new System.Drawing.Point(144, 34);
			this.cmbCredTribut.Name = "cmbCredTribut";
			this.cmbCredTribut.Size = new System.Drawing.Size(216, 22);
			this.cmbCredTribut.TabIndex = 1;
			this.cmbCredTribut.ValueMember = "idCredTributario";
			// 
			// btFiltro
			// 
			this.btFiltro.Location = new System.Drawing.Point(552, 34);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(75, 25);
			this.btFiltro.TabIndex = 11;
			this.btFiltro.Text = "Filtro";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// cmbFechaDesde
			// 
			dateButton2.Caption = "Today";
			this.cmbFechaDesde.DateButtons.Add(dateButton2);
			this.cmbFechaDesde.Format = "dd/MMM/yyyy";
			this.cmbFechaDesde.Location = new System.Drawing.Point(424, 34);
			this.cmbFechaDesde.Name = "cmbFechaDesde";
			this.cmbFechaDesde.NonAutoSizeHeight = 23;
			this.cmbFechaDesde.Size = new System.Drawing.Size(120, 21);
			this.cmbFechaDesde.SpinButtonsVisible = true;
			this.cmbFechaDesde.TabIndex = 7;
			this.cmbFechaDesde.Value = new System.DateTime(2006, 6, 17, 0, 0, 0, 0);
			// 
			// cmbNumero
			// 
			this.cmbNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNumero.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNumero.DataMember = "CompraNumero";
			this.cmbNumero.DataSource = this.cdsCompraTabla;
			ultraGridBand5.AddButtonCaption = "CompraComprobante";
			ultraGridColumn27.Header.VisiblePosition = 5;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn29.Header.VisiblePosition = 7;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 3;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 6;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 8;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 4;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 9;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 10;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 11;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 12;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 13;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 14;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 15;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 16;
			ultraGridColumn43.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn43});
			this.cmbNumero.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbNumero.DisplayMember = "Nombre";
			this.cmbNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNumero.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNumero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNumero.Location = new System.Drawing.Point(376, 86);
			this.cmbNumero.Name = "cmbNumero";
			this.cmbNumero.Size = new System.Drawing.Size(168, 21);
			this.cmbNumero.TabIndex = 12;
			this.cmbNumero.ValueMember = "idTipoFactura";
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.btFactura);
			this.ultraTabPageControl2.Controls.Add(this.chkVacios);
			this.ultraTabPageControl2.Controls.Add(this.btCaducidad);
			this.ultraTabPageControl2.Controls.Add(this.cmbCaducidad);
			this.ultraTabPageControl2.Controls.Add(this.txtAutoriza);
			this.ultraTabPageControl2.Controls.Add(this.btAutorizacion);
			this.ultraTabPageControl2.Controls.Add(this.txtSerie);
			this.ultraTabPageControl2.Controls.Add(this.btSerie);
			this.ultraTabPageControl2.Controls.Add(this.ultraButton1);
			this.ultraTabPageControl2.Controls.Add(this.cmbCredTribR);
			this.ultraTabPageControl2.Controls.Add(this.btReemplazar);
			this.ultraTabPageControl2.Controls.Add(this.CmbComprobR);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(644, 120);
			// 
			// btFactura
			// 
			this.btFactura.Location = new System.Drawing.Point(424, 34);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(75, 25);
			this.btFactura.TabIndex = 23;
			this.btFactura.Text = "Factura";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// chkVacios
			// 
			this.chkVacios.BackColor = System.Drawing.Color.Transparent;
			this.chkVacios.Checked = true;
			this.chkVacios.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkVacios.Location = new System.Drawing.Point(432, 9);
			this.chkVacios.Name = "chkVacios";
			this.chkVacios.Size = new System.Drawing.Size(192, 21);
			this.chkVacios.TabIndex = 22;
			this.chkVacios.Text = "Reemplazar Solo en Vacios";
			// 
			// btCaducidad
			// 
			this.btCaducidad.Location = new System.Drawing.Point(424, 60);
			this.btCaducidad.Name = "btCaducidad";
			this.btCaducidad.Size = new System.Drawing.Size(75, 25);
			this.btCaducidad.TabIndex = 21;
			this.btCaducidad.Text = "Caducidad";
			this.btCaducidad.Click += new System.EventHandler(this.btCaducidad_Click);
			// 
			// cmbCaducidad
			// 
			dateButton3.Caption = "Today";
			this.cmbCaducidad.DateButtons.Add(dateButton3);
			this.cmbCaducidad.Format = "dd/MMM/yyyy";
			this.cmbCaducidad.Location = new System.Drawing.Point(512, 60);
			this.cmbCaducidad.Name = "cmbCaducidad";
			this.cmbCaducidad.NonAutoSizeHeight = 23;
			this.cmbCaducidad.Size = new System.Drawing.Size(120, 21);
			this.cmbCaducidad.SpinButtonsVisible = true;
			this.cmbCaducidad.TabIndex = 20;
			this.cmbCaducidad.Value = new System.DateTime(2006, 6, 17, 0, 0, 0, 0);
			// 
			// txtAutoriza
			// 
			this.txtAutoriza.Location = new System.Drawing.Point(304, 60);
			this.txtAutoriza.Name = "txtAutoriza";
			this.txtAutoriza.Size = new System.Drawing.Size(100, 22);
			this.txtAutoriza.TabIndex = 19;
			// 
			// btAutorizacion
			// 
			this.btAutorizacion.Location = new System.Drawing.Point(224, 60);
			this.btAutorizacion.Name = "btAutorizacion";
			this.btAutorizacion.Size = new System.Drawing.Size(75, 25);
			this.btAutorizacion.TabIndex = 18;
			this.btAutorizacion.Text = "Aut. Fact.";
			this.btAutorizacion.Click += new System.EventHandler(this.btAutorizacion_Click);
			// 
			// txtSerie
			// 
			this.txtSerie.Location = new System.Drawing.Point(96, 60);
			this.txtSerie.Name = "txtSerie";
			this.txtSerie.Size = new System.Drawing.Size(100, 22);
			this.txtSerie.TabIndex = 17;
			// 
			// btSerie
			// 
			this.btSerie.Location = new System.Drawing.Point(16, 60);
			this.btSerie.Name = "btSerie";
			this.btSerie.Size = new System.Drawing.Size(75, 25);
			this.btSerie.TabIndex = 16;
			this.btSerie.Text = "Serie";
			this.btSerie.Click += new System.EventHandler(this.btSerie_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(16, 34);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(152, 25);
			this.ultraButton1.TabIndex = 15;
			this.ultraButton1.Text = "Reemplazar Créd. Tribut.";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// cmbCredTribR
			// 
			this.cmbCredTribR.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCredTribR.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCredTribR.DataMember = "CompraCredTribut";
			this.cmbCredTribR.DataSource = this.cdsCompraTabla;
			appearance11.BackColor = System.Drawing.Color.White;
			this.cmbCredTribR.DisplayLayout.Appearance = appearance11;
			ultraGridBand6.AddButtonCaption = "CompraComprobante";
			ultraGridColumn44.Header.VisiblePosition = 2;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 1;
			ultraGridColumn45.Width = 58;
			ultraGridColumn46.Header.VisiblePosition = 0;
			ultraGridColumn46.Width = 242;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46});
			this.cmbCredTribR.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.cmbCredTribR.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 10F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbCredTribR.DisplayLayout.Override.HeaderAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCredTribR.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCredTribR.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.cmbCredTribR.DisplayMember = "CredTributario";
			this.cmbCredTribR.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCredTribR.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCredTribR.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCredTribR.Location = new System.Drawing.Point(184, 34);
			this.cmbCredTribR.Name = "cmbCredTribR";
			this.cmbCredTribR.Size = new System.Drawing.Size(216, 21);
			this.cmbCredTribR.TabIndex = 13;
			this.cmbCredTribR.ValueMember = "idCredTributario";
			// 
			// btReemplazar
			// 
			this.btReemplazar.Location = new System.Drawing.Point(16, 9);
			this.btReemplazar.Name = "btReemplazar";
			this.btReemplazar.Size = new System.Drawing.Size(152, 24);
			this.btReemplazar.TabIndex = 12;
			this.btReemplazar.Text = "Reemplazar Comprobante";
			this.btReemplazar.Click += new System.EventHandler(this.btReemplazar_Click);
			// 
			// CmbComprobR
			// 
			this.CmbComprobR.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.CmbComprobR.Cursor = System.Windows.Forms.Cursors.Default;
			this.CmbComprobR.DataMember = "CompraComprobante";
			this.CmbComprobR.DataSource = this.cdsCompraTabla;
			appearance16.BackColor = System.Drawing.Color.White;
			this.CmbComprobR.DisplayLayout.Appearance = appearance16;
			ultraGridColumn47.Header.VisiblePosition = 2;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 1;
			ultraGridColumn48.Width = 54;
			ultraGridColumn49.Header.VisiblePosition = 0;
			ultraGridColumn49.Width = 265;
			ultraGridColumn50.Header.VisiblePosition = 3;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 4;
			ultraGridColumn51.Hidden = true;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51});
			this.CmbComprobR.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.CmbComprobR.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 10F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.CmbComprobR.DisplayLayout.Override.HeaderAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.CmbComprobR.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.CmbComprobR.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.CmbComprobR.DisplayMember = "Comprobante";
			this.CmbComprobR.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.CmbComprobR.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.CmbComprobR.Location = new System.Drawing.Point(184, 9);
			this.CmbComprobR.Name = "CmbComprobR";
			this.CmbComprobR.Size = new System.Drawing.Size(216, 22);
			this.CmbComprobR.TabIndex = 3;
			this.CmbComprobR.ValueMember = "idComprobante";
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.grCliente);
			this.ultraTabPageControl3.Controls.Add(this.btClienteProv);
			this.ultraTabPageControl3.Controls.Add(this.btCedula);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(644, 120);
			// 
			// btClienteProv
			// 
			this.btClienteProv.Location = new System.Drawing.Point(24, 43);
			this.btClienteProv.Name = "btClienteProv";
			this.btClienteProv.Size = new System.Drawing.Size(128, 26);
			this.btClienteProv.TabIndex = 2;
			this.btClienteProv.Text = "Cliente o Proveedor";
			this.btClienteProv.Click += new System.EventHandler(this.btClienteProv_Click);
			// 
			// udsCliente
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.udsCliente.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn1,
																													 ultraDataColumn2,
																													 ultraDataColumn3,
																													 ultraDataColumn4});
			// 
			// btCedula
			// 
			this.btCedula.Location = new System.Drawing.Point(24, 9);
			this.btCedula.Name = "btCedula";
			this.btCedula.Size = new System.Drawing.Size(128, 25);
			this.btCedula.TabIndex = 0;
			this.btCedula.Text = "Cédulas Incorrectas";
			this.btCedula.Click += new System.EventHandler(this.btCedula_Click);
			this.btCedula.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCedula_MouseUp);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Compra";
			this.ultraGrid1.DataSource = this.cdsCompra;
			appearance21.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance21;
			ultraGridColumn52.Header.VisiblePosition = 42;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 95;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 85;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 87;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.Caption = "Nombre";
			ultraGridColumn56.Header.VisiblePosition = 0;
			ultraGridColumn56.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn56.Width = 96;
			ultraGridColumn57.Header.VisiblePosition = 40;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 86;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 45;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 46;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 77;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 44;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 43;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 47;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 41;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 49;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 48;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.Caption = "Comprobante";
			ultraGridColumn68.Header.VisiblePosition = 1;
			ultraGridColumn68.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn69.Header.Caption = "CredTributario";
			ultraGridColumn69.Header.VisiblePosition = 2;
			ultraGridColumn69.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn70.Header.VisiblePosition = 83;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 79;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Format = "dd/MMM/yyyy";
			ultraGridColumn72.Header.VisiblePosition = 8;
			ultraGridColumn73.Header.VisiblePosition = 33;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 36;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 70;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 5;
			ultraGridColumn76.Width = 75;
			ultraGridColumn77.Header.VisiblePosition = 56;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 68;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 24;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 25;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 26;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 27;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn83.Header.VisiblePosition = 12;
			ultraGridColumn84.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn84.Header.VisiblePosition = 13;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn85.CellAppearance = appearance22;
			ultraGridColumn85.Format = "#,##0.00";
			ultraGridColumn85.Header.VisiblePosition = 10;
			ultraGridColumn86.Header.VisiblePosition = 51;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 30;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 20;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 22;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 67;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 60;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 58;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 21;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 59;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 71;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 69;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 72;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 63;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 73;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 35;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 76;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 75;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 34;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 61;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn105.CellAppearance = appearance23;
			ultraGridColumn105.Format = "#,##0.00";
			ultraGridColumn105.Header.Caption = "Gravado";
			ultraGridColumn105.Header.VisiblePosition = 7;
			ultraGridColumn105.Width = 67;
			ultraGridColumn106.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn106.CellAppearance = appearance24;
			ultraGridColumn106.Format = "#,##0.00";
			ultraGridColumn106.Header.Caption = "Excento";
			ultraGridColumn106.Header.VisiblePosition = 6;
			ultraGridColumn106.Width = 60;
			ultraGridColumn107.Header.VisiblePosition = 66;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 23;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 39;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 65;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 19;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 18;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 37;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 62;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 31;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 15;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 3;
			ultraGridColumn117.Width = 86;
			ultraGridColumn118.Header.Caption = "Serie";
			ultraGridColumn118.Header.VisiblePosition = 4;
			ultraGridColumn118.Width = 63;
			ultraGridColumn119.Header.VisiblePosition = 29;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 28;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.Caption = "Fue";
			ultraGridColumn121.Header.VisiblePosition = 14;
			ultraGridColumn122.Header.VisiblePosition = 64;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn123.Header.VisiblePosition = 17;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn124.Header.VisiblePosition = 16;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn125.Header.VisiblePosition = 52;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 53;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn127.Header.VisiblePosition = 54;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn128.Header.VisiblePosition = 55;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn129.Header.VisiblePosition = 38;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn130.Header.VisiblePosition = 50;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn131.Header.VisiblePosition = 57;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn132.Header.VisiblePosition = 32;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn133.Header.VisiblePosition = 11;
			ultraGridColumn134.Header.VisiblePosition = 74;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn135.Header.VisiblePosition = 78;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn136.Header.VisiblePosition = 80;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 81;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn138.Header.VisiblePosition = 82;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn139.Header.VisiblePosition = 84;
			ultraGridColumn139.Hidden = true;
			ultraGridColumn140.Format = "dd/MMM/yyyy";
			ultraGridColumn140.Header.Caption = "Caducidad";
			ultraGridColumn140.Header.VisiblePosition = 9;
			ultraGridColumn141.Header.VisiblePosition = 88;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn142.Header.VisiblePosition = 89;
			ultraGridColumn142.Hidden = true;
			ultraGridColumn143.Header.VisiblePosition = 90;
			ultraGridColumn143.Hidden = true;
			ultraGridColumn144.Header.VisiblePosition = 94;
			ultraGridColumn144.Hidden = true;
			ultraGridColumn145.Header.VisiblePosition = 93;
			ultraGridColumn145.Hidden = true;
			ultraGridColumn146.Header.VisiblePosition = 91;
			ultraGridColumn146.Hidden = true;
			ultraGridColumn147.Header.VisiblePosition = 92;
			ultraGridColumn147.Hidden = true;
			ultraGridColumn148.Header.VisiblePosition = 96;
			ultraGridColumn148.Hidden = true;
			ultraGridColumn149.Header.VisiblePosition = 97;
			ultraGridColumn149.Hidden = true;
			ultraGridColumn150.Header.VisiblePosition = 98;
			ultraGridColumn150.Hidden = true;
			ultraGridColumn151.Header.VisiblePosition = 99;
			ultraGridColumn151.Hidden = true;
			ultraGridColumn152.Header.VisiblePosition = 100;
			ultraGridColumn152.Hidden = true;
			ultraGridColumn153.Header.VisiblePosition = 101;
			ultraGridColumn153.Hidden = true;
			ultraGridColumn154.Header.VisiblePosition = 102;
			ultraGridColumn154.Hidden = true;
			ultraGridColumn155.Header.VisiblePosition = 103;
			ultraGridColumn155.Hidden = true;
			ultraGridColumn156.Header.VisiblePosition = 104;
			ultraGridColumn156.Hidden = true;
			ultraGridColumn157.Header.VisiblePosition = 105;
			ultraGridColumn157.Hidden = true;
			ultraGridColumn158.Header.VisiblePosition = 106;
			ultraGridColumn158.Hidden = true;
			ultraGridColumn159.Header.VisiblePosition = 107;
			ultraGridColumn159.Hidden = true;
			ultraGridColumn160.Header.VisiblePosition = 108;
			ultraGridColumn160.Hidden = true;
			ultraGridColumn161.Header.VisiblePosition = 109;
			ultraGridColumn161.Hidden = true;
			ultraGridColumn162.Header.VisiblePosition = 110;
			ultraGridColumn162.Hidden = true;
			ultraGridColumn163.Header.VisiblePosition = 111;
			ultraGridColumn163.Hidden = true;
			ultraGridColumn164.Header.VisiblePosition = 112;
			ultraGridColumn164.Hidden = true;
			ultraGridColumn165.Header.VisiblePosition = 113;
			ultraGridColumn165.Hidden = true;
			ultraGridColumn166.Header.VisiblePosition = 114;
			ultraGridColumn166.Hidden = true;
			ultraGridColumn167.Header.VisiblePosition = 115;
			ultraGridColumn167.Hidden = true;
			ultraGridColumn168.Header.VisiblePosition = 116;
			ultraGridColumn168.Hidden = true;
			ultraGridColumn169.Header.VisiblePosition = 117;
			ultraGridColumn169.Hidden = true;
			ultraGridColumn170.Header.VisiblePosition = 118;
			ultraGridColumn170.Hidden = true;
			ultraGridColumn171.Header.VisiblePosition = 119;
			ultraGridColumn171.Hidden = true;
			ultraGridColumn172.Header.VisiblePosition = 120;
			ultraGridColumn172.Hidden = true;
			ultraGridColumn173.Header.VisiblePosition = 121;
			ultraGridColumn173.Hidden = true;
			ultraGridColumn174.Header.VisiblePosition = 122;
			ultraGridColumn174.Hidden = true;
			ultraGridColumn175.Header.VisiblePosition = 123;
			ultraGridColumn175.Hidden = true;
			ultraGridColumn176.Header.VisiblePosition = 124;
			ultraGridColumn176.Hidden = true;
			ultraGridColumn177.Header.VisiblePosition = 125;
			ultraGridColumn177.Hidden = true;
			ultraGridColumn178.Header.VisiblePosition = 126;
			ultraGridColumn178.Hidden = true;
			ultraGridColumn179.Header.VisiblePosition = 127;
			ultraGridColumn179.Hidden = true;
			ultraGridColumn180.Header.VisiblePosition = 128;
			ultraGridColumn180.Hidden = true;
			ultraGridColumn181.Header.VisiblePosition = 129;
			ultraGridColumn181.Hidden = true;
			ultraGridColumn182.Header.VisiblePosition = 130;
			ultraGridColumn182.Hidden = true;
			ultraGridColumn183.Header.VisiblePosition = 131;
			ultraGridColumn183.Hidden = true;
			ultraGridColumn184.Header.VisiblePosition = 132;
			ultraGridColumn184.Hidden = true;
			ultraGridColumn185.Header.VisiblePosition = 133;
			ultraGridColumn185.Hidden = true;
			ultraGridColumn186.Header.VisiblePosition = 134;
			ultraGridColumn186.Hidden = true;
			ultraGridColumn187.Header.VisiblePosition = 135;
			ultraGridColumn187.Hidden = true;
			ultraGridColumn188.Header.VisiblePosition = 136;
			ultraGridColumn188.Hidden = true;
			ultraGridColumn189.Header.VisiblePosition = 137;
			ultraGridColumn189.Hidden = true;
			ultraGridColumn190.Header.VisiblePosition = 138;
			ultraGridBand8.Columns.AddRange(new object[] {
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
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67,
																										 ultraGridColumn68,
																										 ultraGridColumn69,
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73,
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79,
																										 ultraGridColumn80,
																										 ultraGridColumn81,
																										 ultraGridColumn82,
																										 ultraGridColumn83,
																										 ultraGridColumn84,
																										 ultraGridColumn85,
																										 ultraGridColumn86,
																										 ultraGridColumn87,
																										 ultraGridColumn88,
																										 ultraGridColumn89,
																										 ultraGridColumn90,
																										 ultraGridColumn91,
																										 ultraGridColumn92,
																										 ultraGridColumn93,
																										 ultraGridColumn94,
																										 ultraGridColumn95,
																										 ultraGridColumn96,
																										 ultraGridColumn97,
																										 ultraGridColumn98,
																										 ultraGridColumn99,
																										 ultraGridColumn100,
																										 ultraGridColumn101,
																										 ultraGridColumn102,
																										 ultraGridColumn103,
																										 ultraGridColumn104,
																										 ultraGridColumn105,
																										 ultraGridColumn106,
																										 ultraGridColumn107,
																										 ultraGridColumn108,
																										 ultraGridColumn109,
																										 ultraGridColumn110,
																										 ultraGridColumn111,
																										 ultraGridColumn112,
																										 ultraGridColumn113,
																										 ultraGridColumn114,
																										 ultraGridColumn115,
																										 ultraGridColumn116,
																										 ultraGridColumn117,
																										 ultraGridColumn118,
																										 ultraGridColumn119,
																										 ultraGridColumn120,
																										 ultraGridColumn121,
																										 ultraGridColumn122,
																										 ultraGridColumn123,
																										 ultraGridColumn124,
																										 ultraGridColumn125,
																										 ultraGridColumn126,
																										 ultraGridColumn127,
																										 ultraGridColumn128,
																										 ultraGridColumn129,
																										 ultraGridColumn130,
																										 ultraGridColumn131,
																										 ultraGridColumn132,
																										 ultraGridColumn133,
																										 ultraGridColumn134,
																										 ultraGridColumn135,
																										 ultraGridColumn136,
																										 ultraGridColumn137,
																										 ultraGridColumn138,
																										 ultraGridColumn139,
																										 ultraGridColumn140,
																										 ultraGridColumn141,
																										 ultraGridColumn142,
																										 ultraGridColumn143,
																										 ultraGridColumn144,
																										 ultraGridColumn145,
																										 ultraGridColumn146,
																										 ultraGridColumn147,
																										 ultraGridColumn148,
																										 ultraGridColumn149,
																										 ultraGridColumn150,
																										 ultraGridColumn151,
																										 ultraGridColumn152,
																										 ultraGridColumn153,
																										 ultraGridColumn154,
																										 ultraGridColumn155,
																										 ultraGridColumn156,
																										 ultraGridColumn157,
																										 ultraGridColumn158,
																										 ultraGridColumn159,
																										 ultraGridColumn160,
																										 ultraGridColumn161,
																										 ultraGridColumn162,
																										 ultraGridColumn163,
																										 ultraGridColumn164,
																										 ultraGridColumn165,
																										 ultraGridColumn166,
																										 ultraGridColumn167,
																										 ultraGridColumn168,
																										 ultraGridColumn169,
																										 ultraGridColumn170,
																										 ultraGridColumn171,
																										 ultraGridColumn172,
																										 ultraGridColumn173,
																										 ultraGridColumn174,
																										 ultraGridColumn175,
																										 ultraGridColumn176,
																										 ultraGridColumn177,
																										 ultraGridColumn178,
																										 ultraGridColumn179,
																										 ultraGridColumn180,
																										 ultraGridColumn181,
																										 ultraGridColumn182,
																										 ultraGridColumn183,
																										 ultraGridColumn184,
																										 ultraGridColumn185,
																										 ultraGridColumn186,
																										 ultraGridColumn187,
																										 ultraGridColumn188,
																										 ultraGridColumn189,
																										 ultraGridColumn190});
			summarySettings1.DisplayFormat = "{0:#,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings2.DisplayFormat = "{0:#,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings3.DisplayFormat = "{0:#,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings4.DisplayFormat = "{0:#,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings5.DisplayFormat = "{0:#,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings6.DisplayFormat = "{0:#,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand8.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6});
			ultraGridBand8.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a agrupar";
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance25.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 10F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance26;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance28;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 172);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(648, 233);
			this.ultraGrid1.TabIndex = 4;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			// 
			// cmbComprob1
			// 
			this.cmbComprob1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbComprob1.DataMember = "CompraComprobante";
			this.cmbComprob1.DataSource = this.cdsCompraTabla;
			ultraGridColumn191.Header.VisiblePosition = 2;
			ultraGridColumn191.Hidden = true;
			ultraGridColumn192.Header.VisiblePosition = 1;
			ultraGridColumn192.Width = 55;
			ultraGridColumn193.Header.VisiblePosition = 0;
			ultraGridColumn193.Width = 206;
			ultraGridColumn194.Header.VisiblePosition = 3;
			ultraGridColumn194.Hidden = true;
			ultraGridColumn195.Header.VisiblePosition = 4;
			ultraGridColumn195.Hidden = true;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn191,
																										 ultraGridColumn192,
																										 ultraGridColumn193,
																										 ultraGridColumn194,
																										 ultraGridColumn195});
			this.cmbComprob1.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbComprob1.DisplayMember = "Comprobante";
			this.cmbComprob1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbComprob1.Location = new System.Drawing.Point(120, 224);
			this.cmbComprob1.Name = "cmbComprob1";
			this.cmbComprob1.Size = new System.Drawing.Size(264, 86);
			this.cmbComprob1.TabIndex = 5;
			this.cmbComprob1.ValueMember = "idComprobante";
			this.cmbComprob1.Visible = false;
			// 
			// cmbCredTrib1
			// 
			this.cmbCredTrib1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCredTrib1.DataMember = "CompraCredTribut";
			this.cmbCredTrib1.DataSource = this.cdsCompraTabla;
			ultraGridColumn196.Header.VisiblePosition = 2;
			ultraGridColumn196.Hidden = true;
			ultraGridColumn197.Header.VisiblePosition = 1;
			ultraGridColumn197.Width = 60;
			ultraGridColumn198.Header.VisiblePosition = 0;
			ultraGridColumn198.Width = 121;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn196,
																											ultraGridColumn197,
																											ultraGridColumn198});
			this.cmbCredTrib1.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbCredTrib1.DisplayMember = "CredTributario";
			this.cmbCredTrib1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCredTrib1.Location = new System.Drawing.Point(432, 267);
			this.cmbCredTrib1.Name = "cmbCredTrib1";
			this.cmbCredTrib1.Size = new System.Drawing.Size(200, 78);
			this.cmbCredTrib1.TabIndex = 6;
			this.cmbCredTrib1.ValueMember = "idCredTributario";
			this.cmbCredTrib1.Visible = false;
			// 
			// Marco
			// 
			this.Marco.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.Marco.Controls.Add(this.ultraTabSharedControlsPage1);
			this.Marco.Controls.Add(this.ultraTabPageControl1);
			this.Marco.Controls.Add(this.ultraTabPageControl2);
			this.Marco.Controls.Add(this.ultraTabPageControl3);
			this.Marco.Location = new System.Drawing.Point(16, 9);
			this.Marco.Name = "Marco";
			this.Marco.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.Marco.Size = new System.Drawing.Size(648, 146);
			this.Marco.TabIndex = 14;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Filtro";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Reemplazo";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Errores Cédula";
			this.Marco.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																										ultraTab1,
																																										ultraTab2,
																																										ultraTab3});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(644, 120);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdvCompra
			// 
			this.cdvCompra.BindingContextCtrl = this;
			this.cdvCompra.DataSet = this.cdsCompra;
			this.cdvCompra.TableName = "";
			this.cdvCompra.TableViewName = "Compra";
			// 
			// cmbCliente1
			// 
			this.cmbCliente1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente1.DataMember = "Cliente";
			this.cmbCliente1.DataSource = this.cdsCliente;
			ultraGridColumn199.Header.VisiblePosition = 3;
			ultraGridColumn199.Hidden = true;
			ultraGridColumn200.Header.VisiblePosition = 4;
			ultraGridColumn200.Hidden = true;
			ultraGridColumn201.Header.VisiblePosition = 1;
			ultraGridColumn201.Width = 157;
			ultraGridColumn202.Header.VisiblePosition = 5;
			ultraGridColumn202.Hidden = true;
			ultraGridColumn203.Header.VisiblePosition = 0;
			ultraGridColumn204.Header.VisiblePosition = 2;
			ultraGridColumn204.Hidden = true;
			ultraGridColumn205.Header.VisiblePosition = 6;
			ultraGridColumn206.Header.VisiblePosition = 7;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn199,
																											ultraGridColumn200,
																											ultraGridColumn201,
																											ultraGridColumn202,
																											ultraGridColumn203,
																											ultraGridColumn204,
																											ultraGridColumn205,
																											ultraGridColumn206});
			this.cmbCliente1.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbCliente1.DisplayMember = "Nombre";
			this.cmbCliente1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente1.Location = new System.Drawing.Point(368, 233);
			this.cmbCliente1.Name = "cmbCliente1";
			this.cmbCliente1.Size = new System.Drawing.Size(232, 69);
			this.cmbCliente1.TabIndex = 15;
			this.cmbCliente1.ValueMember = "idCliente";
			this.cmbCliente1.Visible = false;
			// 
			// grCliente
			// 
			this.grCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdBalances.EmbeddedNavigator
			// 
			this.grCliente.EmbeddedNavigator.Name = "";
			this.grCliente.Location = new System.Drawing.Point(168, 8);
			this.grCliente.MainView = this.gridView4;
			this.grCliente.Name = "grCliente";
			this.grCliente.Size = new System.Drawing.Size(464, 104);
			this.grCliente.TabIndex = 19;
			this.grCliente.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																														 this.grvBalances,
																																														 this.gridView4});
			// 
			// grvBalances
			// 
			this.grvBalances.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																											 this.idCliente,
																																											 this.Nombre,
																																											 this.gridColumn3});
			this.grvBalances.GridControl = this.grCliente;
			this.grvBalances.Name = "grvBalances";
			this.grvBalances.OptionsSelection.MultiSelect = true;
			this.grvBalances.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvBalances.OptionsView.ShowGroupPanel = false;
			// 
			// gridView4
			// 
			this.gridView4.GridControl = this.grCliente;
			this.gridView4.Name = "gridView4";
			this.gridView4.OptionsSelection.MultiSelect = true;
			this.gridView4.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.gridView4.OptionsView.ShowGroupPanel = false;
			// 
			// idCliente
			// 
			this.idCliente.Caption = "idCliente";
			this.idCliente.Name = "idCliente";
			this.idCliente.Visible = true;
			this.idCliente.VisibleIndex = 0;
			// 
			// Nombre
			// 
			this.Nombre.Caption = "gridColumn2";
			this.Nombre.Name = "Nombre";
			this.Nombre.Visible = true;
			this.Nombre.VisibleIndex = 1;
			// 
			// gridColumn3
			// 
			this.gridColumn3.Caption = "gridColumn3";
			this.gridColumn3.Name = "gridColumn3";
			this.gridColumn3.Visible = true;
			this.gridColumn3.VisibleIndex = 2;
			// 
			// AnexoTransAvanzado
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(688, 428);
			this.Controls.Add(this.cmbCliente1);
			this.Controls.Add(this.Marco);
			this.Controls.Add(this.cmbCredTrib1);
			this.Controls.Add(this.cmbComprob1);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "AnexoTransAvanzado";
			this.Text = "AnexoTransAvanzado";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.AnexoTransAvanzado_Closing);
			this.Load += new System.EventHandler(this.AnexoTransAvanzado_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.spnBodega.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNumero)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbCaducidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutoriza)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribR)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbComprobR)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprob1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTrib1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.Marco)).EndInit();
			this.Marco.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBalances)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;

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
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//				MessageBox.Show(ex.Message, "Tecla");
			}
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			cdsCompra.Update();
		}

		private string Filtro()
		{
			DateTime dtFechaIni = (DateTime) cmbFechaDesde.Value;
			DateTime dtFechaFin = (DateTime) cmbFechaHasta.Value;
			string stFiltro = " Compra.Fecha >= '" + dtFechaIni.ToString("yyyyMMdd")
				+ "' And Compra.Fecha < '" + dtFechaFin.AddDays(1).ToString("yyyyMMdd") 
				+ "' And Compra.idTipoFactura = " + cmbNumero.Value.ToString();
			if (!chkCliente.Checked)
				stFiltro += " And Compra.idCliente = " + cmbCliente.Value.ToString();
			if (!chkComprob.Checked)
				stFiltro += " And Compra.idComprobante = " + cmbComprobante.Value.ToString();
			if (!chkCredTrib.Checked)
				stFiltro += " And Compra.idCredTributario = " + cmbCredTribut.Value.ToString();
			if (!chkCentroC.Checked)
					stFiltro += " And Compra.idProyecto = " + cmbCentroCosto.Value.ToString();
			if (!chkBodega.Checked)
				stFiltro += " And Compra.Bodega = " + spnBodega.Value.ToString();

			return stFiltro;
		}

		private bool Verifica()
		{
			bool bOk = true;
			errorProvider.SetError(Marco, "");
			errorProvider.SetError(cmbNumero, "");
			errorProvider.SetError(cmbCliente, "");
			errorProvider.SetError(cmbComprobante, "");
			errorProvider.SetError(cmbCredTribut, "");
			errorProvider.SetError(cmbCentroCosto,"");
			errorProvider.SetError(spnBodega,"");

			if (cmbNumero.Text.Trim().Length == 0)
			{
				bOk = false;
				errorProvider.SetError(cmbNumero, "Ingrese Tipo de Documento");
			}
			if (!chkCliente.Checked && cmbCliente.Text.Trim().Length == 0)
			{
				bOk = false;
				errorProvider.SetError(cmbCliente, "Ingrese Cliente");
			}
			if (!chkComprob.Checked && cmbComprobante.Text.Trim().Length == 0)
			{
				bOk = false;
				errorProvider.SetError(cmbComprobante, "Ingrese Comprobante");
			}

			if (!chkCredTrib.Checked&& cmbCredTribut.Text.Trim().Length == 0)
			{
				bOk = false;
				errorProvider.SetError(cmbCredTribut, "Ingrese Credito Tributario");
			}
			if (!chkCentroC.Checked && cmbCentroCosto.Text.Trim().Length == 0)
			{
					bOk = false;
				errorProvider.SetError(cmbCentroCosto, "Ingrese Centro de Costo");
			}
			if (!bOk)
				errorProvider.SetError(Marco, "Error en Ficha Filtro");


			return bOk;
		}
		private void btReemplazar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string stFiltro = " Where " + Filtro();
			string stSelect = "Update Compra set idComprobante = "
				+ CmbComprobR.Value.ToString() + stFiltro;
			if (chkVacios.Checked)
				stSelect += " And IsNull(idComprobante, 0) = 0";
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);
			MessageBox.Show("Datos Reemplazados", "Información");
			btFiltro_Click(this, e);
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string stFiltro = " Where " + Filtro();
			string stSelect = "Update Compra set idCredTributario = "
				+ cmbCredTribR.Value.ToString() + stFiltro;
			if (chkVacios.Checked)
				stSelect += " And IsNull(idCredTributario, 0) = 0";
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);
			MessageBox.Show("Datos Reemplazados", "Información");
			btFiltro_Click(this, e);
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string stFiltro = Filtro();
			C1.Data.FilterCondition fcPadre = new 
			C1.Data.FilterCondition(this.cdsCompra, "Compra", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			C1.Data.FilterCondition fcHija = new 
				C1.Data.FilterCondition(this.cdsCompra, "DetCompra", "idCompra = 0");
			fcTotal.Add(fcHija);

			this.cdsCompra.Fill(fcTotal, false);
			if (this.cdvCompra.Count == 0)
			{
				MessageBox.Show("No existen registros que cumplan la condicion de busqueda",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

		}

		private void btSerie_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string stFiltro = "' Where " + Filtro();
			string stSelect = "Update Compra set SerieFactura = '"
				+ txtSerie.Text.ToString() + stFiltro;
			if (chkVacios.Checked)
				stSelect += " And Len(Rtrim(Ltrim(IsNull(SerieFactura,'')))) = 0";
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);
			MessageBox.Show("Datos Reemplazados", "Información");
			btFiltro_Click(this, e);
		}

		private void btAutorizacion_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string stFiltro = "' Where " + Filtro();
			string stSelect = "Update Compra set AutFactura = '"
				+ txtAutoriza.Text.ToString() + stFiltro;
			if (chkVacios.Checked)
				stSelect += " And Len(Rtrim(Ltrim(IsNull(AutFactura, '')))) = 0";
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);
			MessageBox.Show("Datos Reemplazados", "Información");		
			btFiltro_Click(this, e);
		}

		private void btCaducidad_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			DateTime dtFechaCad = (DateTime) cmbCaducidad.Value;
			string stFiltro = "' Where " + Filtro();
			string stSelect = "Update Compra set FechaCaducidad = '"
				+ dtFechaCad.ToString("yyyyMMdd") + stFiltro;
			if (chkVacios.Checked)
				stSelect += " And IsNull(FechaCaducidad, DateAdd(Year, -100, GetDate())) > '19000101'";
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);
			MessageBox.Show("Datos Reemplazados", "Información");		
			btFiltro_Click(this, e);
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(ultraGrid1, "");
			if (ultraGrid1.ActiveRow == null)
			{
				errorProvider.SetError(ultraGrid1, "Seleccione una Fila");
				return;
			}
			int IdCompra = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
//			Compra miCompra = new Compra((int) cmbNumero.Value, IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(ultraGrid1, "Anexo.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "Anexo.xls");
		}

		private void chkComprob_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbComprobante.Enabled = !chkComprob.Checked;
			if (!chkComprob.Checked) cmbComprobante.Select();
		}

		private void chkCredTrib_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCredTribut.Enabled = !chkCredTrib.Checked;
			if (!chkCredTrib.Checked) cmbCredTribut.Select();
		}

		private void chkCliente_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCliente.Enabled = !chkCliente.Checked;
			if (!chkCliente.Checked) cmbCliente.Select();			
		}

		private void AnexoTransAvanzado_Load(object sender, System.EventArgs e)
		{
			cmbFechaDesde.Value = DateTime.Today;
			cmbFechaHasta.Value = DateTime.Today;
			string stAudita = "Exec AuditaCrear 53, 6, 'AnexoT Avz'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}

		private void AnexoTransAvanzado_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'AnexoT Avz'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}

		private void btCedula_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string stSelect = "Select Max(IsNull(Cliente.Nombre, '')) As Nombre, Max(ct.TipoRuc) As TipoRuc, Max(IsNull(Cliente.Ruc, '')) As RUC, "
				+ "dbo.funClienteVerificaCedula(Max(IsNull(Cliente.Ruc, '')), Max(IsNull(Cliente.idtiporuc, '')) ) as Mensaje, "
				+ "Max(Cliente.idCliente) As idCliente From Compra "
				+ "Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
				+ "Left Join ClienteTipoRuc ct On Cliente.idTipoRuc = ct.idTipoRuc "
				+ "Where " + Filtro() + " Group by Compra.idCliente";
			grCliente.DataSource = Funcion.dvProcedimiento(cdsCentroCosto, stSelect);

		}

		private void btClienteProv_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
//			int idCliente = (int) grCliente.ActiveRow.Cells["idCliente"].Value;
//			bool bProveedor = true;
//			if ((int) cmbNumero.Value == 4) bProveedor = false;
//			Cliente miCliente = new Cliente(bProveedor, idCliente);
//			miCliente.MdiParent = this.MdiParent;
//			miCliente.Show();
		}

		private void cmbComprobante_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void chkCentroC_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCentroCosto.Enabled = !chkCentroC.Checked;
			if (!chkCentroC.Checked) cmbCentroCosto.Select();
		}

		private void chkBodega_CheckedChanged(object sender, System.EventArgs e)
		{
			spnBodega.Enabled = !chkBodega.Checked;
			if (!chkBodega.Checked) spnBodega.Select();
		}

		private void btCedula_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			string stSelect = "Select cl.nombre As Nombre, ct.TipoRuc, cl.Ruc AS RUC, "
				+ "dbo.funClienteVerificaCedula(cl.ruc, cl.idTipoRuc) As Mensaje, idCliente From Cliente cl "
				+ "Left Join ClienteTipoRuc ct On cl.idTipoRuc = ct.idTipoRuc ";
			grCliente.DataSource = Funcion.dvProcedimiento(cdsCentroCosto, stSelect);
		}
	}
}
