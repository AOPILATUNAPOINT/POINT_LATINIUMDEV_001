using System;
using System.IO;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ResumenRetencion.
	/// </summary>
	public class ResumenRetencion : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsRetencion;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRetenciones;
		private C1.Data.C1DataSet cdsTipoCedula;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsDetalleRet;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsErrores;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsArticuloSRI;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsListado;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl pgRetTipo;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet opTipoFecha;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkResumen;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbAplica;
		private Infragistics.Win.Misc.UltraButton btRetencion;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optAnuladas;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArtSRI;
		private Infragistics.Win.Misc.UltraButton btRetIva;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optIngEgr;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtHasta;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtDesde;
		private System.Windows.Forms.CheckBox chkRetIva;
		private System.Windows.Forms.CheckBox chkRetRenta;
		private Infragistics.Win.Misc.UltraButton btReporte;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRetIva;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optRentaIva;
		private Infragistics.Win.Misc.UltraButton btGenRet;
		private Infragistics.Win.Misc.UltraButton btRetLista;
		private Infragistics.Win.Misc.UltraButton btAsiLista;
		private Infragistics.Win.Misc.UltraButton btFactLista;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdListado;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl PagReemplazar;
		private Infragistics.Win.Misc.UltraButton btRentencion;
		private Infragistics.Win.Misc.UltraButton btGeneraRet;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVacios;
		private Infragistics.Win.Misc.UltraButton btReemplazar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerieRet;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutorizaRet;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ResumenRetencion()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSRI", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RetencionIva", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSri");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodFact");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizacion");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retenido", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retenido", 5, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Base", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Base", 6, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodSri");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodRet");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodFact");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serie");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizacion");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Retencion", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetRenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetRenta1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSri");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSri1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenida");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaRetenido");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenida1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaRetenido1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaTotal1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ChequeEfectivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngresoEgreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FacturaIva0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FacturaIva12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FacturaIva01");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FacturaIva121");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIva0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContadoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IncluirIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Beneficiario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoAplica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContable");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base Imponible");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor Retenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo Ret");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo Ret");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Año");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("No Retenciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ret1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ret5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ret8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Iva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Problema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.pgRetTipo = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.opTipoFecha = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.chkResumen = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbAplica = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.btRetencion = new Infragistics.Win.Misc.UltraButton();
			this.optAnuladas = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label10 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbArtSRI = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsArticuloSRI = new C1.Data.C1DataSet();
			this.btRetIva = new Infragistics.Win.Misc.UltraButton();
			this.optIngEgr = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label9 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtDesde = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.chkRetIva = new System.Windows.Forms.CheckBox();
			this.chkRetRenta = new System.Windows.Forms.CheckBox();
			this.btReporte = new Infragistics.Win.Misc.UltraButton();
			this.cmbRetIva = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsRetencion = new C1.Data.C1DataSet();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.optRentaIva = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btGenRet = new Infragistics.Win.Misc.UltraButton();
			this.btRetLista = new Infragistics.Win.Misc.UltraButton();
			this.btAsiLista = new Infragistics.Win.Misc.UltraButton();
			this.btFactLista = new Infragistics.Win.Misc.UltraButton();
			this.grdListado = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsListado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.PagReemplazar = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btRentencion = new Infragistics.Win.Misc.UltraButton();
			this.btGeneraRet = new Infragistics.Win.Misc.UltraButton();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.chkVacios = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btReemplazar = new Infragistics.Win.Misc.UltraButton();
			this.txtSerieRet = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtAutorizaRet = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRetenciones = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsTipoCedula = new C1.Data.C1DataSet();
			this.udsDetalleRet = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsErrores = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.pgRetTipo.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.opTipoFecha)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbAplica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optAnuladas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArtSRI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloSRI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optIngEgr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetencion)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optRentaIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdListado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsListado)).BeginInit();
			this.PagReemplazar.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieRet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizaRet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRetenciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDetalleRet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsErrores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			this.SuspendLayout();
			// 
			// pgRetTipo
			// 
			this.pgRetTipo.Controls.Add(this.opTipoFecha);
			this.pgRetTipo.Controls.Add(this.chkResumen);
			this.pgRetTipo.Controls.Add(this.groupBox1);
			this.pgRetTipo.Controls.Add(this.label10);
			this.pgRetTipo.Controls.Add(this.txtNumero);
			this.pgRetTipo.Controls.Add(this.cmbArtSRI);
			this.pgRetTipo.Controls.Add(this.btRetIva);
			this.pgRetTipo.Controls.Add(this.optIngEgr);
			this.pgRetTipo.Controls.Add(this.label9);
			this.pgRetTipo.Controls.Add(this.label8);
			this.pgRetTipo.Controls.Add(this.dtHasta);
			this.pgRetTipo.Controls.Add(this.dtDesde);
			this.pgRetTipo.Controls.Add(this.chkRetIva);
			this.pgRetTipo.Controls.Add(this.chkRetRenta);
			this.pgRetTipo.Controls.Add(this.btReporte);
			this.pgRetTipo.Controls.Add(this.cmbRetIva);
			this.pgRetTipo.Location = new System.Drawing.Point(-10000, -10000);
			this.pgRetTipo.Name = "pgRetTipo";
			this.pgRetTipo.Size = new System.Drawing.Size(716, 439);
			// 
			// opTipoFecha
			// 
			this.opTipoFecha.CheckedIndex = 0;
			this.opTipoFecha.ItemAppearance = appearance1;
			this.opTipoFecha.ItemOrigin = new System.Drawing.Point(6, 6);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Fecha de Retención";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Fecha de Contable";
			this.opTipoFecha.Items.Add(valueListItem1);
			this.opTipoFecha.Items.Add(valueListItem2);
			this.opTipoFecha.ItemSpacingHorizontal = 10;
			this.opTipoFecha.Location = new System.Drawing.Point(88, 267);
			this.opTipoFecha.Name = "opTipoFecha";
			this.opTipoFecha.Size = new System.Drawing.Size(280, 35);
			this.opTipoFecha.TabIndex = 19;
			this.opTipoFecha.Text = "Fecha de Retención";
			// 
			// chkResumen
			// 
			this.chkResumen.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkResumen.Location = new System.Drawing.Point(344, 241);
			this.chkResumen.Name = "chkResumen";
			this.chkResumen.Size = new System.Drawing.Size(120, 22);
			this.chkResumen.TabIndex = 18;
			this.chkResumen.Text = "Resumen";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cmbAplica);
			this.groupBox1.Controls.Add(this.btRetencion);
			this.groupBox1.Controls.Add(this.optAnuladas);
			this.groupBox1.Location = new System.Drawing.Point(40, 345);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(528, 51);
			this.groupBox1.TabIndex = 17;
			this.groupBox1.TabStop = false;
			// 
			// cmbAplica
			// 
			this.cmbAplica.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			valueListItem3.DataValue = 0;
			valueListItem3.DisplayText = "Aplican";
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "No Aplican";
			valueListItem5.DataValue = 2;
			valueListItem5.DisplayText = "Todas";
			this.cmbAplica.Items.Add(valueListItem3);
			this.cmbAplica.Items.Add(valueListItem4);
			this.cmbAplica.Items.Add(valueListItem5);
			this.cmbAplica.Location = new System.Drawing.Point(288, 17);
			this.cmbAplica.Name = "cmbAplica";
			this.cmbAplica.Size = new System.Drawing.Size(96, 22);
			this.cmbAplica.TabIndex = 2;
			// 
			// btRetencion
			// 
			this.btRetencion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btRetencion.Location = new System.Drawing.Point(400, 17);
			this.btRetencion.Name = "btRetencion";
			this.btRetencion.Size = new System.Drawing.Size(104, 25);
			this.btRetencion.TabIndex = 1;
			this.btRetencion.Text = "Retenciones";
			this.btRetencion.Click += new System.EventHandler(this.btRetencion_Click);
			// 
			// optAnuladas
			// 
			this.optAnuladas.CheckedIndex = 0;
			this.optAnuladas.ItemAppearance = appearance2;
			this.optAnuladas.ItemOrigin = new System.Drawing.Point(10, 4);
			valueListItem6.DataValue = 0;
			valueListItem6.DisplayText = "Todas";
			valueListItem7.DataValue = 1;
			valueListItem7.DisplayText = "Anuladas";
			valueListItem8.DataValue = 2;
			valueListItem8.DisplayText = "No Anuladas";
			this.optAnuladas.Items.Add(valueListItem6);
			this.optAnuladas.Items.Add(valueListItem7);
			this.optAnuladas.Items.Add(valueListItem8);
			this.optAnuladas.Location = new System.Drawing.Point(16, 17);
			this.optAnuladas.Name = "optAnuladas";
			this.optAnuladas.Size = new System.Drawing.Size(248, 26);
			this.optAnuladas.TabIndex = 0;
			this.optAnuladas.Text = "Todas";
			// 
			// label10
			// 
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(88, 241);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(128, 25);
			this.label10.TabIndex = 16;
			this.label10.Text = "Número de Reporte";
			// 
			// txtNumero
			// 
			this.txtNumero.Location = new System.Drawing.Point(240, 241);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(24, 22);
			this.txtNumero.TabIndex = 15;
			// 
			// cmbArtSRI
			// 
			this.cmbArtSRI.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArtSRI.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArtSRI.DataMember = "ArticuloSRI";
			this.cmbArtSRI.DataSource = this.cdsArticuloSRI;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbArtSRI.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbArtSRI.DisplayMember = "Codigo";
			this.cmbArtSRI.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArtSRI.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArtSRI.Enabled = false;
			this.cmbArtSRI.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArtSRI.Location = new System.Drawing.Point(240, 172);
			this.cmbArtSRI.Name = "cmbArtSRI";
			this.cmbArtSRI.Size = new System.Drawing.Size(240, 21);
			this.cmbArtSRI.TabIndex = 14;
			this.cmbArtSRI.ValueMember = "idArticuloSRI";
			// 
			// cdsArticuloSRI
			// 
			this.cdsArticuloSRI.BindingContextCtrl = this;
			this.cdsArticuloSRI.DataLibrary = "LibFacturacion";
			this.cdsArticuloSRI.DataLibraryUrl = "";
			this.cdsArticuloSRI.DataSetDef = "dsArticuloTabla";
			this.cdsArticuloSRI.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticuloSRI.Name = "cdsArticuloSRI";
			this.cdsArticuloSRI.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticuloSRI.SchemaDef = null;
			// 
			// btRetIva
			// 
			this.btRetIva.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btRetIva.Location = new System.Drawing.Point(368, 310);
			this.btRetIva.Name = "btRetIva";
			this.btRetIva.Size = new System.Drawing.Size(104, 25);
			this.btRetIva.TabIndex = 13;
			this.btRetIva.Text = "&Reporte Iva";
			this.btRetIva.Click += new System.EventHandler(this.btRetIva_Click);
			// 
			// optIngEgr
			// 
			this.optIngEgr.CheckedIndex = 0;
			this.optIngEgr.ItemAppearance = appearance3;
			this.optIngEgr.ItemOrigin = new System.Drawing.Point(6, 6);
			valueListItem9.DataValue = 2;
			valueListItem9.DisplayText = "Egreso";
			valueListItem10.DataValue = 1;
			valueListItem10.DisplayText = "Ingreso";
			this.optIngEgr.Items.Add(valueListItem9);
			this.optIngEgr.Items.Add(valueListItem10);
			this.optIngEgr.ItemSpacingHorizontal = 6;
			this.optIngEgr.ItemSpacingVertical = 6;
			this.optIngEgr.Location = new System.Drawing.Point(240, 95);
			this.optIngEgr.Name = "optIngEgr";
			this.optIngEgr.Size = new System.Drawing.Size(136, 60);
			this.optIngEgr.TabIndex = 12;
			this.optIngEgr.Text = "Egreso";
			// 
			// label9
			// 
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(24, 34);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(40, 18);
			this.label9.TabIndex = 11;
			this.label9.Text = "Desde";
			// 
			// label8
			// 
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(296, 34);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(40, 18);
			this.label8.TabIndex = 10;
			this.label8.Text = "Hasta";
			// 
			// dtHasta
			// 
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtHasta.DateTime = new System.DateTime(2006, 3, 14, 0, 0, 0, 0);
			this.dtHasta.FormatString = "dd/MMM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(352, 34);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.Size = new System.Drawing.Size(144, 22);
			this.dtHasta.TabIndex = 9;
			this.dtHasta.Value = new System.DateTime(2006, 3, 14, 0, 0, 0, 0);
			// 
			// dtDesde
			// 
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtDesde.DateTime = new System.DateTime(2006, 3, 14, 0, 0, 0, 0);
			this.dtDesde.FormatString = "dd/MMM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(72, 34);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.Size = new System.Drawing.Size(144, 22);
			this.dtDesde.TabIndex = 8;
			this.dtDesde.Value = new System.DateTime(2006, 3, 14, 0, 0, 0, 0);
			// 
			// chkRetIva
			// 
			this.chkRetIva.BackColor = System.Drawing.Color.Transparent;
			this.chkRetIva.Checked = true;
			this.chkRetIva.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkRetIva.Location = new System.Drawing.Point(80, 207);
			this.chkRetIva.Name = "chkRetIva";
			this.chkRetIva.Size = new System.Drawing.Size(128, 26);
			this.chkRetIva.TabIndex = 7;
			this.chkRetIva.Text = "Retención Iva";
			this.chkRetIva.CheckedChanged += new System.EventHandler(this.chkRetIva_CheckedChanged);
			// 
			// chkRetRenta
			// 
			this.chkRetRenta.BackColor = System.Drawing.Color.Transparent;
			this.chkRetRenta.Checked = true;
			this.chkRetRenta.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkRetRenta.Location = new System.Drawing.Point(80, 172);
			this.chkRetRenta.Name = "chkRetRenta";
			this.chkRetRenta.Size = new System.Drawing.Size(128, 26);
			this.chkRetRenta.TabIndex = 6;
			this.chkRetRenta.Text = "Retención Renta";
			this.chkRetRenta.CheckedChanged += new System.EventHandler(this.chkRetRenta_CheckedChanged);
			// 
			// btReporte
			// 
			this.btReporte.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btReporte.Location = new System.Drawing.Point(96, 310);
			this.btReporte.Name = "btReporte";
			this.btReporte.Size = new System.Drawing.Size(104, 25);
			this.btReporte.TabIndex = 2;
			this.btReporte.Text = "&Reporte Renta";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// cmbRetIva
			// 
			this.cmbRetIva.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRetIva.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRetIva.DataMember = "RetencionIva";
			this.cmbRetIva.DataSource = this.cdsRetencion;
			ultraGridBand2.AddButtonCaption = "Retencion";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbRetIva.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbRetIva.DisplayMember = "Codigo";
			this.cmbRetIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRetIva.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRetIva.Enabled = false;
			this.cmbRetIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRetIva.Location = new System.Drawing.Point(240, 207);
			this.cmbRetIva.Name = "cmbRetIva";
			this.cmbRetIva.Size = new System.Drawing.Size(240, 21);
			this.cmbRetIva.TabIndex = 1;
			this.cmbRetIva.ValueMember = "idRetIva";
			// 
			// cdsRetencion
			// 
			this.cdsRetencion.BindingContextCtrl = this;
			this.cdsRetencion.DataLibrary = "LibContabilidad";
			this.cdsRetencion.DataLibraryUrl = "";
			this.cdsRetencion.DataSetDef = "dsRetencion";
			this.cdsRetencion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRetencion.Name = "cdsRetencion";
			this.cdsRetencion.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsRetencion.SchemaDef = null;
			this.cdsRetencion.BeforeFill += new C1.Data.FillEventHandler(this.cdsRetencion_BeforeFill);
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.optRentaIva);
			this.ultraTabPageControl5.Controls.Add(this.btGenRet);
			this.ultraTabPageControl5.Controls.Add(this.btRetLista);
			this.ultraTabPageControl5.Controls.Add(this.btAsiLista);
			this.ultraTabPageControl5.Controls.Add(this.btFactLista);
			this.ultraTabPageControl5.Controls.Add(this.grdListado);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(716, 439);
			this.ultraTabPageControl5.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl5_Paint);
			// 
			// optRentaIva
			// 
			this.optRentaIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.optRentaIva.CheckedIndex = 0;
			this.optRentaIva.ItemAppearance = appearance4;
			this.optRentaIva.ItemOrigin = new System.Drawing.Point(10, 3);
			valueListItem11.DataValue = 0;
			valueListItem11.DisplayText = "Renta";
			valueListItem12.DataValue = 1;
			valueListItem12.DisplayText = "Iva";
			this.optRentaIva.Items.Add(valueListItem11);
			this.optRentaIva.Items.Add(valueListItem12);
			this.optRentaIva.Location = new System.Drawing.Point(376, 406);
			this.optRentaIva.Name = "optRentaIva";
			this.optRentaIva.Size = new System.Drawing.Size(144, 26);
			this.optRentaIva.TabIndex = 5;
			this.optRentaIva.Text = "Renta";
			// 
			// btGenRet
			// 
			this.btGenRet.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGenRet.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGenRet.Location = new System.Drawing.Point(16, 406);
			this.btGenRet.Name = "btGenRet";
			this.btGenRet.Size = new System.Drawing.Size(75, 25);
			this.btGenRet.TabIndex = 4;
			this.btGenRet.Text = "Generar";
			this.btGenRet.Click += new System.EventHandler(this.btGenRet_Click);
			// 
			// btRetLista
			// 
			this.btRetLista.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btRetLista.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btRetLista.Location = new System.Drawing.Point(296, 406);
			this.btRetLista.Name = "btRetLista";
			this.btRetLista.Size = new System.Drawing.Size(32, 25);
			this.btRetLista.TabIndex = 3;
			this.btRetLista.Text = "R";
			this.btRetLista.Click += new System.EventHandler(this.btRetLista_Click);
			// 
			// btAsiLista
			// 
			this.btAsiLista.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsiLista.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAsiLista.Location = new System.Drawing.Point(256, 406);
			this.btAsiLista.Name = "btAsiLista";
			this.btAsiLista.Size = new System.Drawing.Size(32, 25);
			this.btAsiLista.TabIndex = 2;
			this.btAsiLista.Text = "A";
			this.btAsiLista.Click += new System.EventHandler(this.btAsiLista_Click);
			// 
			// btFactLista
			// 
			this.btFactLista.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btFactLista.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFactLista.Location = new System.Drawing.Point(216, 406);
			this.btFactLista.Name = "btFactLista";
			this.btFactLista.Size = new System.Drawing.Size(32, 25);
			this.btFactLista.TabIndex = 1;
			this.btFactLista.Text = "F";
			this.btFactLista.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// grdListado
			// 
			this.grdListado.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdListado.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdListado.DataMember = "Band 0";
			this.grdListado.DataSource = this.udsListado;
			appearance5.BackColor = System.Drawing.Color.White;
			this.grdListado.DisplayLayout.Appearance = appearance5;
			this.grdListado.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Width = 64;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 62;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Width = 60;
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Width = 64;
			ultraGridColumn17.Header.VisiblePosition = 4;
			ultraGridColumn17.Width = 67;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance6;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 5;
			ultraGridColumn18.Width = 76;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance7;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 6;
			ultraGridColumn19.Width = 76;
			ultraGridColumn20.Header.VisiblePosition = 7;
			ultraGridColumn20.Width = 64;
			ultraGridColumn21.Header.VisiblePosition = 10;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 33;
			ultraGridColumn22.Header.VisiblePosition = 11;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 31;
			ultraGridColumn23.Header.VisiblePosition = 12;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 42;
			ultraGridColumn24.Header.VisiblePosition = 8;
			ultraGridColumn24.Width = 65;
			ultraGridColumn25.Header.VisiblePosition = 9;
			ultraGridColumn25.Width = 76;
			ultraGridBand3.Columns.AddRange(new object[] {
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
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand3.SummaryFooterCaption = "";
			this.grdListado.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.grdListado.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 10F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdListado.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.grdListado.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdListado.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdListado.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.grdListado.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdListado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdListado.Location = new System.Drawing.Point(8, 17);
			this.grdListado.Name = "grdListado";
			this.grdListado.Size = new System.Drawing.Size(695, 372);
			this.grdListado.TabIndex = 0;
			// 
			// udsListado
			// 
			ultraDataColumn1.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.Double);
			ultraDataColumn7.DataType = typeof(System.Double);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			this.udsListado.Band.Columns.AddRange(new object[] {
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
			// PagReemplazar
			// 
			this.PagReemplazar.Controls.Add(this.btRentencion);
			this.PagReemplazar.Controls.Add(this.btGeneraRet);
			this.PagReemplazar.Controls.Add(this.btGrabar);
			this.PagReemplazar.Controls.Add(this.chkVacios);
			this.PagReemplazar.Controls.Add(this.btReemplazar);
			this.PagReemplazar.Controls.Add(this.txtSerieRet);
			this.PagReemplazar.Controls.Add(this.txtAutorizaRet);
			this.PagReemplazar.Controls.Add(this.ultraGrid1);
			this.PagReemplazar.Location = new System.Drawing.Point(-10000, -10000);
			this.PagReemplazar.Name = "PagReemplazar";
			this.PagReemplazar.Size = new System.Drawing.Size(716, 439);
			// 
			// btRentencion
			// 
			this.btRentencion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btRentencion.Location = new System.Drawing.Point(224, 397);
			this.btRentencion.Name = "btRentencion";
			this.btRentencion.Size = new System.Drawing.Size(75, 25);
			this.btRentencion.TabIndex = 7;
			this.btRentencion.Text = "Retención";
			this.btRentencion.Click += new System.EventHandler(this.btRentencion_Click);
			// 
			// btGeneraRet
			// 
			this.btGeneraRet.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGeneraRet.Location = new System.Drawing.Point(128, 397);
			this.btGeneraRet.Name = "btGeneraRet";
			this.btGeneraRet.Size = new System.Drawing.Size(75, 25);
			this.btGeneraRet.TabIndex = 6;
			this.btGeneraRet.Text = "Generar";
			this.btGeneraRet.Click += new System.EventHandler(this.btGeneraRet_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.Location = new System.Drawing.Point(24, 397);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(75, 25);
			this.btGrabar.TabIndex = 5;
			this.btGrabar.Text = "Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// chkVacios
			// 
			this.chkVacios.BackColor = System.Drawing.Color.Transparent;
			this.chkVacios.Checked = true;
			this.chkVacios.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkVacios.Location = new System.Drawing.Point(336, 9);
			this.chkVacios.Name = "chkVacios";
			this.chkVacios.Size = new System.Drawing.Size(184, 17);
			this.chkVacios.TabIndex = 4;
			this.chkVacios.Text = "Reemplazar Solo Vacios";
			// 
			// btReemplazar
			// 
			this.btReemplazar.Location = new System.Drawing.Point(24, 9);
			this.btReemplazar.Name = "btReemplazar";
			this.btReemplazar.Size = new System.Drawing.Size(75, 24);
			this.btReemplazar.TabIndex = 3;
			this.btReemplazar.Text = "Reemplazar";
			this.btReemplazar.Click += new System.EventHandler(this.btReemplazar_Click);
			// 
			// txtSerieRet
			// 
			this.txtSerieRet.Location = new System.Drawing.Point(240, 9);
			this.txtSerieRet.Name = "txtSerieRet";
			this.txtSerieRet.Size = new System.Drawing.Size(64, 22);
			this.txtSerieRet.TabIndex = 2;
			// 
			// txtAutorizaRet
			// 
			this.txtAutorizaRet.Location = new System.Drawing.Point(136, 9);
			this.txtAutorizaRet.Name = "txtAutorizaRet";
			this.txtAutorizaRet.Size = new System.Drawing.Size(100, 22);
			this.txtAutorizaRet.TabIndex = 1;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Retencion";
			this.ultraGrid1.DataSource = this.cdsRetencion;
			appearance12.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance12;
			ultraGridColumn26.Header.VisiblePosition = 29;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 9;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 23;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 27;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 32;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 33;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 31;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 25;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 28;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 24;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 34;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 26;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 52;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 14;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 46;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Format = "dd/MMM/yyyy";
			ultraGridColumn42.Header.VisiblePosition = 0;
			ultraGridColumn43.Header.VisiblePosition = 48;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 4;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance13;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.VisiblePosition = 6;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance14;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 5;
			ultraGridColumn47.Header.VisiblePosition = 43;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.VisiblePosition = 37;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 38;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 39;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 41;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 12;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 36;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 19;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 21;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 20;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 22;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 42;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 45;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 44;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 18;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 17;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 15;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 10;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 7;
			ultraGridColumn66.Header.VisiblePosition = 40;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 3;
			ultraGridColumn67.Width = 81;
			ultraGridColumn68.Header.VisiblePosition = 16;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 35;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 11;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 13;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 8;
			ultraGridColumn73.Header.VisiblePosition = 47;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 49;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 50;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 51;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 1;
			ultraGridColumn78.Header.VisiblePosition = 2;
			ultraGridColumn78.Width = 66;
			ultraGridColumn79.Header.VisiblePosition = 53;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 54;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 55;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 56;
			ultraGridColumn82.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn82});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 10F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance16;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(24, 34);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(663, 355);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// udsRetenciones
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(System.Double);
			ultraDataColumn17.DataType = typeof(System.Double);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			this.udsRetenciones.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn14,
																															 ultraDataColumn15,
																															 ultraDataColumn16,
																															 ultraDataColumn17,
																															 ultraDataColumn18,
																															 ultraDataColumn19,
																															 ultraDataColumn20,
																															 ultraDataColumn21,
																															 ultraDataColumn22});
			// 
			// cdsTipoCedula
			// 
			this.cdsTipoCedula.BindingContextCtrl = this;
			this.cdsTipoCedula.DataLibrary = "LibFacturacion";
			this.cdsTipoCedula.DataLibraryUrl = "";
			this.cdsTipoCedula.DataSetDef = "dsClienteTabla";
			this.cdsTipoCedula.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoCedula.Name = "cdsTipoCedula";
			this.cdsTipoCedula.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipoCedula.SchemaDef = null;
			this.cdsTipoCedula.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoCedula_BeforeFill);
			// 
			// udsDetalleRet
			// 
			ultraDataColumn24.DataType = typeof(System.DateTime);
			ultraDataColumn26.DataType = typeof(System.Double);
			ultraDataColumn27.DataType = typeof(System.Double);
			ultraDataColumn30.DataType = typeof(System.Double);
			ultraDataColumn31.DataType = typeof(int);
			this.udsDetalleRet.Band.Columns.AddRange(new object[] {
																															ultraDataColumn23,
																															ultraDataColumn24,
																															ultraDataColumn25,
																															ultraDataColumn26,
																															ultraDataColumn27,
																															ultraDataColumn28,
																															ultraDataColumn29,
																															ultraDataColumn30,
																															ultraDataColumn31});
			// 
			// udsErrores
			// 
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(System.DateTime);
			ultraDataColumn37.DataType = typeof(System.UInt32);
			this.udsErrores.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn32,
																													 ultraDataColumn33,
																													 ultraDataColumn34,
																													 ultraDataColumn35,
																													 ultraDataColumn36,
																													 ultraDataColumn37});
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
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(716, 439);
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.pgRetTipo);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl5);
			this.ultraTabControl1.Controls.Add(this.PagReemplazar);
			this.ultraTabControl1.Location = new System.Drawing.Point(8, 17);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(720, 465);
			this.ultraTabControl1.TabIndex = 8;
			ultraTab1.TabPage = this.pgRetTipo;
			ultraTab1.Text = "Reportes";
			ultraTab2.TabPage = this.ultraTabPageControl5;
			ultraTab2.Text = "Listado";
			ultraTab3.TabPage = this.PagReemplazar;
			ultraTab3.Text = "Reemplazar";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2,
																																															 ultraTab3});
			// 
			// ResumenRetencion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(736, 496);
			this.Controls.Add(this.ultraTabControl1);
			this.Name = "ResumenRetencion";
			this.Text = "Resumen de Retenciones";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ResumenRetencion_Closing);
			this.Load += new System.EventHandler(this.ResumenRetencion_Load);
			this.pgRetTipo.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.opTipoFecha)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbAplica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optAnuladas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArtSRI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloSRI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optIngEgr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetencion)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optRentaIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdListado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsListado)).EndInit();
			this.PagReemplazar.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtSerieRet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizaRet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRetenciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDetalleRet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsErrores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsRetencion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsRetencion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ResumenRetencion_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'ResumenRetencion'";
			Funcion.EjecutaSQL(cdsArticuloSRI, stAudita, true);
			Cursor = Cursors.Default;
			dtDesde.Value = DateTime.Today;
			dtHasta.Value = DateTime.Today;
			cmbAplica.Value = 0;
			this.drSeteo = this.cdsSeteo.TableViews["Seteo"].Rows[0];
		}

		private string Archivo(int iMes, int iAnio)
		{
			string stMes = iMes.ToString();
			if (iMes < 10) stMes = "0" + stMes;
			return "REOC" + iAnio.ToString() +".ANE";
		}



		private string CambiaCodigoSRI(string stCod)
		{
			string stCodigo = stCod;
			switch (int.Parse(stCod))
			{
				case 301: // Relacion de dependencia
					stCodigo = "302";
					break;
				case 303: // Honorario dietas y pagos a personas naturales
					stCodigo = "302";
					break;
				case 305: // Pago local a extranjeros servicios ocacionalea
					stCodigo = "303";
					break;
				case 307: // Por loteriras rifas apuestas y similares
					stCodigo = "304";
					break;
				case 309: // Arriendo Mercantil
					stCodigo = "305";
					break;
				case 311: // Arrendamiento de bienes inmuebles
					stCodigo = "306";
					break;
				case 313: // Seguros y reaseguros
					stCodigo = "307";
					break;
				case 315: // Rendimientos financieros
					stCodigo = "308";
					break;
				case 317: // Intereses y comisiones en credito financieto
					stCodigo = "302";
					break;
				case 319: // Compra de combustible
					stCodigo = "319";
					break;
				case 321: // Servicios petroleros
					stCodigo = "312";
					break;
				case 323: // Activos fijos
					stCodigo = "323";
					break;
				case 325: // Activos corrientes
					stCodigo = "325";
					break;
				case 327: // Servicios
					stCodigo = "327";
					break;
				case 329: // Comisiones
					stCodigo = "315";
					break;
				case 331: // Pago o creditos a tarjetas de credito
					stCodigo = "316";
					break;
				case 333: // Regalias derechos de autor
					stCodigo = "317";
					break;
				case 335: // Remuneraciones a deportistas o arbitros
					stCodigo = "318";
					break;
				case 337: // Pagos a notarios y registradores de propiedad
					stCodigo = "337";
					break;
				case 339: // Pagos de dividendos anticipados
					stCodigo = "339";
					break;
			}
			return stCodigo;
		}

		private string FormatoNum(double dValor, int iAncho)
		{
			string stTemp, stDec;
			double dDec, iDec;

			dValor += .0001;
			dDec = dValor - Math.Floor(dValor);
			dDec = Math.Round(dDec*100, 0);
			iDec = (int) dDec;

			if (iDec < 10) stDec = "0" + iDec.ToString().Trim();			
			else stDec = iDec.ToString().Trim();

			dValor = Math.Floor(dValor);
			stTemp = dValor.ToString().Trim() + ".";
			stTemp += stDec;

			if (stTemp.Length < iAncho) stTemp = stTemp.PadLeft(iAncho, '0');
			return stTemp;
		}


		C1.Data.C1DataRow drSeteo;
		private string RucEmpresa()
		{
			string stRUC = "";
			bool bVerifica = true;
			if (this.drSeteo["RUC"] == System.DBNull.Value) 
			{
				MessageBox.Show("Ingrese el RUC de su empresa");
				bVerifica = false;
			}
			else
			{
				stRUC = this.drSeteo["RUC"].ToString().Trim();
				string stMensaje = Funcion.VerificaCedula(stRUC, false);
				if (stMensaje.Length > 0) bVerifica = false;
				if (bVerifica && stRUC.Length != 13)
				{
					bVerifica = false;
					MessageBox.Show("El Ruc debe tener 13 digitos");
				}
			}
			if (!bVerifica)
			{
				using (Seteo miSeteo = new Seteo())
				{
					miSeteo.ShowDialog();
				}
				this.cdsSeteo.Fill();
			}
			return stRUC;
		}


		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsTipoCedula_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTipoCedula.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private bool VerificaRep()
		{
			bool bOk = true;
			errorProvider.SetError(cmbRetIva, "");
			//			errorProvider.SetError(cmbRetRenta, "");
			errorProvider.SetError(cmbArtSRI, "");
			if (!chkRetIva.Checked)
				if (cmbRetIva.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbRetIva, "Ingrese Retencion de Iva");
					bOk = false;
				}
			if (!chkRetRenta.Checked)
				if (cmbArtSRI.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbArtSRI, "Ingrese Retencion de Renta");
					bOk = false;
				}
			return bOk;
		}

		private string Filtro()
		{
			DateTime dtDesde1 = (DateTime) dtDesde.Value;
			DateTime dtHasta1 = (DateTime) dtHasta.Value;
			string stFiltro = "Retencion.Fecha >= '" + dtDesde1.ToString("yyyyMMdd")
				+ "' And Retencion.Fecha < '" + dtHasta1.ToString("yyyyMMdd") + "'";
			stFiltro += " And Retencion.IngresoEgreso = " + optIngEgr.Value.ToString();
			if (!chkRetIva.Checked)
			{
				stFiltro += " And Retencion.idRetIva = " + cmbRetIva.Value;
			}
			if (!chkRetRenta.Checked)
			{
				stFiltro += " And RetDetalle.idRetRenta = " + cmbArtSRI.Value;
			}
			return stFiltro;
		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			if (!VerificaRep()) return;
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde1 = (DateTime) dtDesde.Value;
			DateTime dtHasta1 = (DateTime) dtHasta.Value;
			string stTitulo = "Retencion Renta";
			if (stReporte == "RetencionIva") stTitulo= "Retencion Iva";
			stTitulo = " desde " + dtDesde1.ToString("dd/MMM/yyyy")
				+ " Hasta " + dtHasta1.ToString("dd/MMM/yyyy") + " ";
			if ((int) optIngEgr.Value == 2) stTitulo += " Egresos";
			else stTitulo += " Ingresos";

			dtHasta1 = dtHasta1.AddDays(1);
			string stFecha = "Fecha";
			if ((int) opTipoFecha.Value == 1) stFecha = "FechaContable";
			string stFiltro = "{Retencion." + stFecha + "} >= #" + dtDesde1.ToString("MM-dd-yyyy")
				+ "# And {Retencion." + stFecha + "} < #" + dtHasta1.ToString("MM-dd-yyyy") + "#";
			stFiltro += " And {Retencion.IngresoEgreso} = " + optIngEgr.Value.ToString();
			if (!chkRetIva.Checked)
			{
				stFiltro += " And {TmpRetIva.idRetIva} = " + cmbRetIva.Value;
				stTitulo += " CodIva " + cmbRetIva.Text.ToString();
			}
			if (!chkRetRenta.Checked)
			{
				stFiltro += " And {RetDetalle.idRetRenta} = " + cmbArtSRI.Value;
				stTitulo += " CodRenta " + cmbArtSRI.Text.ToString();
			}

			if ((int) cmbAplica.Value == 0)
			{
				stFiltro += " And {Retencion.NoAplica} = false";
				stTitulo += " Aplican";
			}
			if ((int) cmbAplica.Value == 1)
			{
				stFiltro += " And {Retencion.NoAplica} = true";
				stTitulo += " No Aplican";
			}

			if (stReporte != "RetencionIva")
			{
				if ((bool) drSeteo["FormMarzo"] ) stReporte = "RetencionRentaAT";
				else stReporte = "RetencionRenta";
			}

			// Filtro de Anulados
			if ((int) optAnuladas.Value == 1) 
			{
				stFiltro += " And {Retencion.Borrar} = true";
				stTitulo += " Anuladas";
			}
			if ((int) optAnuladas.Value == 2) 
			{
				stFiltro += " And {Retencion.Borrar} = false";
				stTitulo += " No Anuladas";
			}
			// Filtro de Anulados

			stReporte += txtNumero.Text.Trim();
			Reporte miRep = new Reporte(stReporte, stFiltro);
			miRep.MdiParent = MdiParent;
			miRep.Titulo(stTitulo);
			if (chkResumen.Checked) miRep.Resumen();
			miRep.Show();
			stReporte = "RetencionRenta";
			Cursor = Cursors.Default;
		}

		private void chkRetRenta_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbArtSRI.Enabled = !chkRetRenta.Checked;
			if (!chkRetRenta.Checked) cmbArtSRI.Select();
		}

		private void chkRetIva_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbRetIva.Enabled = !chkRetIva.Checked;
			if (!chkRetIva.Checked) cmbRetIva.Select();
		}

		string stReporte = "RetencionRenta";
		private void btRetIva_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaIni = (DateTime) dtDesde.Value;
			DateTime dtFechaFin = (DateTime) dtHasta.Value;
			string stSelect = "EXEC RetencionIvaTotal '" + dtFechaIni.ToString("yyyyMMdd") + "', '"
				+ dtFechaFin.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsRetencion, stSelect, true);
			stReporte = "RetencionIva";
			btReporte_Click(this, e);
		}

		private void optRetTipo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btReemplazar_Click(object sender, System.EventArgs e)
		{
			if (txtSerieRet.Text.Trim().Length == 0 && txtAutorizaRet.Text.Trim().Length == 0)
			{
				MessageBox.Show("No hay nada que Reemplazar", "Información");
				return;
			}
			string stFiltro = "";
			if (txtSerieRet.Text.Trim().Length > 0)
			{
				stFiltro = " SerieRet = '" + txtSerieRet.Text;
				string stSelect = "Update Retencion Set" + stFiltro + "' Where " + Filtro();
				if (chkVacios.Checked)
					stSelect += " And LEN(SerieRet) = 0";
				Funcion.EjecutaSQL(cdsRetencion, stSelect, true);
			}
			if (txtAutorizaRet.Text.Trim().Length > 0)
			{
				stFiltro = " AutorizaRet = '" + txtAutorizaRet.Text;
				string stSelect = "Update Retencion Set" + stFiltro + "' Where " + Filtro();
				if (chkVacios.Checked)
					stSelect += " And LEN(AutorizaRet) = 0";
				Funcion.EjecutaSQL(cdsRetencion, stSelect, true);
			}
			btGeneraRet_Click(cdsRetencion, e);
			MessageBox.Show("Reemplazo Terminado", "Información");
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsRetencion.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error");
			}
			MessageBox.Show("Retenciones Grabadas", "Información");
		}

		private void btGeneraRet_Click(object sender, System.EventArgs e)
		{
			string stFiltro = Filtro();
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsRetencion, "Retencion", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			cdsRetencion.Fill(fcTotal, false);
		}

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

		private void btRetencion_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFechaIni = (DateTime) dtDesde.Value;
			DateTime dtFechaFin = (DateTime) dtHasta.Value;
			string pcTitulo = "Retenciones desde el " + dtFechaIni.ToString("dd/MMM/yyyy") 
				+ " Hasta el " + dtFechaFin.ToString("dd/MMM/yyyy");
			string stReporte = "RetencionResumen";
			string stFiltro = "{Retencion.Fecha} >= #" + dtFechaIni.ToString("MM/dd/yyyy")
				+ "# And {Retencion.Fecha} < #" + dtFechaFin.AddDays(1).ToString("MM/dd/yyyy") + "#";
			
			stFiltro += " And {Retencion.IngresoEgreso} = " + optIngEgr.Value.ToString();
			if ((int) optIngEgr.Value == 1) pcTitulo += " Ingresos";
			else pcTitulo += " Egresos";

			if ((int) optAnuladas.Value == 1) 
			{
				stFiltro += " And {Retencion.Borrar} = true";
				pcTitulo += " Anuladas";
			}
			if ((int) optAnuladas.Value == 2) 
			{
				stFiltro += " And {Retencion.Borrar} = false";
				pcTitulo += " No Anuladas";
			}
			if ((int) cmbAplica.Value == 0)
			{
				stFiltro += " And {Retencion.NoAplica} = false";
				pcTitulo += " Aplican";
			}
			if ((int) cmbAplica.Value == 1)
			{
				stFiltro += " And {Retencion.NoAplica} = true";
				pcTitulo += " No Aplican";
			}

			Reporte miRep = new Reporte(stReporte, stFiltro);
			miRep.MdiParent = this.MdiParent;
			if (chkResumen.Checked) miRep.Resumen();
			miRep.Titulo(pcTitulo);
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void ResumenRetencion_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'ResumenRetencion'";
			Funcion.EjecutaSQL(cdsArticuloSRI, stAudita, true);
		}

		private void ultraTabPageControl5_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void btGenRet_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaDesde = (DateTime) dtDesde.Value;
			DateTime dtFechaHasta = (DateTime) dtHasta.Value;
			dtFechaHasta = dtFechaHasta.AddDays(1);
			udsListado.Rows.Clear();
			int iTipo = 4;
			if ((int) optIngEgr.Value == 1) iTipo = 1;
			string stSelect = "Select IsNull(RetDetalle.Base, 0), IsNull(ArticuloSRI.Codigo, '0'), "
				+ "IsNull(ArticuloSRI.Porcentaje, 0), IsNull(Retencion.Fecha, '19000101'), "
				+ "IsNull(Cliente.Nombre, ''), IsNull(Retencion.Codigo, ''), "
				+ "IsNull(Compra.Numero, ''), IsNull(Asiento.CodAsiento, ''), "
				+ "IsNull(RetDetalle.Retenido, 0), IsNull(RetDetalle.Base, 0), "
				+ "IsNull(Retencion.SerieRet, ''), IsNull(Retencion.AutorizaRet, ''), "
				+ "IsNull(Retencion.idRetencion, 0), IsNull(Compra.IdCompra, 0), IsNull(Asiento.idAsiento, 0) "
				+ " FROM Retencion Retencion "
				+ "LEFT OUTER JOIN RetDetalle ON Retencion.idRetencion=RetDetalle.idRetencion "
				+ "LEFT OUTER JOIN Cliente ON Retencion.idCliente=Cliente.idCliente "
				+ "LEFT OUTER JOIN Asiento ON Retencion.idAsiento=Asiento.idAsiento "
				+ "LEFT OUTER JOIN Compra ON Retencion.idCompra=Compra.idCompra "
				+ "LEFT OUTER JOIN Asiento Asiento_1 ON Compra.idAsiento=Asiento_1.idAsiento "
				+ "LEFT OUTER JOIN ArticuloSRI ON RetDetalle.idRetRenta=ArticuloSRI.idArticuloSRI "
				+ "Where Compra.Fecha >= '" + dtFechaDesde.ToString("yyyyMMdd")
				+ "' And Compra.Fecha < '" + dtFechaHasta.ToString("yyyyMMdd") + "' "
				+ "And Compra.idTipoFactura = " + iTipo.ToString() + " And Compra.Borrar = 0 ";

			if ((int) optRentaIva.Value == 1)
			{
				stSelect = "Select IsNull(Retencion.IvaRetenido, 0), "
					+ "IsNull(RetencionIva.Codigo, '0'), "
					+ "IsNull(RetencionIva.Porcentaje, 0), IsNull(Retencion.Fecha, '19000101'), "
					+ "IsNull(Cliente.Nombre, ''), IsNull(Retencion.Codigo, ''), "
					+ "IsNull(Compra.Numero, ''), IsNull(Asiento.CodAsiento, ''), "
					+ "IsNull(Retencion.IvaRetenido, 0), IsNull(Retencion.IvaTotal, 0), "
					+ "IsNull(Retencion.SerieRet, ''), IsNull(Retencion.AutorizaRet, ''), "
					+ "IsNull(Retencion.idRetencion, 0), IsNull(Compra.IdCompra, 0), IsNull(Asiento.idAsiento, 0), "
					+ "IsNull(Retencion.IvaRetenido1, 0), IsNull(RI1.Codigo, '0'), IsNull(Retencion.IvaTotal1, 0)"
					+ " FROM Retencion Retencion "
					+ "LEFT OUTER JOIN RetencionIva ON Retencion.idRetIva = RetencionIva.idRetIva "
					+ "LEFT OUTER JOIN RetencionIva RI1 ON Retencion.idRetIva1 = RI1.idRetIva "
					+ "LEFT OUTER JOIN Cliente ON Retencion.idCliente=Cliente.idCliente "
					+ "LEFT OUTER JOIN Asiento ON Retencion.idAsiento=Asiento.idAsiento "
					+ "LEFT OUTER JOIN Compra ON Retencion.idCompra=Compra.idCompra "
					+ "LEFT OUTER JOIN Asiento Asiento_1 ON Compra.idAsiento=Asiento_1.idAsiento "
					+ "Where Compra.Fecha >= '" + dtFechaDesde.ToString("yyyyMMdd")
					+ "' And Compra.Fecha < '" + dtFechaHasta.ToString("yyyyMMdd") + "' "
					+ "And Compra.idTipoFactura = " + iTipo.ToString() + " And Compra.Borrar = 0 ";
			}
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsRetencion, stSelect, true);
			int iNumFilas = 0;
			while (miReader.Read())
			{
				double dRetBase1 = miReader.GetDouble(0);
				double dRetBase0 = dRetBase1;
				double dRetenido1 = miReader.GetDouble(9);
				double dRetenido0 = dRetenido1;
				string stCod1 = miReader.GetString(1);
				string stCod0 = stCod1;

				double dRetBase2 = 0;
				double dRetenido2 = 0;
				string stCod2 = "";
				if ((int) optRentaIva.Value == 1)
				{
					dRetenido2 = miReader.GetDouble(15);
					stCod2 = miReader.GetString(16);
					dRetBase2 = miReader.GetDouble(17);
					if (dRetBase1 < 0.01 && dRetBase2 >= 0.01)
					{
						dRetBase0 = dRetBase2;
						dRetenido0 = dRetenido2;
						stCod0 = stCod2;
					}
				}
				udsListado.Rows.Add();
				udsListado.Rows[iNumFilas]["Base"]=dRetBase0;
				udsListado.Rows[iNumFilas]["Retenido"]=dRetenido0;
				udsListado.Rows[iNumFilas]["CodSRI"]=stCod0;
				udsListado.Rows[iNumFilas]["Fecha"]=miReader.GetDateTime(3);
				udsListado.Rows[iNumFilas]["CodRet"]=miReader.GetString(5);
				udsListado.Rows[iNumFilas]["CodFact"]=miReader.GetString(6);
				udsListado.Rows[iNumFilas]["CodAsiento"]=miReader.GetString(7);
				udsListado.Rows[iNumFilas]["Nombre"]=miReader.GetString(4);
				udsListado.Rows[iNumFilas]["Serie"]=miReader.GetString(10);
				udsListado.Rows[iNumFilas]["Autorizacion"]=miReader.GetString(11);
				udsListado.Rows[iNumFilas]["IdCompra"]=miReader.GetInt32(13);
				udsListado.Rows[iNumFilas]["IdAsiento"]=miReader.GetInt32(14);
				udsListado.Rows[iNumFilas]["IdRetencion"]=miReader.GetInt32(12);
				iNumFilas++;
				if (dRetBase1 >= 0.01 && dRetBase2 >= 0.01 && (int) optRentaIva.Value == 1)
				{
					udsListado.Rows.Add();
					udsListado.Rows[iNumFilas]["Base"]=dRetBase2;
					udsListado.Rows[iNumFilas]["Retenido"]=dRetenido2;
					udsListado.Rows[iNumFilas]["CodSRI"]=stCod2;
					udsListado.Rows[iNumFilas]["Fecha"]=miReader.GetDateTime(3);
					udsListado.Rows[iNumFilas]["CodRet"]=miReader.GetString(5);
					udsListado.Rows[iNumFilas]["CodFact"]=miReader.GetString(6);
					udsListado.Rows[iNumFilas]["CodAsiento"]=miReader.GetString(7);
					udsListado.Rows[iNumFilas]["Nombre"]=miReader.GetString(4);
					udsListado.Rows[iNumFilas]["Serie"]=miReader.GetString(10);
					udsListado.Rows[iNumFilas]["Autorizacion"]=miReader.GetString(11);
					udsListado.Rows[iNumFilas]["IdCompra"]=miReader.GetInt32(13);
					udsListado.Rows[iNumFilas]["IdAsiento"]=miReader.GetInt32(14);
					udsListado.Rows[iNumFilas]["IdRetencion"]=miReader.GetInt32(12);
					iNumFilas++;
				}
			}
			grdListado.Refresh();
			miReader.Close();
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			if (grdListado.ActiveRow == null) return;

			int idFactura = (int) grdListado.ActiveRow.Cells["idCompra"].Value;
			if (idFactura == 0) 
			{
				MessageBox.Show("No existe una Factura creada de esta Retencion", "Información");
				return;
			}
			int iTipo = 4;
			if ((int) optIngEgr.Value == 1) iTipo = 1;
//			Compra miCompra = new Compra(iTipo, idFactura);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
		}

		private void btRentencion_Click(object sender, System.EventArgs e)
		{
		}

		private void btRetLista_Click(object sender, System.EventArgs e)
		{
			if (grdListado.ActiveRow == null) return;
			int idRetencion = (int) grdListado.ActiveRow.Cells["idRetencion"].Value;
			int idAsiento = (int) grdListado.ActiveRow.Cells["idAsiento"].Value;
			if (idAsiento > 0) 
			{
//				EgresoAT miEgreso = new EgresoAT(idAsiento);
//				miEgreso.MdiParent = this.MdiParent;
//				miEgreso.Show();
			}
			else
			{
				int idFactura = (int) grdListado.ActiveRow.Cells["idCompra"].Value;
				if (idFactura == 0) 
				{
					MessageBox.Show("No existe una Factura creada de esta Retencion", "Información");
					return;
				}
				int iTipo = 4;
				if ((int) optIngEgr.Value == 1) iTipo = 1;

//				Compra miCompra = new Compra(iTipo, idFactura);
//				miCompra.MdiParent = this.MdiParent;
//				miCompra.Show();
			}
		}

		private void btAsiLista_Click(object sender, System.EventArgs e)
		{
			if (grdListado.ActiveRow == null) return;
			int idAsiento = (int) grdListado.ActiveRow.Cells["idAsiento"].Value;
			if (idAsiento == 0) 
			{
				MessageBox.Show("No existe un asiento creado de ésta Retención", "Información");
				return;
			}
			using (Asiento miAsiento = new Asiento(idAsiento))
			{
				miAsiento.ShowDialog();
			}
		}
	}
}
