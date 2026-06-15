using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de AnexoTransImporta.
	/// </summary>
	public class AnexoTransImporta : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsCompra;
		private DevExpress.XtraEditors.SimpleButton btLeer;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsVenta;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNota;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsImportacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsExportacion;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdImportacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsAnulados;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl6;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdAnulados;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAgregar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public AnexoTransImporta()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("compras", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("codSustento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("devIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tpIdProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipoComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaRegistro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("establecimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("puntoEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("secuencial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("autorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaCaducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("baseImponible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("baseImpGrav");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("porcentajeIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("montoIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("baseImpIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("porcentajeIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("montoIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("montoIvaBienes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("porRetBienes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valorRetBienes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("montoIvaServicios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("porRetServicios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valorRetServicios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("estabRetencion1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ptoEmiRetencion1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("secRetencion1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("autRetencion1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaEmiRet1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("docModificado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaEmiModificado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("estabModificado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ptoEmiModificado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("secModificado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("autModificado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("montoTituloOneroso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("montoTituloGratuito");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "baseImponible", 12, true, "compras", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "baseImpGrav", 13, true, "compras", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "montoIva", 15, true, "compras", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "montoIvaBienes", 19, true, "compras", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "montoIvaServicios", 22, true, "compras", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Count, null, "codSustento", 0, true, "compras", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("codSustento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("devIva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tpIdProv");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProv");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipoComprobante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fechaRegistro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("establecimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("puntoEmision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("secuencial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fechaEmision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("autorizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fechaCaducidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("baseImponible");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("baseImpGrav");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("porcentajeIva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("montoIva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("baseImpIce");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("porcentajeIce");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("montoIce");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("montoIvaBienes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("porRetBienes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("valorRetBienes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("montoIvaServicios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("porRetServicios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("valorRetServicios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("estabRetencion1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ptoEmiRetencion1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("secRetencion1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("autRetencion1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fechaEmiRet1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("docModificado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fechaEmiModificado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("estabModificado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ptoEmiModificado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("secModificado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("autModificado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("montoTituloOneroso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("montoTituloGratuito");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tpIdCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipoComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaRegistro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("numeroComprobantes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("baseImponible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ivaPresuntivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("baseImpGrav");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("porcentajeIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("montoIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("baseImpIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("porcentajeIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("montoIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("montoIvaBienes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("porRetBienes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valorRetBienes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("montoIvaServicios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("porRetServicios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valorRetServicios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("retPresuntiva");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Count, null, "tpIdCliente", 0, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "baseImponible", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "baseImpGrav", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "montoIva", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "baseImpIce", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "montoIce", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "montoIvaBienes", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "montoIvaServicios", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tpIdCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipoComprobante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fechaRegistro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("numeroComprobantes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fechaEmision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("baseImponible");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ivaPresuntivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("baseImpGrav");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("porcentajeIva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("montoIva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("baseImpIce");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("porcentajeIce");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("montoIce");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("montoIvaBienes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("porRetBienes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("valorRetBienes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("montoIvaServicios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("porRetServicios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("valorRetServicios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("retPresuntiva");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("codSustento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("importacionDe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaLiquidacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipoComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFiscalProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valorCIF");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("razonSocialProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipoSujeto");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Count, null, "codSustento", 0, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valorCIF", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("codSustento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("importacionDe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fechaLiquidacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipoComprobante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFiscalProv");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("valorCIF");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("razonSocialProv");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipoSujeto");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipoComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("establecimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("puntoEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("secuencialInicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("secuencialFin");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("autorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaAnulacion");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Count, null, "tipoComprobante", 0, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipoComprobante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("establecimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("puntoEmision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("secuencialInicio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("secuencialFin");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("autorizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fechaAnulacion");
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab6 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsCompra = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsVenta = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.grdImportacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsImportacion = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtNota = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTabPageControl6 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.grdAnulados = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsAnulados = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.btLeer = new DevExpress.XtraEditors.SimpleButton();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.udsExportacion = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkAgregar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCompra)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsVenta)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsImportacion)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).BeginInit();
			this.ultraTabPageControl6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdAnulados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAnulados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udsExportacion)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(2, 24);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(564, 254);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "compras";
			this.ultraGrid1.DataSource = this.udsCompra;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "Band 0";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 85;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 56;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 66;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn36.Header.VisiblePosition = 35;
			ultraGridColumn37.Header.VisiblePosition = 36;
			ultraGridColumn38.Header.VisiblePosition = 37;
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
																										 ultraGridColumn38});
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(14, 11);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(536, 232);
			this.ultraGrid1.TabIndex = 1;
			// 
			// udsCompra
			// 
			this.udsCompra.Band.Columns.AddRange(new object[] {
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
																													ultraDataColumn38});
			this.udsCompra.Band.Key = "compras";
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.ultraGrid2);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(564, 254);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "Band 0";
			this.ultraGrid2.DataSource = this.udsVenta;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance6;
			ultraGridColumn39.Header.VisiblePosition = 0;
			ultraGridColumn40.Header.VisiblePosition = 1;
			ultraGridColumn41.Header.VisiblePosition = 2;
			ultraGridColumn42.Header.VisiblePosition = 3;
			ultraGridColumn43.Header.VisiblePosition = 4;
			ultraGridColumn44.Header.VisiblePosition = 5;
			ultraGridColumn45.Header.VisiblePosition = 6;
			ultraGridColumn46.Header.VisiblePosition = 7;
			ultraGridColumn47.Header.VisiblePosition = 8;
			ultraGridColumn48.Header.VisiblePosition = 9;
			ultraGridColumn49.Header.VisiblePosition = 10;
			ultraGridColumn50.Header.VisiblePosition = 11;
			ultraGridColumn51.Header.VisiblePosition = 12;
			ultraGridColumn52.Header.VisiblePosition = 13;
			ultraGridColumn53.Header.VisiblePosition = 14;
			ultraGridColumn54.Header.VisiblePosition = 15;
			ultraGridColumn55.Header.VisiblePosition = 16;
			ultraGridColumn56.Header.VisiblePosition = 17;
			ultraGridColumn57.Header.VisiblePosition = 18;
			ultraGridColumn58.Header.VisiblePosition = 19;
			ultraGridColumn59.Header.VisiblePosition = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn59});
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings7,
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13,
																																															summarySettings14});
			ultraGridBand2.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(14, 11);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(536, 232);
			this.ultraGrid2.TabIndex = 1;
			// 
			// udsVenta
			// 
			this.udsVenta.Band.Columns.AddRange(new object[] {
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
																												 ultraDataColumn59});
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.grdImportacion);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(564, 254);
			// 
			// grdImportacion
			// 
			this.grdImportacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdImportacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdImportacion.DataMember = "Band 0";
			this.grdImportacion.DataSource = this.udsImportacion;
			appearance11.BackColor = System.Drawing.Color.White;
			this.grdImportacion.DisplayLayout.Appearance = appearance11;
			ultraGridColumn60.Header.VisiblePosition = 0;
			ultraGridColumn60.Width = 63;
			ultraGridColumn61.Header.VisiblePosition = 1;
			ultraGridColumn61.Width = 63;
			ultraGridColumn62.Header.VisiblePosition = 2;
			ultraGridColumn62.Width = 64;
			ultraGridColumn63.Header.VisiblePosition = 3;
			ultraGridColumn63.Width = 65;
			ultraGridColumn64.Header.VisiblePosition = 4;
			ultraGridColumn64.Width = 63;
			ultraGridColumn65.Header.VisiblePosition = 5;
			ultraGridColumn65.Width = 63;
			ultraGridColumn66.Header.VisiblePosition = 6;
			ultraGridColumn66.Width = 63;
			ultraGridColumn67.Header.VisiblePosition = 7;
			ultraGridColumn67.Width = 71;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67});
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings15,
																																															summarySettings16});
			ultraGridBand3.SummaryFooterCaption = "";
			this.grdImportacion.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.grdImportacion.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 10F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdImportacion.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.grdImportacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdImportacion.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdImportacion.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.grdImportacion.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdImportacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdImportacion.Location = new System.Drawing.Point(16, 16);
			this.grdImportacion.Name = "grdImportacion";
			this.grdImportacion.Size = new System.Drawing.Size(536, 224);
			this.grdImportacion.TabIndex = 0;
			// 
			// udsImportacion
			// 
			this.udsImportacion.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn60,
																															 ultraDataColumn61,
																															 ultraDataColumn62,
																															 ultraDataColumn63,
																															 ultraDataColumn64,
																															 ultraDataColumn65,
																															 ultraDataColumn66,
																															 ultraDataColumn67});
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(564, 254);
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.txtNota);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(564, 254);
			// 
			// txtNota
			// 
			this.txtNota.AcceptsReturn = true;
			this.txtNota.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtNota.Location = new System.Drawing.Point(14, 15);
			this.txtNota.Multiline = true;
			this.txtNota.Name = "txtNota";
			this.txtNota.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtNota.Size = new System.Drawing.Size(536, 224);
			this.txtNota.TabIndex = 1;
			// 
			// ultraTabPageControl6
			// 
			this.ultraTabPageControl6.Controls.Add(this.grdAnulados);
			this.ultraTabPageControl6.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl6.Name = "ultraTabPageControl6";
			this.ultraTabPageControl6.Size = new System.Drawing.Size(564, 254);
			// 
			// grdAnulados
			// 
			this.grdAnulados.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdAnulados.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdAnulados.DataMember = "Band 0";
			this.grdAnulados.DataSource = this.udsAnulados;
			appearance16.BackColor = System.Drawing.Color.White;
			this.grdAnulados.DisplayLayout.Appearance = appearance16;
			ultraGridColumn68.Header.VisiblePosition = 0;
			ultraGridColumn69.Header.VisiblePosition = 1;
			ultraGridColumn70.Header.VisiblePosition = 2;
			ultraGridColumn71.Header.VisiblePosition = 3;
			ultraGridColumn72.Header.VisiblePosition = 4;
			ultraGridColumn73.Header.VisiblePosition = 5;
			ultraGridColumn74.Header.VisiblePosition = 6;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn68,
																										 ultraGridColumn69,
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73,
																										 ultraGridColumn74});
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings17});
			ultraGridBand4.SummaryFooterCaption = "";
			this.grdAnulados.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.grdAnulados.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 10F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdAnulados.DisplayLayout.Override.HeaderAppearance = appearance18;
			this.grdAnulados.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdAnulados.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdAnulados.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.grdAnulados.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdAnulados.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdAnulados.Location = new System.Drawing.Point(16, 16);
			this.grdAnulados.Name = "grdAnulados";
			this.grdAnulados.Size = new System.Drawing.Size(536, 216);
			this.grdAnulados.TabIndex = 0;
			// 
			// udsAnulados
			// 
			this.udsAnulados.Band.Columns.AddRange(new object[] {
																														ultraDataColumn68,
																														ultraDataColumn69,
																														ultraDataColumn70,
																														ultraDataColumn71,
																														ultraDataColumn72,
																														ultraDataColumn73,
																														ultraDataColumn74});
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl3);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl4);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl5);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl6);
			this.ultraTabControl1.Location = new System.Drawing.Point(24, 32);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(568, 280);
			this.ultraTabControl1.TabIndex = 1;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Compra";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Venta";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Importacion";
			ultraTab4.TabPage = this.ultraTabPageControl4;
			ultraTab4.Text = "Exportacion";
			ultraTab5.TabPage = this.ultraTabPageControl5;
			ultraTab5.Text = "Errores";
			ultraTab6.TabPage = this.ultraTabPageControl6;
			ultraTab6.Text = "Anulados";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2,
																																															 ultraTab3,
																																															 ultraTab4,
																																															 ultraTab5,
																																															 ultraTab6});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(564, 254);
			// 
			// btLeer
			// 
			this.btLeer.Location = new System.Drawing.Point(32, 8);
			this.btLeer.Name = "btLeer";
			this.btLeer.Size = new System.Drawing.Size(56, 24);
			this.btLeer.TabIndex = 2;
			this.btLeer.Text = "Leer";
			this.btLeer.Click += new System.EventHandler(this.btLeer_Click);
			// 
			// chkAgregar
			// 
			this.chkAgregar.Location = new System.Drawing.Point(112, 8);
			this.chkAgregar.Name = "chkAgregar";
			this.chkAgregar.Size = new System.Drawing.Size(152, 20);
			this.chkAgregar.TabIndex = 3;
			this.chkAgregar.Text = "Agregar a Existentes";
			// 
			// AnexoTransImporta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(616, 326);
			this.Controls.Add(this.chkAgregar);
			this.Controls.Add(this.btLeer);
			this.Controls.Add(this.ultraTabControl1);
			this.Name = "AnexoTransImporta";
			this.Text = "AnexoTransImporta";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.AnexoTransImporta_Closing);
			this.Load += new System.EventHandler(this.AnexoTransImporta_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCompra)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsVenta)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsImportacion)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).EndInit();
			this.ultraTabPageControl6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdAnulados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAnulados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udsExportacion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btLeer_Click(object sender, System.EventArgs e)
		{
			#region Inicio
			this.Cursor = Cursors.WaitCursor;
			string stArchivo = "";
			if (stArchivo.Trim().Length == 0)
			{
				this.openFileDialog.Filter = "Excel (*.Xml)|*.xml";
				this.openFileDialog.InitialDirectory = Application.ExecutablePath;
				if (DialogResult.OK == this.openFileDialog.ShowDialog())
				{
					stArchivo = this.openFileDialog.FileName.ToString();
				}
				else
				{
					MessageBox.Show("Cancelado por el usuario");
					this.Cursor = Cursors.Default;
					return;
				}
			}
			StreamReader srArchivo = File.OpenText(stArchivo);
			if (!chkAgregar.Checked)
			{
				udsCompra.Rows.Clear();
				udsVenta.Rows.Clear();
				udsImportacion.Rows.Clear();
				udsAnulados.Rows.Clear();
				udsExportacion.Rows.Clear();
			}
			int iLineas = -1;
			string stLinea = "";
			string stLinea1 = "";
			int iCol;
			#endregion Inicio
			try
			{
				while (stLinea != null && stLinea != "</iva>")
				{
					stLinea = srArchivo.ReadLine().Trim();
					if (stLinea == "<compras>")
					{
						#region Lectura de compras
						if (chkAgregar.Checked)
							iLineas = udsCompra.Rows.Count - 1;
						while (stLinea != null && stLinea != "</compras>")
						{
							stLinea = srArchivo.ReadLine().Trim();
							if (stLinea == "</compras>") continue;
							if (stLinea == "<detalleCompras>")
							{
								iCol = 0;
								iLineas ++;
								udsCompra.Rows.Add();
								while (stLinea != null && stLinea != "</detalleCompras>")
								{
									stLinea = srArchivo.ReadLine().Trim();
									if (stLinea == "</detalleCompras>") continue;
									#region Lectura de retencion
									if (stLinea.StartsWith("<air />"))
									{
										stLinea = srArchivo.ReadLine().Trim();
										iCol += 5;
									}
									if (stLinea.StartsWith("<air>"))
									{
										while (stLinea != null && stLinea != "</air>")
											stLinea = srArchivo.ReadLine().Trim();

										stLinea = srArchivo.ReadLine().Trim();
									}
									#endregion Lectura de retencion

									string stClave = ultraGrid1.DisplayLayout.Bands[0].Columns[iCol].ToString().Trim();
									if (stLinea.StartsWith("<" + stClave + ">"))
									{
										stClave = stLinea.Substring(stClave.Length + 2, stLinea.Length - 2 * stClave.Length - 5);
										udsCompra.Rows[iLineas][iCol] = stClave;
									}
									else
										stLinea1 += "Error " + stClave + " " + stLinea + "\n\r";
									iCol ++;
								}
							}
						}
						#endregion Lectura de compras
					}
					if (stLinea == "<ventas>")
					{
						#region Lectura de ventas
						if (stLinea == "</ventas>") continue;
						iLineas = -1;
						while (stLinea != null && stLinea != "</ventas>")
						{
							stLinea = srArchivo.ReadLine().Trim();
							if (stLinea == "<detalleVentas>")
							{
								iCol = 0;
								iLineas ++;
								if (iLineas == 53)
								{
									iLineas ++;
									iLineas --;
								}
								udsVenta.Rows.Add();
								while (stLinea != null && stLinea != "</detalleVentas>")
								{
									stLinea = srArchivo.ReadLine().Trim();
									if (stLinea == "</detalleVentas>") 
										continue;
									#region Lectura de retencion
									if (stLinea.StartsWith("<air />"))
									{
										continue;
									}
									if (stLinea.StartsWith("<air>"))
									{
										while (stLinea != null && stLinea != "</air>")
											stLinea = srArchivo.ReadLine().Trim();

										stLinea = srArchivo.ReadLine().Trim();
									}
									#endregion Lectura de retencion

									string stClave = ultraGrid2.DisplayLayout.Bands[0].Columns[iCol].ToString().Trim();
									if (stLinea.StartsWith("<" + stClave + ">"))
									{
										stClave = stLinea.Substring(stClave.Length + 2, stLinea.Length - 2 * stClave.Length - 5);
										udsVenta.Rows[iLineas][iCol] = stClave;
									}
									else
										stLinea1 += "Error " + stClave + " " + stLinea + "\n\r";
									iCol ++;
								}
							}
						}
						#endregion Lectura de Ventas
					}
					if (stLinea == "<importaciones>")
					{
						#region Lectura de Importaciones
						iLineas = -1;
						while (stLinea != null && stLinea != "</importaciones>")
						{
							stLinea = srArchivo.ReadLine().Trim();
							if (stLinea == "</importaciones>") continue;
							if (stLinea == "<detalleImportaciones>")
							{
								iCol = 0;
								iLineas ++;
								if (iLineas == 88)
								{
									iLineas ++;
									iLineas --;
								}
								udsImportacion.Rows.Add();
								while (stLinea != null && stLinea != "</detalleImportaciones>")
								{
									stLinea = srArchivo.ReadLine().Trim();
									if (stLinea == "</detalleImportaciones>") continue;
									#region Lectura de retencion
									if (stLinea.StartsWith("<air />"))
									{
										continue;
									}
									if (stLinea.StartsWith("<air>"))
									{
										while (stLinea != null && stLinea != "</air>")
											stLinea = srArchivo.ReadLine().Trim();

										stLinea = srArchivo.ReadLine().Trim();
									}
									#endregion Lectura de retencion

									string stClave = grdImportacion.DisplayLayout.Bands[0].Columns[iCol].ToString().Trim();
									if (stLinea.StartsWith("<" + stClave + ">"))
									{
										stClave = stLinea.Substring(stClave.Length + 2, stLinea.Length - 2 * stClave.Length - 5);
										udsImportacion.Rows[iLineas][iCol] = stClave;
									}
									else
										stLinea1 += "Error " + stClave + " " + stLinea + "\n\r";
									iCol ++;
								}
							}
						}
						#endregion Lectura de Importaciones
					}
					if (stLinea == "<anulados>")
					{
						#region Lectura de Anulados
						iLineas = -1;
						while (stLinea != null && stLinea != "</anulados>")
						{
							stLinea = srArchivo.ReadLine().Trim();
							if (stLinea == "</anulados>") continue;
							if (stLinea == "<detalleAnulados>")
							{
								iCol = 0;
								iLineas ++;
								udsAnulados.Rows.Add();
								while (stLinea != null && stLinea != "</detalleAnulados>")
								{
									stLinea = srArchivo.ReadLine().Trim();
									if (stLinea == "</detalleAnulados>") continue;

									string stClave = grdAnulados.DisplayLayout.Bands[0].Columns[iCol].ToString().Trim();
									if (stLinea.StartsWith("<" + stClave + ">"))
									{
										stClave = stLinea.Substring(stClave.Length + 2, stLinea.Length - 2 * stClave.Length - 5);
										udsAnulados.Rows[iLineas][iCol] = stClave;
									}
									else
										stLinea1 += "Error " + stClave + " " + stLinea + "\n\r";
									iCol ++;
								}
							}
						}
						#endregion Lectura de Anulados
					}
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Leer Anexo");
			}
			txtNota.Value = stLinea1;
			this.Cursor = Cursors.Default;
			MessageBox.Show("Archivo Cargado", "Información");
		}

		private void AnexoTransImporta_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
		}

		private void AnexoTransImporta_Load(object sender, System.EventArgs e)
		{
		}
	}
}
