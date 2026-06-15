using System;
using System.Globalization;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Data;
using DevExpress.XtraPivotGrid;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Comparacion.
	/// </summary>
	public class Comparacion : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsAsiento;
		private C1.Data.C1DataSet cdsCuenta;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton btAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private C1.Data.C1DataSet cdsSeteoC;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.Misc.UltraButton btMayor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactura;
		private C1.Data.C1DataSet cdsTipoFactura;
		private Infragistics.Win.Misc.UltraButton btGeneraFact;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private Infragistics.Win.Misc.UltraButton btAsiF;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.Misc.UltraButton btRetencion;
		private C1.Data.C1DataSet cdsProyecto;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl tabBalance;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl tabMayores;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl tabFacturas;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.Misc.UltraButton btAsiPago;
		private Infragistics.Win.Misc.UltraButton btFacturaPago;
		private Infragistics.Win.Misc.UltraButton btGeneraPago;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoCobro;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsPago;
		private Infragistics.Win.Misc.UltraButton btPago;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdPago;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.Misc.UltraButton btExcelPago;
		private Infragistics.Win.Misc.UltraButton btExcelFactura;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private DevExpress.XtraEditors.LookUpEdit cmbCentro;
		private DevExpress.XtraEditors.LookUpEdit cmbSubCentro;
		private DevExpress.XtraEditors.CheckEdit chkDesglosar;
		private DevExpress.XtraEditors.CheckEdit chkMovimiento;
		private DevExpress.XtraEditors.CheckEdit chkOtro;
		private DevExpress.XtraEditors.CheckEdit chkLote;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private DevExpress.XtraEditors.SimpleButton btGenAsiento;
		private System.Windows.Forms.Label label5;
		private DevExpress.XtraEditors.LookUpEdit cmbTipoAsiento;
		private DevExpress.XtraEditors.CheckEdit chkAsientoAnulado;
		private DevExpress.XtraEditors.CheckEdit chkAsientoAgrupar;
		private DevExpress.XtraEditors.SimpleButton btAsientoA;
		private DevExpress.XtraEditors.SimpleButton btOrigen;
		private DevExpress.XtraEditors.TextEdit txtCodigo;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private DevExpress.XtraEditors.CheckEdit chkGrupoAgrupado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFact1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactCred;
		private DevExpress.XtraEditors.CheckEdit chkAgrupaCredito;
		private DevExpress.XtraEditors.DateEdit cmbDesde;
		private DevExpress.XtraEditors.DateEdit cmbHasta;
		private DevExpress.XtraEditors.ComboBoxEdit cmbBalanceTipo;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl6;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField1;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField2;
		private DevExpress.XtraPivotGrid.PivotGridControl pgrBalance;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl7;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridLookUpEdit1View;
		private DevExpress.XtraGrid.GridControl grdArticulos;
		private DevExpress.XtraEditors.GridLookUpEdit cmbTipo;
		private DevExpress.XtraGrid.Views.Grid.GridView grvArticulos;
		private DevExpress.XtraEditors.CheckEdit chkAjustar;
		private DevExpress.XtraEditors.CheckEdit chkAgrupar;
		private DevExpress.XtraEditors.SimpleButton btKardex;
		private DevExpress.XtraEditors.SimpleButton btAsiArt;
		private DevExpress.XtraEditors.SimpleButton btArticuloB;
		private DevExpress.XtraEditors.SimpleButton btTansaccion;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField3;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField4;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField5;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView2;
		private DevExpress.XtraGrid.GridControl grdMayores;
		private DevExpress.XtraGrid.Views.Grid.GridView grvMayores;
		private DevExpress.XtraEditors.CheckEdit chkPeriodo;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView4;
		private DevExpress.XtraGrid.GridControl grdBalances;
		private DevExpress.XtraGrid.Views.Grid.GridView grvBalances;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView5;
		private DevExpress.XtraGrid.GridControl grdAsientos;
		private DevExpress.XtraGrid.Views.Grid.GridView grvAsientos;
		private DevExpress.XtraEditors.SimpleButton btBalance;
		private DevExpress.XtraEditors.SimpleButton btPresupuesto;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField6;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField7;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField8;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField9;
		private DevExpress.XtraEditors.SimpleButton btExcelCubo;
		private DevExpress.XtraEditors.CheckEdit chkTotalF;
		private DevExpress.XtraEditors.CheckEdit chkTotalC;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField10;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl8;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView6;
		private DevExpress.XtraEditors.RadioGroup optCobrar;
		private DevExpress.XtraGrid.GridControl grdAntiguedad;
		private DevExpress.XtraGrid.Views.Grid.GridView grvAntiguedad;
		private DevExpress.XtraEditors.SimpleButton btFacturaAntig;
		private DevExpress.XtraEditors.RadioGroup optVencidos;
		private DevExpress.XtraEditors.SimpleButton btErrorAntiguedad;
		private DevExpress.XtraEditors.SimpleButton btPagoAntig;
		private DevExpress.XtraGrid.GridControl grdGrupoFactura;
		private DevExpress.XtraGrid.Views.Grid.GridView grvGrupoFactura;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView7;
		private DevExpress.XtraEditors.TextEdit txtTotalGrupo;
		private DevExpress.XtraEditors.TextEdit txtDescuentoGrupo;
		private DevExpress.XtraEditors.CheckEdit chkAnuladoFactura;
		private DevExpress.XtraGrid.GridControl grdFactura;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView3;
		private DevExpress.XtraGrid.Views.Grid.GridView grvFactura;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl9;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView9;
		private DevExpress.XtraEditors.SimpleButton btGenera104;
		private DevExpress.XtraGrid.GridControl grdForm104;
		private DevExpress.XtraGrid.Views.Grid.GridView grvForm104;
		private DevExpress.XtraEditors.SimpleButton btTotalForm104;
		private DevExpress.XtraEditors.TextEdit txtTotalForm;
		private DevExpress.XtraEditors.TextEdit txtTotalFormNC;
		private DevExpress.XtraEditors.TextEdit txtIvaForm;
		private DevExpress.XtraEditors.SimpleButton btTransForm;
		private DevExpress.XtraEditors.SimpleButton btArtForm;
		private DevExpress.XtraEditors.SimpleButton btExcelForm;
		private DevExpress.XtraEditors.SimpleButton btExcelAntig;
		private DevExpress.XtraEditors.SimpleButton btExcelArt;
		private DevExpress.XtraEditors.SimpleButton btExcelBalance;
		private DevExpress.XtraEditors.SimpleButton btAntiguedad;
		private DevExpress.XtraEditors.SimpleButton btFacturaArticulo;
		private DevExpress.XtraEditors.CheckEdit chkAnuladoArt;
		private Infragistics.Win.Misc.UltraButton btPagoFact;
		private DevExpress.XtraEditors.CheckEdit chkResumenCompra;
		private DevExpress.XtraEditors.TextEdit txtDevolucion;
		private DevExpress.XtraGrid.GridControl grdCredito;
		private DevExpress.XtraGrid.Views.Grid.GridView grvCredito;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView10;
		private DevExpress.XtraEditors.SimpleButton btGeneraFacturaGrupo;
		private DevExpress.XtraEditors.SimpleButton btCreditoCompara;
		private DevExpress.XtraEditors.SimpleButton btCreditoGenera;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private DevExpress.XtraGrid.GridControl grdUtilidad;
		private DevExpress.XtraGrid.Views.Grid.GridView grvUtilidad;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView11;
		private DevExpress.XtraEditors.SimpleButton btUtilidadGenera;
		private DevExpress.XtraEditors.SimpleButton btFactUtilidad;
		private DevExpress.XtraEditors.SimpleButton btGrabaCubo;
		private DevExpress.XtraEditors.SimpleButton btCargaCubo;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField11;
		private DevExpress.XtraEditors.SimpleButton btComparaFG;
		private DevExpress.XtraGrid.GridControl grdRetencion;
		private DevExpress.XtraGrid.Views.Grid.GridView grvRetencion;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView12;
		private DevExpress.XtraEditors.SimpleButton btGeneraRetencion;
		private DevExpress.XtraEditors.SimpleButton btFacturaRet;
		private DevExpress.XtraEditors.SimpleButton btRetencionRet;
		private DevExpress.XtraEditors.RadioGroup rdgTipoRet;
		private DevExpress.XtraEditors.CheckEdit chkAnuladaRet;
		private DevExpress.XtraEditors.TextEdit txtTotalRet;
		private DevExpress.XtraEditors.SimpleButton btTotalRet;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl11;
		private DevExpress.XtraPivotGrid.PivotGridControl pivotGridControl1;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField12;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField13;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField14;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField15;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField16;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField17;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField18;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField19;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField20;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField21;
		private DevExpress.XtraPivotGrid.PivotGridField pivotGridField22;
		private DevExpress.XtraEditors.SimpleButton btFlujo;
		private DevExpress.XtraEditors.SimpleButton btInconsistencia;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl12;
		private DevExpress.XtraEditors.SimpleButton btInventario;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView13;
		private DevExpress.XtraGrid.GridControl grdInventario;
		private DevExpress.XtraGrid.Views.Grid.GridView grvInventario;
		private DevExpress.XtraEditors.ComboBoxEdit cmbCostoPeriodo;
		private DevExpress.XtraEditors.SimpleButton btKardexComp;
		private DevExpress.XtraEditors.SimpleButton btTotalFG;
		private DevExpress.XtraEditors.SimpleButton btFacturaGrupo;
		private DevExpress.XtraEditors.SimpleButton simpleButton1;
		private DevExpress.XtraEditors.SimpleButton simpleButton2;
		private DevExpress.XtraEditors.SimpleButton btCreditoFactura;
		private DevExpress.XtraEditors.SimpleButton btCreditoAsiento;
		private DevExpress.XtraEditors.SimpleButton btPagoCred;
		private DevExpress.XtraEditors.SimpleButton btExcelCredito;
		private DevExpress.XtraEditors.CheckEdit chkCuentas;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.SimpleButton btVencido;
		private System.ComponentModel.IContainer components;

		public Comparacion()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAsiLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pago");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroP");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Pago", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Pago", 10, true);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAsiLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroP");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab6 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab7 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab8 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab9 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab10 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab11 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab12 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab13 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab14 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab15 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.tabBalance = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.chkCuentas = new DevExpress.XtraEditors.CheckEdit();
			this.btExcelBalance = new DevExpress.XtraEditors.SimpleButton();
			this.btPresupuesto = new DevExpress.XtraEditors.SimpleButton();
			this.grdBalances = new DevExpress.XtraGrid.GridControl();
			this.grvBalances = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView4 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.chkPeriodo = new DevExpress.XtraEditors.CheckEdit();
			this.cmbBalanceTipo = new DevExpress.XtraEditors.ComboBoxEdit();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.txtCodigo = new DevExpress.XtraEditors.TextEdit();
			this.chkMovimiento = new DevExpress.XtraEditors.CheckEdit();
			this.chkDesglosar = new DevExpress.XtraEditors.CheckEdit();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.tabMayores = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.grdMayores = new DevExpress.XtraGrid.GridControl();
			this.grvMayores = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView2 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.grdAsientos = new DevExpress.XtraGrid.GridControl();
			this.grvAsientos = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView5 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.btOrigen = new DevExpress.XtraEditors.SimpleButton();
			this.btAsientoA = new DevExpress.XtraEditors.SimpleButton();
			this.chkAsientoAgrupar = new DevExpress.XtraEditors.CheckEdit();
			this.chkAsientoAnulado = new DevExpress.XtraEditors.CheckEdit();
			this.cmbTipoAsiento = new DevExpress.XtraEditors.LookUpEdit();
			this.label5 = new System.Windows.Forms.Label();
			this.btGenAsiento = new DevExpress.XtraEditors.SimpleButton();
			this.tabFacturas = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btPagoFact = new Infragistics.Win.Misc.UltraButton();
			this.grdFactura = new DevExpress.XtraGrid.GridControl();
			this.grvFactura = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView3 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.chkAnuladoFactura = new DevExpress.XtraEditors.CheckEdit();
			this.chkLote = new DevExpress.XtraEditors.CheckEdit();
			this.chkOtro = new DevExpress.XtraEditors.CheckEdit();
			this.btExcelFactura = new Infragistics.Win.Misc.UltraButton();
			this.btRetencion = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.btAsiF = new Infragistics.Win.Misc.UltraButton();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.btGeneraFact = new Infragistics.Win.Misc.UltraButton();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsTipoFactura = new C1.Data.C1DataSet();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btExcelPago = new Infragistics.Win.Misc.UltraButton();
			this.btPago = new Infragistics.Win.Misc.UltraButton();
			this.label4 = new System.Windows.Forms.Label();
			this.btAsiPago = new Infragistics.Win.Misc.UltraButton();
			this.btFacturaPago = new Infragistics.Win.Misc.UltraButton();
			this.btGeneraPago = new Infragistics.Win.Misc.UltraButton();
			this.cmbTipoCobro = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.grdPago = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsPago = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btTotalRet = new DevExpress.XtraEditors.SimpleButton();
			this.txtTotalRet = new DevExpress.XtraEditors.TextEdit();
			this.rdgTipoRet = new DevExpress.XtraEditors.RadioGroup();
			this.chkAnuladaRet = new DevExpress.XtraEditors.CheckEdit();
			this.btRetencionRet = new DevExpress.XtraEditors.SimpleButton();
			this.btFacturaRet = new DevExpress.XtraEditors.SimpleButton();
			this.btGeneraRetencion = new DevExpress.XtraEditors.SimpleButton();
			this.grdRetencion = new DevExpress.XtraGrid.GridControl();
			this.grvRetencion = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView12 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.simpleButton2 = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.btFacturaGrupo = new DevExpress.XtraEditors.SimpleButton();
			this.btTotalFG = new DevExpress.XtraEditors.SimpleButton();
			this.btComparaFG = new DevExpress.XtraEditors.SimpleButton();
			this.btGeneraFacturaGrupo = new DevExpress.XtraEditors.SimpleButton();
			this.txtDevolucion = new DevExpress.XtraEditors.TextEdit();
			this.txtTotalGrupo = new DevExpress.XtraEditors.TextEdit();
			this.grdGrupoFactura = new DevExpress.XtraGrid.GridControl();
			this.grvGrupoFactura = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView7 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.chkGrupoAgrupado = new DevExpress.XtraEditors.CheckEdit();
			this.cmbTipoFact1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtDescuentoGrupo = new DevExpress.XtraEditors.TextEdit();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btExcelCredito = new DevExpress.XtraEditors.SimpleButton();
			this.btPagoCred = new DevExpress.XtraEditors.SimpleButton();
			this.btCreditoAsiento = new DevExpress.XtraEditors.SimpleButton();
			this.btCreditoFactura = new DevExpress.XtraEditors.SimpleButton();
			this.btInconsistencia = new DevExpress.XtraEditors.SimpleButton();
			this.btCreditoGenera = new DevExpress.XtraEditors.SimpleButton();
			this.btCreditoCompara = new DevExpress.XtraEditors.SimpleButton();
			this.grdCredito = new DevExpress.XtraGrid.GridControl();
			this.grvCredito = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView10 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.chkResumenCompra = new DevExpress.XtraEditors.CheckEdit();
			this.chkAgrupaCredito = new DevExpress.XtraEditors.CheckEdit();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbTipoFactCred = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraTabPageControl7 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.chkAnuladoArt = new DevExpress.XtraEditors.CheckEdit();
			this.btFacturaArticulo = new DevExpress.XtraEditors.SimpleButton();
			this.btExcelArt = new DevExpress.XtraEditors.SimpleButton();
			this.btTansaccion = new DevExpress.XtraEditors.SimpleButton();
			this.btAsiArt = new DevExpress.XtraEditors.SimpleButton();
			this.btArticuloB = new DevExpress.XtraEditors.SimpleButton();
			this.btKardex = new DevExpress.XtraEditors.SimpleButton();
			this.chkAgrupar = new DevExpress.XtraEditors.CheckEdit();
			this.chkAjustar = new DevExpress.XtraEditors.CheckEdit();
			this.cmbTipo = new DevExpress.XtraEditors.GridLookUpEdit();
			this.gridLookUpEdit1View = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.grdArticulos = new DevExpress.XtraGrid.GridControl();
			this.grvArticulos = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.ultraTabPageControl6 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btCargaCubo = new DevExpress.XtraEditors.SimpleButton();
			this.btGrabaCubo = new DevExpress.XtraEditors.SimpleButton();
			this.btExcelCubo = new DevExpress.XtraEditors.SimpleButton();
			this.chkTotalF = new DevExpress.XtraEditors.CheckEdit();
			this.chkTotalC = new DevExpress.XtraEditors.CheckEdit();
			this.pgrBalance = new DevExpress.XtraPivotGrid.PivotGridControl();
			this.pivotGridField1 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField2 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField3 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField4 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField5 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField6 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField7 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField8 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField9 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField10 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField11 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.ultraTabPageControl8 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btVencido = new DevExpress.XtraEditors.SimpleButton();
			this.btExcelAntig = new DevExpress.XtraEditors.SimpleButton();
			this.btPagoAntig = new DevExpress.XtraEditors.SimpleButton();
			this.btErrorAntiguedad = new DevExpress.XtraEditors.SimpleButton();
			this.optVencidos = new DevExpress.XtraEditors.RadioGroup();
			this.btFacturaAntig = new DevExpress.XtraEditors.SimpleButton();
			this.optCobrar = new DevExpress.XtraEditors.RadioGroup();
			this.btAntiguedad = new DevExpress.XtraEditors.SimpleButton();
			this.grdAntiguedad = new DevExpress.XtraGrid.GridControl();
			this.grvAntiguedad = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView6 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.ultraTabPageControl9 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btExcelForm = new DevExpress.XtraEditors.SimpleButton();
			this.txtIvaForm = new DevExpress.XtraEditors.TextEdit();
			this.txtTotalFormNC = new DevExpress.XtraEditors.TextEdit();
			this.txtTotalForm = new DevExpress.XtraEditors.TextEdit();
			this.btTotalForm104 = new DevExpress.XtraEditors.SimpleButton();
			this.btTransForm = new DevExpress.XtraEditors.SimpleButton();
			this.btArtForm = new DevExpress.XtraEditors.SimpleButton();
			this.btGenera104 = new DevExpress.XtraEditors.SimpleButton();
			this.grdForm104 = new DevExpress.XtraGrid.GridControl();
			this.grvForm104 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView9 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btFactUtilidad = new DevExpress.XtraEditors.SimpleButton();
			this.btUtilidadGenera = new DevExpress.XtraEditors.SimpleButton();
			this.grdUtilidad = new DevExpress.XtraGrid.GridControl();
			this.grvUtilidad = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView11 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.ultraTabPageControl11 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btFlujo = new DevExpress.XtraEditors.SimpleButton();
			this.pivotGridControl1 = new DevExpress.XtraPivotGrid.PivotGridControl();
			this.pivotGridField12 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField13 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField14 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField15 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField16 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField17 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField18 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField19 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField20 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField21 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.pivotGridField22 = new DevExpress.XtraPivotGrid.PivotGridField();
			this.ultraTabPageControl12 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbCostoPeriodo = new DevExpress.XtraEditors.ComboBoxEdit();
			this.btKardexComp = new DevExpress.XtraEditors.SimpleButton();
			this.grdInventario = new DevExpress.XtraGrid.GridControl();
			this.grvInventario = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView13 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.btInventario = new DevExpress.XtraEditors.SimpleButton();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btAsiento = new Infragistics.Win.Misc.UltraButton();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.btMayor = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.cmbCentro = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbSubCentro = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbDesde = new DevExpress.XtraEditors.DateEdit();
			this.cmbHasta = new DevExpress.XtraEditors.DateEdit();
			this.btBalance = new DevExpress.XtraEditors.SimpleButton();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.tabBalance.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.chkCuentas.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdBalances)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBalances)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkPeriodo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBalanceTipo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkMovimiento.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkDesglosar.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			this.tabMayores.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdMayores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvMayores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdAsientos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvAsientos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAsientoAgrupar.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAsientoAnulado.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoAsiento.Properties)).BeginInit();
			this.tabFacturas.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAnuladoFactura.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkLote.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkOtro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoFactura)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCobro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsPago)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalRet.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoRet.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAnuladaRet.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView12)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDevolucion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdGrupoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvGrupoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupoAgrupado.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFact1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuentoGrupo.Properties)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkResumenCompra.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAgrupaCredito.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactCred)).BeginInit();
			this.ultraTabPageControl7.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.chkAnuladoArt.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAgrupar.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAjustar.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridLookUpEdit1View)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdArticulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvArticulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			this.ultraTabPageControl6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.chkTotalF.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTotalC.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.pgrBalance)).BeginInit();
			this.ultraTabPageControl8.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optVencidos.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optCobrar.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdAntiguedad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvAntiguedad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView6)).BeginInit();
			this.ultraTabPageControl9.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaForm.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFormNC.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalForm.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdForm104)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvForm104)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView9)).BeginInit();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdUtilidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvUtilidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView11)).BeginInit();
			this.ultraTabPageControl11.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.pivotGridControl1)).BeginInit();
			this.ultraTabPageControl12.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbCostoPeriodo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdInventario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvInventario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubCentro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// tabBalance
			// 
			this.tabBalance.Controls.Add(this.chkCuentas);
			this.tabBalance.Controls.Add(this.btExcelBalance);
			this.tabBalance.Controls.Add(this.btPresupuesto);
			this.tabBalance.Controls.Add(this.grdBalances);
			this.tabBalance.Controls.Add(this.chkPeriodo);
			this.tabBalance.Controls.Add(this.cmbBalanceTipo);
			this.tabBalance.Controls.Add(this.labelControl1);
			this.tabBalance.Controls.Add(this.txtCodigo);
			this.tabBalance.Controls.Add(this.chkMovimiento);
			this.tabBalance.Controls.Add(this.chkDesglosar);
			this.tabBalance.Controls.Add(this.txtIdAsiento);
			this.tabBalance.Location = new System.Drawing.Point(1, 23);
			this.tabBalance.Name = "tabBalance";
			this.tabBalance.Size = new System.Drawing.Size(908, 327);
			// 
			// chkCuentas
			// 
			this.chkCuentas.EditValue = true;
			this.chkCuentas.Location = new System.Drawing.Point(624, 8);
			this.chkCuentas.Name = "chkCuentas";
			// 
			// chkCuentas.Properties
			// 
			this.chkCuentas.Properties.AllowGrayed = true;
			this.chkCuentas.Properties.Caption = "Cuentas";
			this.chkCuentas.Size = new System.Drawing.Size(72, 18);
			this.chkCuentas.TabIndex = 20;
			// 
			// btExcelBalance
			// 
			this.btExcelBalance.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btExcelBalance.Location = new System.Drawing.Point(848, 4);
			this.btExcelBalance.Name = "btExcelBalance";
			this.btExcelBalance.Size = new System.Drawing.Size(48, 23);
			this.btExcelBalance.TabIndex = 19;
			this.btExcelBalance.Text = "Excel";
			this.btExcelBalance.Click += new System.EventHandler(this.btExcelBalance_Click);
			// 
			// btPresupuesto
			// 
			this.btPresupuesto.Location = new System.Drawing.Point(712, 4);
			this.btPresupuesto.Name = "btPresupuesto";
			this.btPresupuesto.TabIndex = 18;
			this.btPresupuesto.Text = "Presupuesto";
			this.btPresupuesto.Click += new System.EventHandler(this.btPresupuesto_Click);
			this.btPresupuesto.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btPresupuesto_MouseUp);
			// 
			// grdBalances
			// 
			this.grdBalances.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdBalances.EmbeddedNavigator
			// 
			this.grdBalances.EmbeddedNavigator.Name = "";
			this.grdBalances.Location = new System.Drawing.Point(8, 32);
			this.grdBalances.MainView = this.grvBalances;
			this.grdBalances.Name = "grdBalances";
			this.grdBalances.Size = new System.Drawing.Size(888, 292);
			this.grdBalances.TabIndex = 17;
			this.grdBalances.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															 this.grvBalances,
																																															 this.gridView4});
			this.grdBalances.DoubleClick += new System.EventHandler(this.grdBalances_DoubleClick);
			// 
			// grvBalances
			// 
			this.grvBalances.GridControl = this.grdBalances;
			this.grvBalances.Name = "grvBalances";
			this.grvBalances.OptionsSelection.MultiSelect = true;
			this.grvBalances.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvBalances.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView4
			// 
			this.gridView4.GridControl = this.grdBalances;
			this.gridView4.Name = "gridView4";
			// 
			// chkPeriodo
			// 
			this.chkPeriodo.Location = new System.Drawing.Point(248, 8);
			this.chkPeriodo.Name = "chkPeriodo";
			// 
			// chkPeriodo.Properties
			// 
			this.chkPeriodo.Properties.Caption = "Solo Mes";
			this.chkPeriodo.Size = new System.Drawing.Size(72, 18);
			this.chkPeriodo.TabIndex = 16;
			// 
			// cmbBalanceTipo
			// 
			this.cmbBalanceTipo.Location = new System.Drawing.Point(8, 8);
			this.cmbBalanceTipo.Name = "cmbBalanceTipo";
			// 
			// cmbBalanceTipo.Properties
			// 
			this.cmbBalanceTipo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbBalanceTipo.Properties.Items.AddRange(new object[] {
																																	 "General",
																																	 "Resultados",
																																	 "Comprobación"});
			this.cmbBalanceTipo.Size = new System.Drawing.Size(104, 20);
			this.cmbBalanceTipo.TabIndex = 15;
			// 
			// labelControl1
			// 
			this.labelControl1.Location = new System.Drawing.Point(472, 8);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(40, 16);
			this.labelControl1.TabIndex = 14;
			this.labelControl1.Text = "Código";
			// 
			// txtCodigo
			// 
			this.txtCodigo.Location = new System.Drawing.Point(512, 8);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(96, 20);
			this.txtCodigo.TabIndex = 13;
			// 
			// chkMovimiento
			// 
			this.chkMovimiento.EditValue = true;
			this.chkMovimiento.Location = new System.Drawing.Point(328, 8);
			this.chkMovimiento.Name = "chkMovimiento";
			// 
			// chkMovimiento.Properties
			// 
			this.chkMovimiento.Properties.AllowGrayed = true;
			this.chkMovimiento.Properties.Caption = "Cuentas de Movimiento";
			this.chkMovimiento.Size = new System.Drawing.Size(136, 18);
			this.chkMovimiento.TabIndex = 12;
			// 
			// chkDesglosar
			// 
			this.chkDesglosar.Location = new System.Drawing.Point(120, 8);
			this.chkDesglosar.Name = "chkDesglosar";
			// 
			// chkDesglosar.Properties
			// 
			this.chkDesglosar.Properties.Caption = "Desglosar Proyectos";
			this.chkDesglosar.Size = new System.Drawing.Size(120, 18);
			this.chkDesglosar.TabIndex = 11;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdAsiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "EstadoCuenta.idAsiento"));
			this.txtIdAsiento.Location = new System.Drawing.Point(800, 20);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(24, 22);
			this.txtIdAsiento.TabIndex = 9;
			// 
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsEstadoCuenta";
			this.cdsAsiento.FillOnRequest = false;
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			this.cdsAsiento.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsiento_BeforeFill);
			// 
			// tabMayores
			// 
			this.tabMayores.Controls.Add(this.grdMayores);
			this.tabMayores.Location = new System.Drawing.Point(-10000, -10000);
			this.tabMayores.Name = "tabMayores";
			this.tabMayores.Size = new System.Drawing.Size(909, 328);
			// 
			// grdMayores
			// 
			this.grdMayores.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdMayores.EmbeddedNavigator
			// 
			this.grdMayores.EmbeddedNavigator.Name = "";
			this.grdMayores.Location = new System.Drawing.Point(8, 8);
			this.grdMayores.MainView = this.grvMayores;
			this.grdMayores.Name = "grdMayores";
			this.grdMayores.Size = new System.Drawing.Size(889, 317);
			this.grdMayores.TabIndex = 0;
			this.grdMayores.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															this.grvMayores,
																																															this.gridView2});
			// 
			// grvMayores
			// 
			this.grvMayores.GridControl = this.grdMayores;
			this.grvMayores.Name = "grvMayores";
			this.grvMayores.OptionsSelection.MultiSelect = true;
			this.grvMayores.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvMayores.OptionsView.ShowAutoFilterRow = true;
			this.grvMayores.DoubleClick += new System.EventHandler(this.grvMayores_DoubleClick);
			// 
			// gridView2
			// 
			this.gridView2.GridControl = this.grdMayores;
			this.gridView2.Name = "gridView2";
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.grdAsientos);
			this.ultraTabPageControl3.Controls.Add(this.btOrigen);
			this.ultraTabPageControl3.Controls.Add(this.btAsientoA);
			this.ultraTabPageControl3.Controls.Add(this.chkAsientoAgrupar);
			this.ultraTabPageControl3.Controls.Add(this.chkAsientoAnulado);
			this.ultraTabPageControl3.Controls.Add(this.cmbTipoAsiento);
			this.ultraTabPageControl3.Controls.Add(this.label5);
			this.ultraTabPageControl3.Controls.Add(this.btGenAsiento);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(909, 328);
			// 
			// grdAsientos
			// 
			this.grdAsientos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdAsientos.EmbeddedNavigator
			// 
			this.grdAsientos.EmbeddedNavigator.Name = "";
			this.grdAsientos.Location = new System.Drawing.Point(10, 40);
			this.grdAsientos.MainView = this.grvAsientos;
			this.grdAsientos.Name = "grdAsientos";
			this.grdAsientos.Size = new System.Drawing.Size(889, 285);
			this.grdAsientos.TabIndex = 67;
			this.grdAsientos.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															 this.grvAsientos,
																																															 this.gridView5});
			// 
			// grvAsientos
			// 
			this.grvAsientos.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus;
			this.grvAsientos.GridControl = this.grdAsientos;
			this.grvAsientos.Name = "grvAsientos";
			this.grvAsientos.OptionsSelection.MultiSelect = true;
			this.grvAsientos.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvAsientos.OptionsView.ColumnAutoWidth = false;
			this.grvAsientos.OptionsView.ShowAutoFilterRow = true;
			this.grvAsientos.DoubleClick += new System.EventHandler(this.grvAsientos_DoubleClick);
			// 
			// gridView5
			// 
			this.gridView5.GridControl = this.grdAsientos;
			this.gridView5.Name = "gridView5";
			// 
			// btOrigen
			// 
			this.btOrigen.Location = new System.Drawing.Point(640, 8);
			this.btOrigen.Name = "btOrigen";
			this.btOrigen.Size = new System.Drawing.Size(56, 23);
			this.btOrigen.TabIndex = 66;
			this.btOrigen.Text = "Origen";
			this.btOrigen.Click += new System.EventHandler(this.btOrigen_Click);
			// 
			// btAsientoA
			// 
			this.btAsientoA.Location = new System.Drawing.Point(576, 8);
			this.btAsientoA.Name = "btAsientoA";
			this.btAsientoA.Size = new System.Drawing.Size(56, 23);
			this.btAsientoA.TabIndex = 65;
			this.btAsientoA.Text = "Asiento";
			this.btAsientoA.Click += new System.EventHandler(this.btAsientoA_Click);
			// 
			// chkAsientoAgrupar
			// 
			this.chkAsientoAgrupar.EditValue = true;
			this.chkAsientoAgrupar.Location = new System.Drawing.Point(360, 8);
			this.chkAsientoAgrupar.Name = "chkAsientoAgrupar";
			// 
			// chkAsientoAgrupar.Properties
			// 
			this.chkAsientoAgrupar.Properties.Caption = "Agrupar";
			this.chkAsientoAgrupar.Size = new System.Drawing.Size(64, 18);
			this.chkAsientoAgrupar.TabIndex = 64;
			// 
			// chkAsientoAnulado
			// 
			this.chkAsientoAnulado.EditValue = true;
			this.chkAsientoAnulado.Location = new System.Drawing.Point(248, 8);
			this.chkAsientoAnulado.Name = "chkAsientoAnulado";
			// 
			// chkAsientoAnulado.Properties
			// 
			this.chkAsientoAnulado.Properties.Caption = "Incluir Anulados";
			this.chkAsientoAnulado.Size = new System.Drawing.Size(112, 18);
			this.chkAsientoAnulado.TabIndex = 63;
			// 
			// cmbTipoAsiento
			// 
			this.cmbTipoAsiento.Location = new System.Drawing.Point(56, 8);
			this.cmbTipoAsiento.Name = "cmbTipoAsiento";
			// 
			// cmbTipoAsiento.Properties
			// 
			this.cmbTipoAsiento.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbTipoAsiento.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Tipo", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.cmbTipoAsiento.Properties.DisplayMember = "Tipo";
			this.cmbTipoAsiento.Properties.NullText = "";
			this.cmbTipoAsiento.Properties.PopupWidth = 300;
			this.cmbTipoAsiento.Properties.ShowFooter = false;
			this.cmbTipoAsiento.Properties.ShowHeader = false;
			this.cmbTipoAsiento.Properties.ValueMember = "idTipoAsiento";
			this.cmbTipoAsiento.Size = new System.Drawing.Size(176, 20);
			this.cmbTipoAsiento.TabIndex = 62;
			this.cmbTipoAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cmbTipoAsiento_MouseDown);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(24, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(26, 16);
			this.label5.TabIndex = 9;
			this.label5.Text = "Tipo";
			// 
			// btGenAsiento
			// 
			this.btGenAsiento.Location = new System.Drawing.Point(504, 8);
			this.btGenAsiento.Name = "btGenAsiento";
			this.btGenAsiento.Size = new System.Drawing.Size(56, 23);
			this.btGenAsiento.TabIndex = 4;
			this.btGenAsiento.Text = "Genera";
			this.btGenAsiento.Click += new System.EventHandler(this.btGenAsiento_Click);
			// 
			// tabFacturas
			// 
			this.tabFacturas.Controls.Add(this.btPagoFact);
			this.tabFacturas.Controls.Add(this.grdFactura);
			this.tabFacturas.Controls.Add(this.chkAnuladoFactura);
			this.tabFacturas.Controls.Add(this.chkLote);
			this.tabFacturas.Controls.Add(this.chkOtro);
			this.tabFacturas.Controls.Add(this.btExcelFactura);
			this.tabFacturas.Controls.Add(this.btRetencion);
			this.tabFacturas.Controls.Add(this.ultraButton2);
			this.tabFacturas.Controls.Add(this.ultraButton1);
			this.tabFacturas.Controls.Add(this.btAsiF);
			this.tabFacturas.Controls.Add(this.btFactura);
			this.tabFacturas.Controls.Add(this.btGeneraFact);
			this.tabFacturas.Controls.Add(this.cmbTipoFactura);
			this.tabFacturas.Location = new System.Drawing.Point(-10000, -10000);
			this.tabFacturas.Name = "tabFacturas";
			this.tabFacturas.Size = new System.Drawing.Size(909, 328);
			// 
			// btPagoFact
			// 
			this.btPagoFact.Location = new System.Drawing.Point(736, 8);
			this.btPagoFact.Name = "btPagoFact";
			this.btPagoFact.Size = new System.Drawing.Size(40, 24);
			this.btPagoFact.TabIndex = 20;
			this.btPagoFact.Text = "Pago";
			this.toolTip1.SetToolTip(this.btPagoFact, "Muestra Asiento de Pago");
			this.btPagoFact.Click += new System.EventHandler(this.btPagoFact_Click);
			// 
			// grdFactura
			// 
			this.grdFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdFactura.EmbeddedNavigator
			// 
			this.grdFactura.EmbeddedNavigator.Name = "";
			this.grdFactura.Location = new System.Drawing.Point(8, 40);
			this.grdFactura.MainView = this.grvFactura;
			this.grdFactura.Name = "grdFactura";
			this.grdFactura.Size = new System.Drawing.Size(889, 277);
			this.grdFactura.TabIndex = 19;
			this.grdFactura.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															this.grvFactura,
																																															this.gridView3});
			this.grdFactura.DoubleClick += new System.EventHandler(this.grdFactura_DoubleClick);
			// 
			// grvFactura
			// 
			this.grvFactura.GridControl = this.grdFactura;
			this.grvFactura.Name = "grvFactura";
			this.grvFactura.OptionsSelection.MultiSelect = true;
			this.grvFactura.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvFactura.OptionsView.ColumnAutoWidth = false;
			this.grvFactura.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView3
			// 
			this.gridView3.GridControl = this.grdFactura;
			this.gridView3.Name = "gridView3";
			// 
			// chkAnuladoFactura
			// 
			this.chkAnuladoFactura.EditValue = true;
			this.chkAnuladoFactura.Location = new System.Drawing.Point(392, 8);
			this.chkAnuladoFactura.Name = "chkAnuladoFactura";
			// 
			// chkAnuladoFactura.Properties
			// 
			this.chkAnuladoFactura.Properties.Caption = "Anulado";
			this.chkAnuladoFactura.Size = new System.Drawing.Size(64, 18);
			this.chkAnuladoFactura.TabIndex = 18;
			// 
			// chkLote
			// 
			this.chkLote.EditValue = true;
			this.chkLote.Location = new System.Drawing.Point(280, 8);
			this.chkLote.Name = "chkLote";
			// 
			// chkLote.Properties
			// 
			this.chkLote.Properties.Caption = "Generado Lote";
			this.chkLote.Size = new System.Drawing.Size(104, 18);
			this.chkLote.TabIndex = 17;
			// 
			// chkOtro
			// 
			this.chkOtro.EditValue = true;
			this.chkOtro.Location = new System.Drawing.Point(192, 8);
			this.chkOtro.Name = "chkOtro";
			// 
			// chkOtro.Properties
			// 
			this.chkOtro.Properties.Caption = "Caja Chica";
			this.chkOtro.Size = new System.Drawing.Size(80, 18);
			this.chkOtro.TabIndex = 16;
			// 
			// btExcelFactura
			// 
			this.btExcelFactura.Location = new System.Drawing.Point(792, 8);
			this.btExcelFactura.Name = "btExcelFactura";
			this.btExcelFactura.Size = new System.Drawing.Size(40, 24);
			this.btExcelFactura.TabIndex = 15;
			this.btExcelFactura.Text = "Excel";
			this.btExcelFactura.Click += new System.EventHandler(this.btExcelFactura_Click);
			// 
			// btRetencion
			// 
			this.btRetencion.Location = new System.Drawing.Point(696, 8);
			this.btRetencion.Name = "btRetencion";
			this.btRetencion.Size = new System.Drawing.Size(40, 24);
			this.btRetencion.TabIndex = 8;
			this.btRetencion.Text = "Ret";
			this.toolTip1.SetToolTip(this.btRetencion, "Muestra Asiento de Comprobante");
			this.btRetencion.Click += new System.EventHandler(this.btRetencion_Click);
			// 
			// ultraButton2
			// 
			this.ultraButton2.Location = new System.Drawing.Point(656, 8);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(40, 24);
			this.ultraButton2.TabIndex = 6;
			this.ultraButton2.Text = "A.C.";
			this.toolTip1.SetToolTip(this.ultraButton2, "Muestra Asiento de Comprobante");
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(616, 8);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(40, 24);
			this.ultraButton1.TabIndex = 5;
			this.ultraButton1.Text = "A.R";
			this.toolTip1.SetToolTip(this.ultraButton1, "Muestra Asiento de Retencion");
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// btAsiF
			// 
			this.btAsiF.Location = new System.Drawing.Point(576, 8);
			this.btAsiF.Name = "btAsiF";
			this.btAsiF.Size = new System.Drawing.Size(40, 24);
			this.btAsiF.TabIndex = 4;
			this.btAsiF.Text = "Asi";
			this.toolTip1.SetToolTip(this.btAsiF, "Muestra Asiento de Factura");
			this.btAsiF.Click += new System.EventHandler(this.btAsiF_Click);
			// 
			// btFactura
			// 
			this.btFactura.Location = new System.Drawing.Point(536, 8);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(40, 24);
			this.btFactura.TabIndex = 3;
			this.btFactura.Text = "Fact";
			this.toolTip1.SetToolTip(this.btFactura, "Muestra Factura Seleccionada");
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// btGeneraFact
			// 
			this.btGeneraFact.Location = new System.Drawing.Point(480, 8);
			this.btGeneraFact.Name = "btGeneraFact";
			this.btGeneraFact.Size = new System.Drawing.Size(56, 24);
			this.btGeneraFact.TabIndex = 1;
			this.btGeneraFact.Text = "Genera";
			this.toolTip1.SetToolTip(this.btGeneraFact, "Listado de Facturas / Clic Derecho Resumen de Errores en facturas");
			this.btGeneraFact.Click += new System.EventHandler(this.btGeneraFact_Click);
			this.btGeneraFact.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btGeneraFact_MouseUp);
			// 
			// cmbTipoFactura
			// 
			this.cmbTipoFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFactura.DataMember = "CompraNumero";
			this.cmbTipoFactura.DataSource = this.cdsTipoFactura;
			this.cmbTipoFactura.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 5;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 149;
			ultraGridColumn3.Header.VisiblePosition = 7;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 129;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 133;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 114;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 48;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 46;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 36;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 36;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 23;
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
																										 ultraGridColumn17});
			this.cmbTipoFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoFactura.DisplayMember = "Nombre";
			this.cmbTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFactura.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFactura.Location = new System.Drawing.Point(16, 8);
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(168, 21);
			this.cmbTipoFactura.TabIndex = 0;
			this.cmbTipoFactura.ValueMember = "idTipoFactura";
			// 
			// cdsTipoFactura
			// 
			this.cdsTipoFactura.BindingContextCtrl = this;
			this.cdsTipoFactura.DataLibrary = "LibFacturacion";
			this.cdsTipoFactura.DataLibraryUrl = "";
			this.cdsTipoFactura.DataSetDef = "dsCompraTabla";
			this.cdsTipoFactura.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoFactura.Name = "cdsTipoFactura";
			this.cdsTipoFactura.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipoFactura.SchemaDef = null;
			this.cdsTipoFactura.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoFactura_BeforeFill);
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.btExcelPago);
			this.ultraTabPageControl1.Controls.Add(this.btPago);
			this.ultraTabPageControl1.Controls.Add(this.label4);
			this.ultraTabPageControl1.Controls.Add(this.btAsiPago);
			this.ultraTabPageControl1.Controls.Add(this.btFacturaPago);
			this.ultraTabPageControl1.Controls.Add(this.btGeneraPago);
			this.ultraTabPageControl1.Controls.Add(this.cmbTipoCobro);
			this.ultraTabPageControl1.Controls.Add(this.grdPago);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(909, 328);
			// 
			// btExcelPago
			// 
			this.btExcelPago.Location = new System.Drawing.Point(496, 9);
			this.btExcelPago.Name = "btExcelPago";
			this.btExcelPago.Size = new System.Drawing.Size(48, 24);
			this.btExcelPago.TabIndex = 14;
			this.btExcelPago.Text = "Excel";
			this.toolTip1.SetToolTip(this.btExcelPago, "Muesta Pago Seleccionado");
			this.btExcelPago.Click += new System.EventHandler(this.btExcelPago_Click);
			// 
			// btPago
			// 
			this.btPago.Location = new System.Drawing.Point(440, 9);
			this.btPago.Name = "btPago";
			this.btPago.Size = new System.Drawing.Size(48, 24);
			this.btPago.TabIndex = 13;
			this.btPago.Text = "Pago";
			this.toolTip1.SetToolTip(this.btPago, "Muesta Pago Seleccionado");
			this.btPago.Click += new System.EventHandler(this.btPago_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(16, 9);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(82, 16);
			this.label4.TabIndex = 12;
			this.label4.Text = "Tipo de Factura";
			// 
			// btAsiPago
			// 
			this.btAsiPago.Location = new System.Drawing.Point(392, 9);
			this.btAsiPago.Name = "btAsiPago";
			this.btAsiPago.Size = new System.Drawing.Size(40, 24);
			this.btAsiPago.TabIndex = 11;
			this.btAsiPago.Text = "Asi";
			this.toolTip1.SetToolTip(this.btAsiPago, "Muestra Asiento de Factura");
			this.btAsiPago.Click += new System.EventHandler(this.btAsiPago_Click);
			// 
			// btFacturaPago
			// 
			this.btFacturaPago.Location = new System.Drawing.Point(344, 9);
			this.btFacturaPago.Name = "btFacturaPago";
			this.btFacturaPago.Size = new System.Drawing.Size(40, 24);
			this.btFacturaPago.TabIndex = 10;
			this.btFacturaPago.Text = "Fact";
			this.toolTip1.SetToolTip(this.btFacturaPago, "Muesta Factura Seleccionada");
			this.btFacturaPago.Click += new System.EventHandler(this.btFacturaPago_Click);
			// 
			// btGeneraPago
			// 
			this.btGeneraPago.Location = new System.Drawing.Point(280, 9);
			this.btGeneraPago.Name = "btGeneraPago";
			this.btGeneraPago.Size = new System.Drawing.Size(56, 24);
			this.btGeneraPago.TabIndex = 9;
			this.btGeneraPago.Text = "Generar";
			this.btGeneraPago.Click += new System.EventHandler(this.btGeneraPago_Click);
			// 
			// cmbTipoCobro
			// 
			this.cmbTipoCobro.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoCobro.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoCobro.DataMember = "CompraNumero";
			this.cmbTipoCobro.DataSource = this.cdsTipoFactura;
			this.cmbTipoCobro.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn18.Header.VisiblePosition = 5;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Width = 149;
			ultraGridColumn20.Header.VisiblePosition = 7;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 6;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 8;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 4;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 9;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 129;
			ultraGridColumn28.Header.VisiblePosition = 10;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 133;
			ultraGridColumn29.Header.VisiblePosition = 11;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 114;
			ultraGridColumn30.Header.VisiblePosition = 12;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 48;
			ultraGridColumn31.Header.VisiblePosition = 13;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 46;
			ultraGridColumn32.Header.VisiblePosition = 14;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 36;
			ultraGridColumn33.Header.VisiblePosition = 15;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 36;
			ultraGridColumn34.Header.VisiblePosition = 16;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 23;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn34});
			this.cmbTipoCobro.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipoCobro.DisplayMember = "Nombre";
			this.cmbTipoCobro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoCobro.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoCobro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoCobro.Location = new System.Drawing.Point(104, 9);
			this.cmbTipoCobro.Name = "cmbTipoCobro";
			this.cmbTipoCobro.Size = new System.Drawing.Size(168, 21);
			this.cmbTipoCobro.TabIndex = 8;
			this.cmbTipoCobro.ValueMember = "idTipoFactura";
			// 
			// grdPago
			// 
			this.grdPago.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdPago.DataMember = "Band 0";
			this.grdPago.DataSource = this.udsPago;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdPago.DisplayLayout.Appearance = appearance1;
			this.grdPago.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 56;
			ultraGridColumn36.Header.Caption = "Factura";
			ultraGridColumn36.Header.VisiblePosition = 2;
			ultraGridColumn36.Width = 115;
			ultraGridColumn37.Header.VisiblePosition = 4;
			ultraGridColumn37.Width = 113;
			ultraGridColumn38.Header.VisiblePosition = 3;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 87;
			ultraGridColumn39.Header.VisiblePosition = 5;
			ultraGridColumn39.Width = 136;
			ultraGridColumn40.Header.VisiblePosition = 8;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 66;
			ultraGridColumn41.Header.Caption = "idAsientoLote";
			ultraGridColumn41.Header.VisiblePosition = 6;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 39;
			ultraGridColumn42.Header.VisiblePosition = 7;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 30;
			ultraGridColumn43.Format = "dd/MMM/yyyy";
			ultraGridColumn43.Header.VisiblePosition = 1;
			ultraGridColumn43.Width = 106;
			ultraGridColumn44.Header.VisiblePosition = 9;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 73;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance2;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.VisiblePosition = 10;
			ultraGridColumn45.Width = 108;
			ultraGridColumn46.Header.Caption = "Comprob";
			ultraGridColumn46.Header.VisiblePosition = 11;
			ultraGridColumn46.Width = 135;
			ultraGridColumn47.Header.VisiblePosition = 12;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 71;
			ultraGridColumn48.Header.VisiblePosition = 13;
			ultraGridColumn48.Width = 146;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn48});
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand3.SummaryFooterCaption = "";
			this.grdPago.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.grdPago.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdPago.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grdPago.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			this.grdPago.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.grdPago.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.grdPago.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.grdPago.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdPago.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.grdPago.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdPago.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdPago.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.grdPago.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grdPago.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grdPago.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.grdPago.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.grdPago.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.grdPago.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.grdPago.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.grdPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdPago.Location = new System.Drawing.Point(14, 43);
			this.grdPago.Name = "grdPago";
			this.grdPago.Size = new System.Drawing.Size(880, 269);
			this.grdPago.TabIndex = 3;
			// 
			// udsPago
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DefaultValue = "";
			ultraDataColumn3.DefaultValue = "";
			ultraDataColumn4.DefaultValue = "";
			ultraDataColumn5.DefaultValue = "";
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(System.Double);
			this.udsPago.Band.Columns.AddRange(new object[] {
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
																												ultraDataColumn14});
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.btTotalRet);
			this.ultraTabPageControl2.Controls.Add(this.txtTotalRet);
			this.ultraTabPageControl2.Controls.Add(this.rdgTipoRet);
			this.ultraTabPageControl2.Controls.Add(this.chkAnuladaRet);
			this.ultraTabPageControl2.Controls.Add(this.btRetencionRet);
			this.ultraTabPageControl2.Controls.Add(this.btFacturaRet);
			this.ultraTabPageControl2.Controls.Add(this.btGeneraRetencion);
			this.ultraTabPageControl2.Controls.Add(this.grdRetencion);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(909, 328);
			// 
			// btTotalRet
			// 
			this.btTotalRet.Location = new System.Drawing.Point(376, 8);
			this.btTotalRet.Name = "btTotalRet";
			this.btTotalRet.Size = new System.Drawing.Size(48, 24);
			this.btTotalRet.TabIndex = 25;
			this.btTotalRet.Text = "Total";
			this.btTotalRet.Click += new System.EventHandler(this.btTotalRet_Click);
			// 
			// txtTotalRet
			// 
			this.txtTotalRet.Location = new System.Drawing.Point(432, 8);
			this.txtTotalRet.Name = "txtTotalRet";
			this.txtTotalRet.Size = new System.Drawing.Size(80, 20);
			this.txtTotalRet.TabIndex = 24;
			// 
			// rdgTipoRet
			// 
			this.rdgTipoRet.EditValue = 1;
			this.rdgTipoRet.Location = new System.Drawing.Point(16, 8);
			this.rdgTipoRet.Name = "rdgTipoRet";
			// 
			// rdgTipoRet.Properties
			// 
			this.rdgTipoRet.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			 new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Realizadas"),
																																																			 new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Recibidas")});
			this.rdgTipoRet.Size = new System.Drawing.Size(216, 32);
			this.rdgTipoRet.TabIndex = 23;
			// 
			// chkAnuladaRet
			// 
			this.chkAnuladaRet.EditValue = true;
			this.chkAnuladaRet.Location = new System.Drawing.Point(248, 8);
			this.chkAnuladaRet.Name = "chkAnuladaRet";
			// 
			// chkAnuladaRet.Properties
			// 
			this.chkAnuladaRet.Properties.Caption = "Anulado";
			this.chkAnuladaRet.Size = new System.Drawing.Size(64, 18);
			this.chkAnuladaRet.TabIndex = 22;
			// 
			// btRetencionRet
			// 
			this.btRetencionRet.Location = new System.Drawing.Point(752, 8);
			this.btRetencionRet.Name = "btRetencionRet";
			this.btRetencionRet.Size = new System.Drawing.Size(72, 24);
			this.btRetencionRet.TabIndex = 21;
			this.btRetencionRet.Text = "Retención";
			this.btRetencionRet.Click += new System.EventHandler(this.btRetencionRet_Click);
			// 
			// btFacturaRet
			// 
			this.btFacturaRet.Location = new System.Drawing.Point(672, 8);
			this.btFacturaRet.Name = "btFacturaRet";
			this.btFacturaRet.Size = new System.Drawing.Size(72, 24);
			this.btFacturaRet.TabIndex = 20;
			this.btFacturaRet.Text = "Factura";
			this.btFacturaRet.Click += new System.EventHandler(this.btFacturaRet_Click);
			// 
			// btGeneraRetencion
			// 
			this.btGeneraRetencion.Location = new System.Drawing.Point(320, 8);
			this.btGeneraRetencion.Name = "btGeneraRetencion";
			this.btGeneraRetencion.Size = new System.Drawing.Size(48, 24);
			this.btGeneraRetencion.TabIndex = 19;
			this.btGeneraRetencion.Text = "Genera";
			this.btGeneraRetencion.Click += new System.EventHandler(this.btGeneraRetencion_Click);
			this.btGeneraRetencion.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btGeneraRetencion_MouseUp);
			// 
			// grdRetencion
			// 
			this.grdRetencion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdRetencion.EmbeddedNavigator
			// 
			this.grdRetencion.EmbeddedNavigator.Name = "";
			this.grdRetencion.Location = new System.Drawing.Point(10, 40);
			this.grdRetencion.MainView = this.grvRetencion;
			this.grdRetencion.Name = "grdRetencion";
			this.grdRetencion.Size = new System.Drawing.Size(887, 277);
			this.grdRetencion.TabIndex = 18;
			this.grdRetencion.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.grvRetencion,
																																																this.gridView12});
			this.grdRetencion.DoubleClick += new System.EventHandler(this.grdRetencion_DoubleClick);
			// 
			// grvRetencion
			// 
			this.grvRetencion.GridControl = this.grdRetencion;
			this.grvRetencion.Name = "grvRetencion";
			this.grvRetencion.OptionsSelection.MultiSelect = true;
			this.grvRetencion.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvRetencion.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView12
			// 
			this.gridView12.GridControl = this.grdRetencion;
			this.gridView12.Name = "gridView12";
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.simpleButton2);
			this.ultraTabPageControl4.Controls.Add(this.simpleButton1);
			this.ultraTabPageControl4.Controls.Add(this.btFacturaGrupo);
			this.ultraTabPageControl4.Controls.Add(this.btTotalFG);
			this.ultraTabPageControl4.Controls.Add(this.btComparaFG);
			this.ultraTabPageControl4.Controls.Add(this.btGeneraFacturaGrupo);
			this.ultraTabPageControl4.Controls.Add(this.txtDevolucion);
			this.ultraTabPageControl4.Controls.Add(this.txtTotalGrupo);
			this.ultraTabPageControl4.Controls.Add(this.grdGrupoFactura);
			this.ultraTabPageControl4.Controls.Add(this.chkGrupoAgrupado);
			this.ultraTabPageControl4.Controls.Add(this.cmbTipoFact1);
			this.ultraTabPageControl4.Controls.Add(this.txtDescuentoGrupo);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(909, 328);
			// 
			// simpleButton2
			// 
			this.simpleButton2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.simpleButton2.Location = new System.Drawing.Point(849, 8);
			this.simpleButton2.Name = "simpleButton2";
			this.simpleButton2.Size = new System.Drawing.Size(40, 24);
			this.simpleButton2.TabIndex = 75;
			this.simpleButton2.Text = "Excel";
			this.simpleButton2.ToolTip = "Muesta Pago Seleccionado";
			this.simpleButton2.Click += new System.EventHandler(this.btExcelGrupo_Click);
			// 
			// simpleButton1
			// 
			this.simpleButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.simpleButton1.Location = new System.Drawing.Point(809, 8);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(40, 24);
			this.simpleButton1.TabIndex = 74;
			this.simpleButton1.Text = "Asi";
			this.simpleButton1.ToolTip = "Muestra Asiento de Factura";
			this.simpleButton1.Click += new System.EventHandler(this.btAsientoGrupo_Click);
			// 
			// btFacturaGrupo
			// 
			this.btFacturaGrupo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btFacturaGrupo.Location = new System.Drawing.Point(769, 8);
			this.btFacturaGrupo.Name = "btFacturaGrupo";
			this.btFacturaGrupo.Size = new System.Drawing.Size(40, 24);
			this.btFacturaGrupo.TabIndex = 73;
			this.btFacturaGrupo.Text = "Fact";
			this.btFacturaGrupo.ToolTip = "Muesta Factura Seleccionada";
			this.btFacturaGrupo.Click += new System.EventHandler(this.btFacturaGrupo_Click);
			// 
			// btTotalFG
			// 
			this.btTotalFG.Location = new System.Drawing.Point(416, 8);
			this.btTotalFG.Name = "btTotalFG";
			this.btTotalFG.Size = new System.Drawing.Size(56, 24);
			this.btTotalFG.TabIndex = 72;
			this.btTotalFG.Text = "Total";
			this.btTotalFG.ToolTip = "Total, Devolución, Descuento / Total, Devolución + Descuento, Neto";
			this.btTotalFG.Click += new System.EventHandler(this.btTotalFG_Click);
			this.btTotalFG.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btTotalFG_MouseUp);
			// 
			// btComparaFG
			// 
			this.btComparaFG.Location = new System.Drawing.Point(320, 8);
			this.btComparaFG.Name = "btComparaFG";
			this.btComparaFG.Size = new System.Drawing.Size(56, 24);
			this.btComparaFG.TabIndex = 71;
			this.btComparaFG.Text = "Compara";
			this.btComparaFG.Click += new System.EventHandler(this.btComparaFG_Click);
			// 
			// btGeneraFacturaGrupo
			// 
			this.btGeneraFacturaGrupo.Location = new System.Drawing.Point(256, 8);
			this.btGeneraFacturaGrupo.Name = "btGeneraFacturaGrupo";
			this.btGeneraFacturaGrupo.Size = new System.Drawing.Size(56, 24);
			this.btGeneraFacturaGrupo.TabIndex = 70;
			this.btGeneraFacturaGrupo.Text = "Generar";
			this.btGeneraFacturaGrupo.Click += new System.EventHandler(this.btGeneraFacturaGrupo_Click);
			// 
			// txtDevolucion
			// 
			this.txtDevolucion.Location = new System.Drawing.Point(544, 10);
			this.txtDevolucion.Name = "txtDevolucion";
			// 
			// txtDevolucion.Properties
			// 
			this.txtDevolucion.Properties.Appearance.Options.UseTextOptions = true;
			this.txtDevolucion.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.txtDevolucion.Properties.DisplayFormat.FormatString = "n2";
			this.txtDevolucion.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtDevolucion.Properties.EditFormat.FormatString = "n2";
			this.txtDevolucion.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtDevolucion.Properties.ReadOnly = true;
			this.txtDevolucion.Size = new System.Drawing.Size(72, 20);
			this.txtDevolucion.TabIndex = 69;
			this.txtDevolucion.ToolTip = "Devolucion y Descuento";
			// 
			// txtTotalGrupo
			// 
			this.txtTotalGrupo.Location = new System.Drawing.Point(472, 10);
			this.txtTotalGrupo.Name = "txtTotalGrupo";
			// 
			// txtTotalGrupo.Properties
			// 
			this.txtTotalGrupo.Properties.Appearance.Options.UseTextOptions = true;
			this.txtTotalGrupo.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.txtTotalGrupo.Properties.DisplayFormat.FormatString = "n2";
			this.txtTotalGrupo.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtTotalGrupo.Properties.EditFormat.FormatString = "n2";
			this.txtTotalGrupo.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtTotalGrupo.Properties.ReadOnly = true;
			this.txtTotalGrupo.Size = new System.Drawing.Size(72, 20);
			this.txtTotalGrupo.TabIndex = 31;
			this.txtTotalGrupo.ToolTip = "Total";
			// 
			// grdGrupoFactura
			// 
			this.grdGrupoFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdGrupoFactura.EmbeddedNavigator
			// 
			this.grdGrupoFactura.EmbeddedNavigator.Name = "";
			this.grdGrupoFactura.Location = new System.Drawing.Point(16, 40);
			this.grdGrupoFactura.MainView = this.grvGrupoFactura;
			this.grdGrupoFactura.Name = "grdGrupoFactura";
			this.grdGrupoFactura.Size = new System.Drawing.Size(881, 277);
			this.grdGrupoFactura.TabIndex = 30;
			this.grdGrupoFactura.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																	 this.grvGrupoFactura,
																																																	 this.gridView7});
			// 
			// grvGrupoFactura
			// 
			this.grvGrupoFactura.GridControl = this.grdGrupoFactura;
			this.grvGrupoFactura.Name = "grvGrupoFactura";
			this.grvGrupoFactura.OptionsSelection.MultiSelect = true;
			this.grvGrupoFactura.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvGrupoFactura.OptionsView.ShowAutoFilterRow = true;
			this.grvGrupoFactura.DoubleClick += new System.EventHandler(this.grvGrupoFactura_DoubleClick);
			// 
			// gridView7
			// 
			this.gridView7.GridControl = this.grdGrupoFactura;
			this.gridView7.Name = "gridView7";
			// 
			// chkGrupoAgrupado
			// 
			this.chkGrupoAgrupado.EditValue = true;
			this.chkGrupoAgrupado.Location = new System.Drawing.Point(176, 8);
			this.chkGrupoAgrupado.Name = "chkGrupoAgrupado";
			// 
			// chkGrupoAgrupado.Properties
			// 
			this.chkGrupoAgrupado.Properties.AllowGrayed = true;
			this.chkGrupoAgrupado.Properties.Caption = "Agrupado";
			this.chkGrupoAgrupado.Size = new System.Drawing.Size(72, 18);
			this.chkGrupoAgrupado.TabIndex = 29;
			this.toolTip1.SetToolTip(this.chkGrupoAgrupado, "Agrupa por Cuenta / Agrupa por Cuenta y Proyecto / Sin Agrupacion");
			// 
			// cmbTipoFact1
			// 
			this.cmbTipoFact1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFact1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFact1.DataMember = "CompraNumero";
			this.cmbTipoFact1.DataSource = this.cdsTipoFactura;
			this.cmbTipoFact1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn49.Header.VisiblePosition = 5;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 0;
			ultraGridColumn50.Width = 133;
			ultraGridColumn51.Header.VisiblePosition = 7;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 2;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 3;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 1;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 6;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 8;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 4;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 9;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 129;
			ultraGridColumn59.Header.VisiblePosition = 10;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 133;
			ultraGridColumn60.Header.VisiblePosition = 11;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 114;
			ultraGridColumn61.Header.VisiblePosition = 12;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn61.Width = 48;
			ultraGridColumn62.Header.VisiblePosition = 13;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 46;
			ultraGridColumn63.Header.VisiblePosition = 14;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Width = 36;
			ultraGridColumn64.Header.VisiblePosition = 15;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 36;
			ultraGridColumn65.Header.VisiblePosition = 16;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 23;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn65});
			this.cmbTipoFact1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipoFact1.DisplayMember = "Nombre";
			this.cmbTipoFact1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFact1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoFact1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFact1.Location = new System.Drawing.Point(16, 8);
			this.cmbTipoFact1.Name = "cmbTipoFact1";
			this.cmbTipoFact1.Size = new System.Drawing.Size(152, 21);
			this.cmbTipoFact1.TabIndex = 18;
			this.cmbTipoFact1.ValueMember = "idTipoFactura";
			// 
			// txtDescuentoGrupo
			// 
			this.txtDescuentoGrupo.Location = new System.Drawing.Point(616, 10);
			this.txtDescuentoGrupo.Name = "txtDescuentoGrupo";
			// 
			// txtDescuentoGrupo.Properties
			// 
			this.txtDescuentoGrupo.Properties.Appearance.Options.UseTextOptions = true;
			this.txtDescuentoGrupo.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.txtDescuentoGrupo.Properties.DisplayFormat.FormatString = "n2";
			this.txtDescuentoGrupo.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtDescuentoGrupo.Properties.ReadOnly = true;
			this.txtDescuentoGrupo.Size = new System.Drawing.Size(72, 20);
			this.txtDescuentoGrupo.TabIndex = 67;
			this.txtDescuentoGrupo.ToolTip = "Neto";
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.btExcelCredito);
			this.ultraTabPageControl5.Controls.Add(this.btPagoCred);
			this.ultraTabPageControl5.Controls.Add(this.btCreditoAsiento);
			this.ultraTabPageControl5.Controls.Add(this.btCreditoFactura);
			this.ultraTabPageControl5.Controls.Add(this.btInconsistencia);
			this.ultraTabPageControl5.Controls.Add(this.btCreditoGenera);
			this.ultraTabPageControl5.Controls.Add(this.btCreditoCompara);
			this.ultraTabPageControl5.Controls.Add(this.grdCredito);
			this.ultraTabPageControl5.Controls.Add(this.chkResumenCompra);
			this.ultraTabPageControl5.Controls.Add(this.chkAgrupaCredito);
			this.ultraTabPageControl5.Controls.Add(this.label7);
			this.ultraTabPageControl5.Controls.Add(this.cmbTipoFactCred);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(909, 328);
			// 
			// btExcelCredito
			// 
			this.btExcelCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btExcelCredito.Location = new System.Drawing.Point(849, 8);
			this.btExcelCredito.Name = "btExcelCredito";
			this.btExcelCredito.Size = new System.Drawing.Size(40, 24);
			this.btExcelCredito.TabIndex = 77;
			this.btExcelCredito.Text = "Excel";
			this.btExcelCredito.ToolTip = "Muesta Pago Seleccionado";
			this.btExcelCredito.Click += new System.EventHandler(this.btExcelCredito_Click);
			this.btExcelCredito.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btExcelCredito_MouseUp);
			// 
			// btPagoCred
			// 
			this.btPagoCred.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btPagoCred.Location = new System.Drawing.Point(809, 8);
			this.btPagoCred.Name = "btPagoCred";
			this.btPagoCred.Size = new System.Drawing.Size(40, 24);
			this.btPagoCred.TabIndex = 76;
			this.btPagoCred.Text = "Pago";
			this.btPagoCred.ToolTip = "Muestra Pago de Factura";
			this.btPagoCred.Click += new System.EventHandler(this.btPagoCred_Click);
			// 
			// btCreditoAsiento
			// 
			this.btCreditoAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btCreditoAsiento.Location = new System.Drawing.Point(769, 8);
			this.btCreditoAsiento.Name = "btCreditoAsiento";
			this.btCreditoAsiento.Size = new System.Drawing.Size(40, 24);
			this.btCreditoAsiento.TabIndex = 75;
			this.btCreditoAsiento.Text = "Asi";
			this.btCreditoAsiento.ToolTip = "Muestra Asiento de Factura";
			this.btCreditoAsiento.Click += new System.EventHandler(this.btCreditoAsiento_Click);
			// 
			// btCreditoFactura
			// 
			this.btCreditoFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btCreditoFactura.Location = new System.Drawing.Point(729, 8);
			this.btCreditoFactura.Name = "btCreditoFactura";
			this.btCreditoFactura.Size = new System.Drawing.Size(40, 24);
			this.btCreditoFactura.TabIndex = 74;
			this.btCreditoFactura.Text = "Fact";
			this.btCreditoFactura.ToolTip = "Muesta Factura Seleccionada";
			this.btCreditoFactura.Click += new System.EventHandler(this.btCreditoFactura_Click);
			// 
			// btInconsistencia
			// 
			this.btInconsistencia.Location = new System.Drawing.Point(560, 8);
			this.btInconsistencia.Name = "btInconsistencia";
			this.btInconsistencia.Size = new System.Drawing.Size(64, 23);
			this.btInconsistencia.TabIndex = 69;
			this.btInconsistencia.Text = "Diferencia";
			this.btInconsistencia.Click += new System.EventHandler(this.btInconsistencia_Click);
			this.btInconsistencia.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btInconsistencia_MouseUp);
			// 
			// btCreditoGenera
			// 
			this.btCreditoGenera.Location = new System.Drawing.Point(448, 8);
			this.btCreditoGenera.Name = "btCreditoGenera";
			this.btCreditoGenera.Size = new System.Drawing.Size(56, 23);
			this.btCreditoGenera.TabIndex = 68;
			this.btCreditoGenera.Text = "Genera";
			this.btCreditoGenera.Click += new System.EventHandler(this.btCreditoGenera_Click);
			// 
			// btCreditoCompara
			// 
			this.btCreditoCompara.Location = new System.Drawing.Point(504, 8);
			this.btCreditoCompara.Name = "btCreditoCompara";
			this.btCreditoCompara.Size = new System.Drawing.Size(56, 23);
			this.btCreditoCompara.TabIndex = 67;
			this.btCreditoCompara.Text = "Compara";
			this.btCreditoCompara.Click += new System.EventHandler(this.btCreditoCompara_Click);
			// 
			// grdCredito
			// 
			this.grdCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdCredito.EmbeddedNavigator
			// 
			this.grdCredito.EmbeddedNavigator.Name = "";
			this.grdCredito.Location = new System.Drawing.Point(14, 40);
			this.grdCredito.MainView = this.grvCredito;
			this.grdCredito.Name = "grdCredito";
			this.grdCredito.Size = new System.Drawing.Size(881, 277);
			this.grdCredito.TabIndex = 66;
			this.grdCredito.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															this.grvCredito,
																																															this.gridView10});
			// 
			// grvCredito
			// 
			this.grvCredito.GridControl = this.grdCredito;
			this.grvCredito.Name = "grvCredito";
			this.grvCredito.OptionsSelection.MultiSelect = true;
			this.grvCredito.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvCredito.OptionsView.ColumnAutoWidth = false;
			this.grvCredito.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView10
			// 
			this.gridView10.GridControl = this.grdCredito;
			this.gridView10.Name = "gridView10";
			// 
			// chkResumenCompra
			// 
			this.chkResumenCompra.EditValue = true;
			this.chkResumenCompra.Location = new System.Drawing.Point(320, 8);
			this.chkResumenCompra.Name = "chkResumenCompra";
			// 
			// chkResumenCompra.Properties
			// 
			this.chkResumenCompra.Properties.Caption = "Resumen Compra";
			this.chkResumenCompra.Size = new System.Drawing.Size(112, 18);
			this.chkResumenCompra.TabIndex = 65;
			// 
			// chkAgrupaCredito
			// 
			this.chkAgrupaCredito.EditValue = true;
			this.chkAgrupaCredito.Location = new System.Drawing.Point(224, 8);
			this.chkAgrupaCredito.Name = "chkAgrupaCredito";
			// 
			// chkAgrupaCredito.Properties
			// 
			this.chkAgrupaCredito.Properties.Caption = "Agrupado";
			this.chkAgrupaCredito.Size = new System.Drawing.Size(80, 18);
			this.chkAgrupaCredito.TabIndex = 37;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(15, 9);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(26, 16);
			this.label7.TabIndex = 35;
			this.label7.Text = "Tipo";
			// 
			// cmbTipoFactCred
			// 
			this.cmbTipoFactCred.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactCred.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFactCred.DataMember = "CompraNumero";
			this.cmbTipoFactCred.DataSource = this.cdsTipoFactura;
			this.cmbTipoFactCred.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn66.Header.VisiblePosition = 5;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 0;
			ultraGridColumn67.Width = 149;
			ultraGridColumn68.Header.VisiblePosition = 7;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 2;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 3;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 1;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 6;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 8;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 4;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 9;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn75.Width = 129;
			ultraGridColumn76.Header.VisiblePosition = 10;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.Width = 133;
			ultraGridColumn77.Header.VisiblePosition = 11;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn77.Width = 114;
			ultraGridColumn78.Header.VisiblePosition = 12;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn78.Width = 48;
			ultraGridColumn79.Header.VisiblePosition = 13;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn79.Width = 46;
			ultraGridColumn80.Header.VisiblePosition = 14;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn80.Width = 36;
			ultraGridColumn81.Header.VisiblePosition = 15;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn81.Width = 36;
			ultraGridColumn82.Header.VisiblePosition = 16;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn82.Width = 23;
			ultraGridBand5.Columns.AddRange(new object[] {
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
			this.cmbTipoFactCred.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbTipoFactCred.DisplayMember = "Nombre";
			this.cmbTipoFactCred.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFactCred.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoFactCred.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFactCred.Location = new System.Drawing.Point(47, 7);
			this.cmbTipoFactCred.Name = "cmbTipoFactCred";
			this.cmbTipoFactCred.Size = new System.Drawing.Size(168, 21);
			this.cmbTipoFactCred.TabIndex = 30;
			this.cmbTipoFactCred.ValueMember = "idTipoFactura";
			// 
			// ultraTabPageControl7
			// 
			this.ultraTabPageControl7.Controls.Add(this.chkAnuladoArt);
			this.ultraTabPageControl7.Controls.Add(this.btFacturaArticulo);
			this.ultraTabPageControl7.Controls.Add(this.btExcelArt);
			this.ultraTabPageControl7.Controls.Add(this.btTansaccion);
			this.ultraTabPageControl7.Controls.Add(this.btAsiArt);
			this.ultraTabPageControl7.Controls.Add(this.btArticuloB);
			this.ultraTabPageControl7.Controls.Add(this.btKardex);
			this.ultraTabPageControl7.Controls.Add(this.chkAgrupar);
			this.ultraTabPageControl7.Controls.Add(this.chkAjustar);
			this.ultraTabPageControl7.Controls.Add(this.cmbTipo);
			this.ultraTabPageControl7.Controls.Add(this.grdArticulos);
			this.ultraTabPageControl7.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl7.Name = "ultraTabPageControl7";
			this.ultraTabPageControl7.Size = new System.Drawing.Size(909, 328);
			// 
			// chkAnuladoArt
			// 
			this.chkAnuladoArt.EditValue = true;
			this.chkAnuladoArt.Location = new System.Drawing.Point(320, 8);
			this.chkAnuladoArt.Name = "chkAnuladoArt";
			// 
			// chkAnuladoArt.Properties
			// 
			this.chkAnuladoArt.Properties.Caption = "Anulados";
			this.chkAnuladoArt.Size = new System.Drawing.Size(64, 18);
			this.chkAnuladoArt.TabIndex = 19;
			// 
			// btFacturaArticulo
			// 
			this.btFacturaArticulo.Location = new System.Drawing.Point(472, 8);
			this.btFacturaArticulo.Name = "btFacturaArticulo";
			this.btFacturaArticulo.Size = new System.Drawing.Size(64, 24);
			this.btFacturaArticulo.TabIndex = 18;
			this.btFacturaArticulo.Text = "Factura";
			this.btFacturaArticulo.Click += new System.EventHandler(this.btFacturaArticulo_Click);
			// 
			// btExcelArt
			// 
			this.btExcelArt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btExcelArt.Location = new System.Drawing.Point(849, 8);
			this.btExcelArt.Name = "btExcelArt";
			this.btExcelArt.Size = new System.Drawing.Size(48, 23);
			this.btExcelArt.TabIndex = 17;
			this.btExcelArt.Text = "Excel";
			this.btExcelArt.Click += new System.EventHandler(this.btExcelArt_Click);
			// 
			// btTansaccion
			// 
			this.btTansaccion.Location = new System.Drawing.Point(800, 8);
			this.btTansaccion.Name = "btTansaccion";
			this.btTansaccion.Size = new System.Drawing.Size(40, 24);
			this.btTansaccion.TabIndex = 8;
			this.btTansaccion.Text = "Trans.";
			this.btTansaccion.Click += new System.EventHandler(this.btTransaccion_Click);
			// 
			// btAsiArt
			// 
			this.btAsiArt.Location = new System.Drawing.Point(704, 8);
			this.btAsiArt.Name = "btAsiArt";
			this.btAsiArt.Size = new System.Drawing.Size(40, 24);
			this.btAsiArt.TabIndex = 7;
			this.btAsiArt.Text = "Asi.";
			this.btAsiArt.Click += new System.EventHandler(this.btAsiArt_Click);
			// 
			// btArticuloB
			// 
			this.btArticuloB.Location = new System.Drawing.Point(752, 8);
			this.btArticuloB.Name = "btArticuloB";
			this.btArticuloB.Size = new System.Drawing.Size(40, 24);
			this.btArticuloB.TabIndex = 6;
			this.btArticuloB.Text = "Art.";
			this.btArticuloB.Click += new System.EventHandler(this.btArticuloB_Click);
			// 
			// btKardex
			// 
			this.btKardex.Location = new System.Drawing.Point(400, 8);
			this.btKardex.Name = "btKardex";
			this.btKardex.Size = new System.Drawing.Size(64, 24);
			this.btKardex.TabIndex = 5;
			this.btKardex.Text = "Kardex";
			this.btKardex.Click += new System.EventHandler(this.btKardex_Click);
			// 
			// chkAgrupar
			// 
			this.chkAgrupar.EditValue = true;
			this.chkAgrupar.Location = new System.Drawing.Point(16, 8);
			this.chkAgrupar.Name = "chkAgrupar";
			// 
			// chkAgrupar.Properties
			// 
			this.chkAgrupar.Properties.Caption = "Agrupar";
			this.chkAgrupar.Size = new System.Drawing.Size(64, 18);
			this.chkAgrupar.TabIndex = 4;
			// 
			// chkAjustar
			// 
			this.chkAjustar.EditValue = true;
			this.chkAjustar.Location = new System.Drawing.Point(96, 8);
			this.chkAjustar.Name = "chkAjustar";
			// 
			// chkAjustar.Properties
			// 
			this.chkAjustar.Properties.Caption = "Ajustar";
			this.chkAjustar.Size = new System.Drawing.Size(64, 18);
			this.chkAjustar.TabIndex = 3;
			// 
			// cmbTipo
			// 
			this.cmbTipo.Location = new System.Drawing.Point(168, 8);
			this.cmbTipo.Name = "cmbTipo";
			// 
			// cmbTipo.Properties
			// 
			this.cmbTipo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbTipo.Properties.DisplayMember = "Nombre";
			this.cmbTipo.Properties.NullText = "";
			this.cmbTipo.Properties.ValueMember = "idTipoFactura";
			this.cmbTipo.Properties.View = this.gridLookUpEdit1View;
			this.cmbTipo.Size = new System.Drawing.Size(144, 20);
			this.cmbTipo.TabIndex = 1;
			this.cmbTipo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbTipo_MouseUp);
			// 
			// gridLookUpEdit1View
			// 
			this.gridLookUpEdit1View.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																															 this.gridColumn1});
			this.gridLookUpEdit1View.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus;
			this.gridLookUpEdit1View.Name = "gridLookUpEdit1View";
			this.gridLookUpEdit1View.OptionsSelection.EnableAppearanceFocusedCell = false;
			this.gridLookUpEdit1View.OptionsView.ShowGroupPanel = false;
			// 
			// gridColumn1
			// 
			this.gridColumn1.Caption = "Nombre";
			this.gridColumn1.FieldName = "Nombre";
			this.gridColumn1.Name = "gridColumn1";
			this.gridColumn1.Visible = true;
			this.gridColumn1.VisibleIndex = 0;
			// 
			// grdArticulos
			// 
			this.grdArticulos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdArticulos.EmbeddedNavigator
			// 
			this.grdArticulos.EmbeddedNavigator.Name = "";
			this.grdArticulos.Location = new System.Drawing.Point(16, 40);
			this.grdArticulos.MainView = this.grvArticulos;
			this.grdArticulos.Name = "grdArticulos";
			this.grdArticulos.Size = new System.Drawing.Size(881, 277);
			this.grdArticulos.TabIndex = 0;
			this.grdArticulos.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.grvArticulos,
																																																this.gridView1});
			// 
			// grvArticulos
			// 
			this.grvArticulos.GridControl = this.grdArticulos;
			this.grvArticulos.Name = "grvArticulos";
			this.grvArticulos.OptionsSelection.MultiSelect = true;
			this.grvArticulos.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvArticulos.OptionsView.ColumnAutoWidth = false;
			this.grvArticulos.OptionsView.ShowAutoFilterRow = true;
			this.grvArticulos.DoubleClick += new System.EventHandler(this.grvArticulos_DoubleClick);
			// 
			// gridView1
			// 
			this.gridView1.GridControl = this.grdArticulos;
			this.gridView1.Name = "gridView1";
			// 
			// ultraTabPageControl6
			// 
			this.ultraTabPageControl6.Controls.Add(this.btCargaCubo);
			this.ultraTabPageControl6.Controls.Add(this.btGrabaCubo);
			this.ultraTabPageControl6.Controls.Add(this.btExcelCubo);
			this.ultraTabPageControl6.Controls.Add(this.chkTotalF);
			this.ultraTabPageControl6.Controls.Add(this.chkTotalC);
			this.ultraTabPageControl6.Controls.Add(this.pgrBalance);
			this.ultraTabPageControl6.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl6.Name = "ultraTabPageControl6";
			this.ultraTabPageControl6.Size = new System.Drawing.Size(909, 328);
			// 
			// btCargaCubo
			// 
			this.btCargaCubo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btCargaCubo.Location = new System.Drawing.Point(713, 6);
			this.btCargaCubo.Name = "btCargaCubo";
			this.btCargaCubo.Size = new System.Drawing.Size(80, 23);
			this.btCargaCubo.TabIndex = 17;
			this.btCargaCubo.Text = "Carga Diseño";
			this.btCargaCubo.Click += new System.EventHandler(this.btCargaCubo_Click);
			// 
			// btGrabaCubo
			// 
			this.btGrabaCubo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btGrabaCubo.Location = new System.Drawing.Point(617, 6);
			this.btGrabaCubo.Name = "btGrabaCubo";
			this.btGrabaCubo.Size = new System.Drawing.Size(80, 23);
			this.btGrabaCubo.TabIndex = 16;
			this.btGrabaCubo.Text = "Graba Diseño";
			this.btGrabaCubo.Click += new System.EventHandler(this.btGrabaCubo_Click);
			// 
			// btExcelCubo
			// 
			this.btExcelCubo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btExcelCubo.Location = new System.Drawing.Point(817, 6);
			this.btExcelCubo.Name = "btExcelCubo";
			this.btExcelCubo.TabIndex = 15;
			this.btExcelCubo.Text = "Excel";
			this.btExcelCubo.Click += new System.EventHandler(this.btExcelCubo_Click);
			// 
			// chkTotalF
			// 
			this.chkTotalF.Location = new System.Drawing.Point(16, 8);
			this.chkTotalF.Name = "chkTotalF";
			// 
			// chkTotalF.Properties
			// 
			this.chkTotalF.Properties.Caption = "Total Filas";
			this.chkTotalF.Size = new System.Drawing.Size(104, 18);
			this.chkTotalF.TabIndex = 14;
			this.chkTotalF.CheckedChanged += new System.EventHandler(this.chkTotalF_CheckedChanged);
			// 
			// chkTotalC
			// 
			this.chkTotalC.EditValue = true;
			this.chkTotalC.Location = new System.Drawing.Point(120, 8);
			this.chkTotalC.Name = "chkTotalC";
			// 
			// chkTotalC.Properties
			// 
			this.chkTotalC.Properties.Caption = "Total Columnas";
			this.chkTotalC.Size = new System.Drawing.Size(104, 18);
			this.chkTotalC.TabIndex = 13;
			this.chkTotalC.CheckedChanged += new System.EventHandler(this.chkTotalC_CheckedChanged);
			// 
			// pgrBalance
			// 
			this.pgrBalance.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.pgrBalance.Cursor = System.Windows.Forms.Cursors.Default;
			this.pgrBalance.Fields.AddRange(new DevExpress.XtraPivotGrid.PivotGridField[] {
																																											this.pivotGridField1,
																																											this.pivotGridField2,
																																											this.pivotGridField3,
																																											this.pivotGridField4,
																																											this.pivotGridField5,
																																											this.pivotGridField6,
																																											this.pivotGridField7,
																																											this.pivotGridField8,
																																											this.pivotGridField9,
																																											this.pivotGridField10,
																																											this.pivotGridField11});
			this.pgrBalance.Location = new System.Drawing.Point(8, 32);
			this.pgrBalance.Name = "pgrBalance";
			this.pgrBalance.OptionsView.ShowColumnTotals = false;
			this.pgrBalance.OptionsView.ShowRowGrandTotals = false;
			this.pgrBalance.OptionsView.ShowRowTotals = false;
			this.pgrBalance.Size = new System.Drawing.Size(889, 285);
			this.pgrBalance.TabIndex = 0;
			this.pgrBalance.CellDoubleClick += new DevExpress.XtraPivotGrid.PivotCellEventHandler(this.pgrBalance_CellDoubleClick);
			// 
			// pivotGridField1
			// 
			this.pivotGridField1.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
			this.pivotGridField1.AreaIndex = 0;
			this.pivotGridField1.CellFormat.FormatString = "n2";
			this.pivotGridField1.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField1.EmptyValueText = "0";
			this.pivotGridField1.FieldName = "Saldo";
			this.pivotGridField1.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField1.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField1.Name = "pivotGridField1";
			this.pivotGridField1.TotalsVisibility = DevExpress.XtraPivotGrid.PivotTotalsVisibility.None;
			// 
			// pivotGridField2
			// 
			this.pivotGridField2.Area = DevExpress.XtraPivotGrid.PivotArea.RowArea;
			this.pivotGridField2.AreaIndex = 0;
			this.pivotGridField2.FieldName = "Codigo";
			this.pivotGridField2.Name = "pivotGridField2";
			this.pivotGridField2.Options.AllowExpand = DevExpress.Utils.DefaultBoolean.False;
			// 
			// pivotGridField3
			// 
			this.pivotGridField3.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField3.AreaIndex = 0;
			this.pivotGridField3.FieldName = "Proyecto";
			this.pivotGridField3.Name = "pivotGridField3";
			// 
			// pivotGridField4
			// 
			this.pivotGridField4.Area = DevExpress.XtraPivotGrid.PivotArea.ColumnArea;
			this.pivotGridField4.AreaIndex = 0;
			this.pivotGridField4.Caption = "Centro de Costo";
			this.pivotGridField4.FieldName = "CentroCosto";
			this.pivotGridField4.Name = "pivotGridField4";
			// 
			// pivotGridField5
			// 
			this.pivotGridField5.Area = DevExpress.XtraPivotGrid.PivotArea.RowArea;
			this.pivotGridField5.AreaIndex = 1;
			this.pivotGridField5.Caption = "Descripción";
			this.pivotGridField5.FieldName = "Descripcion";
			this.pivotGridField5.Name = "pivotGridField5";
			this.pivotGridField5.Width = 200;
			// 
			// pivotGridField6
			// 
			this.pivotGridField6.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField6.AreaIndex = 1;
			this.pivotGridField6.CellFormat.FormatString = "n2";
			this.pivotGridField6.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField6.FieldName = "Presupuesto";
			this.pivotGridField6.Name = "pivotGridField6";
			// 
			// pivotGridField7
			// 
			this.pivotGridField7.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField7.AreaIndex = 2;
			this.pivotGridField7.Caption = "Ejecución";
			this.pivotGridField7.CellFormat.FormatString = "n2";
			this.pivotGridField7.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField7.FieldName = "Contable";
			this.pivotGridField7.Name = "pivotGridField7";
			// 
			// pivotGridField8
			// 
			this.pivotGridField8.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField8.AreaIndex = 3;
			this.pivotGridField8.FieldName = "Porc";
			this.pivotGridField8.Name = "pivotGridField8";
			// 
			// pivotGridField9
			// 
			this.pivotGridField9.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField9.AreaIndex = 5;
			this.pivotGridField9.Caption = "Mes";
			this.pivotGridField9.FieldName = "Fecha";
			this.pivotGridField9.GroupInterval = DevExpress.XtraPivotGrid.PivotGroupInterval.DateMonth;
			this.pivotGridField9.Name = "pivotGridField9";
			// 
			// pivotGridField10
			// 
			this.pivotGridField10.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField10.AreaIndex = 4;
			this.pivotGridField10.Caption = "Año";
			this.pivotGridField10.FieldName = "Fecha";
			this.pivotGridField10.GroupInterval = DevExpress.XtraPivotGrid.PivotGroupInterval.DateYear;
			this.pivotGridField10.Name = "pivotGridField10";
			// 
			// pivotGridField11
			// 
			this.pivotGridField11.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
			this.pivotGridField11.AreaIndex = 1;
			this.pivotGridField11.CellFormat.FormatString = "n2";
			this.pivotGridField11.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField11.FieldName = "Periodo";
			this.pivotGridField11.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField11.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField11.Name = "pivotGridField11";
			this.pivotGridField11.TotalCellFormat.FormatString = "n2";
			this.pivotGridField11.TotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField11.TotalValueFormat.FormatString = "n2";
			this.pivotGridField11.TotalValueFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField11.ValueFormat.FormatString = "n2";
			this.pivotGridField11.ValueFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			// 
			// ultraTabPageControl8
			// 
			this.ultraTabPageControl8.Controls.Add(this.btVencido);
			this.ultraTabPageControl8.Controls.Add(this.btExcelAntig);
			this.ultraTabPageControl8.Controls.Add(this.btPagoAntig);
			this.ultraTabPageControl8.Controls.Add(this.btErrorAntiguedad);
			this.ultraTabPageControl8.Controls.Add(this.optVencidos);
			this.ultraTabPageControl8.Controls.Add(this.btFacturaAntig);
			this.ultraTabPageControl8.Controls.Add(this.optCobrar);
			this.ultraTabPageControl8.Controls.Add(this.btAntiguedad);
			this.ultraTabPageControl8.Controls.Add(this.grdAntiguedad);
			this.ultraTabPageControl8.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl8.Name = "ultraTabPageControl8";
			this.ultraTabPageControl8.Size = new System.Drawing.Size(909, 328);
			// 
			// btVencido
			// 
			this.btVencido.Location = new System.Drawing.Point(24, 40);
			this.btVencido.Name = "btVencido";
			this.btVencido.TabIndex = 76;
			this.btVencido.Text = "Vencido";
			this.btVencido.Click += new System.EventHandler(this.btVencido_Click);
			// 
			// btExcelAntig
			// 
			this.btExcelAntig.Location = new System.Drawing.Point(24, 288);
			this.btExcelAntig.Name = "btExcelAntig";
			this.btExcelAntig.Size = new System.Drawing.Size(72, 23);
			this.btExcelAntig.TabIndex = 75;
			this.btExcelAntig.Text = "Excel";
			this.btExcelAntig.Click += new System.EventHandler(this.btExcelAntig_Click);
			// 
			// btPagoAntig
			// 
			this.btPagoAntig.Location = new System.Drawing.Point(24, 256);
			this.btPagoAntig.Name = "btPagoAntig";
			this.btPagoAntig.Size = new System.Drawing.Size(72, 23);
			this.btPagoAntig.TabIndex = 74;
			this.btPagoAntig.Text = "Pago";
			this.btPagoAntig.Click += new System.EventHandler(this.btPagoAntig_Click);
			// 
			// btErrorAntiguedad
			// 
			this.btErrorAntiguedad.Location = new System.Drawing.Point(24, 192);
			this.btErrorAntiguedad.Name = "btErrorAntiguedad";
			this.btErrorAntiguedad.Size = new System.Drawing.Size(72, 23);
			this.btErrorAntiguedad.TabIndex = 73;
			this.btErrorAntiguedad.Text = "Errores";
			this.btErrorAntiguedad.Click += new System.EventHandler(this.btErrorAntiguedad_Click);
			// 
			// optVencidos
			// 
			this.optVencidos.EditValue = 0;
			this.optVencidos.Location = new System.Drawing.Point(16, 128);
			this.optVencidos.Name = "optVencidos";
			// 
			// optVencidos.Properties
			// 
			this.optVencidos.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																				new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Por vencer"),
																																																				new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Vencido")});
			this.optVencidos.Size = new System.Drawing.Size(88, 48);
			this.optVencidos.TabIndex = 72;
			// 
			// btFacturaAntig
			// 
			this.btFacturaAntig.Location = new System.Drawing.Point(24, 224);
			this.btFacturaAntig.Name = "btFacturaAntig";
			this.btFacturaAntig.Size = new System.Drawing.Size(72, 23);
			this.btFacturaAntig.TabIndex = 71;
			this.btFacturaAntig.Text = "Factura";
			this.btFacturaAntig.Click += new System.EventHandler(this.btFacturaAntig_Click);
			// 
			// optCobrar
			// 
			this.optCobrar.EditValue = 1;
			this.optCobrar.Location = new System.Drawing.Point(16, 72);
			this.optCobrar.Name = "optCobrar";
			// 
			// optCobrar.Properties
			// 
			this.optCobrar.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Cobro"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(4, "Pago")});
			this.optCobrar.Size = new System.Drawing.Size(88, 48);
			this.optCobrar.TabIndex = 70;
			// 
			// btAntiguedad
			// 
			this.btAntiguedad.Location = new System.Drawing.Point(24, 8);
			this.btAntiguedad.Name = "btAntiguedad";
			this.btAntiguedad.TabIndex = 69;
			this.btAntiguedad.Text = "Generar";
			this.btAntiguedad.Click += new System.EventHandler(this.btAntigueada_Click);
			// 
			// grdAntiguedad
			// 
			this.grdAntiguedad.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdAntiguedad.EmbeddedNavigator
			// 
			this.grdAntiguedad.EmbeddedNavigator.Name = "";
			this.grdAntiguedad.Location = new System.Drawing.Point(120, 8);
			this.grdAntiguedad.MainView = this.grvAntiguedad;
			this.grdAntiguedad.Name = "grdAntiguedad";
			this.grdAntiguedad.Size = new System.Drawing.Size(777, 309);
			this.grdAntiguedad.TabIndex = 68;
			this.grdAntiguedad.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																 this.grvAntiguedad,
																																																 this.gridView6});
			// 
			// grvAntiguedad
			// 
			this.grvAntiguedad.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus;
			this.grvAntiguedad.GridControl = this.grdAntiguedad;
			this.grvAntiguedad.Name = "grvAntiguedad";
			this.grvAntiguedad.OptionsSelection.MultiSelect = true;
			this.grvAntiguedad.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvAntiguedad.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView6
			// 
			this.gridView6.GridControl = this.grdAntiguedad;
			this.gridView6.Name = "gridView6";
			// 
			// ultraTabPageControl9
			// 
			this.ultraTabPageControl9.Controls.Add(this.btExcelForm);
			this.ultraTabPageControl9.Controls.Add(this.txtIvaForm);
			this.ultraTabPageControl9.Controls.Add(this.txtTotalFormNC);
			this.ultraTabPageControl9.Controls.Add(this.txtTotalForm);
			this.ultraTabPageControl9.Controls.Add(this.btTotalForm104);
			this.ultraTabPageControl9.Controls.Add(this.btTransForm);
			this.ultraTabPageControl9.Controls.Add(this.btArtForm);
			this.ultraTabPageControl9.Controls.Add(this.btGenera104);
			this.ultraTabPageControl9.Controls.Add(this.grdForm104);
			this.ultraTabPageControl9.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl9.Name = "ultraTabPageControl9";
			this.ultraTabPageControl9.Size = new System.Drawing.Size(909, 328);
			// 
			// btExcelForm
			// 
			this.btExcelForm.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btExcelForm.Location = new System.Drawing.Point(849, 8);
			this.btExcelForm.Name = "btExcelForm";
			this.btExcelForm.Size = new System.Drawing.Size(48, 23);
			this.btExcelForm.TabIndex = 16;
			this.btExcelForm.Text = "Excel";
			this.btExcelForm.Click += new System.EventHandler(this.btExcelForm_Click);
			// 
			// txtIvaForm
			// 
			this.txtIvaForm.Location = new System.Drawing.Point(392, 8);
			this.txtIvaForm.Name = "txtIvaForm";
			// 
			// txtIvaForm.Properties
			// 
			this.txtIvaForm.Properties.Appearance.Options.UseTextOptions = true;
			this.txtIvaForm.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.txtIvaForm.Properties.DisplayFormat.FormatString = "n2";
			this.txtIvaForm.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtIvaForm.Properties.ReadOnly = true;
			this.txtIvaForm.Size = new System.Drawing.Size(72, 20);
			this.txtIvaForm.TabIndex = 15;
			this.toolTip1.SetToolTip(this.txtIvaForm, "Impuesto Generado");
			// 
			// txtTotalFormNC
			// 
			this.txtTotalFormNC.Location = new System.Drawing.Point(304, 8);
			this.txtTotalFormNC.Name = "txtTotalFormNC";
			// 
			// txtTotalFormNC.Properties
			// 
			this.txtTotalFormNC.Properties.Appearance.Options.UseTextOptions = true;
			this.txtTotalFormNC.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.txtTotalFormNC.Properties.DisplayFormat.FormatString = "n2";
			this.txtTotalFormNC.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtTotalFormNC.Properties.ReadOnly = true;
			this.txtTotalFormNC.Size = new System.Drawing.Size(72, 20);
			this.txtTotalFormNC.TabIndex = 14;
			this.toolTip1.SetToolTip(this.txtTotalFormNC, "Total Neto = Total Bruto - NC");
			// 
			// txtTotalForm
			// 
			this.txtTotalForm.Location = new System.Drawing.Point(216, 8);
			this.txtTotalForm.Name = "txtTotalForm";
			// 
			// txtTotalForm.Properties
			// 
			this.txtTotalForm.Properties.Appearance.Options.UseTextOptions = true;
			this.txtTotalForm.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.txtTotalForm.Properties.DisplayFormat.FormatString = "n2";
			this.txtTotalForm.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtTotalForm.Properties.ReadOnly = true;
			this.txtTotalForm.Size = new System.Drawing.Size(72, 20);
			this.txtTotalForm.TabIndex = 13;
			this.toolTip1.SetToolTip(this.txtTotalForm, "Total Bruto");
			// 
			// btTotalForm104
			// 
			this.btTotalForm104.Location = new System.Drawing.Point(128, 8);
			this.btTotalForm104.Name = "btTotalForm104";
			this.btTotalForm104.Size = new System.Drawing.Size(64, 24);
			this.btTotalForm104.TabIndex = 12;
			this.btTotalForm104.Text = "Total";
			this.btTotalForm104.Click += new System.EventHandler(this.btTotalForm104_Click);
			// 
			// btTransForm
			// 
			this.btTransForm.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btTransForm.Location = new System.Drawing.Point(753, 8);
			this.btTransForm.Name = "btTransForm";
			this.btTransForm.Size = new System.Drawing.Size(40, 24);
			this.btTransForm.TabIndex = 11;
			this.btTransForm.Text = "Trans.";
			this.btTransForm.Click += new System.EventHandler(this.btTransForm_Click);
			// 
			// btArtForm
			// 
			this.btArtForm.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btArtForm.Location = new System.Drawing.Point(801, 8);
			this.btArtForm.Name = "btArtForm";
			this.btArtForm.Size = new System.Drawing.Size(40, 24);
			this.btArtForm.TabIndex = 9;
			this.btArtForm.Text = "Art.";
			this.btArtForm.Click += new System.EventHandler(this.btArtForm_Click);
			// 
			// btGenera104
			// 
			this.btGenera104.Location = new System.Drawing.Point(16, 8);
			this.btGenera104.Name = "btGenera104";
			this.btGenera104.Size = new System.Drawing.Size(96, 24);
			this.btGenera104.TabIndex = 6;
			this.btGenera104.Text = "Formulario 104";
			this.btGenera104.Click += new System.EventHandler(this.btGenera104_Click);
			// 
			// grdForm104
			// 
			this.grdForm104.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdForm104.EmbeddedNavigator
			// 
			this.grdForm104.EmbeddedNavigator.Name = "";
			this.grdForm104.Location = new System.Drawing.Point(16, 40);
			this.grdForm104.MainView = this.grvForm104;
			this.grdForm104.Name = "grdForm104";
			this.grdForm104.Size = new System.Drawing.Size(881, 277);
			this.grdForm104.TabIndex = 1;
			this.grdForm104.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															this.grvForm104,
																																															this.gridView9});
			// 
			// grvForm104
			// 
			this.grvForm104.GridControl = this.grdForm104;
			this.grvForm104.Name = "grvForm104";
			this.grvForm104.OptionsSelection.MultiSelect = true;
			this.grvForm104.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvForm104.OptionsView.ColumnAutoWidth = false;
			this.grvForm104.OptionsView.ShowAutoFilterRow = true;
			this.grvForm104.DoubleClick += new System.EventHandler(this.grvForm104_DoubleClick);
			// 
			// gridView9
			// 
			this.gridView9.GridControl = this.grdForm104;
			this.gridView9.Name = "gridView9";
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.btFactUtilidad);
			this.ultraTabPageControl10.Controls.Add(this.btUtilidadGenera);
			this.ultraTabPageControl10.Controls.Add(this.grdUtilidad);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(909, 328);
			// 
			// btFactUtilidad
			// 
			this.btFactUtilidad.Location = new System.Drawing.Point(624, 8);
			this.btFactUtilidad.Name = "btFactUtilidad";
			this.btFactUtilidad.Size = new System.Drawing.Size(72, 23);
			this.btFactUtilidad.TabIndex = 72;
			this.btFactUtilidad.Text = "Factura";
			this.btFactUtilidad.Click += new System.EventHandler(this.btFactUtilidad_Click);
			// 
			// btUtilidadGenera
			// 
			this.btUtilidadGenera.Location = new System.Drawing.Point(8, 8);
			this.btUtilidadGenera.Name = "btUtilidadGenera";
			this.btUtilidadGenera.Size = new System.Drawing.Size(64, 24);
			this.btUtilidadGenera.TabIndex = 7;
			this.btUtilidadGenera.Text = "Utilidad";
			this.btUtilidadGenera.Click += new System.EventHandler(this.btUtilidadGenera_Click);
			// 
			// grdUtilidad
			// 
			this.grdUtilidad.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdUtilidad.EmbeddedNavigator
			// 
			this.grdUtilidad.EmbeddedNavigator.Name = "";
			this.grdUtilidad.Location = new System.Drawing.Point(8, 40);
			this.grdUtilidad.MainView = this.grvUtilidad;
			this.grdUtilidad.Name = "grdUtilidad";
			this.grdUtilidad.Size = new System.Drawing.Size(881, 277);
			this.grdUtilidad.TabIndex = 2;
			this.grdUtilidad.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															 this.grvUtilidad,
																																															 this.gridView11});
			// 
			// grvUtilidad
			// 
			this.grvUtilidad.GridControl = this.grdUtilidad;
			this.grvUtilidad.Name = "grvUtilidad";
			this.grvUtilidad.OptionsSelection.MultiSelect = true;
			this.grvUtilidad.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvUtilidad.OptionsView.ColumnAutoWidth = false;
			this.grvUtilidad.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView11
			// 
			this.gridView11.GridControl = this.grdUtilidad;
			this.gridView11.Name = "gridView11";
			// 
			// ultraTabPageControl11
			// 
			this.ultraTabPageControl11.Controls.Add(this.btFlujo);
			this.ultraTabPageControl11.Controls.Add(this.pivotGridControl1);
			this.ultraTabPageControl11.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl11.Name = "ultraTabPageControl11";
			this.ultraTabPageControl11.Size = new System.Drawing.Size(909, 328);
			// 
			// btFlujo
			// 
			this.btFlujo.Location = new System.Drawing.Point(64, 8);
			this.btFlujo.Name = "btFlujo";
			this.btFlujo.Size = new System.Drawing.Size(72, 23);
			this.btFlujo.TabIndex = 73;
			this.btFlujo.Text = "Flujo de Caja";
			this.btFlujo.Click += new System.EventHandler(this.btFlujo_Click);
			// 
			// pivotGridControl1
			// 
			this.pivotGridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.pivotGridControl1.Cursor = System.Windows.Forms.Cursors.Default;
			this.pivotGridControl1.Fields.AddRange(new DevExpress.XtraPivotGrid.PivotGridField[] {
																																														 this.pivotGridField12,
																																														 this.pivotGridField13,
																																														 this.pivotGridField14,
																																														 this.pivotGridField15,
																																														 this.pivotGridField16,
																																														 this.pivotGridField17,
																																														 this.pivotGridField18,
																																														 this.pivotGridField19,
																																														 this.pivotGridField20,
																																														 this.pivotGridField21,
																																														 this.pivotGridField22});
			this.pivotGridControl1.Location = new System.Drawing.Point(10, 40);
			this.pivotGridControl1.Name = "pivotGridControl1";
			this.pivotGridControl1.OptionsView.ShowColumnTotals = false;
			this.pivotGridControl1.OptionsView.ShowRowGrandTotals = false;
			this.pivotGridControl1.OptionsView.ShowRowTotals = false;
			this.pivotGridControl1.Size = new System.Drawing.Size(889, 277);
			this.pivotGridControl1.TabIndex = 1;
			// 
			// pivotGridField12
			// 
			this.pivotGridField12.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
			this.pivotGridField12.AreaIndex = 0;
			this.pivotGridField12.CellFormat.FormatString = "n2";
			this.pivotGridField12.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField12.EmptyValueText = "0";
			this.pivotGridField12.FieldName = "Saldo";
			this.pivotGridField12.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField12.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField12.Name = "pivotGridField12";
			this.pivotGridField12.TotalsVisibility = DevExpress.XtraPivotGrid.PivotTotalsVisibility.None;
			// 
			// pivotGridField13
			// 
			this.pivotGridField13.Area = DevExpress.XtraPivotGrid.PivotArea.RowArea;
			this.pivotGridField13.AreaIndex = 0;
			this.pivotGridField13.FieldName = "Codigo";
			this.pivotGridField13.Name = "pivotGridField13";
			this.pivotGridField13.Options.AllowExpand = DevExpress.Utils.DefaultBoolean.False;
			// 
			// pivotGridField14
			// 
			this.pivotGridField14.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField14.AreaIndex = 0;
			this.pivotGridField14.FieldName = "Proyecto";
			this.pivotGridField14.Name = "pivotGridField14";
			// 
			// pivotGridField15
			// 
			this.pivotGridField15.Area = DevExpress.XtraPivotGrid.PivotArea.ColumnArea;
			this.pivotGridField15.AreaIndex = 0;
			this.pivotGridField15.Caption = "Centro de Costo";
			this.pivotGridField15.FieldName = "CentroCosto";
			this.pivotGridField15.Name = "pivotGridField15";
			// 
			// pivotGridField16
			// 
			this.pivotGridField16.Area = DevExpress.XtraPivotGrid.PivotArea.RowArea;
			this.pivotGridField16.AreaIndex = 1;
			this.pivotGridField16.Caption = "Descripción";
			this.pivotGridField16.FieldName = "Descripcion";
			this.pivotGridField16.Name = "pivotGridField16";
			this.pivotGridField16.Width = 200;
			// 
			// pivotGridField17
			// 
			this.pivotGridField17.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField17.AreaIndex = 1;
			this.pivotGridField17.CellFormat.FormatString = "n2";
			this.pivotGridField17.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField17.FieldName = "Presupuesto";
			this.pivotGridField17.Name = "pivotGridField17";
			// 
			// pivotGridField18
			// 
			this.pivotGridField18.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField18.AreaIndex = 2;
			this.pivotGridField18.Caption = "Ejecución";
			this.pivotGridField18.CellFormat.FormatString = "n2";
			this.pivotGridField18.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField18.FieldName = "Contable";
			this.pivotGridField18.Name = "pivotGridField18";
			// 
			// pivotGridField19
			// 
			this.pivotGridField19.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField19.AreaIndex = 3;
			this.pivotGridField19.FieldName = "Porc";
			this.pivotGridField19.Name = "pivotGridField19";
			// 
			// pivotGridField20
			// 
			this.pivotGridField20.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField20.AreaIndex = 5;
			this.pivotGridField20.Caption = "Mes";
			this.pivotGridField20.FieldName = "Fecha";
			this.pivotGridField20.GroupInterval = DevExpress.XtraPivotGrid.PivotGroupInterval.DateMonth;
			this.pivotGridField20.Name = "pivotGridField20";
			// 
			// pivotGridField21
			// 
			this.pivotGridField21.Area = DevExpress.XtraPivotGrid.PivotArea.FilterArea;
			this.pivotGridField21.AreaIndex = 4;
			this.pivotGridField21.Caption = "Año";
			this.pivotGridField21.FieldName = "Fecha";
			this.pivotGridField21.GroupInterval = DevExpress.XtraPivotGrid.PivotGroupInterval.DateYear;
			this.pivotGridField21.Name = "pivotGridField21";
			// 
			// pivotGridField22
			// 
			this.pivotGridField22.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
			this.pivotGridField22.AreaIndex = 1;
			this.pivotGridField22.CellFormat.FormatString = "n2";
			this.pivotGridField22.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField22.FieldName = "Periodo";
			this.pivotGridField22.GrandTotalCellFormat.FormatString = "n2";
			this.pivotGridField22.GrandTotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField22.Name = "pivotGridField22";
			this.pivotGridField22.TotalCellFormat.FormatString = "n2";
			this.pivotGridField22.TotalCellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField22.TotalValueFormat.FormatString = "n2";
			this.pivotGridField22.TotalValueFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.pivotGridField22.ValueFormat.FormatString = "n2";
			this.pivotGridField22.ValueFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			// 
			// ultraTabPageControl12
			// 
			this.ultraTabPageControl12.Controls.Add(this.cmbCostoPeriodo);
			this.ultraTabPageControl12.Controls.Add(this.btKardexComp);
			this.ultraTabPageControl12.Controls.Add(this.grdInventario);
			this.ultraTabPageControl12.Controls.Add(this.btInventario);
			this.ultraTabPageControl12.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl12.Name = "ultraTabPageControl12";
			this.ultraTabPageControl12.Size = new System.Drawing.Size(909, 328);
			// 
			// cmbCostoPeriodo
			// 
			this.cmbCostoPeriodo.Location = new System.Drawing.Point(8, 8);
			this.cmbCostoPeriodo.Name = "cmbCostoPeriodo";
			// 
			// cmbCostoPeriodo.Properties
			// 
			this.cmbCostoPeriodo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																						new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCostoPeriodo.Properties.Items.AddRange(new object[] {
																																		"Costo Mensual",
																																		"Costo Diario",
																																		"Costo por Factura"});
			this.cmbCostoPeriodo.Size = new System.Drawing.Size(104, 20);
			this.cmbCostoPeriodo.TabIndex = 78;
			// 
			// btKardexComp
			// 
			this.btKardexComp.Location = new System.Drawing.Point(240, 8);
			this.btKardexComp.Name = "btKardexComp";
			this.btKardexComp.Size = new System.Drawing.Size(72, 23);
			this.btKardexComp.TabIndex = 77;
			this.btKardexComp.Text = "Kardex";
			this.btKardexComp.Click += new System.EventHandler(this.simpleButton1_Click);
			// 
			// grdInventario
			// 
			this.grdInventario.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdInventario.EmbeddedNavigator
			// 
			this.grdInventario.EmbeddedNavigator.Name = "";
			this.grdInventario.Location = new System.Drawing.Point(8, 40);
			this.grdInventario.MainView = this.grvInventario;
			this.grdInventario.Name = "grdInventario";
			this.grdInventario.Size = new System.Drawing.Size(881, 277);
			this.grdInventario.TabIndex = 76;
			this.grdInventario.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																 this.grvInventario,
																																																 this.gridView13});
			// 
			// grvInventario
			// 
			this.grvInventario.GridControl = this.grdInventario;
			this.grvInventario.Name = "grvInventario";
			this.grvInventario.OptionsSelection.MultiSelect = true;
			this.grvInventario.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvInventario.OptionsView.ColumnAutoWidth = false;
			this.grvInventario.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView13
			// 
			this.gridView13.GridControl = this.grdInventario;
			this.gridView13.Name = "gridView13";
			// 
			// btInventario
			// 
			this.btInventario.Location = new System.Drawing.Point(128, 8);
			this.btInventario.Name = "btInventario";
			this.btInventario.Size = new System.Drawing.Size(72, 23);
			this.btInventario.TabIndex = 75;
			this.btInventario.Text = "Inventario";
			this.btInventario.Click += new System.EventHandler(this.btInventario_Click);
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
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
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(22, 16);
			this.label1.TabIndex = 4;
			this.label1.Text = "De:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(168, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(14, 16);
			this.label2.TabIndex = 5;
			this.label2.Text = "a:";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// btAsiento
			// 
			this.btAsiento.Location = new System.Drawing.Point(784, 0);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(48, 24);
			this.btAsiento.TabIndex = 8;
			this.btAsiento.Text = "Asiento";
			this.btAsiento.Visible = false;
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.tabBalance);
			this.ultraTabControl1.Controls.Add(this.tabMayores);
			this.ultraTabControl1.Controls.Add(this.tabFacturas);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl3);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl4);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl5);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl6);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl7);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl8);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl9);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl10);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl11);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl12);
			this.ultraTabControl1.Location = new System.Drawing.Point(16, 43);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(912, 353);
			this.ultraTabControl1.TabIndex = 11;
			ultraTab1.TabPage = this.tabBalance;
			ultraTab1.Text = "Balance";
			ultraTab2.TabPage = this.tabMayores;
			ultraTab2.Text = "Mayores";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Asientos";
			ultraTab4.TabPage = this.tabFacturas;
			ultraTab4.Text = "Facturas";
			ultraTab5.TabPage = this.ultraTabPageControl1;
			ultraTab5.Text = "Pagos";
			ultraTab5.Visible = false;
			ultraTab6.TabPage = this.ultraTabPageControl2;
			ultraTab6.Text = "Retención";
			ultraTab7.TabPage = this.ultraTabPageControl4;
			ultraTab7.Text = "Factura - Grupo";
			ultraTab8.TabPage = this.ultraTabPageControl5;
			ultraTab8.Text = "Crédito";
			ultraTab9.TabPage = this.ultraTabPageControl7;
			ultraTab9.Text = "Articulos";
			ultraTab10.TabPage = this.ultraTabPageControl6;
			ultraTab10.Text = "Cubos";
			ultraTab11.TabPage = this.ultraTabPageControl8;
			ultraTab11.Text = "Antigüedad";
			ultraTab12.TabPage = this.ultraTabPageControl9;
			ultraTab12.Text = "104";
			ultraTab13.TabPage = this.ultraTabPageControl10;
			ultraTab13.Text = "Utilidad";
			ultraTab13.Visible = false;
			ultraTab14.TabPage = this.ultraTabPageControl11;
			ultraTab14.Text = "Flujo Caja";
			ultraTab14.Visible = false;
			ultraTab15.TabPage = this.ultraTabPageControl12;
			ultraTab15.Text = "Inventario";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2,
																																															 ultraTab3,
																																															 ultraTab4,
																																															 ultraTab5,
																																															 ultraTab6,
																																															 ultraTab7,
																																															 ultraTab8,
																																															 ultraTab9,
																																															 ultraTab10,
																																															 ultraTab11,
																																															 ultraTab12,
																																															 ultraTab13,
																																															 ultraTab14,
																																															 ultraTab15});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(909, 328);
			// 
			// btMayor
			// 
			this.btMayor.Location = new System.Drawing.Point(784, 24);
			this.btMayor.Name = "btMayor";
			this.btMayor.Size = new System.Drawing.Size(48, 24);
			this.btMayor.TabIndex = 12;
			this.btMayor.Text = "Mayor";
			this.btMayor.Visible = false;
			this.btMayor.Click += new System.EventHandler(this.btMayor_Click);
			// 
			// cmbCentro
			// 
			this.cmbCentro.Location = new System.Drawing.Point(312, 8);
			this.cmbCentro.Name = "cmbCentro";
			// 
			// cmbCentro.Properties
			// 
			this.cmbCentro.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCentro.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.cmbCentro.Properties.DisplayMember = "Nombre";
			this.cmbCentro.Properties.NullText = "";
			this.cmbCentro.Properties.PopupWidth = 300;
			this.cmbCentro.Properties.ShowFooter = false;
			this.cmbCentro.Properties.ShowHeader = false;
			this.cmbCentro.Properties.ValueMember = "idProyecto";
			this.cmbCentro.Size = new System.Drawing.Size(176, 20);
			this.cmbCentro.TabIndex = 61;
			this.cmbCentro.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cmbCentro_MouseDown);
			this.cmbCentro.EditValueChanged += new System.EventHandler(this.cmbCentro_EditValueChanged);
			// 
			// cmbSubCentro
			// 
			this.cmbSubCentro.Location = new System.Drawing.Point(496, 8);
			this.cmbSubCentro.Name = "cmbSubCentro";
			// 
			// cmbSubCentro.Properties
			// 
			this.cmbSubCentro.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbSubCentro.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.cmbSubCentro.Properties.DisplayMember = "Nombre";
			this.cmbSubCentro.Properties.NullText = "";
			this.cmbSubCentro.Properties.PopupWidth = 300;
			this.cmbSubCentro.Properties.ShowFooter = false;
			this.cmbSubCentro.Properties.ShowHeader = false;
			this.cmbSubCentro.Properties.ValueMember = "idSubProyecto";
			this.cmbSubCentro.Size = new System.Drawing.Size(176, 20);
			this.cmbSubCentro.TabIndex = 63;
			this.cmbSubCentro.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cmbSubCentro_MouseDown);
			this.cmbSubCentro.EditValueChanged += new System.EventHandler(this.cmbSubCentro_EditValueChanged);
			// 
			// cmbDesde
			// 
			this.cmbDesde.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.FechaRIni"));
			this.cmbDesde.EditValue = null;
			this.cmbDesde.Location = new System.Drawing.Point(56, 8);
			this.cmbDesde.Name = "cmbDesde";
			// 
			// cmbDesde.Properties
			// 
			this.cmbDesde.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbDesde.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.cmbDesde.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbDesde.Properties.EditFormat.FormatString = "dd/MMM/yyyy";
			this.cmbDesde.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbDesde.Properties.Mask.EditMask = "dd/MMM/yyyy";
			this.cmbDesde.Properties.NullDate = "";
			this.cmbDesde.Size = new System.Drawing.Size(104, 20);
			this.cmbDesde.TabIndex = 64;
			this.cmbDesde.EditValueChanged += new System.EventHandler(this.cmbDesde_EditValueChanged);
			// 
			// cmbHasta
			// 
			this.cmbHasta.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.FechaRFin"));
			this.cmbHasta.EditValue = null;
			this.cmbHasta.Location = new System.Drawing.Point(184, 8);
			this.cmbHasta.Name = "cmbHasta";
			// 
			// cmbHasta.Properties
			// 
			this.cmbHasta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbHasta.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbHasta.Properties.EditFormat.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbHasta.Properties.Mask.EditMask = "dd/MMM/yyyy";
			this.cmbHasta.Properties.NullDate = "";
			this.cmbHasta.Size = new System.Drawing.Size(104, 20);
			this.cmbHasta.TabIndex = 65;
			this.cmbHasta.EditValueChanged += new System.EventHandler(this.cmbHasta_EditValueChanged);
			// 
			// btBalance
			// 
			this.btBalance.Location = new System.Drawing.Point(712, 8);
			this.btBalance.Name = "btBalance";
			this.btBalance.Size = new System.Drawing.Size(56, 23);
			this.btBalance.TabIndex = 66;
			this.btBalance.Text = "Balance";
			this.btBalance.Click += new System.EventHandler(this.btFiltro_Click);
			this.btBalance.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btBalance_MouseUp);
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.DefaultExt = "Xml";
			this.openFileDialog1.Filter = "Archivo de Diseño | *.Xml";
			// 
			// saveFileDialog1
			// 
			this.saveFileDialog1.DefaultExt = "Xml";
			this.saveFileDialog1.Filter = "Archivo de Diseño | *.Xml";
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
			// Comparacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(936, 406);
			this.Controls.Add(this.btBalance);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.cmbSubCentro);
			this.Controls.Add(this.cmbCentro);
			this.Controls.Add(this.btMayor);
			this.Controls.Add(this.ultraTabControl1);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Name = "Comparacion";
			this.Text = "Comparación";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Comparacion_Closing);
			this.Load += new System.EventHandler(this.Comparacion_Load);
			this.tabBalance.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.chkCuentas.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdBalances)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBalances)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkPeriodo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBalanceTipo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkMovimiento.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkDesglosar.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			this.tabMayores.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdMayores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvMayores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdAsientos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvAsientos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAsientoAgrupar.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAsientoAnulado.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoAsiento.Properties)).EndInit();
			this.tabFacturas.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAnuladoFactura.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkLote.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkOtro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoFactura)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCobro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsPago)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtTotalRet.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoRet.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAnuladaRet.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView12)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDevolucion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdGrupoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvGrupoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupoAgrupado.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFact1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuentoGrupo.Properties)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkResumenCompra.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAgrupaCredito.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactCred)).EndInit();
			this.ultraTabPageControl7.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.chkAnuladoArt.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAgrupar.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAjustar.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridLookUpEdit1View)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdArticulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvArticulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			this.ultraTabPageControl6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.chkTotalF.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTotalC.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.pgrBalance)).EndInit();
			this.ultraTabPageControl8.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optVencidos.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optCobrar.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdAntiguedad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvAntiguedad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView6)).EndInit();
			this.ultraTabPageControl9.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtIvaForm.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFormNC.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalForm.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdForm104)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvForm104)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView9)).EndInit();
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdUtilidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvUtilidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView11)).EndInit();
			this.ultraTabPageControl11.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.pivotGridControl1)).EndInit();
			this.ultraTabPageControl12.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbCostoPeriodo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdInventario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvInventario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbCentro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubCentro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private int idProyecto = 0; 
		DateTime dtFechaIni;
		DateTime dtFechaFin;

		private void GeneraBalance(string stFuncion)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;
			int iDesglosar = 0;
			if (chkDesglosar.Checked) iDesglosar = 1;
			int iMovimiento = 0;
			if (chkMovimiento.Checked) iMovimiento = 1;
			if (chkMovimiento.CheckState == System.Windows.Forms.CheckState.Indeterminate) iMovimiento = 2;
			int iTipo = cmbBalanceTipo.SelectedIndex + 1;
			int iSoloMes = 0;
			if (chkPeriodo.Checked) iSoloMes = 1;
			string stCodigo = txtCodigo.Text;
			string stSelect = string.Format("Exec {0}ComparaBalance '{1}', '{2}', {3}, {4}, {5}, {6}, '{7}', {8}, {9}",
				stFuncion, dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"), idProyecto, 
				idSubProyecto, iDesglosar, iMovimiento, stCodigo, iTipo, iSoloMes);
			grvBalances.Columns.Clear();
			grdBalances.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			string stArchivo = Application.StartupPath+"\\CuboBalance2.Xml";
			if (iTipo == 3) stArchivo = Application.StartupPath+"\\CuboBalance.Xml";
			if (File.Exists(stArchivo)) 
				pgrBalance.RestoreLayoutFromXml(stArchivo);
			pgrBalance.DataSource = grdBalances.DataSource;
			FormatoGrillaB("Balance");
			FormatoGrillaB("Balance");
			grvBalances.BestFitColumns();

			if (iDesglosar == 0)
			{
				grvBalances.OptionsView.ShowGroupPanel = false;
				ultraTabControl1.SelectedTab = ultraTabControl1.Tabs[0];
			}
			else
			{
				grvBalances.OptionsView.ShowGroupPanel = true;
				ultraTabControl1.SelectedTab = ultraTabControl1.Tabs[9];
			}

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btBalance.ToolTip= "Tiempo de Balance: " + ts.TotalSeconds.ToString("n2");
			Cursor = Cursors.Default;
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			GeneraBalance("");
		}

		private void cdsAsiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsAsiento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			int iFila = grvMayores.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Asiento", "Información");
				return;
			}
			int idAsiento = (int) grvMayores.GetRowCellValue(iFila, "idAsiento");
			if (idAsiento == 0)
			{
				MessageBox.Show("Fila no contiene Asiento Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}
			Cursor = Cursors.WaitCursor;
			
			Asiento miAsiento = new Asiento(idAsiento);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void Comparacion_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			DateTime dtInicio = DateTime.Today.AddMonths(-1);
			cmbBalanceTipo.SelectedIndex = 2;
			cmbCostoPeriodo.SelectedIndex = 1;
			this.txtIdAsiento.Width = 0;
			string stAudita = "Exec AuditaCrear 53, 6, 'Comparacion'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);

			#region Configuracion de Impresion de Cubos
			pgrBalance.OptionsPrint.PrintColumnHeaders = DevExpress.Utils.DefaultBoolean.False;
			pgrBalance.OptionsPrint.PrintDataHeaders = DevExpress.Utils.DefaultBoolean.False;;
			pgrBalance.OptionsPrint.PrintFilterHeaders = DevExpress.Utils.DefaultBoolean.False;;
			#endregion
			#region Cargar Combos
			cmbSubCentro.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoC,
				"Select Nombre, idSubProyecto From SubProyecto");
			cmbCentro.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoC,
				"Select Nombre, idProyecto From Proyecto");
			cmbTipoAsiento.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoC,
				"Select idTipoAsiento, Tipo from AsientoTipo");
			cmbTipo.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoC, 
				"Select Nombre, idTipoFactura From CompraNumero Where Grupo = 1 And (CobroPago != 0 Or Signo != 0) Order by idTipoFactura");
			#endregion
		}
		private void FormatoGrillaB(string stTipo)
		{
			DevExpress.XtraGrid.Views.Base.ColumnView View = null;
			if (stTipo == "Balance")
			{
				View = grdBalances.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			}
			if (stTipo == "Mayor")
			{
				View = grdMayores.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
				grvMayores.BestFitColumns();
			}
			if (stTipo == "Asiento")
			{
				View = grdAsientos.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			}
			if (stTipo == "Articulo")
			{
				View = grdArticulos.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			}
			if (stTipo == "Inventario")
			{
				View = grdInventario.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			}

			if (stTipo == "GrupoF")
			{
				View = grdGrupoFactura.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			}
			if (stTipo == "Factura")
			{
				View = grdFactura.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			}
			if (stTipo == "Formulario")
			{
				View = grdForm104.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			}
			if (stTipo == "Retencion")
			{
				View = grdRetencion.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			}
			if (stTipo == "Credito")
			{
				View = grdCredito.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			}
			if (stTipo == "Utilidad")
			{
				View = grdUtilidad.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			}

			if (stTipo == "Antiguedad")
			{
				View = grdAntiguedad.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
				foreach (DevExpress.XtraGrid.Columns.GridColumn col in View.Columns)
				{
					if ((col.AbsoluteIndex > 4 && col.AbsoluteIndex < 11) 
						|| col.FieldName.StartsWith("Pago") || col.FieldName.StartsWith("Total"))
					{
						col.DisplayFormat.FormatString = "n2";
						col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
					}
				
				}
			}
			
			foreach (DevExpress.XtraGrid.Columns.GridColumn col in View.Columns)
			{
				col.OptionsColumn.AllowEdit = false;
				if (col.FieldName.StartsWith("id"))
					col.Visible = false;
				if (col.FieldName.StartsWith("Debe") || col.FieldName.StartsWith("Haber")
					|| col.FieldName.StartsWith("Saldo") || col.FieldName.StartsWith("Antes")
					|| col.FieldName.StartsWith("Total") || col.FieldName.StartsWith("Descuento") 
					|| col.FieldName.StartsWith("Neto") || col.FieldName.StartsWith("Iva") 
					|| col.FieldName.StartsWith("Devolucion") || col.FieldName.StartsWith("Utilidad") 
					|| col.FieldName.StartsWith("SubTotal") || col.FieldName.StartsWith("Saldo") 
					|| col.FieldName.StartsWith("Contable") || col.FieldName.StartsWith("Diferencia") 
					|| col.FieldName.StartsWith("Credito") || col.FieldName.StartsWith("Retenido") 
					|| col.FieldName.StartsWith("Vencido") || col.FieldName.StartsWith("Posterior") 
					|| col.FieldName.StartsWith("Precio") || col.FieldName.StartsWith("Ingreso")
					|| col.FieldName.StartsWith("Dif") 
					|| col.FieldName.StartsWith("Costo") || col.FieldName.StartsWith("Periodo") 
					|| col.FieldName.StartsWith("Retenido") || col.FieldName.StartsWith("Base") 
					|| col.FieldName.StartsWith("Egreso") || col.FieldName.StartsWith("Cant"))
				{
					col.DisplayFormat.FormatString = "n2";
					col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
				}
				if (col.FieldName.StartsWith("Fecha") || col.FieldName.StartsWith("Desde")
					|| col.FieldName.StartsWith("Hasta"))
				{
					col.DisplayFormat.FormatString = "dd/MMM/yyyy";
					col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
				}
			}
		}
		private void btMayor_Click(object sender, System.EventArgs e)
		{
			int iFila = grvBalances.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la Cuenta de la Ficha Balance");
				return;
			}

			Cursor = Cursors.WaitCursor;
			int IdCuenta = (int) grvBalances.GetRowCellValue(iFila, "idCuenta");
			int IdProy = (int) grvBalances.GetRowCellValue(iFila, "idProyecto");
			int IdSubProy = (int) grvBalances.GetRowCellValue(iFila, "idSubProyecto");
			int Desglosar = 0;
			if (chkDesglosar.Checked) Desglosar = 1;
			string stCodigo = "Select Codigo From Cuenta Where idCuenta = " + IdCuenta.ToString();
			stCodigo = Funcion.sEscalarSQL(cdsAsiento, stCodigo);
			string stExec = string.Format("Exec ContabilidadMayor '{0:yyyyMMdd}', '{1:yyyyMMdd}', '{2}', {3}, {4}, {5}",
				dtFechaIni, dtFechaFin, stCodigo, IdProy, IdSubProy, Desglosar);
			grvMayores.Columns.Clear();
			grdMayores.DataSource = Funcion.dvProcedimiento(cdsSeteoC, stExec);
			FormatoGrillaB("Mayor");
			FormatoGrillaB("Mayor");
			grvMayores.BestFitColumns();
			ultraTabControl1.SelectedTab = ultraTabControl1.Tabs[1];
			Cursor = Cursors.Default;
		}

		private void btGeneraFact_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;
			int idTipoFactura = 0;
			if (cmbTipoFactura.Text.Trim().Length > 0)
				idTipoFactura = (int) cmbTipoFactura.Value;
			int iLote = 0;
			if (chkLote.Checked) iLote = 1;
			int iOtro = 0;
			if (chkOtro.Checked) iOtro = 1;
			int iAnulado = 0;
			if (chkAnuladoFactura.Checked) iAnulado = 1;
			string stSelect = string.Format("Exec ComparaFacturaAsiento '{0}', '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"),
				idTipoFactura, idProyecto, idSubProyecto, iOtro, iLote, iAnulado);

			grvFactura.Columns.Clear();
			grdFactura.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("Factura");
			FormatoGrillaB("Factura");
			grvFactura.BestFitColumns();
			
			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
//			btGeneraFact.ToolTip = "Tiempo de Facturas: " + ts.TotalSeconds.ToString("n2");
			Cursor = Cursors.Default;
		}

		private void cdsTipoFactura_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTipoFactura.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			int iFila = grvFactura.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			int IdCompra = (int) grvFactura.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				return;
			}
			if (cmbTipoFactura.Text.Trim().Length == 0)
			{
				MessageBox.Show("Seleccione Tipo de Factura", "Información");
				return;
			}

			Cursor = Cursors.WaitCursor;
//			Compra miCompra = new Compra((int) cmbTipoFactura.Value, IdCompra);
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void btAsiF_Click(object sender, System.EventArgs e)
		{
			int iFila = grvFactura.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Asiento", "Información");
				return;
			}
		
			int idAsiento = (int) grvFactura.GetRowCellValue(iFila, "idAsiento");
			if (idAsiento == 0)
			{
				MessageBox.Show("Fila no contiene Asiento Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}

			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(idAsiento);
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			int iFila = grvFactura.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Asiento", "Información");
				return;
			}
		
			int IdAsiento = (int) grvFactura.GetRowCellValue(iFila, "idAsiAR");
			if (IdAsiento == 0)
			{
				MessageBox.Show("No existe un Asiento de Retención", "Información");
				return;
			}

			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(IdAsiento);
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			int iFila = grvFactura.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Asiento", "Información");
				return;
			}
		
			int IdAsiento = (int) grvFactura.GetRowCellValue(iFila, "idAsiAC");
			if (IdAsiento == 0)
			{
				MessageBox.Show("No existe un Asiento de Comprobante", "Información");
				return;
			}

			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(IdAsiento);
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void btRetencion_Click(object sender, System.EventArgs e)
		{
			int iFila = grvFactura.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Retención", "Información");
				return;
			}
		
			int IdCompra = (int) grvFactura.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				return;
			}

			string stCodRet = (string) grvFactura.GetRowCellValue(iFila, "CodReten");
			if (stCodRet.Trim().Length == 0)
			{
				MessageBox.Show("No existe una Retención", "Información");
				return;
			}

			if ((int) cmbTipoFactura.Value == 1)
			{
				Cursor = Cursors.WaitCursor;
				IngresoAT miEgreso = new IngresoAT();
				miEgreso.IdFactura = IdCompra;
				miEgreso.Show();
			}
			else
			{
				Cursor = Cursors.WaitCursor;
//				EgresoAT miEgreso = new EgresoAT();
//				miEgreso.IdFactura = IdCompra;
//				miEgreso.Show();
			}
			Cursor = Cursors.Default;
		}

		private void ugBalance_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
		}

		private void ugMayores_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			btAsiento_Click(this, e);		
		}

		private void Comparacion_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Comparacion'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
		}

		private void btGeneraPago_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			string stSelect = "Select IsNull(c.Numero, ''), "
				+ "RTrim(IsNUll(AP.CodAsiento, '')), "
				+ "Rtrim(IsNull(AL.CodAsiento, '')), "
				+ "RTrim(IsNull(p.Cheque, '')), "
				+ "IsNull(p.idAsiento, 0), "
				+ "IsNull(pL.idAsiento, 0), "
				+ "IsNull(p.idPago, 0), "
				+ "IsNull(c.idCompra, 0), "
				+ "IsNull(p.Fecha, GetDate()), "
				+ "IsNull(p.NumeroIngreso, 0), "
				+ "IsNull(p.Referencia, 0), "
				+ "Rtrim(IsNull(p.Numero, '')), "
				+ "RTrim(IsNull(PL.Codigo, '')), IsNull(p.Pago, 0) from Pago p "
				+ "Inner Join Compra c On c.idCompra = p.idCompra "
				+ "Left Join Asiento AP On AP.idAsiento = P.idAsiento "
				+ "Left Join PagoLote PL On PL.idPagoLote = P.idPagoLote "
				+ "Left Join Asiento AL On AL.idAsiento = PL.idAsiento "
//				+ "Left Join Retencion On Retencion.idCompra = Compra.idCompra "
				+ "Where p.Fecha >= '" + dtFechaIni.ToString("yyyyMMdd")
				+ "' And p.Fecha < '" + dtFechaFin.AddDays(1).ToString("yyyyMMdd")
				+ "' ";
			if (cmbTipoCobro.Text.Trim().Length > 0)
				stSelect += " And c.idTipoFactura = " + cmbTipoCobro.Value.ToString();

			SqlDataReader miReader = Funcion.rEscalarSQL(cdsTipoFactura, stSelect, true);

			udsPago.Rows.Clear();
			int iNumFilas = 0;
			try
			{
				while (miReader.Read())
				{
					string stNumero = miReader.GetString(0);
					string stCodAsiP = miReader.GetString(1);
					string stCodAsiL = miReader.GetString(2);
					string stCheque = miReader.GetString(3);
					int IdAsiento = miReader.GetInt32(4);
					int IdAsiLote = miReader.GetInt32(5);
					int IdPago = miReader.GetInt32(6);
					int IdCompra = miReader.GetInt32(7);
					DateTime dtFecha = miReader.GetDateTime(8);
					string stNumIngreso = miReader.GetString(9);
					string stReferencia = miReader.GetString(10);
					string stNumeroP = miReader.GetString(11);
					string stCodigoLote = miReader.GetString(12);
					double dPago = miReader.GetDouble(13);

					udsPago.Rows.Add();
					udsPago.Rows[iNumFilas]["Numero"]=stNumero;
					if (stCodAsiL.Length > 0)
					{
						if (stCodAsiP == stCodAsiL)
							stCodAsiP = stCodAsiL;
						else
              stCodAsiP = "Er: " + stCodAsiL + " | " + stCodAsiP;
					}
					udsPago.Rows[iNumFilas]["CodAsiento"]=stCodAsiP;
					udsPago.Rows[iNumFilas]["CodAsiLote"]=stCodAsiL;
					udsPago.Rows[iNumFilas]["Cheque"]=stCheque;
					udsPago.Rows[iNumFilas]["idAsiento"]=IdAsiento;
					udsPago.Rows[iNumFilas]["idAsiLote"]=IdAsiLote;
					udsPago.Rows[iNumFilas]["idPago"]=IdPago;
					udsPago.Rows[iNumFilas]["idCompra"]=IdCompra;
					udsPago.Rows[iNumFilas]["Fecha"]=dtFecha;
					udsPago.Rows[iNumFilas]["NumeroIngreso"]=stNumIngreso;
					udsPago.Rows[iNumFilas]["Referencia"]=stReferencia;

					if (stCodigoLote.Length > 0)
					{
						if (stCodigoLote == stNumeroP)
							stNumeroP = stCodigoLote;
						else
							stNumeroP = "Er:" + stCodigoLote + "|" + stNumeroP;
					}
					udsPago.Rows[iNumFilas]["NumeroP"]=stNumeroP;
					udsPago.Rows[iNumFilas]["CodigoLote"]=stCodigoLote;
					udsPago.Rows[iNumFilas]["Pago"]=dPago;
					iNumFilas ++;
				}
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error");
			}
			miReader.Close();
//			btGeneraPago.ToolTip = "Tiempo de Pagos: " + ts.TotalSeconds.ToString("n2");
		}

		private void btFacturaPago_Click(object sender, System.EventArgs e)
		{
			if (grdPago.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Fila de la cual desea ver la Factura", "Información");
				return;
			}
			int IdCompra = (int) grdPago.ActiveRow.Cells["idCompra"].Value;
			if (IdCompra == 0)
			{
				MessageBox.Show("No existe una Factura", "Información");
				return;
			}
			if (cmbTipoCobro.Text.Trim().Length == 0)
			{
				MessageBox.Show("Seleccione Tipo de Factura", "Información");
				return;
			}

			Cursor = Cursors.WaitCursor;
//			Compra miCompra = new Compra((int) cmbTipoCobro.Value, IdCompra);
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void btPago_Click(object sender, System.EventArgs e)
		{
			if (grdPago.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Fila de la cual desea ver la Factura", "Información");
				return;
			}
			int IdCompra = (int) grdPago.ActiveRow.Cells["idCompra"].Value;
			if (IdCompra == 0)
			{
				MessageBox.Show("No existe una Factura", "Información");
				return;
			}
			if (cmbTipoCobro.Text.Trim().Length == 0)
			{
				MessageBox.Show("Seleccione Tipo de Factura", "Información");
				return;
			}

			Cursor = Cursors.WaitCursor;
//			Pagos miPago = new Pagos((int) cmbTipoCobro.Value, IdCompra);
//			miPago.MdiParent = this.MdiParent;
//			miPago.Show();
			Cursor = Cursors.Default;
		}

		private void btAsiPago_Click(object sender, System.EventArgs e)
		{
			if (grdPago.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Fila de la cual desea ver el Asiento", "Información");
				return;
			}
			int IdAsiento = (int) grdPago.ActiveRow.Cells["idAsiento"].Value;
			int IdAsiLote = (int) grdPago.ActiveRow.Cells["idAsiLote"].Value;
			if (IdAsiLote>0) IdAsiento = IdAsiLote;

			if (IdAsiento == 0)
			{
				MessageBox.Show("No existe una Asiento de esta Factura", "Información");
				return;
			}

			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(IdAsiento);
			miAsiento.Show();
			Cursor = Cursors.Default;
		}
		DataSet miDataSet = new DataSet();
		DataView dv = new DataView();

		private void btExcelFactura_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\FacturaAsiento.xls";
			grdFactura.ExportToXls(stDir, true);
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir);		
		}

		private void btExcelPago_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\Pagosiento.xls"; 
			ultraGridExcelExporter1.Export(grdPago, stDir);
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir);		
		}

		private void btExcelRet_Click(object sender, System.EventArgs e)
		{
			
		}

		private void cmbCentro_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbCentro.EditValue = 0;
		}

		private void cmbSubCentro_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbSubCentro.EditValue = 0;
		}

		private void chkDesglosar_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btGenAsiento_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			int idTipoFactura = 0;
			if (cmbTipoFactura.Text.Trim().Length > 0)
				idTipoFactura = (int) cmbTipoFactura.Value;
			int idTipoAsiento = 0;
			if (cmbTipoAsiento.Text.Length > 0) idTipoAsiento = (int) cmbTipoAsiento.EditValue;
			int IncluirAnulado = 0;
			if (chkAsientoAnulado.Checked) IncluirAnulado = 1;
			int Agrupar = 0;
			if (chkAsientoAgrupar.Checked) Agrupar = 1;
			string stSelect = string.Format("Exec ComparaAsiento '{0}', '{1}', {2}, {3}, {4}, {5}, {6}",
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"),
				idTipoAsiento, idProyecto, idSubProyecto, IncluirAnulado, Agrupar);
			grvAsientos.Columns.Clear();
      grdAsientos.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("Asiento");
			FormatoGrillaB("Asiento");
			grvAsientos.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btGenAsiento.ToolTip = "Tiempo de Asientos: " + ts.TotalSeconds.ToString("n2");
		}

		private void btAsientoA_Click(object sender, System.EventArgs e)
		{
			int iFila = grvAsientos.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione Fila de la cual desea ver el Asiento", "Información");
				return;
			}

			Cursor = Cursors.WaitCursor;
			int IdAsiento = (int) grvAsientos.GetRowCellValue(iFila, "idAsiento");
			if (IdAsiento == 0)
			{
				MessageBox.Show("No existe una Asiento de esta Factura", "Información");
				return;
			}

			Cursor = Cursors.Default;
			Asiento miAsiento = new Asiento(IdAsiento);
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void btOrigen_Click(object sender, System.EventArgs e)
		{
			int iFila = grvAsientos.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione Fila de la cual desea ver el Asiento", "Información");
				return;
			}

			Cursor = Cursors.WaitCursor;
			//Busca en Factura
			string stSelect = "";

			// Busca en factura directa
			int idAsi = (int) grvAsientos.GetRowCellValue(iFila, "idAsiento");
			stSelect = "Select Count(*) From Compra Where idAsiento = " + idAsi.ToString();
			if (Funcion.iEscalarSQL(cdsCuenta, stSelect, true) > 0)
			{
				stSelect = "Select idCompra From Compra Where idAsiento = " + idAsi.ToString();
				int idCompra = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				stSelect = "Select idTipoFactura From Compra Where idAsiento = " + idAsi.ToString();
				int IdTipoF = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
//				Compra miCompra = new Compra(IdTipoF, idCompra);
//				miCompra.MdiParent = MdiParent;
//				miCompra.Show();
				Cursor = Cursors.WaitCursor;
				return;
			}
			// Busca en factura generada por lotes
			stSelect = "Select Count(*) From Compra Where idAsientoPer = " + idAsi.ToString();
			if (Funcion.iEscalarSQL(cdsCuenta, stSelect, true) > 0)
			{
				stSelect = "Select top 1 idCompra From Compra Where idAsientoPer = " + idAsi.ToString();
				int idCompra = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				stSelect = "Select idTipoFactura From Compra Where idCompra = " + idCompra.ToString();
				int IdTipoF = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				MessageBox.Show("Asiento de Factura por lotes generada", "Información");
//				Compra miCompra = new Compra(IdTipoF, idCompra);
//				miCompra.MdiParent = MdiParent;
//				miCompra.Show();
				Cursor = Cursors.WaitCursor;
				return;
			}

			// Busca Retencion de factura
			stSelect = "Select Count(*) From Compra Where idAsiRet = " + idAsi.ToString();
			if (Funcion.iEscalarSQL(cdsCuenta, stSelect, true) > 0)
			{
				stSelect = "Select top 1 idCompra From Compra Where idAsiRet = " + idAsi.ToString();
				int idCompra = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				stSelect = "Select idTipoFactura From Compra Where idCompra = " + idCompra.ToString();
				int IdTipoF = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				MessageBox.Show("Asiento de Retencion", "Información");
//				Compra miCompra = new Compra(IdTipoF, idCompra);
//				miCompra.MdiParent = MdiParent;
//				miCompra.Show();
				Cursor = Cursors.WaitCursor;
				return;
			}

			// Busca Pago por Lotes
			stSelect = "Select Count(*) From PagoLote Where idAsiento = " + idAsi.ToString();
			if (Funcion.iEscalarSQL(cdsCuenta, stSelect, true) > 0)
			{
				stSelect = "Select top 1 IsNull(idCompra, 0) From Pago Where idAsiento = " + idAsi.ToString();
				int idCompra = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				stSelect = "Select top 1 idPagoLote From PagoLote Where idAsiento = " + idAsi.ToString();
				int idPagoLote = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				stSelect = "Select top 1 IsNull(idDeposito, 0) From PagoLote Where idAsiento = " + idAsi.ToString();
				int idDeposito = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);

				stSelect = "Select idTipoFactura From Compra Where idCompra = " + idCompra.ToString();
				int IdTipoF = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				if (idDeposito > 0)
				{
					MessageBox.Show("Asiento de Deposito de Pago por lotes", "Información");
					PagoLoteDia miPago = new PagoLoteDia(IdTipoF, idPagoLote);
					miPago.MdiParent = MdiParent;
					miPago.Show();
				}
				else
				{
					MessageBox.Show("Asiento de Pago por lotes", "Información");
					PagoLotes miPago = new PagoLotes(IdTipoF, idPagoLote);
					miPago.MdiParent = MdiParent;
					miPago.Show();
				}
				Cursor = Cursors.WaitCursor;
				return;
			}	
			
			// Busca Pago individual
			stSelect = "Select Count(*) From Pago Where idAsiento = " + idAsi.ToString();
			if (Funcion.iEscalarSQL(cdsCuenta, stSelect, true) > 0)
			{
				stSelect = "Select top 1 idCompra From Pago Where idAsiento = " + idAsi.ToString();
				int idCompra = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				stSelect = "Select idTipoFactura From Compra Where idCompra = " + idCompra.ToString();
				int IdTipoF = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				MessageBox.Show("Asiento de Pago", "Información");
//				Pagos miPago = new Pagos(IdTipoF, idCompra);
//				miPago.MdiParent = MdiParent;
//				miPago.Show();
				Cursor = Cursors.WaitCursor;
				return;
			}
			Cursor = Cursors.Default;
			MessageBox.Show("Asiento No Existe en Modulos", "Información");
		}

		private void cmbTipoAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbTipoAsiento.EditValue = 0;
		}

		private void btGeneraFacturaGrupo_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			int idTipoFactura = 0;
			if (cmbTipoFact1.Text.Trim().Length > 0)
				idTipoFactura = (int) cmbTipoFact1.Value;
//			int idProyecto = 0;
//			if (cmbCentro.Text.Length > 0) idProyecto = (int) cmbCentro.EditValue;
//			int idSubProyecto = 0;
//			if (cmbSubCentro.Text.Length > 0) idProyecto = (int) cmbSubCentro.EditValue;
			int iLote = 0;
			if (chkGrupoAgrupado.Checked) iLote = 1;
			if (chkGrupoAgrupado.CheckState == System.Windows.Forms.CheckState.Indeterminate) iLote = -1;
			string stSelect = string.Format("Exec ComparaFacturaGrupo '{0}', '{1}', {2}, {3}, {4}, {5}",
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"),
				idTipoFactura, idProyecto, idSubProyecto, iLote);

			grvGrupoFactura.Columns.Clear();
			grdGrupoFactura.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("GrupoF");
			FormatoGrillaB("GrupoF");
			grvGrupoFactura.BestFitColumns();

			btTotalFG_Click(this, e);

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btGeneraFacturaGrupo.ToolTip = "Tiempo de Facturas por Grupo: " + ts.TotalSeconds.ToString("n2");
		}

		private void btPagoCred_Click(object sender, System.EventArgs e)
		{
			int iFila = grvCredito.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Pago", "Información");
				return;
			}
		
			int IdCompra = (int) grvCredito.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("No existe una Factura", "Información");
				Cursor = Cursors.Default;
				return;
			}

			if (cmbTipoFactCred.Text.Trim().Length == 0)
			{
				MessageBox.Show("Seleccione Tipo de Factura", "Información");
				return;
			}
//			Pagos miPago = new Pagos((int) cmbTipoFactCred.Value, IdCompra);
//			miPago.MdiParent = MdiParent;
//			miPago.Show();
		}

		private void btKardex_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			int iAgrupado = 0;
			if (chkAgrupar.Checked) iAgrupado = 1;
			int idTipoFactura = 0;
			if (cmbTipo.Text.Length > 0) idTipoFactura = (int) cmbTipo.EditValue;

			string stExec = string.Format("Exec ArticuloKardexAsiento 0, '{0}', '{1}', {2}, {3}, {4}, {5}", 
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"), iAgrupado, idTipoFactura,
				idProyecto, idSubProyecto);
			grvArticulos.Columns.Clear();
			grdArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteoC, stExec);
			FormatoGrillaB("Articulo");
			FormatoGrillaB("Articulo");
			grvArticulos.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btKardex.ToolTip= "Tiempo de Reporte de Kardex: " + ts.TotalSeconds.ToString("n2");
		}

		private void btArticuloB_Click(object sender, System.EventArgs e)
		{
			int iFila = grvArticulos.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Item", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			int idArticulo = (int) grvArticulos.GetRowCellValue(iFila, "idArticulo");
			if (idArticulo == 0)
			{
				MessageBox.Show("Fila no contiene Articulo Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}
//			Articulo miArt = new Articulo(idArticulo);
//			miArt.MdiParent = MdiParent;
//			miArt.Show();
			Cursor = Cursors.Default;

		}

		private void btTransaccion_Click(object sender, System.EventArgs e)
		{
			int iFila = grvArticulos.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) grvArticulos.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}
//			Compra miCompra = new Compra(-IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void cmbCentro_EditValueChanged(object sender, System.EventArgs e)
		{
			idProyecto = 0;
			if (cmbCentro.Text.Length > 0) idProyecto = (int) cmbCentro.EditValue;
		}

		private	int idSubProyecto = 0;
		private void cmbSubCentro_EditValueChanged(object sender, System.EventArgs e)
		{
			idSubProyecto = 0;
			if (cmbSubCentro.Text.Length > 0) idSubProyecto = (int) cmbSubCentro.EditValue;
		}

		private void cmbDesde_EditValueChanged(object sender, System.EventArgs e)
		{
			dtFechaIni = (DateTime) cmbDesde.EditValue;
		}

		private void cmbHasta_EditValueChanged(object sender, System.EventArgs e)
		{
			dtFechaFin = (DateTime) cmbHasta.EditValue;
		}

		private void chkAjustar_CheckedChanged(object sender, System.EventArgs e)
		{
			grvArticulos.OptionsView.ColumnAutoWidth = chkAjustar.Checked;
		}

		private void cmbTipo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			cmbTipo.EditValue = 0;
		}

		private void btAsiArt_Click(object sender, System.EventArgs e)
		{
			int iFila = grvArticulos.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Asiento", "Información");
				return;
			}
		
			int idAsiento = (int) grvArticulos.GetRowCellValue(iFila, "idAsiento");
			if (idAsiento == 0)
			{
				MessageBox.Show("Fila no contiene Asiento Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}
			Cursor = Cursors.WaitCursor;
			Asiento miArt = new Asiento(idAsiento);
			miArt.MdiParent = MdiParent;
			miArt.Show();
			Cursor = Cursors.Default;
		}

		private void pgrBalance_CellDoubleClick(object sender, DevExpress.XtraPivotGrid.PivotCellEventArgs e)
		{
			grdBalances.DataSource = e.CreateDrillDownDataSource();
			ultraTabControl1.SelectedTab = ultraTabControl1.Tabs[0];
		}

		private void grdBalances_DoubleClick(object sender, System.EventArgs e)
		{
			btMayor_Click(this, e);
		}

		private void grvMayores_DoubleClick(object sender, System.EventArgs e)
		{
			btAsiento_Click(this, e);
		}

		private void grvAsientos_DoubleClick(object sender, System.EventArgs e)
		{
			btAsientoA_Click(this, e);
		}

		private void grvArticulos_DoubleClick(object sender, System.EventArgs e)
		{
			if (grvArticulos.FocusedColumn.Caption == "Articulo" || grvArticulos.FocusedColumn.Caption == "CodArt")
				btArticuloB_Click(this, e);
			else if (grvArticulos.FocusedColumn.Caption == "Precio" || grvArticulos.FocusedColumn.Caption == "Numero"
				|| grvArticulos.FocusedColumn.Caption.StartsWith("Cant"))
				btTransaccion_Click(this, e);
			else
				btAsiArt_Click(this, e);
		}

		private void btBalance_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;

			GeneraBalance("fun_");
		}

		string stTitulo = "Titulo Generico";
		private void printableComponentLink1_CreateReportHeaderArea(object sender, DevExpress.XtraPrinting.CreateAreaEventArgs e)
		{
			string reportHeader = stTitulo;
			e.Graph.StringFormat = new DevExpress.XtraPrinting.BrickStringFormat(StringAlignment.Near);
			e.Graph.Font = new Font("Tahoma", 14, FontStyle.Bold);
			RectangleF rec = new RectangleF(0, 0, e.Graph.ClientPageSize.Width, e.Graph.Font.Height * 4 + e.Graph.Font.Height/4);
			e.Graph.DrawString(reportHeader, rec);
		}

		private void btPresupuesto_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;
			string stFiltro = txtCodigo.Text.Trim();
			bool blCuentas = chkCuentas.Checked;
			string stActivado = "Select IsNull(Activado, 0) From SeteoG Where Codigo = 'FCCCodigo'";
			if (Funcion.bEjecutaSQL(cdsSeteoF, stActivado))
			{
				string stExec = string.Format("Exec ComparaPresupuesto '{0}', '{1}', {2}, {3}, '{4}', {5}",
					dtFechaIni.ToString("yyyyMMdd"),  dtFechaFin.ToString("yyyyMMdd"), idProyecto, idSubProyecto, stFiltro, blCuentas);
				grvBalances.Columns.Clear();
				grdBalances.DataSource = Funcion.dvProcedimiento(cdsAsiento, stExec);
				pgrBalance.DataSource = grdBalances.DataSource;
				FormatoGrillaB("Balance");
				FormatoGrillaB("Balance");
				TimeSpan ts = DateTime.Now.Subtract(dtInicio);
				btPresupuesto.ToolTip = "Tiempo de Presupuesto: " + ts.TotalSeconds.ToString("n2");
			}
			else
			{
				string stExec = string.Format("Exec ComparaPresupuesto '{0}', '{1}', {2}, {3}, '{4}'",
					dtFechaIni.ToString("yyyyMMdd"),  dtFechaFin.ToString("yyyyMMdd"), idProyecto, idSubProyecto, stFiltro);
				grvBalances.Columns.Clear();
				grdBalances.DataSource = Funcion.dvProcedimiento(cdsAsiento, stExec);
				string stArchivo = Application.StartupPath+"\\CuboPresupuesto.Xml";
				if (File.Exists(stArchivo)) 
					pgrBalance.RestoreLayoutFromXml(stArchivo);
				pgrBalance.DataSource = grdBalances.DataSource;
				FormatoGrillaB("Balance");
				FormatoGrillaB("Balance");
				grvBalances.BestFitColumns();
			
			{
				grvBalances.OptionsView.ShowGroupPanel = true;
				ultraTabControl1.SelectedTab = ultraTabControl1.Tabs[9];
			}

				TimeSpan ts = DateTime.Now.Subtract(dtInicio);
				btPresupuesto.ToolTip = "Tiempo de Presupuesto: " + ts.TotalSeconds.ToString("n2");
				Cursor = Cursors.Default;
			}
		}
		private void btPresupuesto_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			pgrBalance.ExportToXls("Presupuesto.Xls");
		}

		private void chkTotalC_CheckedChanged(object sender, System.EventArgs e)
		{
			pgrBalance.OptionsView.ShowColumnGrandTotals = chkTotalC.Checked;
			pgrBalance.OptionsView.ShowColumnTotals = chkTotalC.Checked;
		}

		private void chkTotalF_CheckedChanged(object sender, System.EventArgs e)
		{
			pgrBalance.OptionsView.ShowRowTotals = chkTotalF.Checked;
			pgrBalance.OptionsView.ShowRowGrandTotals = chkTotalF.Checked;
		}

		private void btAntigueada_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			DateTime dtFecha = (DateTime)cmbHasta.EditValue;
			int optTipo = (int)optCobrar.EditValue;
			int optVencido = (int)optVencidos.EditValue;
			string stSelect = String.Format("EXEC CxC30Incobrable '{0:yyyyMMdd}', {1}, {2}, {3}, {4}",
				dtFecha, optTipo, optVencido, idProyecto, idSubProyecto);
			grvAntiguedad.Columns.Clear();
			grdAntiguedad.DataSource = Funcion.dvProcedimiento(cdsSeteoC, stSelect);
			FormatoGrillaB("Antiguedad");
			FormatoGrillaB("Antiguedad");
			grvAntiguedad.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btAntiguedad.ToolTip= "Tiempo de Reporte de Antigüedad: " + ts.TotalSeconds.ToString("n2");
		}

		private void btFacturaAntig_Click(object sender, System.EventArgs e)
		{
			int iFila = grvAntiguedad.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) grvAntiguedad.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}
//			Compra miCompra = new Compra(-IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void btErrorAntiguedad_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			int optTipo = (int)optCobrar.EditValue;
			string stSelect = String.Format("EXEC PagoPendienteError {0}",
				optTipo);
			grvAntiguedad.Columns.Clear();
			grdAntiguedad.DataSource = Funcion.dvProcedimiento(cdsSeteoC, stSelect);
			FormatoGrillaB("Antiguedad");
			FormatoGrillaB("Antiguedad");
			grvAntiguedad.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btErrorAntiguedad.ToolTip= "Tiempo de Reporte de Errores de Antigüedad: " + ts.TotalSeconds.ToString("n2");
		}

		private void btPagoAntig_Click(object sender, System.EventArgs e)
		{
			int iFila = grvAntiguedad.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) grvAntiguedad.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}

			Cursor = Cursors.WaitCursor;
//			Pagos miPago = new Pagos((int)optCobrar.EditValue, IdCompra);
//			miPago.MdiParent = this.MdiParent;
//			miPago.Show();
			Cursor = Cursors.Default;
		}

		private void btAsientoGrupo_Click(object sender, System.EventArgs e)
		{
			int iFila = grvGrupoFactura.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Asiento", "Información");
				return;
			}
		
			int idAsiento = (int) grvGrupoFactura.GetRowCellValue(iFila, "idAsiento");
			if (idAsiento == 0)
			{
				MessageBox.Show("Fila no contiene Asiento Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}
			Cursor = Cursors.WaitCursor;
			Asiento miArt = new Asiento(idAsiento);
			miArt.MdiParent = MdiParent;
			miArt.Show();
			Cursor = Cursors.Default;
		}

		private void btFacturaGrupo_Click(object sender, System.EventArgs e)
		{
			int iFila = grvGrupoFactura.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) grvGrupoFactura.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}
//			Compra miCompra = new Compra(-IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void btExcelGrupo_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\FacturaGrupo.xls";
			grdGrupoFactura.ExportToXls(stDir, true);
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir);		
		}

		private void btGenera104_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			string stExec = string.Format("Exec Form104Mostrar1 '{0}', '{1}'", 
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"));
			grvForm104.Columns.Clear();
			grdForm104.DataSource = Funcion.dvProcedimiento(cdsSeteoC, stExec);
			grvForm104.BestFitColumns();
			FormatoGrillaB("Formulario");
			FormatoGrillaB("Formulario");
			grvForm104.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btGenera104.ToolTip= "Tiempo de Formulario 104: " + ts.TotalSeconds.ToString("n2");
		}

		private void btTotalForm104_Click(object sender, System.EventArgs e)
		{
			double dTotalPos = 0;
			double dTotalNC = 0;
			double dImpuesto = 0;
			for (int i=0; i< grvForm104.RowCount; i++)
			{
				double dTotal = (double) grvForm104.GetRowCellValue(i, "Total");
				if (dTotal > 0) dTotalPos += dTotal;
				dTotalNC += dTotal;
				dImpuesto += (double) grvForm104.GetRowCellValue(i, "Iva");
			}
			txtTotalForm.EditValue = dTotalPos;
			txtTotalFormNC.EditValue = dTotalNC;
			txtIvaForm.EditValue = dImpuesto;
		}

		private void grvForm104_DoubleClick(object sender, System.EventArgs e)
		{
			if (grvForm104.FocusedColumn.Caption == "Articulo" || grvForm104.FocusedColumn.Caption == "Codigo")
				btArtForm_Click(this, e);
			else
				btTransForm_Click(this, e);
		}

		private void btArtForm_Click(object sender, System.EventArgs e)
		{
			int iFila = grvForm104.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Item", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			int idArticulo = (int) grvForm104.GetRowCellValue(iFila, "idArticulo");
			if (idArticulo == 0)
			{
				MessageBox.Show("Fila no contiene Articulo Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}
//			Articulo miArt = new Articulo(idArticulo);
//			miArt.MdiParent = MdiParent;
//			miArt.Show();
			Cursor = Cursors.Default;
		}

		private void btTransForm_Click(object sender, System.EventArgs e)
		{
			int iFila = grvForm104.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) grvForm104.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}
//			Compra miCompra = new Compra(-IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void btExcelCubo_Click(object sender, System.EventArgs e)
		{
			string stEmpresa = Funcion.sEscalarSQL(cdsAsiento, "Select Empresa From seteo");
			stTitulo = string.Format("{3}\nEjecución Presupuestaria\ndel {0} al {1}\nExpresado en Dólares Americanos", 
				dtFechaIni.ToString("dd/MMM/yyyy"), 
				dtFechaFin.ToString("dd/MMM/yyyy"), cmbBalanceTipo.Text, stEmpresa);

			DevExpress.XtraPrinting.PrintingSystem ps = new DevExpress.XtraPrinting.PrintingSystem();
			DevExpress.XtraPrinting.PrintableComponentLink link = new DevExpress.XtraPrinting.PrintableComponentLink(ps);
			link.Component = pgrBalance;
			link.CreateReportHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(
				printableComponentLink1_CreateReportHeaderArea);
			link.CreateDocument();
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo en Excel?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				string stDir = MenuLatinium.stDirInicio + "\\Presupuesto.xls";
				ps.ExportToXls(stDir, true);
				System.Diagnostics.Process.Start(stDir);
			}
			else
				link.ShowPreview();
		}

		private void btExcelForm_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\Formulario.xls";
			grdForm104.ExportToXls(stDir, true);
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir);
		}

		private void btExcelAntig_Click(object sender, System.EventArgs e)
		{
			string stEmpresa = Funcion.sEscalarSQL(cdsAsiento, "Select Empresa From seteo");

			stTitulo = string.Format("{3}\nAntigüedad de Cartera\nAl {1}\nExpresado en Dólares Americanos", 
				dtFechaIni.ToString("dd/MMM/yyyy"), 
				dtFechaFin.ToString("dd/MMM/yyyy"), cmbBalanceTipo.Text, stEmpresa);
			DevExpress.XtraPrinting.PrintingSystem ps = new DevExpress.XtraPrinting.PrintingSystem();
			DevExpress.XtraPrinting.PrintableComponentLink link = new DevExpress.XtraPrinting.PrintableComponentLink(ps);
			link.Component = grdAntiguedad;
			link.CreateReportHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(
				printableComponentLink1_CreateReportHeaderArea);
			link.CreateDocument();
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo en Excel?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				string stDir = MenuLatinium.stDirInicio + "\\CXC306090.xls";
				ps.ExportToXls(stDir, true);
				System.Diagnostics.Process.Start(stDir);
			}
			else
				link.ShowPreview();
		}

		private void btExcelBalance_Click(object sender, System.EventArgs e)
		{
			string stEmpresa = Funcion.sEscalarSQL(cdsAsiento, "Select Empresa From seteo");

			stTitulo = string.Format("{3}\nBalance {2}\ndel {0} al {1}\nExpresado en Dólares Americanos", 
				dtFechaIni.ToString("dd/MMM/yyyy"), 
				dtFechaFin.ToString("dd/MMM/yyyy"), cmbBalanceTipo.Text, stEmpresa);
			DevExpress.XtraPrinting.PrintingSystem ps = new DevExpress.XtraPrinting.PrintingSystem();
			DevExpress.XtraPrinting.PrintableComponentLink link = new DevExpress.XtraPrinting.PrintableComponentLink(ps);
			link.Component = grdBalances;
			link.CreateReportHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(
				printableComponentLink1_CreateReportHeaderArea);
			link.CreateDocument();
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo en Excel?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				string stDir = MenuLatinium.stDirInicio + "\\Balance.xls";
				ps.ExportToXls(stDir, true);
				System.Diagnostics.Process.Start(stDir);
			}
			else
				link.ShowPreview();
		}

		private void btGeneraFact_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			Cursor = Cursors.Default;
			if (e.Button != MouseButtons.Right) return;
			int idTipoFactura = 0;
			if (cmbTipoFactura.Text.Trim().Length > 0)
				idTipoFactura = (int) cmbTipoFactura.Value;
			string stSelect = string.Format("Exec CompraNumerosFaltantes '{0}', '{1}', {2}",
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"),	idTipoFactura);

			grvFactura.Columns.Clear();
			grdFactura.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("Factura");
			FormatoGrillaB("Factura");
			grvFactura.BestFitColumns();
			Cursor = Cursors.WaitCursor;
		}

		private void btExcelArt_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\Kardex.xls";
			grdArticulos.ExportToXls(stDir, true);
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir);		

		}

		private void btFacturaArticulo_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			int idTipoFactura = 0;
			if (cmbTipo.Text.Trim().Length > 0)
				idTipoFactura = (int) cmbTipo.EditValue;
			int Anulado = 0;
			if (chkAnuladoArt.Checked) Anulado = 1;
			string stSelect = string.Format("Exec CompraArticulo'{0}', '{1}', {2}, {3}, {4}, {5}",
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"),	idTipoFactura,
				idProyecto, idSubProyecto, Anulado);

			grvArticulos.Columns.Clear();
			grdArticulos.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("Articulo");
			FormatoGrillaB("Articulo");
			grvArticulos.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btFacturaArticulo.ToolTip= "Tiempo de Reporte de Facturas con Articulos: " + ts.TotalSeconds.ToString("n2");
		}

		private void grdFactura_DoubleClick(object sender, System.EventArgs e)
		{
			if (grvFactura.FocusedColumn.Caption == "Asiento" || grvFactura.FocusedColumn.Caption == "Origen Asi.")
				btAsiF_Click(this, e);
			else if (grvFactura.FocusedColumn.Caption == "FechaRet" || grvFactura.FocusedColumn.Caption == "Retenido" 
				|| grvFactura.FocusedColumn.Caption == "CodReten"
				|| grvFactura.FocusedColumn.Caption.StartsWith("Retencion"))
				btRetencion_Click(this, e);
			else
				btFactura_Click(this, e);
		}

		private void btPagoFact_Click(object sender, System.EventArgs e)
		{
			int iFila = grvFactura.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Retención", "Información");
				return;
			}

			int IdCompra = (int) grvFactura.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("No existe una Factura", "Información");
				return;
			}

			int idTipoFactura = Funcion.iEscalarSQL(cdsAsiento, "Select idTipoFactura From Compra Where idCompra = " + IdCompra);
			Cursor = Cursors.WaitCursor;
//			Pagos miPago = new Pagos(idTipoFactura, IdCompra);
//			miPago.MdiParent = this.MdiParent;
//			miPago.Show();
			Cursor = Cursors.Default;
		}

		private void btTotalFG_Click(object sender, System.EventArgs e)
		{
			double dTotalDebe = 0;
			double dTotalDevolucion = 0;
			double dTotalHaber = 0;
			for (int i=0; i< grvGrupoFactura.RowCount; i++)
			{
				dTotalDebe += (double) grvGrupoFactura.GetRowCellValue(i, "Total");
				dTotalDevolucion += (double) grvGrupoFactura.GetRowCellValue(i, "Devolucion");
				dTotalHaber += (double) grvGrupoFactura.GetRowCellValue(i, "Descuento");
			}
			txtTotalGrupo.EditValue = dTotalDebe.ToString("n2");
			txtDevolucion.EditValue = (dTotalDevolucion + dTotalHaber).ToString("n2");
			txtDescuentoGrupo.EditValue = (dTotalDebe - dTotalDevolucion - dTotalHaber).ToString("n2");
		}

		private void btTotalFG_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			double dTotalDebe = 0;
			double dTotalDevolucion = 0;
			double dTotalHaber = 0;
			for (int i=0; i< grvGrupoFactura.RowCount; i++)
			{
				dTotalDebe += (double) grvGrupoFactura.GetRowCellValue(i, "Total");
				dTotalDevolucion += (double) grvGrupoFactura.GetRowCellValue(i, "Devolucion");
				dTotalHaber += (double) grvGrupoFactura.GetRowCellValue(i, "Descuento");
			}
			txtTotalGrupo.EditValue = dTotalDebe.ToString("n2");;
			txtDevolucion.EditValue = dTotalDevolucion.ToString("n2");
			txtDescuentoGrupo.EditValue = dTotalHaber.ToString("n2");
		}

		private void btCreditoFactura_Click(object sender, System.EventArgs e)
		{
			int iFila = grvCredito.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione Fila de la cual desea ver la Factura", "Información");
				return;
			}
		
			int IdCompra = (int) grvCredito.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("Fila no contiene Facturas Asociadas",	"Información");
				Cursor = Cursors.Default;
				return;
			}

			if (cmbTipoFactCred.Text.Trim().Length == 0)
			{
				MessageBox.Show("Seleccione Tipo de Factura", "Información");
				return;
			}

			Cursor = Cursors.WaitCursor;
//			Compra miCompra = new Compra(-IdCompra);
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void btCreditoAsiento_Click(object sender, System.EventArgs e)
		{
			int iFila = grvCredito.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Asiento", "Información");
				return;
			}
		
			int idAsiento = (int) grvCredito.GetRowCellValue(iFila, "idAsiento");
			if (idAsiento == 0)
			{
				MessageBox.Show("Fila no contiene Asiento Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}

			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(idAsiento);
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void btCreditoGenera_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			int idTipoFactura = 0;
			if (cmbTipoFactCred.Text.Trim().Length > 0)
				idTipoFactura = (int) cmbTipoFactCred.Value;
			int idProyecto = 0;
			if (cmbCentro.Text.Length > 0) idProyecto = (int) cmbCentro.EditValue;
			int idSubProyecto = 0;
			if (cmbSubCentro.Text.Length > 0) idProyecto = (int) cmbSubCentro.EditValue;
			int iLote = 0;
			if (chkAgrupaCredito.Checked) iLote = 1;
			int iResumen = 0;
			if (chkResumenCompra.Checked) iResumen = 1;

			string stSelect = string.Format("Exec ComparaCxC_Contabilidad '{0}', '{1}', {2}, {3}, {4}, {5}, {6}",
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"),
				idTipoFactura, idProyecto, idSubProyecto, iLote, iResumen);

			grvCredito.Columns.Clear();
			grdCredito.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("Credito");
			FormatoGrillaB("Credito");
			grvCredito.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btCreditoGenera.ToolTip= "Tiempo de Credito: " + ts.TotalSeconds.ToString("n2");
		}

		private void btCreditoCompara_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			int idTipoFactura = 0;
			if (cmbTipoFactCred.Text.Trim().Length > 0)
				idTipoFactura = (int) cmbTipoFactCred.Value;
//			int idProyecto = 0;
//			if (cmbCentro.Text.Length > 0) idProyecto = (int) cmbCentro.EditValue;
//			int idSubProyecto = 0;
//			if (cmbSubCentro.Text.Length > 0) idProyecto = (int) cmbSubCentro.EditValue;
//			int iLote = 0;
//			if (chkAgrupaCredito.Checked) iLote = 1;
//			int iResumen = 0;
//			if (chkResumenCompra.Checked) iResumen = 1;

			string stCodigo = txtCodigo.Text.Trim();
			string stSelect = string.Format("Select Count(*) From Cuenta Where Codigo = '{0}'", stCodigo);
			int idCuenta = Funcion.iEscalarSQL(cdsAsiento, stSelect);
			if (idCuenta == 0)
			{
				MessageBox.Show("No existe una cuenta con codigo " + stCodigo + " definido en la pestaña Balance");
				return;
			}
			stSelect = string.Format("Select idCuenta From Cuenta Where Codigo = '{0}'", stCodigo);
			idCuenta = Funcion.iEscalarSQL(cdsAsiento, stSelect);

			stSelect = string.Format("Exec fun_ComparaCreditoContabilidad {0}, '{1}', '{2}', {3}",
				idCuenta, dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"), idTipoFactura);

			grvCredito.Columns.Clear();
			grdCredito.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("Credito");
			FormatoGrillaB("Credito");
			grvCredito.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btCreditoCompara.ToolTip= "Tiempo de Comparación de Crédito: " + ts.TotalSeconds.ToString("n2");
		}


		private void btUtilidadGenera_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			string stSelect = string.Format("Exec Fun_Utilidad '{0}', '{1}'",
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"));

			grvUtilidad.Columns.Clear();
			grdUtilidad.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("Utilidad");
			FormatoGrillaB("Utilidad");
			grvUtilidad.BestFitColumns();
			
			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btUtilidadGenera.ToolTip = "Tiempo de Utilidad: " + ts.TotalSeconds.ToString("n2");
		}

		private void btFactUtilidad_Click(object sender, System.EventArgs e)
		{
			int iFila = grvUtilidad.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Factura", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) grvUtilidad.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}
//			Compra miCompra = new Compra(-IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void btGrabaCubo_Click(object sender, System.EventArgs e)
		{
			saveFileDialog1.InitialDirectory = Application.StartupPath;
			if (saveFileDialog1.ShowDialog() == DialogResult.OK)
			{
				pgrBalance.SaveLayoutToXml(saveFileDialog1.FileName);
			}
		}

		private void btCargaCubo_Click(object sender, System.EventArgs e)
		{
			openFileDialog1.InitialDirectory = Application.StartupPath;
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				pgrBalance.RestoreLayoutFromXml(openFileDialog1.FileName);
			}
		}


		private void btExcelCredito_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
		
		}

		private void btExcelCredito_Click(object sender, System.EventArgs e)
		{
			string stEmpresa = Funcion.sEscalarSQL(cdsAsiento, "Select Empresa From seteo");

			stTitulo = string.Format("{3}\nCredito\nDesde El {0} al {1}\nExpresado en Dólares Americanos", 
				dtFechaIni.ToString("dd/MMM/yyyy"), 
				dtFechaFin.ToString("dd/MMM/yyyy"), cmbBalanceTipo.Text, stEmpresa);
			DevExpress.XtraPrinting.PrintingSystem ps = new DevExpress.XtraPrinting.PrintingSystem();
			DevExpress.XtraPrinting.PrintableComponentLink link = new DevExpress.XtraPrinting.PrintableComponentLink(ps);
			link.Component = grdCredito;
			link.CreateReportHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(
				printableComponentLink1_CreateReportHeaderArea);
			link.CreateDocument();
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo en Excel?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				string stDir = MenuLatinium.stDirInicio + "\\Credito.xls";
				ps.ExportToXls(stDir, true);
				System.Diagnostics.Process.Start(stDir);
			}
			else
				link.ShowPreview();

		}

		private void btComparaFG_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			int idTipoFactura = 0;
			if (cmbTipoFact1.Text.Trim().Length > 0)
				idTipoFactura = (int) cmbTipoFact1.Value;
			int idProyecto = 0;
			if (cmbCentro.Text.Length > 0) idProyecto = (int) cmbCentro.EditValue;
			int idSubProyecto = 0;
			if (cmbSubCentro.Text.Length > 0) idProyecto = (int) cmbSubCentro.EditValue;
			int iLote = 0;
			if (chkGrupoAgrupado.Checked) iLote = 1;
			if (chkGrupoAgrupado.CheckState == System.Windows.Forms.CheckState.Indeterminate) iLote = -1;
			string stSelect = string.Format("Exec ComparaFacturaGrupoComp '{0}', '{1}', {2}, {3}, {4}, {5}",
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"),
				idTipoFactura, idProyecto, idSubProyecto, iLote);

			grvGrupoFactura.Columns.Clear();
			grdGrupoFactura.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("GrupoF");
			FormatoGrillaB("GrupoF");
			grvGrupoFactura.BestFitColumns();

//			btTotalFG_Click(this, e);

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btGeneraFacturaGrupo.ToolTip = "Tiempo de Facturas por Grupo: " + ts.TotalSeconds.ToString("n2");

		}

		private void btGeneraRetencion_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			int idTipoFactura = (int) rdgTipoRet.EditValue;
			int iAnulado = 0;
			if (chkAnuladaRet.Checked) iAnulado = 1;
			string stSelect = string.Format("Exec ComparaRetencionAsiento '{0}', '{1}', {2}, {3}, {4}, {5}",
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"),
				idTipoFactura, idProyecto, idSubProyecto, iAnulado);

			grvRetencion.Columns.Clear();
			grdRetencion.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("Retencion");
			FormatoGrillaB("Retencion");
			grvRetencion.BestFitColumns();
			
			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btGeneraRetencion.ToolTip = "Tiempo de Retenciones: " + ts.TotalSeconds.ToString("n2");
		}

		private void btFacturaRet_Click(object sender, System.EventArgs e)
		{
			int iFila = grvRetencion.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			int IdCompra = (int) grvRetencion.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("No existe una Factura", "Información");
				return;
			}

			Cursor = Cursors.WaitCursor;
//			Compra miCompra = new Compra(-IdCompra);
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void btRetencionRet_Click(object sender, System.EventArgs e)
		{
			int iFila = grvRetencion.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
			int IdRetencion = (int) grvRetencion.GetRowCellValue(iFila, "idRetencion");
			int IdCompra = (int) grvRetencion.GetRowCellValue(iFila,"idCompra");
			int IdTipoFactura = (int) grvRetencion.GetRowCellValue(iFila, "idTipoFactura");
			if (IdRetencion == 0)
			{
				MessageBox.Show("No existe una Retención", "Información");
				return;
			}

			if (IdTipoFactura == 1)
			{
				Cursor = Cursors.WaitCursor;
				IngresoAT miEgreso = new IngresoAT();
				miEgreso.IdFactura = IdCompra;
				miEgreso.Show();
			}
			else
			{
				Cursor = Cursors.WaitCursor;
//				EgresoAT miEgreso = new EgresoAT(IdRetencion.ToString());
//				miEgreso.IdFactura = IdCompra;
//				miEgreso.Show();
			}
			Cursor = Cursors.Default;
		}

		private void btGeneraRetencion_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			DateTime dtInicio = DateTime.Now;

			string stSelect = "EXEC RetencionesVerifica";
			grvRetencion.Columns.Clear();
			grdRetencion.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("Retencion");
			FormatoGrillaB("Retencion");
			grvRetencion.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btGeneraRetencion.ToolTip = "Tiempo de Retenciones: " + ts.TotalSeconds.ToString("n2");
		}

		private void btTotalRet_Click(object sender, System.EventArgs e)
		{
			double dTotalRet = 0;
			for (int i=0; i< grvRetencion.RowCount; i++)
			{ 
				dTotalRet += (double) grvRetencion.GetRowCellValue(i, "Retenido");
			}
			txtTotalRet.EditValue = dTotalRet;
		}

		private void btFlujo_Click(object sender, System.EventArgs e)
		{
			FlujoCaja miFlujo = new FlujoCaja();
			miFlujo.MdiParent = this.MdiParent;
			miFlujo.Show();
		}

		private void btInconsistencia_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			string stCodigo = txtCodigo.Text.Trim();
			string stSelect = string.Format("Exec ComparaCXC_Asiento_Detalle '{0}', '{1}', '{2}'",
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"), stCodigo);

			grvCredito.Columns.Clear();
			grdCredito.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("Credito");
			FormatoGrillaB("Credito");
			grvCredito.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btInconsistencia.ToolTip= "Tiempo de Comparación de Inconsistencias: " + ts.TotalSeconds.ToString("n2");
		}

		private void btInconsistencia_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			DateTime dtInicio = DateTime.Now;
			string stCodigo = txtCodigo.Text.Trim();
			int idTipoFactura = 0;
			if (cmbTipoFactCred.Text.Trim().Length > 0)
				idTipoFactura = (int) cmbTipoFactCred.Value;
			string stSelect = string.Format("Exec ListadoValorDif '{0}', '{1}', {2}",
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"), idTipoFactura);

			grvCredito.Columns.Clear();
			grdCredito.DataSource = Funcion.dvProcedimiento(cdsAsiento, stSelect);
			FormatoGrillaB("Credito");
			FormatoGrillaB("Credito");
			grvCredito.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btInconsistencia.ToolTip= "Tiempo de Comparación de Inconsistencias: " + ts.TotalSeconds.ToString("n2");
		}

		private void btInventario_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			string stCodigo = txtCodigo.Text.Trim();
			string stSelect = string.Format("Select Count(*) From Cuenta Where Codigo = '{0}'", stCodigo);
			int idCuenta = Funcion.iEscalarSQL(cdsAsiento, stSelect);
			if (idCuenta == 0)
			{
				MessageBox.Show("No existe una cuenta con codigo " + stCodigo + " definido en la pestaña Balance");
				return;
			}

			int idTipo = 0;
			if (cmbCostoPeriodo.Text.Length > 0) idTipo = (int) cmbCostoPeriodo.SelectedIndex;

			string stExec = string.Format("Exec ComparaInventarioAsiento '{0}', '{1}', '{2}', {3}", 
				stCodigo, dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"), idTipo + 1);
			grvInventario.Columns.Clear();
			grdInventario.DataSource = Funcion.dvProcedimiento(cdsSeteoC, stExec);
			FormatoGrillaB("Inventario");
			FormatoGrillaB("Inventario");
			grvInventario.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btInventario.ToolTip= "Tiempo de Reporte de Inventario: " + ts.TotalSeconds.ToString("n2");
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			int iAgrupado = 1;
			int idTipoFactura = 0;
			int iFila = grvInventario.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) grvForm104.GetRowCellValue(iFila, "Fecha");

			string stExec = string.Format("Exec ArticuloKardexAsiento 0, '{0}', '{1}', {2}, {3}, {4}, {5}", 
				dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"), iAgrupado, idTipoFactura,
				idProyecto, idSubProyecto);
			grvInventario.Columns.Clear();
			grdInventario.DataSource = Funcion.dvProcedimiento(cdsSeteoC, stExec);
			FormatoGrillaB("Inventario");
			FormatoGrillaB("Inventario");
			grvInventario.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btKardexComp.ToolTip= "Tiempo de Reporte de Kardex: " + ts.TotalSeconds.ToString("n2");

		}

		private void grdRetencion_DoubleClick(object sender, System.EventArgs e)
		{
			if (grvRetencion.FocusedColumn.Caption == "Numero" || grvRetencion.FocusedColumn.Caption == "Nombre"
				|| grvRetencion.FocusedColumn.Caption == "FechaFact")
				btFacturaRet_Click(this, e);
			else
				btRetencionRet_Click(this, e);

		}

		private void grvGrupoFactura_DoubleClick(object sender, System.EventArgs e)
		{
			if (grvGrupoFactura.FocusedColumn.Caption == "Codigo" || grvGrupoFactura.FocusedColumn.Caption == "Código")
				btAsientoGrupo_Click(this, e);
			else
				btFacturaGrupo_Click(this, e);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);	
		}

		private void btVencido_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			DateTime dtFecha = (DateTime)cmbHasta.EditValue;
			int optTipo = (int)optCobrar.EditValue;
			string stSelect = String.Format("EXEC PagoXCobrar '{0:yyyyMMdd}', {1}, {2}, {3}",
				dtFecha, optTipo, idProyecto, idSubProyecto);
			grvAntiguedad.Columns.Clear();
			grdAntiguedad.DataSource = Funcion.dvProcedimiento(cdsSeteoC, stSelect);
			FormatoGrillaB("Antiguedad");
			FormatoGrillaB("Antiguedad");
			grvAntiguedad.BestFitColumns();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			btAntiguedad.ToolTip= "Tiempo de Reporte de Antigüedad: " + ts.TotalSeconds.ToString("n2");
		
		}
	}
}
