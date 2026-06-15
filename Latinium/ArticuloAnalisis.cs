using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ArticuloAnalisis.
	/// </summary>
	public class ArticuloAnalisis : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsCompra;
		private DevExpress.XtraEditors.SimpleButton btGenerar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsAnalisis;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private C1.Data.C1DataSet cdsNumero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbOperador;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsArticulo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsCliente;
		private DevExpress.XtraEditors.SimpleButton btArticulo;
		private DevExpress.XtraEditors.SimpleButton btCliente;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl Marco;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl marcoArticulo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTituloArt;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdArticulo;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCliente;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl marcoGenerar;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl marcoClientes;
		private DevExpress.XtraEditors.SimpleButton btExcelGen;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdGenerar;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTituloCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTituloGenerar;
		private DevExpress.XtraEditors.SimpleButton btExcelArt;
		private DevExpress.XtraEditors.SimpleButton btExcelCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdCliente;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private System.Windows.Forms.Label label3;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btSalir;
		private System.Windows.Forms.TextBox txtOrigen;
		private C1.Data.C1DataSet cdsGrupo;
		private DevExpress.XtraEditors.SimpleButton btCotizacionEgreso;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private DevExpress.XtraEditors.SimpleButton btEgresoCotiza;
		private DevExpress.XtraEditors.SimpleButton btDetEgresoCot;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReporteCompara;
		private DevExpress.XtraEditors.SimpleButton ultraButton1;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataView cdvReporte;
		private C1.Data.C1DataSet cdsReporte;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private DevExpress.XtraEditors.SimpleButton btImprimir;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkClienteSector;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbClienteSector;
		private C1.Data.C1DataView cdvSector;
		private C1.Data.C1DataSet cdsClienteTabla;
		private DevExpress.XtraEditors.SimpleButton btNoImporta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubCentro;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSubCentro;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCentro;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentro;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataView cdvSubProyecto;
		private DevExpress.XtraEditors.SimpleButton ultraButton2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private System.Windows.Forms.GroupBox groupBox2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private DevExpress.XtraEditors.SimpleButton ultraButton3;
		private System.Windows.Forms.GroupBox groupBox3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private DevExpress.XtraEditors.CheckEdit chkResumen;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaDesde;
		private C1.Data.C1DataSet c1DataSet1;
		private C1.Data.C1DataView cdvCompra;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNumOP;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProyecto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumOP;
		private System.Windows.Forms.GroupBox groupBox4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesdeFifo;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHastaFifo;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticulo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkArticulo;
		private DevExpress.XtraEditors.SimpleButton btnImprimeFifo;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private System.Windows.Forms.ErrorProvider errorProvider1;
//	private Acceso miAcceso;
		public ArticuloAnalisis()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteSector", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ClteCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointAutorizado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucionFinanaciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idObservaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HaceRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaConsignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraDesde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraHasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaSwift");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Beneficiario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasConsignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoProveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoBarras");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NResolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetieneIVA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BloqueadoCrediPoint");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Courier");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Barrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProveedorIESS");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PLNSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FacturaBajoCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito4");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargaSeriales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem14 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem15 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem16 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Utilidad");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Utilidad");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("E_idEducativo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantConf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SeriePropiaAlterna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ManejaSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaDscto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValTransfTemporal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioManual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PeriodoGarantia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodArtAntiguo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoReemplazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tamaño");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Combo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vida_Util");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIVA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPreciador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoBarras");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Inventario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpuestoIVA");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CajaChica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton5 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton6 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CajaChica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Compras");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ventas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Egresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DevCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DevVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pedidos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab6 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbClienteSector = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvSector = new C1.Data.C1DataView();
			this.cdsClienteTabla = new C1.Data.C1DataSet();
			this.chkClienteSector = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btCotizacionEgreso = new DevExpress.XtraEditors.SimpleButton();
			this.txtOrigen = new System.Windows.Forms.TextBox();
			this.btSalir = new DevExpress.XtraEditors.SimpleButton();
			this.label3 = new System.Windows.Forms.Label();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.chkCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.optFiltro = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btCliente = new DevExpress.XtraEditors.SimpleButton();
			this.btArticulo = new DevExpress.XtraEditors.SimpleButton();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsNumero = new C1.Data.C1DataSet();
			this.cmbOperador = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.marcoGenerar = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtTituloGenerar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btExcelGen = new DevExpress.XtraEditors.SimpleButton();
			this.grdGenerar = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.marcoArticulo = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btExcelArt = new DevExpress.XtraEditors.SimpleButton();
			this.txtTituloArt = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.grdArticulo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsArticulo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.marcoClientes = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btExcelCliente = new DevExpress.XtraEditors.SimpleButton();
			this.txtTituloCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.grdCliente = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsCliente = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.chkArticulo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbArticulo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cmbHastaFifo = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbDesdeFifo = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnImprimeFifo = new DevExpress.XtraEditors.SimpleButton();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraButton2 = new DevExpress.XtraEditors.SimpleButton();
			this.cmbSubCentro = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvSubProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.chkSubCentro = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCentro = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbCentro = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.btNoImporta = new DevExpress.XtraEditors.SimpleButton();
			this.ultraButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.txtReporteCompara = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btDetEgresoCot = new DevExpress.XtraEditors.SimpleButton();
			this.btEgresoCotiza = new DevExpress.XtraEditors.SimpleButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.btImprimir = new DevExpress.XtraEditors.SimpleButton();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvReporte = new C1.Data.C1DataView();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.chkNumOP = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtFechaHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.chkResumen = new DevExpress.XtraEditors.CheckEdit();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.chkProyecto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraButton3 = new DevExpress.XtraEditors.SimpleButton();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.txtNumOP = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsGrupo = new C1.Data.C1DataSet();
			this.udsAnalisis = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.Marco = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.c1DataSet1 = new C1.Data.C1DataSet();
			this.cdvCompra = new C1.Data.C1DataView();
			this.errorProvider1 = new System.Windows.Forms.ErrorProvider();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbClienteSector)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSector)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClienteTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOperador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			this.marcoGenerar.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtTituloGenerar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdGenerar)).BeginInit();
			this.marcoArticulo.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtTituloArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsArticulo)).BeginInit();
			this.marcoClientes.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtTituloCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCliente)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			this.groupBox4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHastaFifo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesdeFifo)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubCentro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReporteCompara)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkResumen.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumOP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAnalisis)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Marco)).BeginInit();
			this.Marco.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.c1DataSet1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.cmbClienteSector);
			this.ultraTabPageControl2.Controls.Add(this.chkClienteSector);
			this.ultraTabPageControl2.Controls.Add(this.btCotizacionEgreso);
			this.ultraTabPageControl2.Controls.Add(this.txtOrigen);
			this.ultraTabPageControl2.Controls.Add(this.btSalir);
			this.ultraTabPageControl2.Controls.Add(this.label3);
			this.ultraTabPageControl2.Controls.Add(this.optTipo);
			this.ultraTabPageControl2.Controls.Add(this.chkCliente);
			this.ultraTabPageControl2.Controls.Add(this.cmbCliente);
			this.ultraTabPageControl2.Controls.Add(this.optFiltro);
			this.ultraTabPageControl2.Controls.Add(this.btCliente);
			this.ultraTabPageControl2.Controls.Add(this.btArticulo);
			this.ultraTabPageControl2.Controls.Add(this.cmbHasta);
			this.ultraTabPageControl2.Controls.Add(this.cmbDesde);
			this.ultraTabPageControl2.Controls.Add(this.label2);
			this.ultraTabPageControl2.Controls.Add(this.label1);
			this.ultraTabPageControl2.Controls.Add(this.cmbTipo);
			this.ultraTabPageControl2.Controls.Add(this.cmbOperador);
			this.ultraTabPageControl2.Controls.Add(this.txtValor);
			this.ultraTabPageControl2.Controls.Add(this.btGenerar);
			this.ultraTabPageControl2.Controls.Add(this.btAceptar);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(660, 413);
			// 
			// cmbClienteSector
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbClienteSector.Appearance = appearance1;
			this.cmbClienteSector.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbClienteSector.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbClienteSector.DataSource = this.cdvSector;
			this.cmbClienteSector.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 73;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 197;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbClienteSector.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbClienteSector.DisplayMember = "Sector";
			this.cmbClienteSector.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbClienteSector.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbClienteSector.Location = new System.Drawing.Point(384, 95);
			this.cmbClienteSector.Name = "cmbClienteSector";
			this.cmbClienteSector.Size = new System.Drawing.Size(216, 21);
			this.cmbClienteSector.TabIndex = 86;
			this.cmbClienteSector.ValueMember = "idSector";
			// 
			// cdvSector
			// 
			this.cdvSector.BindingContextCtrl = this;
			this.cdvSector.DataSet = this.cdsClienteTabla;
			this.cdvSector.Sort = "Sector";
			this.cdvSector.TableName = "";
			this.cdvSector.TableViewName = "ClienteSector";
			// 
			// cdsClienteTabla
			// 
			this.cdsClienteTabla.BindingContextCtrl = this;
			this.cdsClienteTabla.DataLibrary = "LibFacturacion";
			this.cdsClienteTabla.DataLibraryUrl = "";
			this.cdsClienteTabla.DataSetDef = "dsClienteTabla";
			this.cdsClienteTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsClienteTabla.Name = "cdsClienteTabla";
			this.cdsClienteTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsClienteTabla.SchemaDef = null;
			this.cdsClienteTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsClienteTabla_BeforeFill);
			// 
			// chkClienteSector
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkClienteSector.Appearance = appearance2;
			this.chkClienteSector.BackColor = System.Drawing.Color.Transparent;
			this.chkClienteSector.Checked = true;
			this.chkClienteSector.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkClienteSector.Location = new System.Drawing.Point(224, 95);
			this.chkClienteSector.Name = "chkClienteSector";
			this.chkClienteSector.Size = new System.Drawing.Size(96, 21);
			this.chkClienteSector.TabIndex = 23;
			this.chkClienteSector.Text = "Todo Sector";
			// 
			// btCotizacionEgreso
			// 
			this.btCotizacionEgreso.Location = new System.Drawing.Point(216, 267);
			this.btCotizacionEgreso.Name = "btCotizacionEgreso";
			this.btCotizacionEgreso.Size = new System.Drawing.Size(160, 26);
			this.btCotizacionEgreso.TabIndex = 22;
			this.btCotizacionEgreso.Text = "Cotización Egreso Compara";
			this.btCotizacionEgreso.Click += new System.EventHandler(this.btCotizacionEgreso_Click);
			// 
			// txtOrigen
			// 
			this.txtOrigen.Location = new System.Drawing.Point(400, 276);
			this.txtOrigen.Name = "txtOrigen";
			this.txtOrigen.TabIndex = 21;
			this.txtOrigen.Text = "Roberto";
			// 
			// btSalir
			// 
			this.btSalir.Location = new System.Drawing.Point(384, 345);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(72, 25);
			this.btSalir.TabIndex = 20;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// label3
			// 
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(224, 121);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(144, 24);
			this.label3.TabIndex = 19;
			this.label3.Text = "Tipo de Transaccion";
			// 
			// optTipo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipo.Appearance = appearance3;
			this.optTipo.CheckedIndex = 0;
			this.optTipo.ItemAppearance = appearance4;
			this.optTipo.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Generacion Global";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Rating de artículos";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Rating de Clientes";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Venta Neta Articulo";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Venta Neta Cliente";
			valueListItem6.DataValue = 5;
			valueListItem6.DisplayText = "Inventario Mov.";
			valueListItem7.DataValue = 6;
			valueListItem7.DisplayText = "Cotización Egreso Comp.";
			valueListItem8.DataValue = 7;
			valueListItem8.DisplayText = "Personalizado";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.Items.Add(valueListItem3);
			this.optTipo.Items.Add(valueListItem4);
			this.optTipo.Items.Add(valueListItem5);
			this.optTipo.Items.Add(valueListItem6);
			this.optTipo.Items.Add(valueListItem7);
			this.optTipo.Items.Add(valueListItem8);
			this.optTipo.ItemSpacingVertical = 10;
			this.optTipo.Location = new System.Drawing.Point(32, 34);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(168, 224);
			this.optTipo.TabIndex = 18;
			this.optTipo.Text = "Generacion Global";
			this.optTipo.ValueChanged += new System.EventHandler(this.optTipo_ValueChanged);
			// 
			// chkCliente
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCliente.Appearance = appearance5;
			this.chkCliente.BackColor = System.Drawing.Color.Transparent;
			this.chkCliente.Checked = true;
			this.chkCliente.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCliente.Location = new System.Drawing.Point(224, 69);
			this.chkCliente.Name = "chkCliente";
			this.chkCliente.Size = new System.Drawing.Size(96, 21);
			this.chkCliente.TabIndex = 17;
			this.chkCliente.Text = "Todo Cliente";
			// 
			// cmbCliente
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCliente.Appearance = appearance6;
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataSource = this.cdvCliente;
			ultraGridBand2.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 104;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 8;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Width = 262;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 17;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 8;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 17;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn17.Header.VisiblePosition = 13;
			ultraGridColumn18.Header.VisiblePosition = 14;
			ultraGridColumn19.Header.VisiblePosition = 15;
			ultraGridColumn20.Header.VisiblePosition = 16;
			ultraGridColumn21.Header.VisiblePosition = 17;
			ultraGridColumn22.Header.VisiblePosition = 18;
			ultraGridColumn23.Header.VisiblePosition = 19;
			ultraGridColumn24.Header.VisiblePosition = 20;
			ultraGridColumn25.Header.VisiblePosition = 21;
			ultraGridColumn26.Header.VisiblePosition = 22;
			ultraGridColumn27.Header.VisiblePosition = 23;
			ultraGridColumn28.Header.VisiblePosition = 24;
			ultraGridColumn29.Header.VisiblePosition = 25;
			ultraGridColumn30.Header.VisiblePosition = 26;
			ultraGridColumn31.Header.VisiblePosition = 27;
			ultraGridColumn32.Header.VisiblePosition = 28;
			ultraGridColumn33.Header.VisiblePosition = 29;
			ultraGridColumn34.Header.VisiblePosition = 30;
			ultraGridColumn35.Header.VisiblePosition = 31;
			ultraGridColumn36.Header.VisiblePosition = 32;
			ultraGridColumn37.Header.VisiblePosition = 33;
			ultraGridColumn38.Header.VisiblePosition = 34;
			ultraGridColumn39.Header.VisiblePosition = 35;
			ultraGridColumn40.Header.VisiblePosition = 36;
			ultraGridColumn41.Header.VisiblePosition = 37;
			ultraGridColumn42.Header.VisiblePosition = 38;
			ultraGridColumn43.Header.VisiblePosition = 39;
			ultraGridColumn44.Header.VisiblePosition = 40;
			ultraGridColumn45.Header.VisiblePosition = 41;
			ultraGridColumn46.Header.VisiblePosition = 42;
			ultraGridColumn47.Header.VisiblePosition = 43;
			ultraGridColumn48.Header.VisiblePosition = 44;
			ultraGridColumn49.Header.VisiblePosition = 45;
			ultraGridColumn50.Header.VisiblePosition = 46;
			ultraGridColumn51.Header.VisiblePosition = 47;
			ultraGridColumn52.Header.VisiblePosition = 48;
			ultraGridColumn53.Header.VisiblePosition = 49;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn53});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(384, 69);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(218, 21);
			this.cmbCliente.TabIndex = 16;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
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
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// optFiltro
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro.Appearance = appearance7;
			this.optFiltro.CheckedIndex = 0;
			this.optFiltro.ItemAppearance = appearance8;
			this.optFiltro.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem9.DataValue = 0;
			valueListItem9.DisplayText = "Todo Artículo";
			valueListItem10.DataValue = 1;
			valueListItem10.DisplayText = "Artículo con Saldo";
			valueListItem11.DataValue = 2;
			valueListItem11.DisplayText = "Artículo con Movimiento";
			this.optFiltro.Items.Add(valueListItem9);
			this.optFiltro.Items.Add(valueListItem10);
			this.optFiltro.Items.Add(valueListItem11);
			this.optFiltro.ItemSpacingVertical = 10;
			this.optFiltro.Location = new System.Drawing.Point(216, 164);
			this.optFiltro.Name = "optFiltro";
			this.optFiltro.Size = new System.Drawing.Size(160, 94);
			this.optFiltro.TabIndex = 15;
			this.optFiltro.Text = "Todo Artículo";
			this.optFiltro.ValueChanged += new System.EventHandler(this.optFiltro_ValueChanged);
			// 
			// btCliente
			// 
			this.btCliente.Location = new System.Drawing.Point(16, 370);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(144, 26);
			this.btCliente.TabIndex = 14;
			this.btCliente.Text = "Movimiento de Clientes";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			// 
			// btArticulo
			// 
			this.btArticulo.Location = new System.Drawing.Point(16, 336);
			this.btArticulo.Name = "btArticulo";
			this.btArticulo.Size = new System.Drawing.Size(144, 26);
			this.btArticulo.TabIndex = 13;
			this.btArticulo.Text = "Artículos Condición";
			this.btArticulo.Click += new System.EventHandler(this.btArticulo_Click);
			// 
			// cmbHasta
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbHasta.Appearance = appearance9;
			dateButton1.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton1);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(440, 34);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 24;
			this.cmbHasta.Size = new System.Drawing.Size(112, 21);
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 3;
			this.cmbHasta.Value = new System.DateTime(2005, 12, 24, 0, 0, 0, 0);
			// 
			// cmbDesde
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDesde.Appearance = appearance10;
			dateButton2.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton2);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(272, 34);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 24;
			this.cmbDesde.Size = new System.Drawing.Size(112, 21);
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 2;
			this.cmbDesde.Value = new System.DateTime(2005, 12, 24, 0, 0, 0, 0);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(392, 34);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 5;
			this.label2.Text = "Hasta:";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(224, 34);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 4;
			this.label1.Text = "Desde:";
			// 
			// cmbTipo
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance11;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataMember = "CompraNumero";
			this.cmbTipo.DataSource = this.cdsNumero;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn54.Header.VisiblePosition = 5;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridColumn55.Width = 93;
			ultraGridColumn56.Header.VisiblePosition = 7;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 3;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 2;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 1;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 6;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 8;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 4;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 9;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Width = 114;
			ultraGridColumn64.Header.VisiblePosition = 10;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 117;
			ultraGridColumn65.Header.VisiblePosition = 11;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 99;
			ultraGridColumn66.Header.VisiblePosition = 12;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn66.Width = 24;
			ultraGridColumn67.Header.VisiblePosition = 13;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn67.Width = 23;
			ultraGridColumn68.Header.VisiblePosition = 14;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 43;
			ultraGridColumn69.Header.VisiblePosition = 15;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.Width = 43;
			ultraGridColumn70.Header.VisiblePosition = 16;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 27;
			ultraGridColumn71.Header.VisiblePosition = 17;
			ultraGridColumn71.Width = 35;
			ultraGridColumn72.Header.VisiblePosition = 18;
			ultraGridColumn72.Width = 41;
			ultraGridColumn73.Header.VisiblePosition = 19;
			ultraGridColumn73.Width = 28;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn73});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipo.DisplayMember = "Nombre";
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(384, 121);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(216, 21);
			this.cmbTipo.TabIndex = 7;
			this.cmbTipo.ValueMember = "idTipoFactura";
			// 
			// cdsNumero
			// 
			this.cdsNumero.BindingContextCtrl = this;
			this.cdsNumero.DataLibrary = "LibFacturacion";
			this.cdsNumero.DataLibraryUrl = "";
			this.cdsNumero.DataSetDef = "dsCompraTabla";
			this.cdsNumero.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsNumero.Name = "cdsNumero";
			this.cdsNumero.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsNumero.SchemaDef = null;
			this.cdsNumero.BeforeFill += new C1.Data.FillEventHandler(this.cdsNumero_BeforeFill);
			// 
			// cmbOperador
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbOperador.Appearance = appearance12;
			this.cmbOperador.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem12.DataValue = 0;
			valueListItem12.DisplayText = "Mayor";
			valueListItem13.DataValue = 1;
			valueListItem13.DisplayText = "Igual";
			valueListItem14.DataValue = 2;
			valueListItem14.DisplayText = "Menor";
			valueListItem15.DataValue = 3;
			valueListItem15.DisplayText = "Mayor o Igual";
			valueListItem16.DataValue = 4;
			valueListItem16.DisplayText = "Menor o Igual";
			this.cmbOperador.Items.Add(valueListItem12);
			this.cmbOperador.Items.Add(valueListItem13);
			this.cmbOperador.Items.Add(valueListItem14);
			this.cmbOperador.Items.Add(valueListItem15);
			this.cmbOperador.Items.Add(valueListItem16);
			this.cmbOperador.Location = new System.Drawing.Point(392, 190);
			this.cmbOperador.Name = "cmbOperador";
			this.cmbOperador.Size = new System.Drawing.Size(88, 22);
			this.cmbOperador.TabIndex = 10;
			// 
			// txtValor
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance13;
			this.txtValor.Location = new System.Drawing.Point(488, 190);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(112, 22);
			this.txtValor.TabIndex = 11;
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(16, 302);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(144, 25);
			this.btGenerar.TabIndex = 0;
			this.btGenerar.Text = "Generar Global";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(192, 345);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(72, 25);
			this.btAceptar.TabIndex = 6;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// marcoGenerar
			// 
			this.marcoGenerar.Controls.Add(this.txtTituloGenerar);
			this.marcoGenerar.Controls.Add(this.btExcelGen);
			this.marcoGenerar.Controls.Add(this.grdGenerar);
			this.marcoGenerar.Location = new System.Drawing.Point(-10000, -10769);
			this.marcoGenerar.Name = "marcoGenerar";
			this.marcoGenerar.Size = new System.Drawing.Size(660, 413);
			// 
			// txtTituloGenerar
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTituloGenerar.Appearance = appearance14;
			this.txtTituloGenerar.Location = new System.Drawing.Point(16, 9);
			this.txtTituloGenerar.Name = "txtTituloGenerar";
			this.txtTituloGenerar.Size = new System.Drawing.Size(552, 22);
			this.txtTituloGenerar.TabIndex = 4;
			// 
			// btExcelGen
			// 
			this.btExcelGen.Location = new System.Drawing.Point(576, 9);
			this.btExcelGen.Name = "btExcelGen";
			this.btExcelGen.Size = new System.Drawing.Size(64, 24);
			this.btExcelGen.TabIndex = 3;
			this.btExcelGen.Text = "Excel";
			this.btExcelGen.Click += new System.EventHandler(this.btExcelGen_Click);
			// 
			// grdGenerar
			// 
			this.grdGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdGenerar.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdGenerar.DisplayLayout.GroupByBox.Prompt = "Arrastre columna que desea Agrupar";
			this.grdGenerar.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grdGenerar.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.grdGenerar.DisplayLayout.Override.CardAreaAppearance = appearance15;
			this.grdGenerar.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.grdGenerar.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 10F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdGenerar.DisplayLayout.Override.HeaderAppearance = appearance16;
			this.grdGenerar.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdGenerar.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdGenerar.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.grdGenerar.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grdGenerar.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grdGenerar.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.grdGenerar.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.grdGenerar.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.grdGenerar.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.grdGenerar.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.grdGenerar.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdGenerar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdGenerar.Location = new System.Drawing.Point(14, 34);
			this.grdGenerar.Name = "grdGenerar";
			this.grdGenerar.Size = new System.Drawing.Size(631, 372);
			this.grdGenerar.TabIndex = 2;
			this.grdGenerar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdGenerar_KeyDown);
			// 
			// marcoArticulo
			// 
			this.marcoArticulo.Controls.Add(this.btExcelArt);
			this.marcoArticulo.Controls.Add(this.txtTituloArt);
			this.marcoArticulo.Controls.Add(this.grdArticulo);
			this.marcoArticulo.Location = new System.Drawing.Point(-10000, -10769);
			this.marcoArticulo.Name = "marcoArticulo";
			this.marcoArticulo.Size = new System.Drawing.Size(660, 413);
			// 
			// btExcelArt
			// 
			this.btExcelArt.Location = new System.Drawing.Point(584, 9);
			this.btExcelArt.Name = "btExcelArt";
			this.btExcelArt.Size = new System.Drawing.Size(64, 24);
			this.btExcelArt.TabIndex = 4;
			this.btExcelArt.Text = "Excel";
			this.btExcelArt.Click += new System.EventHandler(this.btExcelArt_Click);
			// 
			// txtTituloArt
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTituloArt.Appearance = appearance19;
			this.txtTituloArt.Location = new System.Drawing.Point(16, 9);
			this.txtTituloArt.Name = "txtTituloArt";
			this.txtTituloArt.Size = new System.Drawing.Size(568, 22);
			this.txtTituloArt.TabIndex = 1;
			// 
			// grdArticulo
			// 
			this.grdArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdArticulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdArticulo.DataMember = "Band 0";
			this.grdArticulo.DataSource = this.udsArticulo;
			appearance20.BackColor = System.Drawing.Color.White;
			this.grdArticulo.DisplayLayout.Appearance = appearance20;
			this.grdArticulo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn74.Header.VisiblePosition = 0;
			ultraGridColumn74.Width = 157;
			ultraGridColumn75.Header.VisiblePosition = 1;
			ultraGridColumn75.Width = 149;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellAppearance = appearance21;
			ultraGridColumn76.Format = "#,##0.00";
			ultraGridColumn76.Header.VisiblePosition = 2;
			ultraGridColumn76.Width = 147;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn77.CellAppearance = appearance22;
			ultraGridColumn77.Format = "#,##0.00";
			ultraGridColumn77.Header.VisiblePosition = 3;
			ultraGridColumn77.Width = 158;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77});
			this.grdArticulo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.grdArticulo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdArticulo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grdArticulo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance23.BackColor = System.Drawing.Color.Transparent;
			this.grdArticulo.DisplayLayout.Override.CardAreaAppearance = appearance23;
			this.grdArticulo.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.grdArticulo.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance24.FontData.BoldAsString = "True";
			appearance24.FontData.Name = "Arial";
			appearance24.FontData.SizeInPoints = 10F;
			appearance24.ForeColor = System.Drawing.Color.White;
			appearance24.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdArticulo.DisplayLayout.Override.HeaderAppearance = appearance24;
			this.grdArticulo.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdArticulo.DisplayLayout.Override.RowSelectorAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdArticulo.DisplayLayout.Override.SelectedRowAppearance = appearance26;
			this.grdArticulo.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grdArticulo.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grdArticulo.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.grdArticulo.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.grdArticulo.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.grdArticulo.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.grdArticulo.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.grdArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdArticulo.Location = new System.Drawing.Point(16, 34);
			this.grdArticulo.Name = "grdArticulo";
			this.grdArticulo.Size = new System.Drawing.Size(631, 355);
			this.grdArticulo.TabIndex = 0;
			this.grdArticulo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdArticulo_KeyDown);
			// 
			// udsArticulo
			// 
			ultraDataColumn3.DataType = typeof(System.Double);
			ultraDataColumn4.DataType = typeof(System.Double);
			this.udsArticulo.Band.Columns.AddRange(new object[] {
																														ultraDataColumn1,
																														ultraDataColumn2,
																														ultraDataColumn3,
																														ultraDataColumn4});
			// 
			// marcoClientes
			// 
			this.marcoClientes.Controls.Add(this.btExcelCliente);
			this.marcoClientes.Controls.Add(this.txtTituloCliente);
			this.marcoClientes.Controls.Add(this.grdCliente);
			this.marcoClientes.Location = new System.Drawing.Point(-10000, -10769);
			this.marcoClientes.Name = "marcoClientes";
			this.marcoClientes.Size = new System.Drawing.Size(660, 413);
			// 
			// btExcelCliente
			// 
			this.btExcelCliente.Location = new System.Drawing.Point(576, 9);
			this.btExcelCliente.Name = "btExcelCliente";
			this.btExcelCliente.Size = new System.Drawing.Size(64, 24);
			this.btExcelCliente.TabIndex = 4;
			this.btExcelCliente.Text = "Excel";
			this.btExcelCliente.Click += new System.EventHandler(this.btExcelCliente_Click);
			// 
			// txtTituloCliente
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTituloCliente.Appearance = appearance27;
			this.txtTituloCliente.Location = new System.Drawing.Point(24, 9);
			this.txtTituloCliente.Name = "txtTituloCliente";
			this.txtTituloCliente.Size = new System.Drawing.Size(544, 22);
			this.txtTituloCliente.TabIndex = 2;
			// 
			// grdCliente
			// 
			this.grdCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdCliente.DataMember = "Band 0";
			this.grdCliente.DataSource = this.udsCliente;
			appearance28.BackColor = System.Drawing.Color.White;
			this.grdCliente.DisplayLayout.Appearance = appearance28;
			this.grdCliente.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn78.Header.VisiblePosition = 0;
			ultraGridColumn78.Width = 182;
			ultraGridColumn79.Header.VisiblePosition = 1;
			ultraGridColumn79.Width = 103;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance29;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.Caption = "Porcentaje";
			ultraGridColumn80.Header.VisiblePosition = 5;
			ultraGridColumn80.Width = 74;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn81.CellAppearance = appearance30;
			ultraGridColumn81.Format = "#,##0.00";
			ultraGridColumn81.Header.VisiblePosition = 2;
			ultraGridColumn81.Width = 80;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn82.CellAppearance = appearance31;
			ultraGridColumn82.Format = "#,##0.00";
			ultraGridColumn82.Header.VisiblePosition = 3;
			ultraGridColumn82.Width = 79;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn83.CellAppearance = appearance32;
			ultraGridColumn83.Format = "#,##0.00";
			ultraGridColumn83.Header.VisiblePosition = 4;
			ultraGridColumn83.Width = 101;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn78,
																										 ultraGridColumn79,
																										 ultraGridColumn80,
																										 ultraGridColumn81,
																										 ultraGridColumn82,
																										 ultraGridColumn83});
			this.grdCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.grdCliente.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdCliente.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grdCliente.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance33.BackColor = System.Drawing.Color.Transparent;
			this.grdCliente.DisplayLayout.Override.CardAreaAppearance = appearance33;
			this.grdCliente.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.grdCliente.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance34.FontData.BoldAsString = "True";
			appearance34.FontData.Name = "Arial";
			appearance34.FontData.SizeInPoints = 10F;
			appearance34.ForeColor = System.Drawing.Color.White;
			appearance34.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdCliente.DisplayLayout.Override.HeaderAppearance = appearance34;
			this.grdCliente.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCliente.DisplayLayout.Override.RowSelectorAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCliente.DisplayLayout.Override.SelectedRowAppearance = appearance36;
			this.grdCliente.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grdCliente.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grdCliente.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.grdCliente.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.grdCliente.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.grdCliente.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.grdCliente.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.grdCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdCliente.Location = new System.Drawing.Point(16, 34);
			this.grdCliente.Name = "grdCliente";
			this.grdCliente.Size = new System.Drawing.Size(640, 362);
			this.grdCliente.TabIndex = 0;
			this.grdCliente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdCliente_KeyDown);
			// 
			// udsCliente
			// 
			ultraDataColumn7.DataType = typeof(System.Double);
			ultraDataColumn8.DataType = typeof(System.Double);
			ultraDataColumn9.DataType = typeof(System.Double);
			ultraDataColumn10.DataType = typeof(System.Double);
			this.udsCliente.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn5,
																													 ultraDataColumn6,
																													 ultraDataColumn7,
																													 ultraDataColumn8,
																													 ultraDataColumn9,
																													 ultraDataColumn10});
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.groupBox4);
			this.ultraTabPageControl1.Controls.Add(this.groupBox2);
			this.ultraTabPageControl1.Controls.Add(this.cmbSubCentro);
			this.ultraTabPageControl1.Controls.Add(this.chkSubCentro);
			this.ultraTabPageControl1.Controls.Add(this.chkCentro);
			this.ultraTabPageControl1.Controls.Add(this.cmbCentro);
			this.ultraTabPageControl1.Controls.Add(this.btNoImporta);
			this.ultraTabPageControl1.Controls.Add(this.ultraButton1);
			this.ultraTabPageControl1.Controls.Add(this.txtReporteCompara);
			this.ultraTabPageControl1.Controls.Add(this.btDetEgresoCot);
			this.ultraTabPageControl1.Controls.Add(this.btEgresoCotiza);
			this.ultraTabPageControl1.Controls.Add(this.groupBox1);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10769);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(660, 413);
			// 
			// groupBox4
			// 
			this.groupBox4.BackColor = System.Drawing.Color.Transparent;
			this.groupBox4.Controls.Add(this.chkArticulo);
			this.groupBox4.Controls.Add(this.cmbArticulo);
			this.groupBox4.Controls.Add(this.cmbHastaFifo);
			this.groupBox4.Controls.Add(this.label4);
			this.groupBox4.Controls.Add(this.label5);
			this.groupBox4.Controls.Add(this.cmbDesdeFifo);
			this.groupBox4.Controls.Add(this.btnImprimeFifo);
			this.groupBox4.Location = new System.Drawing.Point(48, 190);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(448, 94);
			this.groupBox4.TabIndex = 97;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "REPORTE METODO  FIFO";
			this.groupBox4.Visible = false;
			// 
			// chkArticulo
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkArticulo.Appearance = appearance37;
			this.chkArticulo.BackColor = System.Drawing.Color.Transparent;
			this.chkArticulo.Checked = true;
			this.chkArticulo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkArticulo.Location = new System.Drawing.Point(40, 60);
			this.chkArticulo.Name = "chkArticulo";
			this.chkArticulo.Size = new System.Drawing.Size(96, 22);
			this.chkArticulo.TabIndex = 95;
			this.chkArticulo.Text = "Todo Artículo";
			this.chkArticulo.CheckedChanged += new System.EventHandler(this.chkArticulo_CheckedChanged);
			// 
			// cmbArticulo
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArticulo.Appearance = appearance38;
			this.cmbArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo.DataSource = this.cdvArticulo;
			this.cmbArticulo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand6.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn84.Header.VisiblePosition = 0;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn84.Width = 28;
			ultraGridColumn85.Header.VisiblePosition = 1;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn85.Width = 19;
			ultraGridColumn86.Header.VisiblePosition = 2;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn86.Width = 19;
			ultraGridColumn87.Header.VisiblePosition = 3;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn87.Width = 16;
			ultraGridColumn88.Header.VisiblePosition = 4;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn88.Width = 24;
			ultraGridColumn89.Header.VisiblePosition = 6;
			ultraGridColumn89.Width = 8;
			ultraGridColumn90.Header.VisiblePosition = 5;
			ultraGridColumn90.Width = 8;
			ultraGridColumn91.Header.VisiblePosition = 7;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Width = 32;
			ultraGridColumn92.Header.VisiblePosition = 8;
			ultraGridColumn92.Width = 17;
			ultraGridColumn93.Header.VisiblePosition = 9;
			ultraGridColumn93.Width = 8;
			ultraGridColumn94.Header.VisiblePosition = 10;
			ultraGridColumn94.Width = 17;
			ultraGridColumn95.Header.VisiblePosition = 11;
			ultraGridColumn95.Width = 8;
			ultraGridColumn96.Header.VisiblePosition = 12;
			ultraGridColumn96.Width = 8;
			ultraGridColumn97.Header.VisiblePosition = 13;
			ultraGridColumn97.Width = 8;
			ultraGridColumn98.Header.VisiblePosition = 14;
			ultraGridColumn98.Width = 17;
			ultraGridColumn99.Header.VisiblePosition = 15;
			ultraGridColumn99.Width = 8;
			ultraGridColumn100.Header.VisiblePosition = 16;
			ultraGridColumn100.Width = 8;
			ultraGridColumn101.Header.VisiblePosition = 17;
			ultraGridColumn101.Width = 8;
			ultraGridColumn102.Header.VisiblePosition = 18;
			ultraGridColumn102.Width = 8;
			ultraGridColumn103.Header.VisiblePosition = 19;
			ultraGridColumn103.Width = 8;
			ultraGridColumn104.Header.VisiblePosition = 20;
			ultraGridColumn104.Width = 17;
			ultraGridColumn105.Header.VisiblePosition = 21;
			ultraGridColumn105.Width = 17;
			ultraGridColumn106.Header.VisiblePosition = 22;
			ultraGridColumn106.Width = 17;
			ultraGridColumn107.Header.VisiblePosition = 23;
			ultraGridColumn107.Width = 8;
			ultraGridColumn108.Header.VisiblePosition = 24;
			ultraGridColumn108.Width = 8;
			ultraGridColumn109.Header.VisiblePosition = 25;
			ultraGridColumn109.Width = 8;
			ultraGridColumn110.Header.VisiblePosition = 26;
			ultraGridColumn110.Width = 17;
			ultraGridColumn111.Header.VisiblePosition = 27;
			ultraGridColumn111.Width = 8;
			ultraGridColumn112.Header.VisiblePosition = 28;
			ultraGridColumn112.Width = 8;
			ultraGridColumn113.Header.VisiblePosition = 29;
			ultraGridColumn113.Width = 8;
			ultraGridColumn114.Header.VisiblePosition = 30;
			ultraGridColumn114.Width = 8;
			ultraGridColumn115.Header.VisiblePosition = 31;
			ultraGridColumn115.Width = 17;
			ultraGridColumn116.Header.VisiblePosition = 32;
			ultraGridColumn116.Width = 8;
			ultraGridColumn117.Header.VisiblePosition = 33;
			ultraGridColumn117.Width = 8;
			ultraGridColumn118.Header.VisiblePosition = 34;
			ultraGridColumn118.Width = 8;
			ultraGridColumn119.Header.VisiblePosition = 35;
			ultraGridColumn119.Width = 17;
			ultraGridColumn120.Header.VisiblePosition = 36;
			ultraGridColumn120.Width = 8;
			ultraGridColumn121.Header.VisiblePosition = 37;
			ultraGridColumn121.Width = 17;
			ultraGridColumn122.Header.VisiblePosition = 38;
			ultraGridColumn122.Width = 17;
			ultraGridColumn123.Header.VisiblePosition = 39;
			ultraGridColumn123.Width = 17;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn123});
			this.cmbArticulo.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbArticulo.DisplayMember = "Codigo";
			this.cmbArticulo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArticulo.Enabled = false;
			this.cmbArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticulo.Location = new System.Drawing.Point(144, 60);
			this.cmbArticulo.Name = "cmbArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(216, 21);
			this.cmbArticulo.TabIndex = 94;
			this.cmbArticulo.ValueMember = "idArticulo";
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Codigo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			// 
			// cmbHastaFifo
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbHastaFifo.Appearance = appearance39;
			dateButton3.Caption = "Today";
			this.cmbHastaFifo.DateButtons.Add(dateButton3);
			this.cmbHastaFifo.Format = "dd/MMM/yyyy";
			this.cmbHastaFifo.Location = new System.Drawing.Point(248, 17);
			this.cmbHastaFifo.Name = "cmbHastaFifo";
			this.cmbHastaFifo.NonAutoSizeHeight = 24;
			this.cmbHastaFifo.Size = new System.Drawing.Size(112, 21);
			this.cmbHastaFifo.SpinButtonsVisible = true;
			this.cmbHastaFifo.TabIndex = 6;
			this.cmbHastaFifo.Value = new System.DateTime(2005, 12, 24, 0, 0, 0, 0);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(200, 26);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(37, 16);
			this.label4.TabIndex = 8;
			this.label4.Text = "Hasta:";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(32, 26);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(39, 16);
			this.label5.TabIndex = 7;
			this.label5.Text = "Desde:";
			// 
			// cmbDesdeFifo
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDesdeFifo.Appearance = appearance40;
			dateButton4.Caption = "Today";
			this.cmbDesdeFifo.DateButtons.Add(dateButton4);
			this.cmbDesdeFifo.Format = "dd/MMM/yyyy";
			this.cmbDesdeFifo.Location = new System.Drawing.Point(80, 17);
			this.cmbDesdeFifo.Name = "cmbDesdeFifo";
			this.cmbDesdeFifo.NonAutoSizeHeight = 24;
			this.cmbDesdeFifo.Size = new System.Drawing.Size(112, 21);
			this.cmbDesdeFifo.SpinButtonsVisible = true;
			this.cmbDesdeFifo.TabIndex = 3;
			this.cmbDesdeFifo.Value = new System.DateTime(2005, 12, 24, 0, 0, 0, 0);
			// 
			// btnImprimeFifo
			// 
			this.btnImprimeFifo.Location = new System.Drawing.Point(376, 52);
			this.btnImprimeFifo.Name = "btnImprimeFifo";
			this.btnImprimeFifo.Size = new System.Drawing.Size(64, 26);
			this.btnImprimeFifo.TabIndex = 98;
			this.btnImprimeFifo.Text = "Imprimir";
			this.btnImprimeFifo.Click += new System.EventHandler(this.btnImprimeFifo_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.BackColor = System.Drawing.Color.Transparent;
			this.groupBox2.Controls.Add(this.ultraLabel2);
			this.groupBox2.Controls.Add(this.ultraButton2);
			this.groupBox2.Location = new System.Drawing.Point(48, 345);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(448, 51);
			this.groupBox2.TabIndex = 96;
			this.groupBox2.TabStop = false;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel2.Location = new System.Drawing.Point(16, 26);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(184, 17);
			this.ultraLabel2.TabIndex = 95;
			this.ultraLabel2.Text = "Demanda Insatisfecha UMCO";
			// 
			// ultraButton2
			// 
			this.ultraButton2.Location = new System.Drawing.Point(320, 17);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(104, 26);
			this.ultraButton2.TabIndex = 94;
			this.ultraButton2.Text = "Reporte UMCO";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// cmbSubCentro
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubCentro.Appearance = appearance41;
			this.cmbSubCentro.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubCentro.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubCentro.DataSource = this.cdvSubProyecto;
			this.cmbSubCentro.DisplayLayout.AutoFitColumns = true;
			ultraGridBand7.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn124.Header.VisiblePosition = 0;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn124.Width = 25;
			ultraGridColumn125.Header.VisiblePosition = 1;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn125.Width = 19;
			ultraGridColumn126.Header.VisiblePosition = 2;
			ultraGridColumn126.Width = 197;
			ultraGridColumn127.Header.VisiblePosition = 3;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn127.Width = 27;
			ultraGridColumn128.Header.VisiblePosition = 4;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn128.Width = 27;
			ultraGridColumn129.Header.VisiblePosition = 5;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn129.Width = 27;
			ultraGridColumn130.Header.VisiblePosition = 6;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn130.Width = 28;
			ultraGridColumn131.Header.VisiblePosition = 7;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn131.Width = 18;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn124,
																										 ultraGridColumn125,
																										 ultraGridColumn126,
																										 ultraGridColumn127,
																										 ultraGridColumn128,
																										 ultraGridColumn129,
																										 ultraGridColumn130,
																										 ultraGridColumn131});
			this.cmbSubCentro.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbSubCentro.DisplayMember = "Nombre";
			this.cmbSubCentro.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubCentro.Enabled = false;
			this.cmbSubCentro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubCentro.Location = new System.Drawing.Point(416, 78);
			this.cmbSubCentro.Name = "cmbSubCentro";
			this.cmbSubCentro.Size = new System.Drawing.Size(216, 21);
			this.cmbSubCentro.TabIndex = 93;
			this.cmbSubCentro.ValueMember = "idSubProyecto";
			// 
			// cdvSubProyecto
			// 
			this.cdvSubProyecto.BindingContextCtrl = this;
			this.cdvSubProyecto.DataSet = this.cdsProyecto;
			this.cdvSubProyecto.TableName = "";
			this.cdvSubProyecto.TableViewName = "SubProyecto";
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
			// chkSubCentro
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSubCentro.Appearance = appearance42;
			this.chkSubCentro.BackColor = System.Drawing.Color.Transparent;
			this.chkSubCentro.Checked = true;
			this.chkSubCentro.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkSubCentro.Location = new System.Drawing.Point(296, 78);
			this.chkSubCentro.Name = "chkSubCentro";
			this.chkSubCentro.Size = new System.Drawing.Size(96, 21);
			this.chkSubCentro.TabIndex = 92;
			this.chkSubCentro.Text = "Todo Proyecto";
			this.chkSubCentro.CheckedChanged += new System.EventHandler(this.chkSubCentro_CheckedChanged);
			// 
			// chkCentro
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCentro.Appearance = appearance43;
			this.chkCentro.BackColor = System.Drawing.Color.Transparent;
			this.chkCentro.Checked = true;
			this.chkCentro.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCentro.Location = new System.Drawing.Point(296, 52);
			this.chkCentro.Name = "chkCentro";
			this.chkCentro.Size = new System.Drawing.Size(96, 21);
			this.chkCentro.TabIndex = 91;
			this.chkCentro.Text = "Todo Centro";
			this.chkCentro.CheckedChanged += new System.EventHandler(this.chkCentro_CheckedChanged);
			// 
			// cmbCentro
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCentro.Appearance = appearance44;
			this.cmbCentro.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentro.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentro.DataSource = this.cdvProyecto;
			ultraGridBand8.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn132.Header.VisiblePosition = 0;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn133.Header.VisiblePosition = 2;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn134.Header.VisiblePosition = 1;
			ultraGridColumn134.Width = 262;
			ultraGridColumn135.Header.VisiblePosition = 3;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn136.Header.VisiblePosition = 4;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 5;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn138.Header.VisiblePosition = 6;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn139.Header.VisiblePosition = 7;
			ultraGridColumn139.Hidden = true;
			ultraGridColumn140.Header.VisiblePosition = 8;
			ultraGridColumn140.Hidden = true;
			ultraGridColumn141.Header.VisiblePosition = 9;
			ultraGridColumn142.Header.VisiblePosition = 10;
			ultraGridColumn143.Header.VisiblePosition = 11;
			ultraGridColumn144.Header.VisiblePosition = 12;
			ultraGridBand8.Columns.AddRange(new object[] {
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
																										 ultraGridColumn144});
			this.cmbCentro.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbCentro.DisplayMember = "Nombre";
			this.cmbCentro.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentro.Enabled = false;
			this.cmbCentro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentro.Location = new System.Drawing.Point(416, 52);
			this.cmbCentro.Name = "cmbCentro";
			this.cmbCentro.Size = new System.Drawing.Size(218, 21);
			this.cmbCentro.TabIndex = 90;
			this.cmbCentro.ValueMember = "idProyecto";
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// btNoImporta
			// 
			this.btNoImporta.Location = new System.Drawing.Point(56, 138);
			this.btNoImporta.Name = "btNoImporta";
			this.btNoImporta.Size = new System.Drawing.Size(216, 34);
			this.btNoImporta.TabIndex = 89;
			this.btNoImporta.Text = "Totales Req <-> Ing. Egr.";
			this.btNoImporta.Click += new System.EventHandler(this.btNoImporta_Click);
			this.btNoImporta.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btNoImporta_MouseUp);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(480, 155);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(144, 35);
			this.ultraButton1.TabIndex = 3;
			this.ultraButton1.Text = "Gastos";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click_1);
			// 
			// txtReporteCompara
			// 
			this.txtReporteCompara.Location = new System.Drawing.Point(480, 121);
			this.txtReporteCompara.Name = "txtReporteCompara";
			this.txtReporteCompara.Size = new System.Drawing.Size(144, 22);
			this.txtReporteCompara.TabIndex = 2;
			this.txtReporteCompara.Text = "Gastos";
			this.txtReporteCompara.ValueChanged += new System.EventHandler(this.txtReporteCompara_ValueChanged);
			// 
			// btDetEgresoCot
			// 
			this.btDetEgresoCot.Location = new System.Drawing.Point(56, 95);
			this.btDetEgresoCot.Name = "btDetEgresoCot";
			this.btDetEgresoCot.Size = new System.Drawing.Size(216, 34);
			this.btDetEgresoCot.TabIndex = 1;
			this.btDetEgresoCot.Text = "Detalle: Egreso <-> Cotización";
			this.btDetEgresoCot.Click += new System.EventHandler(this.btDetEgresoCot_Click);
			// 
			// btEgresoCotiza
			// 
			this.btEgresoCotiza.Location = new System.Drawing.Point(56, 52);
			this.btEgresoCotiza.Name = "btEgresoCotiza";
			this.btEgresoCotiza.Size = new System.Drawing.Size(216, 34);
			this.btEgresoCotiza.TabIndex = 0;
			this.btEgresoCotiza.Text = "Totales: Egreso <-> Cotización";
			this.btEgresoCotiza.Click += new System.EventHandler(this.btEgresoCotiza_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.BackColor = System.Drawing.Color.Transparent;
			this.groupBox1.Controls.Add(this.ultraLabel1);
			this.groupBox1.Controls.Add(this.btImprimir);
			this.groupBox1.Controls.Add(this.cmbReporte);
			this.groupBox1.Location = new System.Drawing.Point(48, 284);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(448, 52);
			this.groupBox1.TabIndex = 88;
			this.groupBox1.TabStop = false;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Location = new System.Drawing.Point(16, 17);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(136, 17);
			this.ultraLabel1.TabIndex = 88;
			this.ultraLabel1.Text = "Reportes Personalizados ";
			// 
			// btImprimir
			// 
			this.btImprimir.Location = new System.Drawing.Point(344, 17);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(80, 26);
			this.btImprimir.TabIndex = 87;
			this.btImprimir.Text = "Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DataSource = this.cdvReporte;
			ultraGridBand9.AddButtonCaption = "Reporte";
			this.cmbReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbReporte.DisplayMember = "Reporte";
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(168, 17);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(160, 21);
			this.cmbReporte.TabIndex = 86;
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
			this.cdvReporte.RowFilter = "Tipo = 10 And idGrupo = 11";
			this.cdvReporte.TableName = "";
			this.cdvReporte.TableViewName = "Reporte";
			// 
			// cdsReporte
			// 
			this.cdsReporte.BindingContextCtrl = this;
			this.cdsReporte.DataLibrary = "LibEmpresa";
			this.cdsReporte.DataLibraryUrl = "";
			this.cdsReporte.DataSetDef = "dsReporte";
			this.cdsReporte.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsReporte.Name = "cdsReporte";
			this.cdsReporte.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsReporte.SchemaDef = null;
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.groupBox3);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10769);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(660, 413);
			this.ultraTabPageControl3.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl3_Paint);
			// 
			// groupBox3
			// 
			this.groupBox3.BackColor = System.Drawing.Color.Transparent;
			this.groupBox3.Controls.Add(this.chkNumOP);
			this.groupBox3.Controls.Add(this.dtFechaHasta);
			this.groupBox3.Controls.Add(this.dtFechaDesde);
			this.groupBox3.Controls.Add(this.chkResumen);
			this.groupBox3.Controls.Add(this.ultraLabel6);
			this.groupBox3.Controls.Add(this.ultraLabel5);
			this.groupBox3.Controls.Add(this.chkProyecto);
			this.groupBox3.Controls.Add(this.ultraButton3);
			this.groupBox3.Controls.Add(this.cmbProyecto);
			this.groupBox3.Controls.Add(this.ultraLabel3);
			this.groupBox3.Controls.Add(this.txtNumOP);
			this.groupBox3.Location = new System.Drawing.Point(8, 9);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(368, 232);
			this.groupBox3.TabIndex = 98;
			this.groupBox3.TabStop = false;
			// 
			// chkNumOP
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkNumOP.Appearance = appearance45;
			this.chkNumOP.BackColor = System.Drawing.Color.Transparent;
			this.chkNumOP.Checked = true;
			this.chkNumOP.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkNumOP.Location = new System.Drawing.Point(16, 106);
			this.chkNumOP.Name = "chkNumOP";
			this.chkNumOP.Size = new System.Drawing.Size(112, 21);
			this.chkNumOP.TabIndex = 106;
			this.chkNumOP.Text = "Todo Número OP";
			this.chkNumOP.CheckedChanged += new System.EventHandler(this.chkNumOP_CheckedChanged);
			// 
			// dtFechaHasta
			// 
			this.dtFechaHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton5.Caption = "Today";
			this.dtFechaHasta.DateButtons.Add(dateButton5);
			this.dtFechaHasta.Format = "dd/MMM/yyyy";
			this.dtFechaHasta.Location = new System.Drawing.Point(232, 60);
			this.dtFechaHasta.Name = "dtFechaHasta";
			this.dtFechaHasta.NonAutoSizeHeight = 23;
			this.dtFechaHasta.Size = new System.Drawing.Size(112, 21);
			this.dtFechaHasta.SpinButtonsVisible = true;
			this.dtFechaHasta.TabIndex = 105;
			this.dtFechaHasta.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.AllowMonthSelection = true;
			this.dtFechaDesde.AllowWeekSelection = true;
			this.dtFechaDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton6.Caption = "Today";
			this.dtFechaDesde.DateButtons.Add(dateButton6);
			this.dtFechaDesde.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.dtFechaDesde.Format = "dd/MMM/yyyy";
			this.dtFechaDesde.Location = new System.Drawing.Point(56, 60);
			this.dtFechaDesde.Name = "dtFechaDesde";
			this.dtFechaDesde.NonAutoSizeHeight = 23;
			this.dtFechaDesde.Size = new System.Drawing.Size(112, 21);
			this.dtFechaDesde.SpinButtonsVisible = true;
			this.dtFechaDesde.TabIndex = 104;
			this.dtFechaDesde.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// chkResumen
			// 
			this.chkResumen.Location = new System.Drawing.Point(16, 198);
			this.chkResumen.Name = "chkResumen";
			// 
			// chkResumen.Properties
			// 
			this.chkResumen.Properties.Caption = "Resumen";
			this.chkResumen.Size = new System.Drawing.Size(75, 18);
			this.chkResumen.TabIndex = 103;
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel6.Location = new System.Drawing.Point(184, 60);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(40, 18);
			this.ultraLabel6.TabIndex = 102;
			this.ultraLabel6.Text = "Hasta:";
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel5.Location = new System.Drawing.Point(16, 60);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(40, 18);
			this.ultraLabel5.TabIndex = 101;
			this.ultraLabel5.Text = "Desde:";
			// 
			// chkProyecto
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkProyecto.Appearance = appearance46;
			this.chkProyecto.BackColor = System.Drawing.Color.Transparent;
			this.chkProyecto.Checked = true;
			this.chkProyecto.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkProyecto.Location = new System.Drawing.Point(16, 146);
			this.chkProyecto.Name = "chkProyecto";
			this.chkProyecto.Size = new System.Drawing.Size(96, 22);
			this.chkProyecto.TabIndex = 96;
			this.chkProyecto.Text = "Todo Centro";
			this.chkProyecto.CheckedChanged += new System.EventHandler(this.chkProyecto_CheckedChanged);
			// 
			// ultraButton3
			// 
			this.ultraButton3.Location = new System.Drawing.Point(128, 198);
			this.ultraButton3.Name = "ultraButton3";
			this.ultraButton3.Size = new System.Drawing.Size(96, 26);
			this.ultraButton3.TabIndex = 97;
			this.ultraButton3.Text = "Reporte";
			this.ultraButton3.Click += new System.EventHandler(this.ultraButton3_Click);
			// 
			// cmbProyecto
			// 
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance47;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Compra.idProyecto"));
			this.cmbProyecto.DataSource = this.cdvProyecto;
			ultraGridBand10.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn145.Header.VisiblePosition = 0;
			ultraGridColumn145.Hidden = true;
			ultraGridColumn146.Header.VisiblePosition = 2;
			ultraGridColumn146.Hidden = true;
			ultraGridColumn147.Header.VisiblePosition = 1;
			ultraGridColumn147.Width = 262;
			ultraGridColumn148.Header.VisiblePosition = 3;
			ultraGridColumn148.Hidden = true;
			ultraGridColumn149.Header.VisiblePosition = 4;
			ultraGridColumn149.Hidden = true;
			ultraGridColumn150.Header.VisiblePosition = 5;
			ultraGridColumn150.Hidden = true;
			ultraGridColumn151.Header.VisiblePosition = 6;
			ultraGridColumn151.Hidden = true;
			ultraGridColumn152.Header.VisiblePosition = 7;
			ultraGridColumn152.Hidden = true;
			ultraGridColumn153.Header.VisiblePosition = 8;
			ultraGridColumn153.Hidden = true;
			ultraGridColumn154.Header.VisiblePosition = 9;
			ultraGridColumn155.Header.VisiblePosition = 10;
			ultraGridColumn156.Header.VisiblePosition = 11;
			ultraGridColumn157.Header.VisiblePosition = 12;
			ultraGridBand10.Columns.AddRange(new object[] {
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
																											ultraGridColumn157});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Enabled = false;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(128, 138);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(218, 21);
			this.cmbProyecto.TabIndex = 95;
			this.cmbProyecto.ValueMember = "idProyecto";
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
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel3.Location = new System.Drawing.Point(24, 26);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(136, 17);
			this.ultraLabel3.TabIndex = 5;
			this.ultraLabel3.Text = "REPORTE DE COSTOS";
			// 
			// txtNumOP
			// 
			this.txtNumOP.Enabled = false;
			this.txtNumOP.Location = new System.Drawing.Point(128, 103);
			this.txtNumOP.Name = "txtNumOP";
			this.txtNumOP.Size = new System.Drawing.Size(104, 22);
			this.txtNumOP.TabIndex = 99;
			// 
			// cdsGrupo
			// 
			this.cdsGrupo.BindingContextCtrl = this;
			this.cdsGrupo.DataLibrary = "LibFacturacion";
			this.cdsGrupo.DataLibraryUrl = "";
			this.cdsGrupo.DataSetDef = "dsArticuloTabla";
			this.cdsGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupo.Name = "cdsGrupo";
			this.cdsGrupo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupo.SchemaDef = null;
			// 
			// udsAnalisis
			// 
			ultraDataColumn12.DefaultValue = "";
			ultraDataColumn13.DataType = typeof(System.Double);
			ultraDataColumn13.DefaultValue = 0;
			ultraDataColumn14.DataType = typeof(System.Double);
			ultraDataColumn15.DataType = typeof(System.Double);
			ultraDataColumn16.DataType = typeof(System.Double);
			ultraDataColumn17.DataType = typeof(System.Double);
			ultraDataColumn18.Tag = typeof(System.Double);
			ultraDataColumn19.DataType = typeof(System.Double);
			ultraDataColumn20.DataType = typeof(System.Double);
			ultraDataColumn20.DefaultValue = 0;
			ultraDataColumn21.DataType = typeof(int);
			this.udsAnalisis.Band.Columns.AddRange(new object[] {
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
																														ultraDataColumn22});
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// Marco
			// 
			this.Marco.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.Marco.Controls.Add(this.ultraTabSharedControlsPage1);
			this.Marco.Controls.Add(this.marcoGenerar);
			this.Marco.Controls.Add(this.ultraTabPageControl2);
			this.Marco.Controls.Add(this.marcoArticulo);
			this.Marco.Controls.Add(this.marcoClientes);
			this.Marco.Controls.Add(this.ultraTabPageControl1);
			this.Marco.Controls.Add(this.ultraTabPageControl3);
			this.Marco.Location = new System.Drawing.Point(8, 9);
			this.Marco.Name = "Marco";
			this.Marco.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.Marco.Size = new System.Drawing.Size(664, 439);
			this.Marco.TabIndex = 13;
			ultraTab1.TabPage = this.ultraTabPageControl2;
			ultraTab1.Text = "Condiciones";
			ultraTab2.TabPage = this.marcoGenerar;
			ultraTab2.Text = "Generar";
			ultraTab3.TabPage = this.marcoArticulo;
			ultraTab3.Text = "Artículos";
			ultraTab4.TabPage = this.marcoClientes;
			ultraTab4.Text = "Clientes";
			ultraTab5.TabPage = this.ultraTabPageControl1;
			ultraTab5.Text = "Comparación";
			ultraTab6.TabPage = this.ultraTabPageControl3;
			ultraTab6.Text = "STI";
			this.Marco.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																										ultraTab1,
																																										ultraTab2,
																																										ultraTab3,
																																										ultraTab4,
																																										ultraTab5,
																																										ultraTab6});
			this.Marco.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.Marco_SelectedTabChanged);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(660, 413);
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// c1DataSet1
			// 
			this.c1DataSet1.BindingContextCtrl = this;
			this.c1DataSet1.DataLibrary = "LibFacturacion";
			this.c1DataSet1.DataLibraryUrl = "";
			this.c1DataSet1.DataSetDef = "dsCompra";
			this.c1DataSet1.FillOnRequest = false;
			this.c1DataSet1.Locale = new System.Globalization.CultureInfo("es-EC");
			this.c1DataSet1.Name = "c1DataSet1";
			this.c1DataSet1.SchemaClassName = "LibFacturacion.DataClass";
			this.c1DataSet1.SchemaDef = null;
			// 
			// cdvCompra
			// 
			this.cdvCompra.BindingContextCtrl = this;
			this.cdvCompra.DataSet = this.cdsCompra;
			this.cdvCompra.Sort = "idCompra";
			this.cdvCompra.TableName = "";
			this.cdvCompra.TableViewName = "Compra";
			// 
			// errorProvider1
			// 
			this.errorProvider1.ContainerControl = this;
			// 
			// ArticuloAnalisis
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(680, 454);
			this.Controls.Add(this.Marco);
			this.Name = "ArticuloAnalisis";
			this.Text = "Analisis de Articulos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ArticuloAnalisis_Closing);
			this.Load += new System.EventHandler(this.ArticuloAnalisis_Load);
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbClienteSector)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSector)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClienteTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOperador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			this.marcoGenerar.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtTituloGenerar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdGenerar)).EndInit();
			this.marcoArticulo.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtTituloArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsArticulo)).EndInit();
			this.marcoClientes.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtTituloCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCliente)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHastaFifo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesdeFifo)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbSubCentro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReporteCompara)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkResumen.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumOP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAnalisis)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.Marco)).EndInit();
			this.Marco.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.c1DataSet1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BorraErrores()
		{
			errorProvider.SetError(cmbTipo, "");
			errorProvider.SetError(txtValor, "");
			errorProvider.SetError(cmbOperador, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
				if (cmbTipo.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbTipo, "Ingrese Tipo");
					bOk = false;
				}
			if ((int) optFiltro.Value > 0)
			{
				if (cmbOperador.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbOperador, "Ingrese Operador");
					bOk = false;
				}
				try
				{
					double dValor = (double) txtValor.Value;
				}
				catch
				{
					errorProvider.SetError(txtValor, "Ingrese Valor");
					bOk = false;
				}
			}
				return bOk;
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show(
        "Desea Generar movimientos de todos los articulos", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			Cursor = Cursors.WaitCursor;
//			int iNumFilas = -1;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;

			int idCliente = 0;
			if (!chkCliente.Checked) idCliente = (int) cmbCliente.Value;
			string stSelect = "Exec ArticuloAnalisis '" + dtDesde.ToString("yyyyMMdd")
				+ "', '" + dtHasta.ToString("yyyyMMdd") + "', " + idCliente;

			grdGenerar.DataSource = Funcion.dvProcedimiento(cdsArticulo, stSelect);
			#region Borrar
//			udsAnalisis.Rows.Clear();
//			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCliente, stSelect, true);
//
//			try
//			{
//				string stCodAntes = "xxxxx";
//				string stTipoF = "";
//				while (miReader.Read())
//				{
//					string stArticulo = miReader.GetString(0);
//					if (stArticulo.Length == 0) continue;
//					string stCodArt = miReader.GetString(1);
//					double dCant = miReader.GetDouble(2);
//					int ITipoF = miReader.GetInt32(4);
//					if (stCodAntes != stCodArt)
//					{
//            iNumFilas ++;
//						udsAnalisis.Rows.Add();
//						udsAnalisis.Rows[iNumFilas]["Articulo"]=stArticulo;
//						udsAnalisis.Rows[iNumFilas]["CodArticulo"]=stCodArt;
//						int IGrupo = miReader.GetInt32(6);
//						udsAnalisis.Rows[iNumFilas]["Grupo"]=IGrupo;
//						double dPedido = miReader.GetDouble(7);
//						udsAnalisis.Rows[iNumFilas]["Pedidos"]=dPedido;
//						udsAnalisis.Rows[iNumFilas]["Saldo"]=miReader.GetDouble(8);
//						udsAnalisis.Rows[iNumFilas]["Costo"]=miReader.GetDouble(9);
//
//						stCodAntes = stCodArt;
//					}
//					switch(ITipoF)
//					{
//						case 1: 
//							stTipoF = "Ventas";
//							break;
//						case 4:
//							stTipoF = "Compras";
//							break;
//						case 5:
//							stTipoF = "DevVenta";
//							break;
//						case 6:
//							stTipoF = "DevCompra";
//							break;
//						case 8:
//							stTipoF = "Egresos";
//							break;
//						case 9:
//							stTipoF = "Ingresos";
//							break;
//						default:
//							stTipoF = "nada";
//							break;
//					}
//					if (stTipoF != "nada")
//            udsAnalisis.Rows[iNumFilas][stTipoF]= Math.Abs(dCant);
//				}
//			}
//			catch(System.Exception ex)
//			{
//				MessageBox.Show(ex.Message, "Llenar Grilla");
//			}
			#endregion Borrar
//			miReader.Close();
			Cursor = Cursors.Default;
			MessageBox.Show("Movimientos Generados", "Información");
			Marco.SelectedTab = Marco.Tabs[1];
		}

		private string Operador()
		{
			string stOperador = "";
			switch((int) cmbOperador.Value)
			{
				case 0:
					stOperador = ">";
					break;
				case 1:
					stOperador = "=";
					break;
				case 2:
					stOperador = "<";
					break;
				case 3:
					stOperador = ">=";
					break;
				case 4:
					stOperador = "<=";
					break;
				case 5:
					stOperador = "<>";
					break;
			}
			return stOperador;
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			switch( (int) optTipo.Value)
			{
				case 0:
					btGenerar_Click(this, e);
					break;
				case 1:
					btArticulo_Click(this, e);
					break;
				case 2:
					btCliente_Click(this, e);
					break;
				case 3:
				case 4:
				{
					DateTime dtDesde = (DateTime) cmbDesde.Value;
					DateTime dtHasta = (DateTime) cmbHasta.Value;

					ParameterFields paramFields = new ParameterFields ();

					ParameterField paramField = new ParameterField ();
					ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
					paramField.ParameterFieldName = "@FechaIni";
					discreteVal.Value = dtDesde.ToString("yyyy-MM-dd");
					paramField.CurrentValues.Add (discreteVal);
					paramFields.Add (paramField);

					ParameterField paramField1 = new ParameterField ();
					ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
					paramField1.ParameterFieldName = "@FechaFin";
					discreteVal1.Value = dtHasta.ToString("yyyy-MM-dd");
					paramField1.CurrentValues.Add (discreteVal1);
					
					paramFields.Add (paramField1);

					string stReporte = "VentaNetaArticulo.rpt";
					if ((int) optTipo.Value == 4) stReporte = "VentaNetaCliente.rpt";
					Reporte miReporte = new Reporte(stReporte, "");
					miReporte.MdiParent = this.MdiParent;
					//					if (chkResumen.Checked) miReporte.Resumen();
					miReporte.crVista.ParameterFieldInfo = paramFields;
					miReporte.Show();					
				}
					break;
				case 5:
				case 7:
				{
					DateTime dtDesde = (DateTime) cmbDesde.Value;
					DateTime dtHasta = (DateTime) cmbHasta.Value;

					ParameterFields paramFields = new ParameterFields ();

					ParameterField paramField = new ParameterField ();
					ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
					paramField.ParameterFieldName = "@FechaDesde";
					discreteVal.Value = dtDesde.ToString("yyyy-MM-dd");
					paramField.CurrentValues.Add (discreteVal);
					paramFields.Add (paramField);

					ParameterField paramField1 = new ParameterField ();
					ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
					paramField1.ParameterFieldName = "@FechaHasta";
					discreteVal1.Value = dtHasta.AddDays(1).ToString("yyyy-MM-dd");
					paramField1.CurrentValues.Add (discreteVal1);
					
					paramFields.Add (paramField1);

					string stReporte = "InventarioMov.rpt";
					if ((int) optTipo.Value == 7) stReporte = "rptTotalBioderma.rpt";

					Reporte miReporte = new Reporte(stReporte, "");
					miReporte.MdiParent = this.MdiParent;
					//					if (chkResumen.Checked) miReporte.Resumen();
					miReporte.crVista.ParameterFieldInfo = paramFields;
					miReporte.Show();					
				}
					break;
				case 6:
					btCotizacionEgreso_Click(this, e);
					break;
			}
			Cursor = Cursors.Default;
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsNumero_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsNumero.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btArticulo_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;

			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stTitulo= "Articulos " + cmbTipo.Text + " Desde " + 
				dtDesde.ToString("dd/MMM/yyyy") + " hasta " 
				+ dtHasta.ToString("dd/MMM/yyyy");

			dtHasta = dtHasta.AddDays(1);
			string stWhere = "Where Compra.Fecha >= '" + dtDesde.ToString("yyyyMMdd") + "' "
				+ "And Compra.Fecha < '" + dtHasta.ToString("yyyyMMdd") + "' "
				+ " And Compra.idTipoFactura = " + cmbTipo.Value.ToString() + " ";

			if (!chkCliente.Checked) 
			{
				stWhere += "And Compra.idCliente = " + cmbCliente.Value.ToString();
				stTitulo += " Cliente " + cmbCliente.Text;
			}

			if (DialogResult.No == MessageBox.Show(stTitulo,
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;

			Cursor = Cursors.WaitCursor;
			string stSelect ="SELECT Articulo.Articulo, "
				+ "Articulo.Codigo, IsNull(Articulo.Saldo,0),"
				+ "Articulo.idArticulo "
				+ "FROM Articulo";
			if ((int) optFiltro.Value == 1)
				stSelect += " Where Saldo " + Operador() + txtValor.Value.ToString();

			string stBusca ="SELECT sum(ISNULL(DetCompra.Cantidad,0)) "
				+ "FROM DetCompra INNER JOIN Compra "
				+ "ON Compra.idCompra=DetCompra.idCompra "
				+ stWhere+ " And DetCompra.idArticulo = ";

			this.cdsCompra.Schema.Connections[0].Open();
			SqlDataReader miReader;
			C1.Data.SchemaObjects.Connection miConeccion = cdsCompra.Schema.Connections[0];
			SqlCommand cmdCant = new SqlCommand (stSelect
				, (SqlConnection) miConeccion.DbConnection);
			udsArticulo.Rows.Clear();
			int iNumFilas = 0;
			miReader = cmdCant.ExecuteReader();
			try
			{
				while (miReader.Read())
				{
					string stArticulo = miReader.GetString(0);
					string stCodArt = miReader.GetString(1);
					double dSaldo = miReader.GetDouble(2);
					int IdArt = miReader.GetInt32(3);
					udsArticulo.Rows.Add();
					udsArticulo.Rows[iNumFilas]["Cantidad"]=IdArt;
					udsArticulo.Rows[iNumFilas]["Saldo"]=dSaldo;
					udsArticulo.Rows[iNumFilas]["Articulo"]=stArticulo;
					udsArticulo.Rows[iNumFilas]["Codigo"]=stCodArt;
					iNumFilas++;
				}

				miReader.Close();
				this.cdsCompra.Schema.Connections[0].Close();
				for (int i=0; i<udsArticulo.Rows.Count; i++)
				{
					double dCant = Funcion.dEscalarSQL(cdsCompra, 
						stBusca + udsArticulo.Rows[i]["Cantidad"].ToString(), false);
					if ((int) optFiltro.Value == 2 && !VerificaValor(dCant))
						udsArticulo.Rows.RemoveAt(i--);				
					else
            udsArticulo.Rows[i]["Cantidad"]=dCant;
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Llenar Grilla");
			}
			txtTituloArt.Value = stTitulo;
			Cursor = Cursors.Default;
			MessageBox.Show("Articulos Generados", "Información");
			Marco.SelectedTab = Marco.Tabs[2];
		}

		private bool VerificaValor(double dCant)
		{
			bool bOk = true;
			double dValor = (double) txtValor.Value;
			switch((int) cmbOperador.Value)
			{
				case 0:
					if (dCant <= dValor) bOk = false;
					break;
				case 1:
					if (dCant != dValor) bOk = false;
					break;
				case 2:
					if (dCant >= dValor) bOk = false;
					break;
				case 3:
					if (dCant < dValor) bOk = false;
					break;
				case 4:
					if (dCant > dValor) bOk = false;
					break;
				case 5:
					if (dCant != dValor) bOk = false;
					break;
			}
			return bOk;
		}

		private void ArticuloAnalisis_Load(object sender, System.EventArgs e)
		{
			dtFechaDesde.Value= DateTime.Today;
			dtFechaHasta.Value=DateTime.Today;


			cmbDesde.Value = DateTime.Today;
			cmbHasta.Value = DateTime.Today;
			string stAudita = "Exec AuditaCrear 53, 6, 'Art. Analisis'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
			txtOrigen.Width = 0;
			btGenerar.Width = 0;
			btArticulo.Width = 0;
			btCliente.Width = 0;
			cmbTipo.Text = "Venta";
			optTipo.Value = 0;
			
			Marco.Tabs[5].Visible=false;


			if (Funcion.bEjecutaSQL(cdsCliente, "Exec SeteoGExiste 'FTRSTI'")) 				
			{					
				Marco.Tabs[5].Visible=true;
			
			}

			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'ACTSFIFO'")) 				
			{
				groupBox4.Visible=true;
				cdvArticulo.RowFilter = "idTipoGrupo in(1,6)";

			}

		}

		private void btCliente_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;

			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stTitulo= cmbTipo.Text + " a Clientes Desde " + 
				dtDesde.ToString("dd/MMM/yyyy") + " hasta " 
				+ dtHasta.ToString("dd/MMM/yyyy");

			if (DialogResult.No == MessageBox.Show(stTitulo,
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			
			Cursor = Cursors.WaitCursor;

			string stRating = string.Format("Exec RatingCliente '{0}', '{1}', 4, 1",
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));

			SqlDataReader drC = Funcion.rEscalarSQL(cdsCliente, stRating, true);
			udsCliente.Rows.Clear();
			int iNumFilas = 0;
			while (drC.Read())
			{
				udsCliente.Rows.Add();
				udsCliente.Rows[iNumFilas]["Nombre"]=drC.GetString(0);
				udsCliente.Rows[iNumFilas]["Codigo"]=drC.GetString(1);
				udsCliente.Rows[iNumFilas]["Total"]=drC.GetDouble(2);
				udsCliente.Rows[iNumFilas]["Costo"]=drC.GetDouble(3);
				udsCliente.Rows[iNumFilas]["Utilidad"]=drC.GetDouble(4);
				udsCliente.Rows[iNumFilas]["Cantidad"]=drC.GetDouble(5);
				iNumFilas++;
			}
			drC.Close();
			txtTituloCliente.Value = stTitulo;
			Cursor = Cursors.Default;
			Marco.SelectedTab = Marco.Tabs[3];
			MessageBox.Show("Cliente Generados", "Información");
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btExcelGen_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\";
			if (txtTituloGenerar.Text.Trim().Length > 0)
				stDir += txtTituloGenerar.Text + ".xls";
			else
				stDir += "Movimiento.xls";

			ultraGridExcelExporter1.Export(grdGenerar, stDir);
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir);
		}

		private void btExcelArt_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(grdArticulo, "RatingArticulo.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver el Rating de Articulos Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "RatingArticulo.xls");
		}

		private void btExcelCliente_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(grdCliente, "RatingCliente.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver el rating de Clientes Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "RatingCliente.xls");
		}

		private void optTipo_ValueChanged(object sender, System.EventArgs e)
		{
			chkCliente.Enabled = true;
			cmbCliente.Enabled = true;
			cmbTipo.Enabled = false;
			optFiltro.Enabled = false;
			cmbOperador.Enabled = false;
			txtValor.Enabled = false;
			switch( (int) optTipo.Value)
			{
				case 0:
					break;
				case 1:
					chkCliente.Enabled = true;
					cmbCliente.Enabled = true;
					cmbTipo.Enabled = true;
					optFiltro.Enabled = true;
					cmbOperador.Enabled = true;
					txtValor.Enabled = true;
					break;
				case 2:
					cmbTipo.Enabled = true;
					chkCliente.Enabled = true;
					cmbCliente.Enabled = true;
					break;
			}
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			//MessageBox.Show(Funcion.CreaHash(txtOrigen.Text));
			Close();
		}

		private void ArticuloAnalisis_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Art. Analisis'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}

		private void grdGenerar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
			{
				if (grdGenerar.ActiveCell == null) return;
				try
				{
					if (grdGenerar.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
					switch(e.KeyValue)
					{
						case (int) Keys.PageDown:
							if(grdGenerar.ActiveCell.DroppedDown == false)
								grdGenerar.ActiveCell.DroppedDown = true;
							break;

						case (int) Keys.Up:
							if (grdGenerar.ActiveCell.DroppedDown) return;

							grdGenerar.PerformAction(UltraGridAction.ExitEditMode, false, false);
							grdGenerar.PerformAction(UltraGridAction.AboveCell, false, false);
							e.Handled = true;
							grdGenerar.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Left:
							grdGenerar.PerformAction(UltraGridAction.ExitEditMode, false, false);
							grdGenerar.PerformAction(UltraGridAction.PrevCellByTab, false, false);
							e.Handled = true;
							grdGenerar.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Down:
							if (grdGenerar.ActiveCell.DroppedDown) return;
			
							grdGenerar.PerformAction(UltraGridAction.ExitEditMode, false, false);
							grdGenerar.PerformAction(UltraGridAction.BelowCell, false, false);
							e.Handled = true;
							grdGenerar.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Enter:
							grdGenerar.PerformAction(UltraGridAction.ExitEditMode, false, false);
							grdGenerar.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							grdGenerar.PerformAction(UltraGridAction.EnterEditMode, false, false);
							if(grdGenerar.ActiveCell.DroppedDown == false)
								grdGenerar.ActiveCell.DroppedDown = true;
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



		private void grdArticulo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
			{
				if (grdArticulo.ActiveCell == null) return;
				try
				{
					if (grdArticulo.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
					switch(e.KeyValue)
					{
						case (int) Keys.PageDown:
							if(grdArticulo.ActiveCell.DroppedDown == false)
								grdArticulo.ActiveCell.DroppedDown = true;
							break;

						case (int) Keys.Up:
							if (grdArticulo.ActiveCell.DroppedDown) return;

							grdArticulo.PerformAction(UltraGridAction.ExitEditMode, false, false);
							grdArticulo.PerformAction(UltraGridAction.AboveCell, false, false);
							e.Handled = true;
							grdArticulo.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Left:
							grdArticulo.PerformAction(UltraGridAction.ExitEditMode, false, false);
							grdArticulo.PerformAction(UltraGridAction.PrevCellByTab, false, false);
							e.Handled = true;
							grdArticulo.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Down:
							if (grdArticulo.ActiveCell.DroppedDown) return;
			
							grdArticulo.PerformAction(UltraGridAction.ExitEditMode, false, false);
							grdArticulo.PerformAction(UltraGridAction.BelowCell, false, false);
							e.Handled = true;
							grdArticulo.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Enter:
							grdArticulo.PerformAction(UltraGridAction.ExitEditMode, false, false);
							grdArticulo.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							grdArticulo.PerformAction(UltraGridAction.EnterEditMode, false, false);
							if(grdArticulo.ActiveCell.DroppedDown == false)
								grdArticulo.ActiveCell.DroppedDown = true;
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



		private void grdCliente_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
			{
				if (grdCliente.ActiveCell == null) return;
				try
				{
					if (grdCliente.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
					switch(e.KeyValue)
					{
						case (int) Keys.PageDown:
							if(grdCliente.ActiveCell.DroppedDown == false)
								grdCliente.ActiveCell.DroppedDown = true;
							break;

						case (int) Keys.Up:
							if (grdCliente.ActiveCell.DroppedDown) return;

							grdCliente.PerformAction(UltraGridAction.ExitEditMode, false, false);
							grdCliente.PerformAction(UltraGridAction.AboveCell, false, false);
							e.Handled = true;
							grdCliente.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Left:
							grdCliente.PerformAction(UltraGridAction.ExitEditMode, false, false);
							grdCliente.PerformAction(UltraGridAction.PrevCellByTab, false, false);
							e.Handled = true;
							grdCliente.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Down:
							if (grdCliente.ActiveCell.DroppedDown) return;
			
							grdCliente.PerformAction(UltraGridAction.ExitEditMode, false, false);
							grdCliente.PerformAction(UltraGridAction.BelowCell, false, false);
							e.Handled = true;
							grdCliente.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Enter:
							grdCliente.PerformAction(UltraGridAction.ExitEditMode, false, false);
							grdCliente.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							grdCliente.PerformAction(UltraGridAction.EnterEditMode, false, false);
							if(grdCliente.ActiveCell.DroppedDown == false)
								grdCliente.ActiveCell.DroppedDown = true;
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

		private void Marco_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}

		private void btCotizacionEgreso_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;

			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@FechaDesde";
			discreteVal.Value = dtDesde.ToString("yyyy-MM-dd");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@FechaHasta";
			discreteVal1.Value = dtHasta.ToString("yyyy-MM-dd");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);

			ParameterField paramField2 = new ParameterField ();
			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			paramField2.ParameterFieldName = "@FechaHasta";
			discreteVal2.Value = dtHasta.ToString("yyyy-MM-dd");
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);

			ParameterField paramField3 = new ParameterField ();
			ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
			paramField3.ParameterFieldName = "@IdTipo";
			discreteVal3.Value = 0;
			paramField3.CurrentValues.Add (discreteVal3);
			paramFields.Add (paramField3);

			string stReporte = "CotizaEgresoCompara.rpt";
			if ((int) optTipo.Value == 4) stReporte = "VentaNetaCliente.rpt";
			Reporte miReporte = new Reporte(stReporte, "");
			miReporte.MdiParent = this.MdiParent;
			//					if (chkResumen.Checked) miReporte.Resumen();
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Show();					
		}

		private void btEgresoCotiza_Click(object sender, System.EventArgs e)
		{
			ImprimeReporteNombre("CotizaEgresoCompTotal.rpt");
		}

		private void ImprimeReporteNombre(string stNombre)
		{
			DateTime dtDesde;
			DateTime dtHasta;
			
			dtDesde = (DateTime) cmbDesde.Value;
			dtHasta = (DateTime) cmbHasta.Value;

			if (stNombre.StartsWith("STICostos"))
			{
				dtDesde = (DateTime) dtFechaDesde.Value;
				dtHasta = (DateTime) dtFechaHasta.Value;
			}	
			if (stNombre.ToUpper().StartsWith("KARDEXFIFO"))
			{
				dtDesde = (DateTime) cmbDesdeFifo.Value;
				dtHasta = (DateTime) cmbHastaFifo.Value;
			}	
			
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@FechaDesde";
			discreteVal.Value = dtDesde.ToString("yyyy-MM-dd");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@FechaHasta";
			discreteVal1.Value = dtHasta.ToString("yyyy-MM-dd");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);

			if (stNombre.StartsWith("STICostos"))
			{
				ParameterField paramField5 = new ParameterField ();
				ParameterDiscreteValue discreteVal5 = new ParameterDiscreteValue ();
				paramField5.ParameterFieldName = "@idProyecto";
				if (chkProyecto.Checked)
					discreteVal5.Value = (int) 0;
				else
					discreteVal5.Value = (int) cmbProyecto.Value;
				paramField5.CurrentValues.Add (discreteVal5);
				paramFields.Add (paramField5);
				
				////////////////////////////////////////////////
				
				ParameterField paramFieldOP = new ParameterField ();
				ParameterDiscreteValue discreteValOP = new ParameterDiscreteValue ();
				paramFieldOP.ParameterFieldName = "@NumOP";
				if (chkNumOP.Checked)
					discreteValOP.Value = 0;
				else
					discreteValOP.Value = (string) txtNumOP.Value;
				paramFieldOP.CurrentValues.Add (discreteValOP);
				paramFields.Add (paramFieldOP);				

			}

			if (stNombre.ToUpper().StartsWith("GASTO"))
			{
				ParameterField paramFieldCl = new ParameterField ();
				ParameterDiscreteValue discreteValCl = new ParameterDiscreteValue ();
				paramFieldCl.ParameterFieldName = "@idCliente";
				if (chkCliente.Checked)
					discreteValCl.Value = 0;
				else
					discreteValCl.Value = (int) cmbCliente.Value;
				paramFieldCl.CurrentValues.Add (discreteValCl);
				paramFields.Add (paramFieldCl);
			}

			if (stNombre.ToUpper().StartsWith("GASTO"))
			{
				ParameterField paramFieldCl = new ParameterField ();
				ParameterDiscreteValue discreteValCl = new ParameterDiscreteValue ();
				paramFieldCl.ParameterFieldName = "@idCliente";
				if (chkCliente.Checked)
					discreteValCl.Value = 0;
				else
          discreteValCl.Value = (int) cmbCliente.Value;
				paramFieldCl.CurrentValues.Add (discreteValCl);
				paramFields.Add (paramFieldCl);

				if (Funcion.bEjecutaSQL(cdsClienteTabla,"SeteoGExiste 'FVASR'"))
				{
					ParameterField paramField3 = new ParameterField ();
					ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
					paramField3.ParameterFieldName = "@idSubProyecto";
					if (chkSubCentro.Checked)
						discreteVal3.Value = (int) 0;
					else
						discreteVal3.Value = (int) cmbSubCentro.Value;
					paramField3.CurrentValues.Add (discreteVal3);
					paramFields.Add (paramField3);			

				}
			}

			
			if (stNombre.StartsWith("CotizaReqEgreso"))
			{
				ParameterField paramField2 = new ParameterField ();
				ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
				paramField2.ParameterFieldName = "@idProyecto";
				if (chkCentro.Checked)
					discreteVal2.Value = (int) 0;
				else
					discreteVal2.Value = (int) cmbCentro.Value;
				paramField2.CurrentValues.Add (discreteVal2);
				paramFields.Add (paramField2);

				ParameterField paramField3 = new ParameterField ();
				ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
				paramField3.ParameterFieldName = "@idSubProyecto";
				if (chkSubCentro.Checked)
					discreteVal3.Value = (int) 0;
				else
					discreteVal3.Value = (int) cmbSubCentro.Value;
				paramField3.CurrentValues.Add (discreteVal3);
				paramFields.Add (paramField3);			

			}
			if (stNombre.ToUpper().StartsWith("KARDEXFIFO"))
			{
				
				ParameterField paramField2 = new ParameterField ();
				ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
				paramField2.ParameterFieldName = "@vidArticulo";
				if (chkArticulo.Checked)
					discreteVal2.Value = (int) 0;
				else
					discreteVal2.Value = (int) cmbArticulo.Value;
				paramField2.CurrentValues.Add (discreteVal2);
				paramFields.Add (paramField2);

			}

			string stReporte = stNombre;
			if (stReporte.Trim().Length == 0) stReporte = txtReporteCompara.Text.Trim() + ".rpt";
			Reporte miReporte = new Reporte(stReporte, "");
			miReporte.MdiParent = this.MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			if (stNombre.StartsWith("STICostos"))
			{
				if (chkResumen.Checked) miReporte.Resumen();
			}
			miReporte.Show();					
		}

		private void btDetEgresoCot_Click(object sender, System.EventArgs e)
		{
			ImprimeReporteNombre("CotizaEgresoCompDetalle.rpt");
		}

		private void ultraButton1_Click_1(object sender, System.EventArgs e)
		{
			string stReporte = txtReporteCompara.Value.ToString() + ".rpt";
			ImprimeReporteNombre(stReporte);
		}

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbReporte, "");

			if (cmbReporte.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(cmbReporte, "Ingrese Reporte");
				return;
			}
			Cursor = Cursors.WaitCursor;
			string stReporte = cmbReporte.Text.ToString() + ".rpt";
			ImprimeReporteNombre(stReporte);		
			Cursor = Cursors.Default;
		}

		private void cdsClienteTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsClienteTabla.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void txtReporteCompara_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btNoImporta_Click(object sender, System.EventArgs e)
		{
			ImprimeReporteNombre("CotizaReqEgreso.rpt");
		}

		private void chkSubCentro_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbSubCentro.Enabled = !this.chkSubCentro.Checked;
		}

		private void chkCentro_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbCentro.Enabled = !this.chkCentro.Checked;
		}

		private void btNoImporta_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			ImprimeReporteNombre("CotizaReqEgresoDetalle.Rpt");
		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("713", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene Acceso a Demanda Insatisfecha....",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			Cursor = Cursors.WaitCursor;
			DemandInsatisfUmco miDemanda = new DemandInsatisfUmco();
			miDemanda.MdiParent = this.MdiParent;
			miDemanda.Show();
			Cursor = Cursors.Default;
			
	
		}

		

		private void ultraButton3_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;		
			
			
			string stSelect = "SELECT Count(c.idCompra) FROM Compra c LEFT JOIN detCompra dc on dc.idcompra = c.idcompra WHERE c.borrar = 0 and c.idtipoFactura in (1,8,9) and dc.RefNumero= '"+txtNumOP.Text +"'";

 			int Cuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);

			this.errorProvider1.SetError(this.txtNumOP,"");
			if (Cuenta <= 0)
			{
				this.errorProvider1.SetError(this.txtNumOP, "Numero No Existe");
				this.Cursor = Cursors.Arrow;			
				return;
			}		
			ImprimeReporteNombre("STICostos");
			this.Cursor = Cursors.Arrow;
		}

		private void chkNumOP_CheckedChanged(object sender, System.EventArgs e)
		{
			this.txtNumOP.Enabled = !this.chkNumOP.Checked;
		}

		private void chkProyecto_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbProyecto.Enabled = !this.chkProyecto.Checked;

		}

		private void ultraTabPageControl3_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void chkArticulo_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbArticulo.Enabled = !this.chkArticulo.Checked;
		}

		private void btnImprimeFifo_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ImprimeReporteNombre("KARDEXFIFO");
			Cursor = Cursors.Default;
		}

		private void optFiltro_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		


	}
}
