using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de GuiaRemisionFact.
	/// </summary>
	public class GuiaRemisionFact : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvGuiaRemisionG;
		private DevExpress.XtraEditors.CheckEdit chkCliente;
		private DevExpress.XtraEditors.CheckEdit chkFecha;
		private DevExpress.XtraEditors.DateEdit dtFechaDesde;
		private DevExpress.XtraEditors.DateEdit dtFechaHasta;
		private DevExpress.XtraEditors.CheckEdit chkGuia;
		private DevExpress.XtraEditors.CheckEdit chkRuta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPuntoPartida;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPuntoLlegada;
		private C1.Data.C1DataView cdvGuiaPunto;
		private C1.Data.C1DataSet cdsGuiaPunto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGuia1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGuia2;
		private C1.Data.C1DataSet cdsGuiaRemisionG;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPunto;
		private C1.Data.C1DataSet cdsGuiaRemisionPadre;
		private C1.Data.C1DataView cdvGuiaRemisionPadre;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidGuiaRemisionPadre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumero;

		C1.Data.C1DataRow drGuiaPadre;
		private Infragistics.Win.Misc.UltraButton btnAsiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private DevExpress.XtraEditors.DateEdit dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbLlegada;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbChofer;
		private DevExpress.XtraEditors.CheckEdit chkChofer;
		private C1.Data.C1DataView cdvChofer;
		int editar=0;
		public GuiaRemisionFact()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ClteCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointAutorizado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucionFinanaciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idObservaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HaceRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaConsignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito2");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionPunto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPunto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionPunto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPunto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionG", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fInicioTraslado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fTermTraslado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NroGuia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("venta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("compra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("transformacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("consignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("traslado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("devolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("importacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("exportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("otros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fEmision2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPuntoPartida", -1, "cmbPunto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("rucCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPuntoLlegada", -1, "cmbPunto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nombreTrasnp");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("rucTransp");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("diesel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("adelanto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("contratado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransportista");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudadEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCarro1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCarro2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCarro3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TonelajePedido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CombAutorizado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CombConsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Placa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturado", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TextoTonelaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPadreC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreClienteD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RucClienteD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClienteD");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionPunto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPunto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ClteCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointAutorizado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucionFinanaciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idObservaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HaceRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaConsignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito2");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionPunto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPunto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ClteCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointAutorizado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucionFinanaciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idObservaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HaceRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaConsignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito2");
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cdvGuiaRemisionG = new C1.Data.C1DataView();
			this.chkGuia = new DevExpress.XtraEditors.CheckEdit();
			this.chkFecha = new DevExpress.XtraEditors.CheckEdit();
			this.chkCliente = new DevExpress.XtraEditors.CheckEdit();
			this.dtFechaHasta = new DevExpress.XtraEditors.DateEdit();
			this.dtFechaDesde = new DevExpress.XtraEditors.DateEdit();
			this.chkRuta = new DevExpress.XtraEditors.CheckEdit();
			this.cmbPuntoPartida = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsGuiaPunto = new C1.Data.C1DataSet();
			this.cmbPuntoLlegada = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvGuiaPunto = new C1.Data.C1DataView();
			this.txtGuia1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGuia2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsGuiaRemisionG = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsGuiaRemisionPadre = new C1.Data.C1DataSet();
			this.cmbPunto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvGuiaRemisionPadre = new C1.Data.C1DataView();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtidGuiaRemisionPadre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnAsiento = new Infragistics.Win.Misc.UltraButton();
			this.cmbCliente1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.dtFecha = new DevExpress.XtraEditors.DateEdit();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbLlegada = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.chkChofer = new DevExpress.XtraEditors.CheckEdit();
			this.cmbChofer = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvChofer = new C1.Data.C1DataView();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaRemisionG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGuia.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkFecha.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkRuta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPuntoPartida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaPunto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPuntoLlegada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaPunto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGuia1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGuia2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaRemisionG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaRemisionPadre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPunto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaRemisionPadre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidGuiaRemisionPadre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLlegada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkChofer.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbChofer)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvChofer)).BeginInit();
			this.SuspendLayout();
			// 
			// btFiltro
			// 
			appearance1.Image = 5;
			this.btFiltro.Appearance = appearance1;
			this.btFiltro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFiltro.Location = new System.Drawing.Point(32, 121);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(48, 25);
			this.btFiltro.TabIndex = 10;
			this.btFiltro.Text = "Filtrar";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataSource = this.cdvCliente;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Width = 137;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 322;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
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
																										 ultraGridColumn24});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(144, 9);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(232, 21);
			this.cmbCliente.TabIndex = 8;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.RowFilter = "proveedor=0  or ambos=1";
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.CaseSensitive = false;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cdvGuiaRemisionG
			// 
			this.cdvGuiaRemisionG.BindingContextCtrl = this;
			this.cdvGuiaRemisionG.DataSet = null;
			this.cdvGuiaRemisionG.TableName = "";
			this.cdvGuiaRemisionG.TableViewName = "GuiaRemisionG";
			// 
			// chkGuia
			// 
			this.chkGuia.EditValue = true;
			this.chkGuia.Location = new System.Drawing.Point(32, 60);
			this.chkGuia.Name = "chkGuia";
			// 
			// chkGuia.Properties
			// 
			this.chkGuia.Properties.Caption = "Toda Guia";
			this.chkGuia.Size = new System.Drawing.Size(88, 19);
			this.chkGuia.TabIndex = 72;
			this.chkGuia.CheckedChanged += new System.EventHandler(this.chkGuia_CheckedChanged);
			// 
			// chkFecha
			// 
			this.chkFecha.EditValue = true;
			this.chkFecha.Location = new System.Drawing.Point(32, 34);
			this.chkFecha.Name = "chkFecha";
			// 
			// chkFecha.Properties
			// 
			this.chkFecha.Properties.Caption = "Toda Fecha";
			this.chkFecha.Size = new System.Drawing.Size(88, 19);
			this.chkFecha.TabIndex = 70;
			this.chkFecha.CheckedChanged += new System.EventHandler(this.chkFecha_CheckedChanged);
			// 
			// chkCliente
			// 
			this.chkCliente.Location = new System.Drawing.Point(32, 9);
			this.chkCliente.Name = "chkCliente";
			// 
			// chkCliente.Properties
			// 
			this.chkCliente.Properties.Caption = "Todo Cliente";
			this.chkCliente.Size = new System.Drawing.Size(96, 19);
			this.chkCliente.TabIndex = 68;
			this.chkCliente.Validated += new System.EventHandler(this.chkCliente_Validated);
			this.chkCliente.CheckedChanged += new System.EventHandler(this.chkCliente_CheckedChanged);
			// 
			// dtFechaHasta
			// 
			this.dtFechaHasta.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaHasta.Location = new System.Drawing.Point(280, 34);
			this.dtFechaHasta.Name = "dtFechaHasta";
			// 
			// dtFechaHasta.Properties
			// 
			this.dtFechaHasta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaHasta.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFechaHasta.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaHasta.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFechaHasta.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaHasta.Size = new System.Drawing.Size(96, 20);
			this.dtFechaHasta.TabIndex = 75;
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaDesde.Location = new System.Drawing.Point(144, 34);
			this.dtFechaDesde.Name = "dtFechaDesde";
			// 
			// dtFechaDesde.Properties
			// 
			this.dtFechaDesde.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaDesde.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFechaDesde.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaDesde.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFechaDesde.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaDesde.Size = new System.Drawing.Size(96, 20);
			this.dtFechaDesde.TabIndex = 76;
			// 
			// chkRuta
			// 
			this.chkRuta.EditValue = true;
			this.chkRuta.Location = new System.Drawing.Point(32, 86);
			this.chkRuta.Name = "chkRuta";
			// 
			// chkRuta.Properties
			// 
			this.chkRuta.Properties.Caption = "Toda Ruta";
			this.chkRuta.Size = new System.Drawing.Size(112, 19);
			this.chkRuta.TabIndex = 77;
			this.chkRuta.CheckedChanged += new System.EventHandler(this.chkRuta_CheckedChanged);
			// 
			// cmbPuntoPartida
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPuntoPartida.Appearance = appearance2;
			this.cmbPuntoPartida.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPuntoPartida.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPuntoPartida.DataSource = this.cdsGuiaPunto;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbPuntoPartida.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbPuntoPartida.DisplayMember = "Nombre";
			this.cmbPuntoPartida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPuntoPartida.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPuntoPartida.Enabled = false;
			this.cmbPuntoPartida.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPuntoPartida.Location = new System.Drawing.Point(144, 86);
			this.cmbPuntoPartida.Name = "cmbPuntoPartida";
			this.cmbPuntoPartida.Size = new System.Drawing.Size(184, 21);
			this.cmbPuntoPartida.TabIndex = 79;
			this.cmbPuntoPartida.ValueMember = "idGuiaRemisionPunto";
			// 
			// cdsGuiaPunto
			// 
			this.cdsGuiaPunto.BindingContextCtrl = this;
			this.cdsGuiaPunto.DataLibrary = "LibFacturacion";
			this.cdsGuiaPunto.DataLibraryUrl = "";
			this.cdsGuiaPunto.DataSetDef = "dsGuiaRemisionPunto";
			this.cdsGuiaPunto.EnforceConstraints = false;
			this.cdsGuiaPunto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGuiaPunto.Name = "cdsGuiaPunto";
			this.cdsGuiaPunto.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGuiaPunto.SchemaDef = null;
			// 
			// cmbPuntoLlegada
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPuntoLlegada.Appearance = appearance3;
			this.cmbPuntoLlegada.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPuntoLlegada.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPuntoLlegada.DataSource = this.cdsGuiaPunto;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbPuntoLlegada.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPuntoLlegada.DisplayMember = "Nombre";
			this.cmbPuntoLlegada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPuntoLlegada.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPuntoLlegada.Enabled = false;
			this.cmbPuntoLlegada.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPuntoLlegada.Location = new System.Drawing.Point(352, 86);
			this.cmbPuntoLlegada.Name = "cmbPuntoLlegada";
			this.cmbPuntoLlegada.Size = new System.Drawing.Size(192, 21);
			this.cmbPuntoLlegada.TabIndex = 78;
			this.cmbPuntoLlegada.ValueMember = "idGuiaRemisionPunto";
			// 
			// cdvGuiaPunto
			// 
			this.cdvGuiaPunto.BindingContextCtrl = this;
			this.cdvGuiaPunto.DataSet = this.cdsGuiaPunto;
			this.cdvGuiaPunto.Sort = "nombre";
			this.cdvGuiaPunto.TableName = "";
			this.cdvGuiaPunto.TableViewName = "GuiaRemisionPunto";
			// 
			// txtGuia1
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGuia1.Appearance = appearance4;
			this.txtGuia1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGuia1.Enabled = false;
			this.txtGuia1.FormatString = "";
			this.txtGuia1.Location = new System.Drawing.Point(144, 60);
			this.txtGuia1.MaskInput = "nnnnnnnnn";
			this.txtGuia1.Name = "txtGuia1";
			this.txtGuia1.PromptChar = ' ';
			this.txtGuia1.Size = new System.Drawing.Size(96, 22);
			this.txtGuia1.TabIndex = 120;
			// 
			// txtGuia2
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGuia2.Appearance = appearance5;
			this.txtGuia2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGuia2.Enabled = false;
			this.txtGuia2.FormatString = "";
			this.txtGuia2.Location = new System.Drawing.Point(280, 60);
			this.txtGuia2.MaskInput = "nnnnnnnnn";
			this.txtGuia2.Name = "txtGuia2";
			this.txtGuia2.PromptChar = ' ';
			this.txtGuia2.Size = new System.Drawing.Size(96, 22);
			this.txtGuia2.TabIndex = 122;
			// 
			// cdsGuiaRemisionG
			// 
			this.cdsGuiaRemisionG.BindingContextCtrl = this;
			this.cdsGuiaRemisionG.DataLibrary = "LibFacturacion";
			this.cdsGuiaRemisionG.DataLibraryUrl = "";
			this.cdsGuiaRemisionG.DataSetDef = "dsGuiaRemisionG";
			this.cdsGuiaRemisionG.EnforceConstraints = false;
			this.cdsGuiaRemisionG.FillOnRequest = false;
			this.cdsGuiaRemisionG.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGuiaRemisionG.Name = "cdsGuiaRemisionG";
			this.cdsGuiaRemisionG.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGuiaRemisionG.SchemaDef = null;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "GuiaRemisionG";
			this.ultraGrid1.DataSource = this.cdsGuiaRemisionPadre;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "_GuiaRemisionDetalle";
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 82;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 16;
			ultraGridColumn31.Header.VisiblePosition = 10;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 16;
			ultraGridColumn32.Header.Caption = "Emision";
			ultraGridColumn32.Header.VisiblePosition = 13;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 78;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn33.Header.Caption = "Guia";
			ultraGridColumn33.Header.VisiblePosition = 5;
			ultraGridColumn33.Width = 31;
			ultraGridColumn34.Header.VisiblePosition = 11;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 16;
			ultraGridColumn35.Header.VisiblePosition = 12;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 16;
			ultraGridColumn36.Header.VisiblePosition = 14;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 16;
			ultraGridColumn37.Header.VisiblePosition = 15;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 16;
			ultraGridColumn38.Header.VisiblePosition = 16;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 16;
			ultraGridColumn39.Header.VisiblePosition = 17;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 16;
			ultraGridColumn40.Header.VisiblePosition = 18;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 16;
			ultraGridColumn41.Header.VisiblePosition = 19;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 16;
			ultraGridColumn42.Header.VisiblePosition = 20;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 16;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn43.Format = "";
			ultraGridColumn43.Header.Caption = "Emision";
			ultraGridColumn43.Header.VisiblePosition = 4;
			ultraGridColumn43.Width = 48;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn44.Header.Caption = "Partida";
			ultraGridColumn44.Header.VisiblePosition = 7;
			ultraGridColumn44.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridColumn44.Width = 46;
			ultraGridColumn45.Header.VisiblePosition = 40;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 92;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn46.Header.Caption = "Llegada";
			ultraGridColumn46.Header.VisiblePosition = 8;
			ultraGridColumn46.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridColumn46.Width = 60;
			ultraGridColumn47.Header.VisiblePosition = 21;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 13;
			ultraGridColumn48.Header.VisiblePosition = 22;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 16;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn49.Header.Caption = "Observaciones";
			ultraGridColumn49.Header.VisiblePosition = 9;
			ultraGridColumn49.Width = 170;
			ultraGridColumn50.Header.VisiblePosition = 23;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 12;
			ultraGridColumn51.Header.VisiblePosition = 24;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 12;
			ultraGridColumn52.Header.VisiblePosition = 25;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 16;
			ultraGridColumn53.Header.VisiblePosition = 26;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 14;
			ultraGridColumn54.Header.VisiblePosition = 27;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 16;
			ultraGridColumn55.Header.VisiblePosition = 28;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 16;
			ultraGridColumn56.Header.VisiblePosition = 29;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 9;
			ultraGridColumn57.Header.VisiblePosition = 30;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 9;
			ultraGridColumn58.Header.VisiblePosition = 31;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 9;
			ultraGridColumn59.Header.VisiblePosition = 32;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 9;
			ultraGridColumn60.Header.Caption = "Cliente";
			ultraGridColumn60.Header.VisiblePosition = 6;
			ultraGridColumn60.Width = 73;
			ultraGridColumn61.Header.VisiblePosition = 33;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn61.Width = 15;
			ultraGridColumn62.Header.VisiblePosition = 34;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 16;
			ultraGridColumn63.Header.VisiblePosition = 35;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Width = 17;
			ultraGridColumn64.Header.VisiblePosition = 36;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 16;
			ultraGridColumn65.Header.VisiblePosition = 37;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 16;
			ultraGridColumn66.Header.Caption = "Fact";
			ultraGridColumn66.Header.VisiblePosition = 3;
			ultraGridColumn66.Width = 29;
			ultraGridColumn67.Header.VisiblePosition = 38;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn67.Width = 16;
			ultraGridColumn68.Header.VisiblePosition = 39;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 12;
			ultraGridColumn69.Header.VisiblePosition = 41;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.Width = 11;
			ultraGridColumn70.Header.VisiblePosition = 2;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 80;
			ultraGridColumn71.Header.VisiblePosition = 42;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn71.Width = 77;
			ultraGridColumn72.Header.VisiblePosition = 43;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn72.Width = 80;
			ultraGridColumn73.Header.VisiblePosition = 44;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn73.Width = 116;
			ultraGridColumn74.Header.VisiblePosition = 45;
			ultraGridColumn74.Width = 68;
			ultraGridColumn75.Header.VisiblePosition = 46;
			ultraGridColumn75.Width = 74;
			ultraGridColumn76.Header.VisiblePosition = 47;
			ultraGridColumn76.Width = 68;
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
																										 ultraGridColumn76});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 207);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(688, 163);
			this.ultraGrid1.TabIndex = 125;
			// 
			// cdsGuiaRemisionPadre
			// 
			this.cdsGuiaRemisionPadre.BindingContextCtrl = this;
			this.cdsGuiaRemisionPadre.DataLibrary = "LibFacturacion";
			this.cdsGuiaRemisionPadre.DataLibraryUrl = "";
			this.cdsGuiaRemisionPadre.DataSetDef = "dsGuiaRemisionPadre";
			this.cdsGuiaRemisionPadre.EnforceConstraints = false;
			this.cdsGuiaRemisionPadre.FillOnRequest = false;
			this.cdsGuiaRemisionPadre.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGuiaRemisionPadre.Name = "cdsGuiaRemisionPadre";
			this.cdsGuiaRemisionPadre.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGuiaRemisionPadre.SchemaDef = null;
			// 
			// cmbPunto
			// 
			this.cmbPunto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPunto.DataMember = "GuiaRemisionPunto";
			this.cmbPunto.DataSource = this.cdsGuiaPunto;
			ultraGridBand5.AddButtonCaption = "CompraComprobante";
			ultraGridColumn77.Header.VisiblePosition = 0;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 1;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn77,
																										 ultraGridColumn78});
			this.cmbPunto.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbPunto.DisplayMember = "Nombre";
			this.cmbPunto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPunto.Location = new System.Drawing.Point(304, 250);
			this.cmbPunto.Name = "cmbPunto";
			this.cmbPunto.Size = new System.Drawing.Size(136, 77);
			this.cmbPunto.TabIndex = 126;
			this.cmbPunto.ValueMember = "idGuiaRemisionPunto";
			this.cmbPunto.Visible = false;
			// 
			// cdvGuiaRemisionPadre
			// 
			this.cdvGuiaRemisionPadre.BindingContextCtrl = this;
			this.cdvGuiaRemisionPadre.DataSet = this.cdsGuiaRemisionPadre;
			this.cdvGuiaRemisionPadre.RowFilter = "tipo=1";
			this.cdvGuiaRemisionPadre.TableName = "";
			this.cdvGuiaRemisionPadre.TableViewName = "GuiaRemisionPadre";
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_GuiaRemisionPadre";
			this.barraDato1.DataNombreId = "idGuiaRemisionPadre";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsGuiaRemisionPadre;
			this.barraDato1.DataTabla = "GuiaRemisionPadre";
			this.barraDato1.DataTablaHija = "GuiaRemisionG";
			this.barraDato1.Location = new System.Drawing.Point(192, 388);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 127;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtidGuiaRemisionPadre
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidGuiaRemisionPadre.Appearance = appearance11;
			this.txtidGuiaRemisionPadre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionPadre, "_GuiaRemisionPadre.idGuiaRemisionPadre"));
			this.txtidGuiaRemisionPadre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidGuiaRemisionPadre.Enabled = false;
			this.txtidGuiaRemisionPadre.FormatString = "";
			this.txtidGuiaRemisionPadre.Location = new System.Drawing.Point(40, 155);
			this.txtidGuiaRemisionPadre.MaskInput = "nnnnnnnnn";
			this.txtidGuiaRemisionPadre.Name = "txtidGuiaRemisionPadre";
			this.txtidGuiaRemisionPadre.PromptChar = ' ';
			this.txtidGuiaRemisionPadre.Size = new System.Drawing.Size(48, 22);
			this.txtidGuiaRemisionPadre.TabIndex = 128;
			this.txtidGuiaRemisionPadre.ValueChanged += new System.EventHandler(this.ultraNumericEditor2_ValueChanged);
			// 
			// txtNumero
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance12;
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionPadre, "_GuiaRemisionPadre.Numero"));
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.FormatString = "";
			this.txtNumero.Location = new System.Drawing.Point(208, 146);
			this.txtNumero.MaskInput = "nnnnnnnnn";
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.PromptChar = ' ';
			this.txtNumero.Size = new System.Drawing.Size(96, 22);
			this.txtNumero.TabIndex = 129;
			this.txtNumero.ValueChanged += new System.EventHandler(this.ultraNumericEditor3_ValueChanged);
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance13.Image = 5;
			this.btnAsiento.Appearance = appearance13;
			this.btnAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnAsiento.Location = new System.Drawing.Point(504, 388);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(56, 26);
			this.btnAsiento.TabIndex = 132;
			this.btnAsiento.Text = "Venta";
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// cmbCliente1
			// 
			this.cmbCliente1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionPadre, "_GuiaRemisionPadre.idCliente"));
			this.cmbCliente1.DataSource = this.cdvCliente;
			ultraGridColumn79.Header.VisiblePosition = 1;
			ultraGridColumn79.Width = 137;
			ultraGridColumn80.Header.VisiblePosition = 3;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 0;
			ultraGridColumn81.Width = 322;
			ultraGridColumn82.Header.VisiblePosition = 4;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 5;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 2;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 6;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 7;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 8;
			ultraGridColumn88.Header.VisiblePosition = 9;
			ultraGridColumn89.Header.VisiblePosition = 10;
			ultraGridColumn90.Header.VisiblePosition = 11;
			ultraGridColumn91.Header.VisiblePosition = 12;
			ultraGridColumn92.Header.VisiblePosition = 13;
			ultraGridColumn93.Header.VisiblePosition = 14;
			ultraGridColumn94.Header.VisiblePosition = 15;
			ultraGridColumn95.Header.VisiblePosition = 16;
			ultraGridColumn96.Header.VisiblePosition = 17;
			ultraGridColumn97.Header.VisiblePosition = 18;
			ultraGridColumn98.Header.VisiblePosition = 19;
			ultraGridColumn99.Header.VisiblePosition = 20;
			ultraGridColumn100.Header.VisiblePosition = 21;
			ultraGridColumn101.Header.VisiblePosition = 22;
			ultraGridColumn102.Header.VisiblePosition = 23;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn102});
			this.cmbCliente1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCliente1.DisplayMember = "Nombre";
			this.cmbCliente1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente1.Location = new System.Drawing.Point(144, 172);
			this.cmbCliente1.Name = "cmbCliente1";
			this.cmbCliente1.Size = new System.Drawing.Size(400, 21);
			this.cmbCliente1.TabIndex = 133;
			this.cmbCliente1.ValueMember = "idCliente";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(400, 146);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 136;
			this.label1.Text = "Fecha";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(144, 146);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(47, 16);
			this.label2.TabIndex = 137;
			this.label2.Text = " Numero";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// dtFecha
			// 
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionPadre, "_GuiaRemisionPadre.Fecha"));
			this.dtFecha.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFecha.Location = new System.Drawing.Point(448, 146);
			this.dtFecha.Name = "dtFecha";
			// 
			// dtFecha.Properties
			// 
			this.dtFecha.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFecha.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFecha.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFecha.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFecha.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFecha.Size = new System.Drawing.Size(96, 20);
			this.dtFecha.TabIndex = 139;
			// 
			// txtIdCompra
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdCompra.Appearance = appearance14;
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionPadre, "_GuiaRemisionPadre.idGuiaRemisionPadre"));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.FormatString = "";
			this.txtIdCompra.Location = new System.Drawing.Point(440, 52);
			this.txtIdCompra.MaskInput = "nnnnnnnnn";
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(48, 22);
			this.txtIdCompra.TabIndex = 140;
			// 
			// cmbLlegada
			// 
			this.cmbLlegada.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLlegada.DataMember = "GuiaRemisionPunto";
			this.cmbLlegada.DataSource = this.cdsGuiaPunto;
			ultraGridBand7.AddButtonCaption = "CompraComprobante";
			ultraGridColumn103.Header.VisiblePosition = 0;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 1;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn103,
																										 ultraGridColumn104});
			this.cmbLlegada.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbLlegada.DisplayMember = "Nombre";
			this.cmbLlegada.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLlegada.Location = new System.Drawing.Point(480, 250);
			this.cmbLlegada.Name = "cmbLlegada";
			this.cmbLlegada.Size = new System.Drawing.Size(136, 77);
			this.cmbLlegada.TabIndex = 141;
			this.cmbLlegada.ValueMember = "idGuiaRemisionPunto";
			this.cmbLlegada.Visible = false;
			// 
			// chkChofer
			// 
			this.chkChofer.Location = new System.Drawing.Point(392, 8);
			this.chkChofer.Name = "chkChofer";
			// 
			// chkChofer.Properties
			// 
			this.chkChofer.Properties.Caption = "Todo Chofer";
			this.chkChofer.Size = new System.Drawing.Size(80, 19);
			this.chkChofer.TabIndex = 143;
			this.chkChofer.CheckedChanged += new System.EventHandler(this.chkChofer_CheckedChanged);
			// 
			// cmbChofer
			// 
			this.cmbChofer.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbChofer.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbChofer.DataSource = this.cdvChofer;
			ultraGridColumn105.Header.VisiblePosition = 1;
			ultraGridColumn105.Width = 137;
			ultraGridColumn106.Header.VisiblePosition = 3;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 0;
			ultraGridColumn107.Width = 322;
			ultraGridColumn108.Header.VisiblePosition = 4;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 5;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 2;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 6;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 7;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 8;
			ultraGridColumn114.Header.VisiblePosition = 9;
			ultraGridColumn115.Header.VisiblePosition = 10;
			ultraGridColumn116.Header.VisiblePosition = 11;
			ultraGridColumn117.Header.VisiblePosition = 12;
			ultraGridColumn118.Header.VisiblePosition = 13;
			ultraGridColumn119.Header.VisiblePosition = 14;
			ultraGridColumn120.Header.VisiblePosition = 15;
			ultraGridColumn121.Header.VisiblePosition = 16;
			ultraGridColumn122.Header.VisiblePosition = 17;
			ultraGridColumn123.Header.VisiblePosition = 18;
			ultraGridColumn124.Header.VisiblePosition = 19;
			ultraGridColumn125.Header.VisiblePosition = 20;
			ultraGridColumn126.Header.VisiblePosition = 21;
			ultraGridColumn127.Header.VisiblePosition = 22;
			ultraGridColumn128.Header.VisiblePosition = 23;
			ultraGridBand8.Columns.AddRange(new object[] {
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
																										 ultraGridColumn128});
			this.cmbChofer.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbChofer.DisplayMember = "Nombre";
			this.cmbChofer.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbChofer.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbChofer.Location = new System.Drawing.Point(480, 8);
			this.cmbChofer.Name = "cmbChofer";
			this.cmbChofer.Size = new System.Drawing.Size(216, 21);
			this.cmbChofer.TabIndex = 142;
			this.cmbChofer.ValueMember = "idCliente";
			// 
			// cdvChofer
			// 
			this.cdvChofer.BindingContextCtrl = this;
			this.cdvChofer.DataSet = this.cdsCliente;
			this.cdvChofer.RowFilter = "chofer=1 and (proveedor=1  or ambos=1)";
			this.cdvChofer.Sort = "Nombre";
			this.cdvChofer.TableName = "";
			this.cdvChofer.TableViewName = "Cliente";
			// 
			// GuiaRemisionFact
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(704, 418);
			this.Controls.Add(this.chkChofer);
			this.Controls.Add(this.cmbChofer);
			this.Controls.Add(this.cmbLlegada);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbCliente1);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbPunto);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtGuia2);
			this.Controls.Add(this.txtGuia1);
			this.Controls.Add(this.cmbPuntoPartida);
			this.Controls.Add(this.cmbPuntoLlegada);
			this.Controls.Add(this.chkRuta);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.dtFechaHasta);
			this.Controls.Add(this.chkGuia);
			this.Controls.Add(this.chkFecha);
			this.Controls.Add(this.chkCliente);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.txtidGuiaRemisionPadre);
			this.Name = "GuiaRemisionFact";
			this.Text = "Guia Remision Ventas";
			this.Load += new System.EventHandler(this.GuiaRemisionFact_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaRemisionG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGuia.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkFecha.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkRuta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPuntoPartida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaPunto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPuntoLlegada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaPunto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGuia1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGuia2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaRemisionG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaRemisionPadre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPunto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaRemisionPadre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidGuiaRemisionPadre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLlegada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkChofer.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbChofer)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvChofer)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void GuiaRemisionFact_Load(object sender, System.EventArgs e)
		{
			txtIdCompra.Width = 0;
			txtidGuiaRemisionPadre.Width = 0;
			this.barraDato1.stFiltroExt  = " And tipo = 1";
			dtFechaDesde.Enabled =false;
			dtFechaHasta.Enabled =false; 
		  this.barraDato1.HabilitaControles(false);  
		  
			this.barraDato1.PosUltima(); 
		}

		private void BorraFiltro()
			{
			errorProvider.SetError(cmbCliente, "");
			errorProvider.SetError(cmbPuntoPartida, "");
			errorProvider.SetError(cmbPuntoLlegada, "");
			errorProvider.SetError(cmbChofer, "");
		}
		private bool VerificaFiltro()
		{
			bool bOk = true;

			BorraFiltro();
			if (chkChofer.Checked==false)
				if (cmbChofer.Text.Trim().Length ==0)
				{
					errorProvider.SetError(cmbChofer, "Ingrese Chofer");
					bOk = false;
				}

			if (chkCliente.Checked==false)
				if (cmbCliente.Text.Trim().Length ==0)
				{
					errorProvider.SetError(cmbCliente, "Ingrese Cliente");
					bOk = false;
				}

			if (chkRuta.Checked==false)
			{
				if (cmbPuntoPartida.Text.Trim().Length ==0)
				{
					errorProvider.SetError(cmbPuntoPartida, "Ingrese Partida");
					bOk = false;
				}
				if (cmbPuntoLlegada.Text.Trim().Length ==0)
				{
					errorProvider.SetError(cmbPuntoLlegada, "Ingrese Punto de Llegada");
					bOk = false;
				}
			}


			if (chkGuia.Checked==false)
			{
				if (txtGuia1.Text.Trim().Length ==0)
				{
					errorProvider.SetError(txtGuia1, "Ingrese Guia");
					bOk = false;
				}
				if (txtGuia2.Text.Trim().Length ==0)
				{
					errorProvider.SetError(txtGuia2, "Ingrese Guia");
					bOk = false;
				}
			}

			return bOk;
		}


		private void cdsGuiaRemisionG_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGuiaRemisionG.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			if (!VerificaFiltro()) return;
			//aqui facturado
			string  stFiltro = " facturado=0 and borrar=0 and idguiaremisionpadre is null";
			
			if (chkChofer.Checked==false)
				stFiltro += " and idTransportista=" + cmbChofer.Value.ToString();
			
			if (chkCliente.Checked==false)
					stFiltro += " and idcliente=" + cmbCliente.Value.ToString();

			if (chkRuta.Checked==false)
				stFiltro += " and idpuntopartida=" + cmbPuntoPartida.Value.ToString() + " and idpuntollegada=" + cmbPuntoLlegada.Value.ToString();						

			if (chkGuia.Checked==false)
				stFiltro += " and nroguia>=" + txtGuia1.Value.ToString() + " and nroguia<=" + txtGuia2.Value.ToString() ;						
		
			DateTime dtFechaIni = (DateTime) dtFechaDesde.EditValue;
			DateTime dtFechaFin = (DateTime) dtFechaHasta.EditValue;
			dtFechaFin=dtFechaFin.AddDays(1); 
			if (chkFecha.Checked==false)
				stFiltro += " and fEmision2 >='" + dtFechaIni.ToString("yyyyMMdd") + "' and fEmision2 <'" + dtFechaFin.ToString("yyyyMMdd") +"'";
			

			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsGuiaRemisionPadre, "GuiaRemisionPadre", "idGuiaRemisionPadre = 0 ");
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			C1.Data.FilterCondition fcHija = new 
				C1.Data.FilterCondition(this.cdsGuiaRemisionPadre , "GuiaRemisionG", stFiltro );
			fcTotal.Add(fcHija);
			
			this.cdsGuiaRemisionPadre.Fill(fcTotal, false);
	
		}

		private void chkChofer_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbChofer.Enabled =  !this.chkChofer.Checked;  
		}
		private void chkCliente_CheckedChanged(object sender, System.EventArgs e)
		{
					this.cmbCliente.Enabled =  !this.chkCliente.Checked;  
		}

		private void chkFecha_CheckedChanged(object sender, System.EventArgs e)
		{
				this.dtFechaDesde.Enabled=!this.chkFecha.Checked;  
				this.dtFechaHasta.Enabled=!this.chkFecha.Checked;
		}

		private void chkGuia_CheckedChanged(object sender, System.EventArgs e)
		{
				this.txtGuia1.Enabled  =  !this.chkGuia.Checked;
				this.txtGuia2.Enabled  =  !this.chkGuia.Checked;
		}

		private void chkRuta_CheckedChanged(object sender, System.EventArgs e)
		{
				this.cmbPuntoPartida.Enabled  =  !this.chkRuta.Checked;
				this.cmbPuntoLlegada.Enabled  =  !this.chkRuta.Checked; 
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using (BuscaGuiaGruasV miBusqueda = new BuscaGuiaGruasV())
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					if (miBusqueda.idBusca.Text.Length == 0) return;
					int id = int.Parse(miBusqueda.idBusca.Text);
					barraDato1.IdRegistro = id;
					barraDato1.ProximoId(5);
				}
			}
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro(); 
			editar=0; 

		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();  
			BorraFiltro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();  
			editar=1;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			BorraFiltro();
			if (!Verifica()) return;
			
			if (editar==1)
			{
				this.barraDato1.GrabaRegistro(); 
			}
			else
			{
				DateTime dtFechaIni1 = (DateTime) dtFecha.EditValue ;
				string stSelect= "insert into guiaremisionpadre (tipo,numero,fecha,idcliente) values(1,"
				 + txtNumero.Value.ToString() 		+",'"		
					+ dtFechaIni1.ToString("yyyyMMdd") + "',"
					+ cmbCliente1.Value.ToString() + ")";   
				
				this.cdsGuiaRemisionPadre.Update(); 
        
				this.barraDato1.DeshacerRegistro();  
				
				Funcion.EjecutaSQL(cdsGuiaPunto, stSelect, true);	
				this.barraDato1.PosUltima();
				
				stSelect="update guiaremisiong set idguiaremisionpadre=" 
				+ txtidGuiaRemisionPadre.Value.ToString() + " where facturado=1 and idguiaremisionpadre is null";
				Funcion.EjecutaSQL(cdsGuiaPunto, stSelect, true);	

				this.barraDato1.PosUltima();
			}
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (this.dtFecha.Text.Trim().Length == 0)
			{
				this.errorProvider.SetError(this.dtFecha , "Ingrese fecha ");
				bOk = false;
			}
			if (this.txtNumero.Text.Trim().Length == 0)
			{
				this.errorProvider.SetError(this.txtNumero , "Ingrese numero de factura");
				bOk = false;
			}
			if (this.cmbCliente1.Text.Length == 0)
			{
				this.errorProvider.SetError(this.cmbCliente1, "Ingrese cliente");
				bOk = false;
			}

			return bOk;
		}

		private void BorraErrores()
		{
			this.errorProvider.SetError(this.dtFecha, "");
			this.errorProvider.SetError(this.txtNumero , "");
			this.errorProvider.SetError(this.cmbCliente1, "");

		}




		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drGuiaPadre = cdsGuiaRemisionG.TableViews["GuiaRemisionPadre"].Rows[0];
			}
			catch{};
		}

		private void ultraNumericEditor2_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor3_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkCliente_Validated(object sender, System.EventArgs e)
		{
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar) return;
			if (DialogResult.No == MessageBox.Show("¿Desea Generar un Factura de Venta?", 
        "Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			DateTime dtFechaIni = (DateTime) dtFecha.EditValue ;
			string	stExec = "Exec CreaFactura '" + cmbCliente1.Text.Trim().ToString()  + "','001001'"
				+ ",'" + txtNumero.Value.ToString() +"','" + dtFechaIni.ToString("yyyyMMdd") + "',0,0,0,0,0,1,"
				+ "@CONTADOCREDITO=2,@CREAR = 0";

			int IdCompra = Funcion.iEscalarSQL(cdsCliente, stExec, true);
			if (IdCompra == 0)
			{
				MessageBox.Show("Factura de Venta NO Generada porque ya Existe", 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
	
				bool bFact = bool.Parse(dr.Cells[37].Value.ToString());
				if (bFact) //esta tomada en cuenta 
				{
					if (dr.Cells[27].Value != DBNull.Value)
					{		
						string sel="select nombre from guiaremisionpunto where idguiaremisionpunto="+dr.Cells[15].Value;
						string origen=Funcion.sEscalarSQL(cdsCliente,sel );  
						sel="select nombre from guiaremisionpunto where idguiaremisionpunto="+dr.Cells[17].Value;
						string destino=Funcion.sEscalarSQL(cdsCliente,sel );  
						sel="select nombre from guiaremisionpunto where idguiaremisionpunto="+dr.Cells[17].Value;


						DateTime dtFechaEmision = (DateTime) dr.Cells["fEmision2"].Value;
						string stTonelaje = "";
						if (dr.Cells[42].Value != DBNull.Value) stTonelaje = dr.Cells[42].Value.ToString() ;
						string minota =  stTonelaje + System.Environment.NewLine
							+ "EL DIA: " + dtFechaEmision.ToString("dd/MMM/yyyy") + System.Environment.NewLine 
							+ "LUGAR DE ORIGEN : "  + origen.ToString() + System.Environment.NewLine 
							+ "LUGAR DE DESTINO: "  + destino.ToString() + System.Environment.NewLine
							+ "NUMERO DE GUIA: " + dr.Cells[4].Value;

						int Carro1=int.Parse(dr.Cells[27].Value.ToString()) ;	
						 stExec = "Exec CreaFacturaDetalle 1,'',12,0," + "@IdProyecto=" + Carro1.ToString()  
							+ ",@IdSubProyecto=0, @IdCompra=" + IdCompra.ToString() + ",@Nota='" + minota.ToString() + "'";
						Funcion.EjecutaSQL(cdsCliente,stExec);  
					}

					if (dr.Cells[28].Value != DBNull.Value)
					{		
						string sel="select nombre from guiaremisionpunto where idguiaremisionpunto="+dr.Cells[15].Value;
						string origen=Funcion.sEscalarSQL(cdsCliente,sel );  
						sel="select nombre from guiaremisionpunto where idguiaremisionpunto="+dr.Cells[17].Value;
						string destino=Funcion.sEscalarSQL(cdsCliente,sel );  
						sel="select nombre from guiaremisionpunto where idguiaremisionpunto="+dr.Cells[17].Value;

						DateTime dtFechaEmision = (DateTime) dr.Cells["fEmision2"].Value;
						string minota =  dr.Cells[42].Value + System.Environment.NewLine
							+ "EL DIA: " + dtFechaEmision.ToString("dd/MMM/yyyy") + System.Environment.NewLine 
							+ "LUGAR DE ORIGEN : "  + origen.ToString() + System.Environment.NewLine 
							+ "LUGAR DE DESTINO: "  + destino.ToString() + System.Environment.NewLine
							+ "NUMERO DE GUIA: " + dr.Cells[4].Value;
						
						int Carro2=int.Parse(dr.Cells[28].Value.ToString()) ;	
				    stExec = "Exec CreaFacturaDetalle 1,'',12,0," + "@IdProyecto=" +Carro2.ToString()   
							+ ",@IdSubProyecto=0, @IdCompra=" +IdCompra.ToString() + ",@Nota='" + minota.ToString() +"'";
						Funcion.EjecutaSQL(cdsCliente,stExec);
					}
					if (dr.Cells[29].Value != DBNull.Value)
					{		
						string sel="select nombre from guiaremisionpunto where idguiaremisionpunto="+dr.Cells[15].Value;
						string origen=Funcion.sEscalarSQL(cdsCliente,sel );  
						sel="select nombre from guiaremisionpunto where idguiaremisionpunto="+dr.Cells[17].Value;
						string destino=Funcion.sEscalarSQL(cdsCliente,sel );  
						sel="select nombre from guiaremisionpunto where idguiaremisionpunto="+dr.Cells[17].Value;

						DateTime dtFechaEmision = (DateTime) dr.Cells["fEmision2"].Value;

						string minota =  dr.Cells[42].Value + System.Environment.NewLine
							+ "EL DIA: " +  dtFechaEmision.ToString("dd/MMM/yyyy") + System.Environment.NewLine 
							+ "LUGAR DE ORIGEN : "  + origen.ToString() + System.Environment.NewLine 
							+ "LUGAR DE DESTINO: "  + destino.ToString() + System.Environment.NewLine
							+ "NUMERO DE GUIA: " + dr.Cells[4].Value;

						int Carro3=int.Parse(dr.Cells[29].Value.ToString()) ;	
				    stExec = "Exec CreaFacturaDetalle 1,'',12,0," + "@IdProyecto=" +Carro3.ToString()    
							+ ",@IdSubProyecto=0, @IdCompra=" +IdCompra.ToString() + ",@Nota='" + minota.ToString() + "'";
						Funcion.EjecutaSQL(cdsCliente,stExec);
					}
				}
			}
			MessageBox.Show("Factura de Venta Generada", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
//			string strFiltro = "{GuiaRemisionG.idGuiaRemisionG} = " 
//				+ this.txtIdGuiaRemisionG.Value.ToString();
//
			Cursor = Cursors.WaitCursor;
			Reporte miReporte = new Reporte("GuiaRemisionTotal.rpt");
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		
		
	}
}
