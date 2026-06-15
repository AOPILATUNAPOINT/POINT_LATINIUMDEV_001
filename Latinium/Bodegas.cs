using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Bodegas.
	/// </summary>
	public class Bodegas : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private C1.Data.C1DataSet cdsBodega;
		private C1.Data.C1DataSet cdsNumero;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsProyecto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTiposDeAlmacen;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSectores;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbNivelSocioEconomico;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbUbicacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Bodegas()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Bodega", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad", -1, "ultraDropDown2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura", -1, "ultraDropDown1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TransferenciaMultiple");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OrdenTM");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumNotaCreditoFiscal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumNotaCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutGuiaRemision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorGuiaRemision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoDeAlmacen", -1, "cmbTiposDeAlmacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Inventario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FS");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DirectorioPrecios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DirectorioPCM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DirectorioPCP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DirectorioPC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Compra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PRC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NCP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promociones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImprimePrecios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PTC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector", -1, "cmbSectores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Compras");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodProvinciaTelf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Spiff");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Metros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUbicacion", -1, "cmbUbicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSocioEconomica", -1, "cmbNivelSocioEconomico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jefe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subencargado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cajera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Senior");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Junior");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vigilante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalPersonal");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargaSeriales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCiudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSector");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CajaChica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoDeAlmacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDeAlmacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoDeAlmacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDeAlmacen");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNivelSocioEconomico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nivel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNivelSocioEconomico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nivel");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLocalUbicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLocalUbicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ubicacion");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsBodega = new C1.Data.C1DataSet();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.cdsNumero = new C1.Data.C1DataSet();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDropDown2 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbTiposDeAlmacen = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSectores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbNivelSocioEconomico = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbUbicacion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiposDeAlmacen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSectores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNivelSocioEconomico)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Bodega";
			this.ultraGrid1.DataSource = this.cdsBodega;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Width = 120;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 30;
			ultraGridColumn4.Header.VisiblePosition = 8;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 176;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.Caption = "Ciudad";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn5.Width = 120;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.Header.Caption = "Dirección";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 110;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.Header.Caption = "Tipo de factura";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn7.Width = 121;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.Header.Caption = "Código";
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Width = 50;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.Header.Caption = "Centro de Costo";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Width = 150;
			ultraGridColumn10.Header.Caption = "ACTIVO";
			ultraGridColumn10.Header.VisiblePosition = 14;
			ultraGridColumn10.Width = 17;
			ultraGridColumn11.Header.Caption = "MATRIZ";
			ultraGridColumn11.Header.VisiblePosition = 15;
			ultraGridColumn11.Width = 17;
			ultraGridColumn12.Header.Caption = "TRANSFERENCIA";
			ultraGridColumn12.Header.VisiblePosition = 16;
			ultraGridColumn12.Width = 17;
			ultraGridColumn13.Header.VisiblePosition = 11;
			ultraGridColumn13.Width = 38;
			ultraGridColumn14.Header.VisiblePosition = 12;
			ultraGridColumn14.Width = 100;
			ultraGridColumn15.Header.VisiblePosition = 13;
			ultraGridColumn15.Width = 70;
			ultraGridColumn16.Header.Caption = "FACTURA";
			ultraGridColumn16.Header.VisiblePosition = 19;
			ultraGridColumn16.Width = 17;
			ultraGridColumn17.Header.VisiblePosition = 9;
			ultraGridColumn17.Width = 70;
			ultraGridColumn18.Header.Caption = "E-Mail";
			ultraGridColumn18.Header.VisiblePosition = 10;
			ultraGridColumn18.Width = 52;
			ultraGridColumn19.Header.VisiblePosition = 28;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 160;
			ultraGridColumn20.Header.Caption = "TRANSF. MULTIPLE";
			ultraGridColumn20.Header.VisiblePosition = 17;
			ultraGridColumn20.Width = 25;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance2;
			ultraGridColumn21.Header.Caption = "Orden";
			ultraGridColumn21.Header.VisiblePosition = 29;
			ultraGridColumn21.Width = 29;
			ultraGridColumn22.Header.Caption = "Num Factura Automatico";
			ultraGridColumn22.Header.VisiblePosition = 30;
			ultraGridColumn22.Width = 36;
			ultraGridColumn23.Header.VisiblePosition = 31;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 66;
			ultraGridColumn24.Header.VisiblePosition = 32;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 124;
			ultraGridColumn25.Header.VisiblePosition = 33;
			ultraGridColumn25.Width = 52;
			ultraGridColumn26.Header.VisiblePosition = 34;
			ultraGridColumn26.Width = 58;
			ultraGridColumn27.Header.VisiblePosition = 23;
			ultraGridColumn27.Width = 34;
			ultraGridColumn28.Header.Caption = "Tipo De Almacen";
			ultraGridColumn28.Header.VisiblePosition = 35;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn28.Width = 84;
			ultraGridColumn29.Header.VisiblePosition = 25;
			ultraGridColumn29.Width = 56;
			ultraGridColumn30.Header.Caption = "Factura con Serial";
			ultraGridColumn30.Header.VisiblePosition = 26;
			ultraGridColumn30.Width = 27;
			ultraGridColumn31.Header.VisiblePosition = 36;
			ultraGridColumn31.Width = 97;
			ultraGridColumn32.Header.VisiblePosition = 37;
			ultraGridColumn33.Header.VisiblePosition = 38;
			ultraGridColumn34.Header.VisiblePosition = 39;
			ultraGridColumn35.Header.VisiblePosition = 20;
			ultraGridColumn36.Header.VisiblePosition = 18;
			ultraGridColumn37.Header.VisiblePosition = 21;
			ultraGridColumn38.Header.VisiblePosition = 22;
			ultraGridColumn39.Header.VisiblePosition = 24;
			ultraGridColumn40.Header.Caption = "I P";
			ultraGridColumn40.Header.VisiblePosition = 27;
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn42.Header.VisiblePosition = 41;
			ultraGridColumn43.Header.VisiblePosition = 42;
			ultraGridColumn44.Header.Caption = "Sector";
			ultraGridColumn44.Header.VisiblePosition = 43;
			ultraGridColumn44.Width = 100;
			ultraGridColumn45.Header.VisiblePosition = 44;
			ultraGridColumn46.Header.VisiblePosition = 45;
			ultraGridColumn47.Header.VisiblePosition = 46;
			ultraGridColumn48.Header.VisiblePosition = 47;
			ultraGridColumn49.Header.VisiblePosition = 48;
			ultraGridColumn50.Header.Caption = "Ubicacion";
			ultraGridColumn50.Header.VisiblePosition = 49;
			ultraGridColumn50.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn51.Header.Caption = "Socio Economico";
			ultraGridColumn51.Header.VisiblePosition = 50;
			ultraGridColumn51.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn52.Header.VisiblePosition = 51;
			ultraGridColumn53.Header.VisiblePosition = 52;
			ultraGridColumn54.Header.VisiblePosition = 53;
			ultraGridColumn55.Header.VisiblePosition = 54;
			ultraGridColumn56.Header.VisiblePosition = 55;
			ultraGridColumn57.Header.VisiblePosition = 56;
			ultraGridColumn58.Header.VisiblePosition = 57;
			ultraGridColumn59.Header.VisiblePosition = 58;
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
																										 ultraGridColumn59});
			appearance3.FontData.SizeInPoints = 8F;
			ultraGridBand1.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(1, 0);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1344, 390);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// cdsBodega
			// 
			this.cdsBodega.BindingContextCtrl = this;
			this.cdsBodega.DataLibrary = "LibFacturacion";
			this.cdsBodega.DataLibraryUrl = "";
			this.cdsBodega.DataSetDef = "dsBodega";
			this.cdsBodega.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsBodega.Name = "cdsBodega";
			this.cdsBodega.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsBodega.SchemaDef = null;
			this.cdsBodega.BeforeFill += new C1.Data.FillEventHandler(this.cdsBodega_BeforeFill);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(8, 398);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(104, 25);
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.Location = new System.Drawing.Point(120, 398);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(104, 25);
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cdsNumero
			// 
			this.cdsNumero.BindingContextCtrl = this;
			this.cdsNumero.DataLibrary = "LibFacturacion";
			this.cdsNumero.DataLibraryUrl = "";
			this.cdsNumero.DataSetDef = "dsCompraTabla";
			this.cdsNumero.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsNumero.Name = "cdsNumero";
			this.cdsNumero.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsNumero.SchemaDef = null;
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DataMember = "CompraNumero";
			this.ultraDropDown1.DataSource = this.cdsNumero;
			ultraGridColumn60.Header.VisiblePosition = 1;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 0;
			ultraGridColumn61.Width = 177;
			ultraGridColumn62.Header.VisiblePosition = 3;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 2;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 4;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 5;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 6;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 7;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 8;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 9;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 10;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 11;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 12;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 13;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 14;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 15;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 16;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 17;
			ultraGridColumn78.Header.VisiblePosition = 18;
			ultraGridColumn79.Header.VisiblePosition = 19;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn79});
			this.ultraDropDown1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraDropDown1.DisplayMember = "Nombre";
			this.ultraDropDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown1.Location = new System.Drawing.Point(40, 304);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(200, 86);
			this.ultraDropDown1.TabIndex = 3;
			this.ultraDropDown1.ValueMember = "idTipoFactura";
			this.ultraDropDown1.Visible = false;
			// 
			// ultraDropDown2
			// 
			this.ultraDropDown2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown2.DataSource = this.ultraDataSource3;
			ultraGridColumn80.Header.VisiblePosition = 0;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 1;
			ultraGridColumn81.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn80,
																										 ultraGridColumn81});
			this.ultraDropDown2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraDropDown2.DisplayMember = "Ciudad";
			this.ultraDropDown2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown2.Location = new System.Drawing.Point(360, 88);
			this.ultraDropDown2.Name = "ultraDropDown2";
			this.ultraDropDown2.Size = new System.Drawing.Size(200, 86);
			this.ultraDropDown2.TabIndex = 4;
			this.ultraDropDown2.ValueMember = "idCiudad";
			this.ultraDropDown2.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
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
			// cmbProyecto
			// 
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataMember = "Proyecto";
			this.cmbProyecto.DataSource = this.cdsProyecto;
			ultraGridColumn82.Header.VisiblePosition = 1;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 2;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 0;
			ultraGridColumn84.Width = 149;
			ultraGridColumn85.Header.VisiblePosition = 3;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 4;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 5;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 6;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 7;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 8;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 9;
			ultraGridColumn92.Header.VisiblePosition = 10;
			ultraGridColumn93.Header.VisiblePosition = 11;
			ultraGridColumn94.Header.VisiblePosition = 12;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn94});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(265, 201);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(176, 86);
			this.cmbProyecto.TabIndex = 10;
			this.cmbProyecto.ValueMember = "idProyecto";
			this.cmbProyecto.Visible = false;
			// 
			// cmbTiposDeAlmacen
			// 
			this.cmbTiposDeAlmacen.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTiposDeAlmacen.DataSource = this.ultraDataSource2;
			this.cmbTiposDeAlmacen.DisplayLayout.AutoFitColumns = true;
			ultraGridBand5.AddButtonCaption = "CompraNumero";
			ultraGridColumn95.Header.VisiblePosition = 0;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.Caption = "Tipo";
			ultraGridColumn96.Header.VisiblePosition = 1;
			ultraGridColumn96.Width = 216;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn95,
																										 ultraGridColumn96});
			this.cmbTiposDeAlmacen.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbTiposDeAlmacen.DisplayMember = "TipoDeAlmacen";
			this.cmbTiposDeAlmacen.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTiposDeAlmacen.Location = new System.Drawing.Point(573, 201);
			this.cmbTiposDeAlmacen.Name = "cmbTiposDeAlmacen";
			this.cmbTiposDeAlmacen.Size = new System.Drawing.Size(235, 86);
			this.cmbTiposDeAlmacen.TabIndex = 11;
			this.cmbTiposDeAlmacen.ValueMember = "idTipoDeAlmacen";
			this.cmbTiposDeAlmacen.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// cmbSectores
			// 
			this.cmbSectores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSectores.DataSource = this.ultraDataSource1;
			ultraGridColumn97.Header.VisiblePosition = 0;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 1;
			ultraGridColumn98.Width = 150;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn97,
																										 ultraGridColumn98});
			this.cmbSectores.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbSectores.DisplayMember = "Sector";
			this.cmbSectores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSectores.Location = new System.Drawing.Point(904, 200);
			this.cmbSectores.Name = "cmbSectores";
			this.cmbSectores.Size = new System.Drawing.Size(200, 86);
			this.cmbSectores.TabIndex = 12;
			this.cmbSectores.ValueMember = "idSector";
			this.cmbSectores.Visible = false;
			// 
			// cmbNivelSocioEconomico
			// 
			this.cmbNivelSocioEconomico.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNivelSocioEconomico.DataSource = this.ultraDataSource4;
			ultraGridColumn99.Header.VisiblePosition = 0;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 1;
			ultraGridColumn100.Width = 150;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn99,
																										 ultraGridColumn100});
			this.cmbNivelSocioEconomico.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbNivelSocioEconomico.DisplayMember = "Nivel";
			this.cmbNivelSocioEconomico.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNivelSocioEconomico.Location = new System.Drawing.Point(1032, 320);
			this.cmbNivelSocioEconomico.Name = "cmbNivelSocioEconomico";
			this.cmbNivelSocioEconomico.Size = new System.Drawing.Size(200, 86);
			this.cmbNivelSocioEconomico.TabIndex = 14;
			this.cmbNivelSocioEconomico.ValueMember = "idNivelSocioEconomico";
			this.cmbNivelSocioEconomico.Visible = false;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// cmbUbicacion
			// 
			this.cmbUbicacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUbicacion.DataSource = this.ultraDataSource5;
			this.cmbUbicacion.DisplayLayout.AutoFitColumns = true;
			ultraGridBand8.AddButtonCaption = "CompraNumero";
			ultraGridColumn101.Header.VisiblePosition = 0;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn101.Width = 110;
			ultraGridColumn102.Header.VisiblePosition = 1;
			ultraGridColumn102.Width = 216;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn101,
																										 ultraGridColumn102});
			this.cmbUbicacion.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbUbicacion.DisplayMember = "Ubicacion";
			this.cmbUbicacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUbicacion.Location = new System.Drawing.Point(704, 320);
			this.cmbUbicacion.Name = "cmbUbicacion";
			this.cmbUbicacion.Size = new System.Drawing.Size(235, 86);
			this.cmbUbicacion.TabIndex = 13;
			this.cmbUbicacion.ValueMember = "idLocalUbicacion";
			this.cmbUbicacion.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// Bodegas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1346, 430);
			this.Controls.Add(this.cmbNivelSocioEconomico);
			this.Controls.Add(this.cmbUbicacion);
			this.Controls.Add(this.cmbSectores);
			this.Controls.Add(this.cmbTiposDeAlmacen);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.ultraDropDown2);
			this.Controls.Add(this.ultraDropDown1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "Bodegas";
			this.Text = "Bodegas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Bodegas_KeyDown);
			this.Load += new System.EventHandler(this.Bodegas_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiposDeAlmacen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSectores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNivelSocioEconomico)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void cdsBodega_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsBodega.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsBodega.Update();
				string stExec = "Exec BodegaLlenaNombre ";
				Funcion.EjecutaSQL(cdsBodega, stExec, true);
				Close();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void Bodegas_Load(object sender, System.EventArgs e)
		{
			this.ultraDropDown2.DataSource = Funcion.dvProcedimiento(cdsBodega, "Select idCiudad, Ciudad From ClienteCiudad Order By Ciudad");
			this.cmbTiposDeAlmacen.DataSource = Funcion.dvProcedimiento(cdsBodega, "SELECT idTipoDeAlmacen, TipoDeAlmacen FROM TipoDeAlmacen Order By idTipoDeAlmacen");
			this.cmbSectores.DataSource = Funcion.dvProcedimiento(cdsBodega, "SELECT idSector, Sector FROM ClienteSector");
			this.cmbUbicacion.DataSource = Funcion.dvProcedimiento(cdsBodega, "Select idLocalUbicacion, Ubicacion From LocalUbicacion");
			this.cmbNivelSocioEconomico.DataSource = Funcion.dvProcedimiento(cdsBodega, "Select idNivelSocioEconomico, Nivel From NivelSocioEconomico");
		}

		private void Bodegas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				switch(e.KeyValue)
				{
						#region Page Down Abre lista desplegable
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
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
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Tecla");
			}
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "TELEFONO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "Establecimiento" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "PuntoEmision")
				Funcion.CamposNumericos(sender, e);
		}
	}
}
