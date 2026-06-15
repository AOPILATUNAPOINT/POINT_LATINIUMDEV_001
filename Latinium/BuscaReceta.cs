using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaAsiento.
	/// </summary>
	public class BuscaReceta : DevExpress.XtraEditors.XtraForm
	{
		public System.Windows.Forms.TextBox idBusca;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataView cdvProducto;
		private C1.Data.C1DataSet cdsProducto;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataView cdvArtCod;
		private C1.Data.C1DataSet cdsArticulo;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgTipoBusqueda;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgOpcion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMarca1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private C1.Data.C1DataSet cdsArticuloTabla;
		private C1.Data.C1DataView cdvGrupo;
		private C1.Data.C1DataView cdvMarca;
		private C1.Data.C1DataView cdvSubGrupo;
		private System.ComponentModel.IContainer components=null;

		public BuscaReceta()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloMarca", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSubGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas8");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caracteristica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMinima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMaxima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gramos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConsumoProm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoReposicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroParte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRequerida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcCre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision30");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision60");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision90");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision120");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PaqueteUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CambiarPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredefCR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoFecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aportable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.cdvProducto = new C1.Data.C1DataView();
			this.cdsProducto = new C1.Data.C1DataSet();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdvArtCod = new C1.Data.C1DataView();
			this.idBusca = new System.Windows.Forms.TextBox();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.rdgTipoBusqueda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rdgOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbMarca1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvMarca = new C1.Data.C1DataView();
			this.cdsArticuloTabla = new C1.Data.C1DataSet();
			this.cmbSubGrupo1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvSubGrupo = new C1.Data.C1DataView();
			this.cmbGrupo1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvGrupo = new C1.Data.C1DataView();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cdvProducto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProducto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCod)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// cdvProducto
			// 
			this.cdvProducto.BindingContextCtrl = this;
			this.cdvProducto.DataSet = this.cdsProducto;
			this.cdvProducto.TableName = "";
			this.cdvProducto.TableViewName = "Producto";
			// 
			// cdsProducto
			// 
			this.cdsProducto.BindingContextCtrl = this;
			this.cdsProducto.DataLibrary = "LibFacturacion";
			this.cdsProducto.DataLibraryUrl = "";
			this.cdsProducto.DataSetDef = "dsProducto";
			this.cdsProducto.FillOnRequest = false;
			this.cdsProducto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProducto.Name = "cdsProducto";
			this.cdsProducto.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsProducto.SchemaDef = null;
			this.cdsProducto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProducto_BeforeFill);
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.RowFilter = "idTipoGrupo In (6, 8)";
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulo";
			this.cdsArticulo.FillOnRequest = false;
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// cdvArtCod
			// 
			this.cdvArtCod.BindingContextCtrl = this;
			this.cdvArtCod.DataSet = this.cdsArticulo;
			this.cdvArtCod.Sort = "Codigo";
			this.cdvArtCod.TableName = "";
			this.cdvArtCod.TableViewName = "Articulo";
			// 
			// idBusca
			// 
			this.idBusca.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdvArticulo, "idArticulo"));
			this.idBusca.Location = new System.Drawing.Point(456, 17);
			this.idBusca.Name = "idBusca";
			this.idBusca.Size = new System.Drawing.Size(48, 20);
			this.idBusca.TabIndex = 23;
			this.idBusca.Text = "";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(444, 121);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 57;
			this.btCancelar.Text = "&Cancelar";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(444, 86);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 56;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(444, 52);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(75, 24);
			this.btBuscar.TabIndex = 55;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// rdgTipoBusqueda
			// 
			this.rdgTipoBusqueda.CheckedIndex = 2;
			this.rdgTipoBusqueda.ItemAppearance = appearance1;
			this.rdgTipoBusqueda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Que Empiece";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Exacta";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "En cualquier parte";
			this.rdgTipoBusqueda.Items.Add(valueListItem1);
			this.rdgTipoBusqueda.Items.Add(valueListItem2);
			this.rdgTipoBusqueda.Items.Add(valueListItem3);
			this.rdgTipoBusqueda.ItemSpacingVertical = 10;
			this.rdgTipoBusqueda.Location = new System.Drawing.Point(292, 52);
			this.rdgTipoBusqueda.Name = "rdgTipoBusqueda";
			this.rdgTipoBusqueda.Size = new System.Drawing.Size(136, 86);
			this.rdgTipoBusqueda.TabIndex = 54;
			this.rdgTipoBusqueda.Text = "En cualquier parte";
			// 
			// rdgOpcion
			// 
			this.rdgOpcion.ItemAppearance = appearance2;
			this.rdgOpcion.ItemOrigin = new System.Drawing.Point(6, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Artículo";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Código";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Cod. Alterno";
			valueListItem7.DataValue = 3;
			valueListItem7.DisplayText = "Grupo";
			valueListItem8.DataValue = 4;
			valueListItem8.DisplayText = "SubGrupo";
			valueListItem9.DataValue = 5;
			valueListItem9.DisplayText = "Marca";
			this.rdgOpcion.Items.Add(valueListItem4);
			this.rdgOpcion.Items.Add(valueListItem5);
			this.rdgOpcion.Items.Add(valueListItem6);
			this.rdgOpcion.Items.Add(valueListItem7);
			this.rdgOpcion.Items.Add(valueListItem8);
			this.rdgOpcion.Items.Add(valueListItem9);
			this.rdgOpcion.ItemSpacingVertical = 9;
			this.rdgOpcion.Location = new System.Drawing.Point(12, 17);
			this.rdgOpcion.Name = "rdgOpcion";
			this.rdgOpcion.Size = new System.Drawing.Size(112, 155);
			this.rdgOpcion.TabIndex = 53;
			this.rdgOpcion.ValueChanged += new System.EventHandler(this.rdgOpcion_ValueChanged);
			// 
			// txtNumero
			// 
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(132, 69);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(128, 22);
			this.txtNumero.TabIndex = 52;
			// 
			// txtCodigo
			// 
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(132, 43);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(128, 22);
			this.txtCodigo.TabIndex = 51;
			// 
			// txtNombre
			// 
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(132, 17);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(296, 22);
			this.txtNombre.TabIndex = 50;
			// 
			// cmbMarca1
			// 
			this.cmbMarca1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMarca1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMarca1.DataSource = this.cdvMarca;
			this.cmbMarca1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 109;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbMarca1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbMarca1.DisplayMember = "Marca";
			this.cmbMarca1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMarca1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMarca1.Enabled = false;
			this.cmbMarca1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMarca1.Location = new System.Drawing.Point(132, 146);
			this.cmbMarca1.Name = "cmbMarca1";
			this.cmbMarca1.Size = new System.Drawing.Size(128, 21);
			this.cmbMarca1.TabIndex = 49;
			this.cmbMarca1.ValueMember = "idMarca";
			// 
			// cdvMarca
			// 
			this.cdvMarca.BindingContextCtrl = this;
			this.cdvMarca.DataSet = this.cdsArticuloTabla;
			this.cdvMarca.TableName = "";
			this.cdvMarca.TableViewName = "ArticuloMarca";
			// 
			// cdsArticuloTabla
			// 
			this.cdsArticuloTabla.BindingContextCtrl = this;
			this.cdsArticuloTabla.DataLibrary = "LibFacturacion";
			this.cdsArticuloTabla.DataLibraryUrl = "";
			this.cdsArticuloTabla.DataSetDef = "dsArticuloTabla";
			this.cdsArticuloTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticuloTabla.Name = "cdsArticuloTabla";
			this.cdsArticuloTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticuloTabla.SchemaDef = null;
			// 
			// cmbSubGrupo1
			// 
			this.cmbSubGrupo1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo1.DataSource = this.cdvSubGrupo;
			this.cmbSubGrupo1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 159;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 109;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 209;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbSubGrupo1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSubGrupo1.DisplayMember = "SubGrupo";
			this.cmbSubGrupo1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubGrupo1.Enabled = false;
			this.cmbSubGrupo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo1.Location = new System.Drawing.Point(132, 121);
			this.cmbSubGrupo1.Name = "cmbSubGrupo1";
			this.cmbSubGrupo1.Size = new System.Drawing.Size(128, 21);
			this.cmbSubGrupo1.TabIndex = 48;
			this.cmbSubGrupo1.ValueMember = "idSubGrupo";
			// 
			// cdvSubGrupo
			// 
			this.cdvSubGrupo.BindingContextCtrl = this;
			this.cdvSubGrupo.DataSet = this.cdsArticuloTabla;
			this.cdvSubGrupo.TableName = "";
			this.cdvSubGrupo.TableViewName = "ArticuloSubGrupo";
			// 
			// cmbGrupo1
			// 
			this.cmbGrupo1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo1.DataSource = this.cdvGrupo;
			this.cmbGrupo1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 84;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 73;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 64;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 66;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 66;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 66;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 42;
			ultraGridColumn14.Header.VisiblePosition = 8;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 19;
			ultraGridColumn15.Header.VisiblePosition = 9;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 19;
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 21;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 89;
			ultraGridColumn18.Header.VisiblePosition = 11;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 22;
			ultraGridColumn19.Header.VisiblePosition = 12;
			ultraGridColumn19.Width = 45;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn19});
			this.cmbGrupo1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbGrupo1.DisplayMember = "Grupo";
			this.cmbGrupo1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo1.Enabled = false;
			this.cmbGrupo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo1.Location = new System.Drawing.Point(132, 95);
			this.cmbGrupo1.Name = "cmbGrupo1";
			this.cmbGrupo1.Size = new System.Drawing.Size(128, 21);
			this.cmbGrupo1.TabIndex = 47;
			this.cmbGrupo1.ValueMember = "idGrupoArticulo";
			// 
			// cdvGrupo
			// 
			this.cdvGrupo.BindingContextCtrl = this;
			this.cdvGrupo.DataSet = this.cdsArticuloTabla;
			this.cdvGrupo.TableName = "";
			this.cdvGrupo.TableViewName = "ArticuloGrupo";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvArticulo;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance3;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn20.Header.VisiblePosition = 9;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 100;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 93;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 94;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 48;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 52;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 46;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 49;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 54;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 50;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 53;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 47;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 51;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 91;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 99;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn35.Header.VisiblePosition = 13;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 2;
			ultraGridColumn36.Width = 130;
			ultraGridColumn37.Header.VisiblePosition = 73;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 63;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 26;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 5;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn41.Header.VisiblePosition = 11;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 45;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 64;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 65;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 66;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 67;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 68;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 69;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 21;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 20;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 71;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 36;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 84;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 81;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 82;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 83;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 35;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 43;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 34;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 76;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 79;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 37;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 55;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 56;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 57;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 58;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 59;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 60;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 78;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 62;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 33;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 32;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 23;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 70;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 22;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 41;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 42;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 39;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 40;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 24;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 25;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 77;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 72;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 29;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 30;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 31;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 28;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 6;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 8;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 10;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 12;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 14;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 15;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 16;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 17;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 18;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 74;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 3;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 4;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 61;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 27;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 38;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn103.Header.VisiblePosition = 7;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 44;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 80;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 19;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 75;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 85;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 86;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 87;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 88;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 89;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 90;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 95;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 96;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 97;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 98;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn118.Header.VisiblePosition = 101;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 102;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 103;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 92;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 0;
			ultraGridColumn122.Width = 203;
			ultraGridColumn123.Header.VisiblePosition = 1;
			ultraGridColumn123.Width = 136;
			ultraGridColumn124.Header.VisiblePosition = 104;
			ultraGridColumn125.Header.VisiblePosition = 105;
			ultraGridColumn126.Header.VisiblePosition = 106;
			ultraGridColumn127.Header.VisiblePosition = 107;
			ultraGridColumn128.Header.VisiblePosition = 108;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn128});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 181);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(512, 172);
			this.ultraGrid1.TabIndex = 46;
			// 
			// BuscaReceta
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(536, 384);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.cmbMarca1);
			this.Controls.Add(this.cmbSubGrupo1);
			this.Controls.Add(this.cmbGrupo1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.idBusca);
			this.MinimizeBox = false;
			this.Name = "BuscaReceta";
			this.Text = "Búsqueda de Receta";
			this.Load += new System.EventHandler(this.BuscaAsiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdvProducto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProducto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCod)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BuscaAsiento_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			this.idBusca.Width =0;
			rdgOpcion.Value = 0;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void cdsProducto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProducto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void BorraErrores()
		{
			errorProvider.SetError(txtCodigo, "");
			errorProvider.SetError(txtNombre, "");
			errorProvider.SetError(txtNumero, "");
			errorProvider.SetError(cmbSubGrupo1, "");
			errorProvider.SetError(cmbGrupo1, "");
			errorProvider.SetError(cmbMarca1, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();

			if ((int) rdgOpcion.Value == 0)
				if (txtNombre.Text.Trim().Length == 0)
				{
					errorProvider.SetError(txtNombre, "Ingrese Articulo");
					bOk = false;
				}

			if ((int) rdgOpcion.Value == 1)
				if (txtCodigo.Text.Trim().Length == 0)
				{
					errorProvider.SetError(txtCodigo, "Ingrese Codigo");
					bOk = false;
				}

			if ((int) rdgOpcion.Value == 2)
				if (txtNumero.Text.Trim().Length == 0)
				{
					errorProvider.SetError(txtNumero, "Ingrese Codigo Alterno");
					bOk = false;
				}

			if ((int) rdgOpcion.Value == 3)
				if (cmbGrupo1.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbGrupo1, "Ingrese Grupo");
					bOk = false;
				}

			if ((int) rdgOpcion.Value == 4)
				if (cmbSubGrupo1.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbSubGrupo1, "Ingrese SubGrupo");
					bOk = false;
				}

			if ((int) rdgOpcion.Value == 5)
				if (cmbMarca1.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbMarca1, "Ingrese Marca");
					bOk = false;
				}
			return bOk;
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string strFiltro;
			string strAntes = "";
			string strDespues="'";

			//		Opciones de comodines
			if ((int) this.rdgTipoBusqueda.Value == 0) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.Value == 2) strAntes = "%";
			if ((int) this.rdgTipoBusqueda.Value == 2) strDespues = "%'";

			strFiltro = "";
			string stBusca = "";
			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					if (txtNombre.Value == null)
					{
						errorProvider.SetError(txtNombre, "Ingrese Articulo");
						return;
					}
					errorProvider.SetError(txtNombre, "");

					stBusca = txtNombre.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Articulo like '" + strAntes + stBusca + strDespues;
					break;
				case 1:
					if (txtCodigo.Value == null)
					{
						errorProvider.SetError(txtCodigo, "Ingrese Codigo");
						return;
					}
					errorProvider.SetError(txtCodigo, "");

					stBusca = txtCodigo.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Codigo LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 2:
					if (txtNumero.Value == null)
					{
						errorProvider.SetError(txtNumero, "Ingrese Codigo Alterno");
						return;
					}
					errorProvider.SetError(txtNumero, "");

					stBusca = txtCodigo.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "CodAlterno LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 3: // Grupo
					if (cmbGrupo1.Value == null)
					{
						errorProvider.SetError(cmbGrupo1, "Ingrese Gruppo");
						return;
					}
					errorProvider.SetError(cmbGrupo1, "");

					strFiltro = "idGrupo = " + cmbGrupo1.Value.ToString();
					break;
				case 4: // SubGrupo
					if (cmbSubGrupo1.Value == null)
					{
						errorProvider.SetError(cmbSubGrupo1, "Ingrese SubGrupo");
						return;
					}
					errorProvider.SetError(cmbSubGrupo1, "");
					strFiltro = "idSubGrupo = " + cmbSubGrupo1.Value.ToString();
					break;
				case 5: // Marca
					if (cmbMarca1.Value == null)
					{
						errorProvider.SetError(cmbMarca1, "Ingrese Marca");
						return;
					}
					errorProvider.SetError(cmbMarca1, "");

					strFiltro = "idMarca = " + cmbMarca1.Value.ToString();					
					break;
			}
			//			MessageBox.Show(strFiltro);
//			if (chkEnResultados.Checked && stBuscaEnResultado.Length > 0)
//				strFiltro += " And " + stBuscaEnResultado;
//			stBuscaEnResultado = strFiltro;
			string stFiltro = strFiltro + " And idTipoGrupo In (6, 8) ";
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsArticulo, "Articulo", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			this.cdsArticulo.Fill(fcTotal, false);
			if (this.cdvArticulo.Count == 0)
			{
				MessageBox.Show("No existen registros que cumplan la condicion de busqueda",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			//			this.cdvCompra.RowFilter = strFiltro;
			if (this.cdvArticulo.Count == 1)
			{
				this.ultraGrid1.Select();
				this.DialogResult = DialogResult.OK;
			}
			else
			{
				this.AcceptButton=this.btAceptar;
				this.ultraGrid1.Select();
			}
		}

		private void rdgOpcion_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtNombre.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.txtNumero.Enabled = false;
			this.cmbGrupo1.Enabled = false;
			this.cmbMarca1.Enabled = false;
			this.cmbSubGrupo1.Enabled = false;

			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					this.txtNombre.Enabled = true;
					txtNombre.Select();
					txtNombre.Focus();
					this.txtNombre.SelectAll();
					break;
				case 1:
					this.txtCodigo.Enabled = true;
					this.txtCodigo.SelectAll();
					txtCodigo.Focus();
					break;
				case 2:
					this.txtNumero.Enabled = true;
					this.txtNumero.SelectAll();
					txtNumero.Focus();
					break;
				case 3:
					this.cmbGrupo1.Enabled = true;
					this.cmbGrupo1.Select();
					break;
				case 4:
					this.cmbSubGrupo1.Enabled = true;
					this.cmbSubGrupo1.Select();
					break;
				case 5:
					this.cmbMarca1.Enabled = true;
					this.cmbMarca1.Select();
					break;
			}
		}
	}
}
